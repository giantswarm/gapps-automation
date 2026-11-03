/**
 * Meeting Recording Consent Poll
 *
 * Remote controls a bot (meet-bot/, signed in as the bot Workspace account) that joins recorded
 * Google Meet meetings visibly, asks the participants in the meeting chat whether they consent to
 * being recorded and collects per-person responses via a small web app.
 *
 * Why a bot:
 *   Meeting conversations (the Google Chat backed in-meeting chat) are system generated group DMs.
 *   Chat apps cannot be added to them and there is no API link between them and calendar events, so
 *   reliably posting into them via the Chat API is not possible. Meet has no API to join a call as a
 *   participant either, so the bot drives the Meet web UI in a browser (like note taker bots do).
 *
 * Architecture:
 *   - Time-based trigger scans organizer calendars for upcoming meetings with a Meet link that will
 *     be recorded (automatic recording/transcription/notes enabled, or the opt-in marker in the
 *     event description)
 *   - Creates the poll state (keyed by event ID) and tells the bot to join (idempotent PUT, re-sent
 *     on every run until shortly after the meeting started, so bot restarts are survived)
 *   - The bot greets the meeting with the consent question and a link to this script's web app
 *   - doGet() records responses (identifying the user via Session.getActiveUser()) and tells the bot
 *     about the first objection right away; the tally is posted when the bot leaves
 *
 * Note on async: lib.js is assembled with async/await stripped, so all lib calls are synchronous in
 * Apps Script. This file is written synchronously as well, which doGet() requires anyway.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */

/** The prefix for properties specific to this script in the project. */
const PROPERTY_PREFIX = 'ConsentPoll.';

/** Personio clientId and clientSecret, separated by '|'. */
const PERSONIO_TOKEN_KEY = PROPERTY_PREFIX + 'personioToken';

/** Service account credentials (in JSON format, as downloaded from Google Management Console). */
const SERVICE_ACCOUNT_CREDENTIALS_KEY = PROPERTY_PREFIX + 'serviceAccountCredentials';

/** Filter for allowed domains (to avoid working and failing on users present on foreign domains). */
const ALLOWED_DOMAINS_KEY = PROPERTY_PREFIX + 'allowedDomains';

/** White-list to restrict operation to a few tester email accounts (organizers).
 *
 * Must be one email or a comma separated list.
 *
 * Default: null or empty
 */
const EMAIL_WHITELIST_KEY = PROPERTY_PREFIX + 'emailWhiteList';

/** Lookahead minutes for upcoming meeting detection.
 *
 * Default: 15 minutes
 */
const LOOKAHEAD_MINUTES_KEY = PROPERTY_PREFIX + 'lookaheadMinutes';

/** Minimum number of attendees for a qualifying meeting.
 *
 * Default: 2
 */
const MIN_ATTENDEES_KEY = PROPERTY_PREFIX + 'minAttendees';

/** Which meetings the bot joins: 'recorded' or 'all'.
 *
 * 'recorded': automatic recording, transcription or notes are enabled for the Meet space (Meet REST API,
 * impersonating the organizer), or the event description contains OPT_IN_MARKER.
 *
 * Default: 'recorded'
 */
const JOIN_MODE_KEY = PROPERTY_PREFIX + 'joinMode';

/** Base URL of the meet-bot control API (origin only, e.g. https://meet-bot.example.com). */
const BOT_URL_KEY = PROPERTY_PREFIX + 'botUrl';

/** Shared HMAC secret of the meet-bot control API (MEETBOT_SECRET, at least 32 characters). */
const BOT_SECRET_KEY = PROPERTY_PREFIX + 'botSecret';

/** Email of the bot account, excluded from the attendees.
 *
 * Default: the account executing the script (Session.getEffectiveUser())
 */
const BOT_EMAIL_KEY = PROPERTY_PREFIX + 'botEmail';

/** Minutes the bot stays in the meeting (it leaves earlier if the meeting ends).
 *
 * Default: 10 minutes
 */
const STAY_MINUTES_KEY = PROPERTY_PREFIX + 'stayMinutes';

/** Optional purpose of recordings, shown in the consent question (e.g. "for AI generated meeting minutes"). */
const RECORDING_PURPOSE_KEY = PROPERTY_PREFIX + 'recordingPurpose';

/** Optional link to the recording policy, shown in the consent question. */
const POLICY_URL_KEY = PROPERTY_PREFIX + 'policyUrl';

/** Public URL of this script's web app deployment (the /exec URL).
 *
 * Default: ScriptApp.getService().getUrl()
 */
const WEB_APP_URL_KEY = PROPERTY_PREFIX + 'webAppUrl';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'checkUpcomingMeetings';

/** Prefix for consent state stored in ScriptProperties (keyed by event ID). */
const STATE_KEY_PREFIX = 'state.';

/** Maximum age for state entries before cleanup (48 hours). */
const STATE_MAX_AGE_MS = 48 * 60 * 60 * 1000;

/** Event description marker to request the bot for a meeting without automatic recording. */
const OPT_IN_MARKER = '#consent-poll';

/** The bot joins this long before the meeting starts. */
const JOIN_LEAD_MS = 60 * 1000;

/** The bot is still sent to meetings that started up to this long ago (late trigger runs, bot restarts). */
const JOIN_GRACE_MS = 10 * 60 * 1000;

/** Web app query parameter names. */
const PARAM_EVENT = 'e';
const PARAM_TOKEN = 't';
const PARAM_CONSENT = 'c';


// ---------------------------------------------------------------------------
// Standard Setup (from personio-to-group/PersonioToGroup.js pattern)
// ---------------------------------------------------------------------------

/** Uninstall triggers. */
function uninstall() {
    TriggerUtil.uninstall(TRIGGER_HANDLER_FUNCTION);
}

/** Install periodic execution trigger. */
function install(delayMinutes) {
    TriggerUtil.install(TRIGGER_HANDLER_FUNCTION, delayMinutes);
}

/** Allow setting properties. */
function setProperties(properties, deleteAllOthers) {
    TriggerUtil.setProperties(properties, deleteAllOthers);
}

/** List the bot's sessions (for checking the setup via `clasp run listBotSessions`). */
function listBotSessions() {
    const sessions = callBot_('GET', '/v1/sessions');
    Logger.log(JSON.stringify(sessions, null, 2));
    return sessions;
}


// ---------------------------------------------------------------------------
// Core Trigger: checkUpcomingMeetings
// ---------------------------------------------------------------------------

/** Main entry point: scans organizer calendars for upcoming recorded meetings and sends the bot. */
function checkUpcomingMeetings() {

    const allowedDomains = (getScriptProperties_().getProperty(ALLOWED_DOMAINS_KEY) || '')
        .split(',')
        .map(d => d.trim())
        .filter(d => !!d);

    const emailWhiteList = getEmailWhiteList_();
    const isEmailAllowed = email => (!emailWhiteList.length || emailWhiteList.includes(email))
        && allowedDomains.some(domain => email.endsWith('@' + domain));

    const lookaheadMinutes = getLookaheadMinutes_();
    const minAttendees = getMinAttendees_();
    const joinMode = getJoinMode_();
    const webAppUrl = getWebAppUrl_();
    if (!webAppUrl) {
        throw new Error('Web app URL unknown, deploy the web app or set ' + WEB_APP_URL_KEY);
    }
    getBotConfig_();   // fail early if the bot is not configured

    Logger.log('Configured to handle accounts %s on domains %s, lookahead %s min, min attendees %s, join mode %s, web app %s',
        emailWhiteList.length ? emailWhiteList : '(all)', allowedDomains, '' + lookaheadMinutes, '' + minAttendees, joinMode, webAppUrl);

    const personioCreds = getPersonioCreds_();
    const personio = PersonioClientV1.withApiCredentials(personioCreds.clientId, personioCreds.clientSecret);
    const employees = personio.getPersonioJson('/company/employees');
    const activeEmployees = employees.filter(employee =>
        employee.attributes.status.value !== 'inactive' && isEmailAllowed(employee.attributes.email.value)
    );

    Logger.log('Processing %s active employees', '' + activeEmployees.length);

    const creds = getServiceAccountCredentials_();
    const now = new Date();
    const timeMin = now.toISOString();
    const timeMax = Util.addDateMillies(new Date(now), lookaheadMinutes * 60 * 1000).toISOString();

    // Per-run caches
    const context = {
        creds: creds,
        webAppUrl: webAppUrl,
        directory: null            // lazily loaded: {idToPerson: {}, emailToPerson: {}}
    };

    let firstError = null;
    let processedCount = 0;

    for (const employee of activeEmployees) {
        const email = employee.attributes.email.value;

        try {
            const calendar = CalendarClient.withImpersonatingService(creds, email);
            const events = calendar.list('primary', {
                singleEvents: true,
                showDeleted: false,
                timeMin: timeMin,
                timeMax: timeMax
            });

            for (const event of events) {
                // Only process events where this employee is the organizer (dedup across calendars)
                if ((event.organizer?.email || '').toLowerCase() !== email.toLowerCase()) {
                    continue;
                }

                // Ongoing events are listed as well, join those only shortly after they started
                const eventStart = new Date(event.start?.dateTime || event.start?.date).getTime();
                if (!isQualifyingEvent_(event, minAttendees) || now.getTime() > eventStart + JOIN_GRACE_MS) {
                    continue;
                }

                try {
                    if (joinMode === 'all' || isRecordedMeeting_(context, email, event)) {
                        sendBot_(context, event);
                    }
                } catch (e) {
                    Logger.log('Failed to send the bot to event %s (%s): %s', event.id, event.summary, e);
                    firstError = firstError || e;
                }
            }
        } catch (e) {
            Logger.log('Failed to process calendar for user %s: %s', email, e);
            firstError = firstError || e;
        }

        ++processedCount;
    }

    Logger.log('Completed scanning %s of %s accounts', '' + processedCount, '' + activeEmployees.length);

    // Cleanup old state entries
    cleanupOldState_();

    if (firstError) {
        throw firstError;
    }
}


// ---------------------------------------------------------------------------
// Event Qualification
// ---------------------------------------------------------------------------

/** Check if a calendar event qualifies for a consent poll. */
function isQualifyingEvent_(event, minAttendees) {
    // Must not be cancelled, must have a time (all-day events are no meetings)
    if (event.status === 'cancelled' || !event.start?.dateTime) {
        return false;
    }

    // Must have enough attendees
    const attendees = event.attendees || [];
    if (attendees.length < minAttendees) {
        return false;
    }

    // Must have a Google Meet video link
    return !!getMeetUri_(event);
}


/** The event's Google Meet link, or null. */
function getMeetUri_(event) {
    const entryPoint = (event.conferenceData?.entryPoints || [])
        .find(ep => ep.entryPointType === 'video' && (ep.uri || '').startsWith('https://meet.google.com/'));
    return entryPoint ? entryPoint.uri : null;
}


/** Check if a meeting will be recorded: opt-in marker or automatic artifacts enabled for its Meet space.
 *
 * Gemini notes and transcripts are included, they process what participants say just like recordings.
 */
function isRecordedMeeting_(context, organizerEmail, event) {
    if ((event.description || '').toLowerCase().includes(OPT_IN_MARKER)) {
        return true;
    }

    const meetingCode = event.conferenceData?.conferenceId
        || getMeetUri_(event).replace(/^https:\/\/meet\.google\.com\//, '').replace(/[?#].*$/, '');

    let artifactConfig;
    try {
        const meet = MeetClient.withImpersonatingService(context.creds, organizerEmail);
        artifactConfig = meet.getSpace('spaces/' + meetingCode)?.config?.artifactConfig || {};
    } catch (e) {
        // e.g. a Meet link owned by someone else, nothing we can find out about recording
        Logger.log('Failed to get Meet space %s of event %s (%s) as %s, assuming no recording: %s',
            meetingCode, event.id, event.summary, organizerEmail, e);
        return false;
    }
    const automatic = [
        artifactConfig.recordingConfig?.autoRecordingGeneration,
        artifactConfig.transcriptionConfig?.autoTranscriptionGeneration,
        artifactConfig.smartNotesConfig?.autoSmartNotesGeneration
    ];

    Logger.log('Meet space %s of event %s (%s): automatic recording/transcription/notes %s',
        meetingCode, event.id, event.summary, automatic.map(v => v || '-').join('/'));

    return automatic.includes('ON');
}


// ---------------------------------------------------------------------------
// Directory (email -> name resolution)
// ---------------------------------------------------------------------------

/** Load the domain directory once per run (impersonating the given user). */
function getDirectory_(context, impersonatedEmail) {
    if (context.directory) {
        return context.directory;
    }

    const directoryClient = DirectoryClient.withImpersonatingService(context.creds, impersonatedEmail);
    const people = directoryClient.listDirectoryPeople();

    const idToPerson = {};
    const emailToPerson = {};
    for (const person of people) {
        const id = (person.resourceName || '').replace(/^people\//, '');
        const emails = (person.emailAddresses || []).map(e => (e.value || '').toLowerCase()).filter(e => !!e);
        const name = (person.names || []).map(n => n.displayName).find(n => !!n) || emails[0] || id;
        if (!id || !emails.length) {
            continue;
        }
        const entry = {id: id, email: emails[0], name: name};
        idToPerson[id] = entry;
        for (const email of emails) {
            emailToPerson[email] = entry;
        }
    }

    Logger.log('Loaded %s directory people', '' + Object.keys(idToPerson).length);
    context.directory = {idToPerson: idToPerson, emailToPerson: emailToPerson};
    return context.directory;
}


// ---------------------------------------------------------------------------
// Bot Remote Control
// ---------------------------------------------------------------------------

/** Create or refresh the poll state of an event and tell the bot to join (idempotent). */
function sendBot_(context, event) {
    const organizerEmail = event.organizer.email;
    const directory = getDirectory_(context, organizerEmail);
    const botEmail = getBotEmail_();

    const attendees = (event.attendees || [])
        .map(a => (a.email || '').toLowerCase())
        .filter(e => !!e && !e.endsWith('calendar.google.com') && e !== botEmail);   // drop rooms/resources and the bot

    const names = {};
    for (const email of attendees) {
        const person = directory.emailToPerson[email];
        names[email] = person ? person.name : email;
    }

    const state = updateState_(event.id, state => Object.assign(state || {
        eventId: event.id,
        token: Util.generateUUIDv4(),
        organizerEmail: organizerEmail,
        responses: {},
        createdAt: Date.now()
    }, {
        meetingTitle: event.summary || '(No title)',
        meetingTime: event.start.dateTime,
        meetingUri: getMeetUri_(event),
        attendees: attendees,
        names: names
    }));

    const start = new Date(event.start.dateTime).getTime();
    const end = new Date(event.end?.dateTime || start + 60 * 60 * 1000).getTime();
    const joinAt = Math.max(Date.now(), start - JOIN_LEAD_MS);
    const leaveAt = Math.max(Math.min(joinAt + getStayMinutes_() * 60 * 1000, end), joinAt + 60 * 1000);

    const session = callBot_('PUT', botSessionPath_(event.id), {
        meetingUri: state.meetingUri,
        title: state.meetingTitle,
        joinAt: new Date(joinAt).toISOString(),
        leaveAt: new Date(leaveAt).toISOString(),
        greeting: buildGreetingText_(state, context.webAppUrl),
        farewell: buildTallyText_(state, context.webAppUrl)
    });

    Logger.log('Bot session for event %s (%s) by %s: %s, joining at %s',
        event.id, event.summary, organizerEmail, session.state, session.joinAt);
}


/** Tell the bot about a new response: announce the first objection or full consent, update the tally. */
function notifyBot_(previousState, state) {
    const webAppUrl = getWebAppUrl_();
    const before = countResponses_(previousState);
    const after = countResponses_(state);

    let say = undefined;
    if (after.declined > 0 && before.declined === 0) {
        say = '❌ A participant does not consent to being recorded. Please do not record this meeting.';
    } else if (after.declined === 0 && after.allConsented && !before.allConsented) {
        say = '✅ All ' + after.invited + ' invited participants consented to being recorded.';
    }

    callBot_('PATCH', botSessionPath_(state.eventId), {
        say: say,
        farewell: buildTallyText_(state, webAppUrl)
    });
}


function botSessionPath_(eventId) {
    return '/v1/sessions/' + encodeURIComponent(eventId);
}


/** Call the meet-bot control API with a signed request (see meet-bot/src/auth.mjs), returns the parsed response. */
function callBot_(method, path, payload) {
    const bot = getBotConfig_();
    const body = payload === undefined ? '' : JSON.stringify(payload);
    const timestamp = '' + Math.floor(Date.now() / 1000);
    const signature = Utilities.computeHmacSha256Signature([timestamp, method, path, body].join('\n'), bot.secret,
        Utilities.Charset.UTF_8)
        .map(b => ('0' + (b & 0xff).toString(16)).slice(-2))
        .join('');

    const options = {
        method: method.toLowerCase(),
        headers: {
            'X-MeetBot-Timestamp': timestamp,
            'X-MeetBot-Signature': 'sha256=' + signature
        },
        muteHttpExceptions: true
    };
    if (body) {
        options.contentType = 'application/json';
        options.payload = body;
    }

    const response = UrlFetchApp.fetch(bot.url + path, options);
    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code < 200 || code >= 300) {
        throw new Error('meet-bot ' + method + ' ' + path + ' failed with HTTP ' + code + ': ' + text);
    }
    return text ? JSON.parse(text) : null;
}


// ---------------------------------------------------------------------------
// Message Text (plain text, Meet's chat links URLs automatically)
// ---------------------------------------------------------------------------

/** The consent question the bot posts after joining. */
function buildGreetingText_(state, webAppUrl) {
    const domain = state.organizerEmail.split('@')[1];
    const purpose = (getScriptProperties_().getProperty(RECORDING_PURPOSE_KEY) || '').trim();
    const policyUrl = (getScriptProperties_().getProperty(POLICY_URL_KEY) || '').trim();

    const lines = [];
    lines.push('👋 Hi, I am the recording consent bot. I do not record or listen to anything.');
    lines.push('🎥 "' + state.meetingTitle + '" may be recorded' + (purpose ? ' ' + purpose : '') + '.');
    lines.push('Do you consent to being recorded? Please answer here: ' + buildPollUrl_(webAppUrl, state));
    lines.push('Saying no has no disadvantages for you. External guests: the link only works for ' + domain
        + ' accounts, please answer in this chat.');
    if (policyUrl) {
        lines.push('Recording policy: ' + policyUrl);
    }
    return lines.join('\n');
}


/** The summary the bot posts before leaving (kept up to date on every response). */
function buildTallyText_(state, webAppUrl) {
    const counts = countResponses_(state);
    const lines = [];
    lines.push('📋 Recording consent for "' + state.meetingTitle + '": ' + counts.responded + '/' + counts.invited
        + ' responded, ' + counts.consented + ' consented, ' + counts.declined + ' declined.');
    if (counts.declined > 0) {
        lines.push('❌ At least one participant does not consent. Please do not record this meeting.');
    } else if (counts.allConsented) {
        lines.push('✅ All invited participants consented.');
    } else {
        lines.push('⏳ Not everyone answered yet: please only record with everyone\'s consent.');
    }
    lines.push('You can still answer or change your answer: ' + buildPollUrl_(webAppUrl, state));
    lines.push('👋 Bye!');
    return lines.join('\n');
}


/** Count responses (responders not on the invite, e.g. forwarded invitations, count towards the total). */
function countResponses_(state) {
    const responses = state?.responses || {};
    const attendees = state?.attendees || [];
    const emails = Object.keys(responses);
    const consented = emails.filter(email => responses[email].consent).length;
    const invited = attendees.length + emails.filter(email => !attendees.includes(email)).length;
    return {
        invited: invited,
        responded: emails.length,
        consented: consented,
        declined: emails.length - consented,
        allConsented: invited > 0 && attendees.every(email => responses[email]?.consent) && consented === emails.length
    };
}


/** Build the link to the poll page, with consent = true/false a link that records the answer. */
function buildPollUrl_(webAppUrl, state, consent) {
    const params = {
        [PARAM_EVENT]: state.eventId,
        [PARAM_TOKEN]: state.token
    };
    if (consent !== undefined) {
        params[PARAM_CONSENT] = consent ? '1' : '0';
    }
    return webAppUrl + UrlFetchJsonClient.buildQuery(params);
}


/** Resolve a display name for an email from the state. */
function displayName_(state, email) {
    return (state.responses?.[email]?.name) || (state.names?.[email]) || email;
}


// ---------------------------------------------------------------------------
// Web App: Response Handling
// ---------------------------------------------------------------------------

/** Web app entry point.
 *
 * Deployment: execute as the deploying user, access restricted to the domain, so that
 * Session.getActiveUser().getEmail() identifies the responding person.
 *
 * Query parameters: e=<eventId>, t=<token>, c=1|0 (omit c to show the question and status).
 */
function doGet(e) {
    const params = (e && e.parameter) || {};
    const eventId = params[PARAM_EVENT] || '';
    const token = params[PARAM_TOKEN] || '';
    const consentParam = params[PARAM_CONSENT];

    let userEmail = '';
    try {
        userEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
    } catch (err) {
        Logger.log('Failed to determine active user: %s', err);
    }

    if (!eventId) {
        return renderPage_('Recording Consent', '<p>Missing poll reference.</p>');
    }

    const state = loadState_(eventId);
    if (!state || !token || state.token !== token) {
        return renderPage_('Recording Consent', '<p>This consent poll is no longer active or the link is invalid.</p>');
    }

    const webAppUrl = getWebAppUrl_();

    if (consentParam === undefined) {
        return renderPage_('Recording Consent', buildQuestionHtml_(state, webAppUrl) + buildStatusHtml_(state, userEmail));
    }

    if (!userEmail) {
        return renderPage_('Recording Consent',
            '<p>Could not determine your identity. Please open this link in a browser signed in with your '
            + escapeHtml_(state.organizerEmail.split('@')[1]) + ' account.</p>');
    }

    const consent = consentParam === '1';
    const updated = recordResponse_(eventId, userEmail, consent);
    if (!updated) {
        return renderPage_('Recording Consent', '<p>This consent poll is no longer active.</p>');
    }

    let notice = '';
    try {
        notifyBot_(updated.previous, updated.state);
    } catch (err) {
        Logger.log('Failed to notify the bot about a response for event %s: %s', eventId, err);
        notice = consent ? '' : '<p class="muted">Your response was recorded, but the bot could not inform the meeting. '
            + 'Please tell the organizer.</p>';
    }

    const headline = consent
        ? '<p class="ok">✅ Thank you, your consent has been recorded.</p>'
        : '<p class="no">❌ Thank you, your objection has been recorded. The meeting has been informed (without your name).</p>';

    return renderPage_('Recording Consent', headline + notice + buildQuestionHtml_(updated.state, webAppUrl)
        + buildStatusHtml_(updated.state, userEmail));
}


/** Record a response under a script lock, returns {previous, state} or null if the poll is gone. */
function recordResponse_(eventId, userEmail, consent) {
    let previous = null;
    const state = updateState_(eventId, state => {
        if (!state) {
            return null;
        }
        previous = JSON.parse(JSON.stringify(state));
        state.responses = state.responses || {};
        state.responses[userEmail] = {
            consent: consent,
            name: state.names?.[userEmail] || userEmail,
            respondedAt: Date.now()
        };
        return state;
    });
    return state ? {previous: previous, state: state} : null;
}


/** Render the question with the answer buttons. */
function buildQuestionHtml_(state, webAppUrl) {
    const purpose = (getScriptProperties_().getProperty(RECORDING_PURPOSE_KEY) || '').trim();
    const policyUrl = (getScriptProperties_().getProperty(POLICY_URL_KEY) || '').trim();

    // HtmlService pages run in an iframe, links must navigate the top window
    return '<h2>' + escapeHtml_(state.meetingTitle) + '</h2>'
        + '<p class="muted">' + escapeHtml_(formatMeetingTime_(state.meetingTime)) + '</p>'
        + '<p>This meeting may be recorded' + (purpose ? ' ' + escapeHtml_(purpose) : '') + '. '
        + 'Do you consent to being recorded? Saying no has no disadvantages for you.'
        + (policyUrl ? ' See the <a target="_top" href="' + escapeHtml_(policyUrl) + '">recording policy</a>.' : '')
        + '</p><p>'
        + '<a class="button ok" target="_top" href="' + escapeHtml_(buildPollUrl_(webAppUrl, state, true)) + '">✅ I consent</a> '
        + '<a class="button no" target="_top" href="' + escapeHtml_(buildPollUrl_(webAppUrl, state, false)) + '">❌ I do not consent</a>'
        + '</p>';
}


/** Render the status part of the web app page. */
function buildStatusHtml_(state, userEmail) {
    const responses = state.responses || {};
    const attendees = state.attendees || [];
    const rows = [];

    const all = attendees.slice();
    for (const email of Object.keys(responses)) {
        if (!all.includes(email)) {
            all.push(email);
        }
    }

    for (const email of all) {
        const resp = responses[email];
        const status = !resp ? '⏳ pending' : (resp.consent ? '✅ consented' : '❌ declined');
        const me = email === userEmail ? ' <span class="muted">(you)</span>' : '';
        rows.push('<tr><td>' + escapeHtml_(displayName_(state, email)) + me + '</td><td>' + status + '</td></tr>');
    }

    return '<p>' + Object.keys(responses).length + '/' + all.length + ' responded</p>'
        + '<table>' + rows.join('') + '</table>';
}


/** Wrap body HTML into a minimal page. */
function renderPage_(title, bodyHtml) {
    const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
        + '<title>' + escapeHtml_(title) + '</title>'
        + '<style>body{font-family:system-ui,sans-serif;max-width:40em;margin:2em auto;padding:0 1em;color:#222}'
        + 'table{border-collapse:collapse}td{padding:.25em .75em .25em 0}.ok{color:#1b7f3b;font-weight:bold}'
        + '.no{color:#b3261e;font-weight:bold}.muted{color:#666}'
        + '.button{display:inline-block;padding:.5em 1em;margin:0 .5em .5em 0;border:2px solid;border-radius:.5em;text-decoration:none}</style></head>'
        + '<body><h1>' + escapeHtml_(title) + '</h1>' + bodyHtml + '</body></html>';
    return HtmlService.createHtmlOutput(html).setTitle(title);
}


/** Escape text for HTML output. */
function escapeHtml_(text) {
    return ('' + (text == null ? '' : text))
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}


// ---------------------------------------------------------------------------
// State Management
// ---------------------------------------------------------------------------

/** Load consent poll state from ScriptProperties. */
function loadState_(eventId) {
    const raw = getScriptProperties_().getProperty(STATE_KEY_PREFIX + eventId);
    if (!raw) {
        return null;
    }
    try {
        return JSON.parse(raw);
    } catch (e) {
        Logger.log('Failed to parse state for event %s: %s', eventId, e);
        return null;
    }
}

/** Update consent poll state under a script lock: updater(state or null) returns the new state (null: no change). */
function updateState_(eventId, updater) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const state = updater(loadState_(eventId));
        if (state) {
            getScriptProperties_().setProperty(STATE_KEY_PREFIX + eventId, JSON.stringify(state));
        }
        return state;
    } finally {
        lock.releaseLock();
    }
}

/** Clean up state entries older than STATE_MAX_AGE_MS. */
function cleanupOldState_() {
    const now = Date.now();
    const properties = getScriptProperties_().getProperties() || {};

    for (const key in properties) {
        if (key.startsWith(STATE_KEY_PREFIX)) {
            try {
                const state = JSON.parse(properties[key]);
                const createdAt = state.createdAt || state.pollSentAt;
                if (!createdAt || (now - createdAt) > STATE_MAX_AGE_MS) {
                    getScriptProperties_().deleteProperty(key);
                    Logger.log('Cleaned up old state: %s', key);
                }
            } catch (e) {
                // Malformed state, clean it up
                getScriptProperties_().deleteProperty(key);
            }
        }
    }
}


// ---------------------------------------------------------------------------
// Helper Functions
// ---------------------------------------------------------------------------

/** Get script properties. */
function getScriptProperties_() {
    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties) {
        throw new Error('ScriptProperties not accessible');
    }
    return scriptProperties;
}

/** Get the service account credentials. */
function getServiceAccountCredentials_() {
    const creds = getScriptProperties_().getProperty(SERVICE_ACCOUNT_CREDENTIALS_KEY);
    if (!creds) {
        throw new Error('No service account credentials at script property ' + SERVICE_ACCOUNT_CREDENTIALS_KEY);
    }
    return JSON.parse(creds);
}

/** Get the Personio token. */
function getPersonioCreds_() {
    const credentialFields = (getScriptProperties_().getProperty(PERSONIO_TOKEN_KEY) || '|')
        .split('|')
        .map(field => field.trim());
    return {clientId: credentialFields[0], clientSecret: credentialFields[1]};
}

/** Get the email account white-list (optional, leave empty to process all suitable accounts). */
function getEmailWhiteList_() {
    return (getScriptProperties_().getProperty(EMAIL_WHITELIST_KEY) || '').trim()
        .split(',').map(email => email.trim()).filter(email => !!email);
}

/** Get the lookahead minutes or the default (15 minutes). */
function getLookaheadMinutes_() {
    return getPositiveIntProperty_(LOOKAHEAD_MINUTES_KEY, 15);
}

/** Get the minimum number of attendees or the default (2). */
function getMinAttendees_() {
    return getPositiveIntProperty_(MIN_ATTENDEES_KEY, 2);
}

/** Get the minutes the bot stays in a meeting or the default (10 minutes). */
function getStayMinutes_() {
    return getPositiveIntProperty_(STAY_MINUTES_KEY, 10);
}

/** Read a positive integer property, else the default. */
function getPositiveIntProperty_(key, defaultValue) {
    const raw = (getScriptProperties_().getProperty(key) || '').trim();
    const value = Math.abs(Math.round(+raw));
    if (!value || Number.isNaN(value)) {
        return defaultValue;
    }
    return value;
}

/** Get the join mode ('recorded' or 'all', default 'recorded'). */
function getJoinMode_() {
    const mode = (getScriptProperties_().getProperty(JOIN_MODE_KEY) || '').trim().toLowerCase();
    return mode === 'all' ? 'all' : 'recorded';
}

/** Get the bot control API base URL and secret, throws if not configured. */
function getBotConfig_() {
    const url = (getScriptProperties_().getProperty(BOT_URL_KEY) || '').trim().replace(/\/+$/, '');
    const secret = (getScriptProperties_().getProperty(BOT_SECRET_KEY) || '').trim();
    if (!url.startsWith('https://') || secret.length < 32) {
        throw new Error('Bot not configured, set ' + BOT_URL_KEY + ' (https://...) and ' + BOT_SECRET_KEY + ' (32+ characters)');
    }
    return {url: url, secret: secret};
}

/** Get the bot account email (property override, else the executing account). */
function getBotEmail_() {
    const configured = (getScriptProperties_().getProperty(BOT_EMAIL_KEY) || '').trim().toLowerCase();
    if (configured) {
        return configured;
    }
    try {
        return (Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
    } catch (e) {
        Logger.log('Failed to determine the executing account: %s', e);
        return '';
    }
}

/** Get the web app URL (property override, else the deployed web app URL, else null). */
function getWebAppUrl_() {
    const configured = (getScriptProperties_().getProperty(WEB_APP_URL_KEY) || '').trim();
    if (configured) {
        return configured;
    }
    try {
        return ScriptApp.getService().getUrl() || null;
    } catch (e) {
        Logger.log('Failed to determine web app URL: %s', e);
        return null;
    }
}

/** Format a meeting time for display. */
function formatMeetingTime_(isoTime) {
    if (!isoTime) {
        return '(Unknown time)';
    }
    try {
        const date = new Date(isoTime);
        return Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEE, dd MMM yyyy HH:mm');
    } catch (e) {
        return isoTime;
    }
}
