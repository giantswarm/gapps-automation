/**
 * Meeting Recording Consent Poll
 *
 * Posts a recording consent request into the Google Meet meeting conversation (the Google Chat
 * backed in-meeting chat) and collects per-person responses via a small web app.
 *
 * Why text only:
 *   Meeting conversations are system generated group DMs. Chat apps cannot be added to them and
 *   cards (cardsV2) can only be posted by Chat apps. Messages posted with user authentication
 *   may only contain text. So we post as the meeting organizer (domain-wide delegation) and put
 *   the consent/decline actions behind links to this script's web app, which identifies the
 *   responding user via Session.getActiveUser() and updates the message text with the tally.
 *
 * Architecture:
 *   - Time-based trigger scans organizer calendars for upcoming meetings with a Meet link
 *   - Finds the meeting conversation by impersonating the organizer and matching Chat group DM
 *     members (users/{id}, resolved via the People API directory) against the event attendees
 *   - Posts a text message as the organizer (chat.messages scope) with consent/decline links
 *   - doGet() records responses and patches the message text in place with the current tally
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

/** Public URL of this script's web app deployment (the /exec URL).
 *
 * Default: ScriptApp.getService().getUrl()
 */
const WEB_APP_URL_KEY = PROPERTY_PREFIX + 'webAppUrl';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'checkUpcomingMeetings';

/** Prefix for consent state stored in ScriptProperties (keyed by event ID). */
const STATE_KEY_PREFIX = 'state.';

/** Prefix for sent dedup markers stored in ScriptProperties (keyed by event ID). */
const SENT_KEY_PREFIX = 'sent.';

/** Maximum age for state entries before cleanup (48 hours). */
const STATE_MAX_AGE_MS = 48 * 60 * 60 * 1000;

/** Meeting conversations are created by Google Chat up to 7 days before the meeting. */
const SPACE_CREATE_WINDOW_BEFORE_MS = 8 * 24 * 60 * 60 * 1000;

/** Allow for conversations created late (first message sent during the meeting). */
const SPACE_CREATE_WINDOW_AFTER_MS = 60 * 60 * 1000;

/** Minimum fraction of chat members that must be event attendees. */
const MIN_MEMBER_OVERLAP = 0.8;

/** Chat API base URL. */
const CHAT_API_BASE = 'https://chat.googleapis.com/v1';

/** Scopes used via domain-wide delegation (impersonating the organizer). */
const SCOPE_CHAT_SPACES_READONLY = 'https://www.googleapis.com/auth/chat.spaces.readonly';
const SCOPE_CHAT_MEMBERSHIPS_READONLY = 'https://www.googleapis.com/auth/chat.memberships.readonly';
const SCOPE_CHAT_MESSAGES = 'https://www.googleapis.com/auth/chat.messages';

/** Maximum number of attendee names to show in the message before truncating. */
const MAX_DISPLAY_NAMES = 20;

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


// ---------------------------------------------------------------------------
// Core Trigger: checkUpcomingMeetings
// ---------------------------------------------------------------------------

/** Main entry point: scans organizer calendars for upcoming meetings and posts consent requests. */
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
    const webAppUrl = getWebAppUrl_();

    Logger.log('Configured to handle accounts %s on domains %s, lookahead %s min, min attendees %s, web app %s',
        emailWhiteList.length ? emailWhiteList : '(all)', allowedDomains, '' + lookaheadMinutes, '' + minAttendees, webAppUrl);

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
        directory: null,            // lazily loaded: {idToPerson: {}, emailToPerson: {}}
        usedSpaces: loadUsedSpaces_()
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

                if (!isQualifyingEvent_(event, minAttendees)) {
                    continue;
                }

                // Skip if already sent
                if (getScriptProperties_().getProperty(SENT_KEY_PREFIX + event.id)) {
                    continue;
                }

                try {
                    sendConsentPoll_(context, event);
                } catch (e) {
                    Logger.log('Failed to send consent poll for event %s (%s): %s', event.id, event.summary, e);
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
    // Must not be cancelled
    if (event.status === 'cancelled') {
        return false;
    }

    // Must have enough attendees
    const attendees = event.attendees || [];
    if (attendees.length < minAttendees) {
        return false;
    }

    // Must have a Google Meet video link
    return (event.conferenceData?.entryPoints || [])
        .some(ep => ep.entryPointType === 'video' && (ep.uri || '').includes('meet.google.com'));
}


// ---------------------------------------------------------------------------
// Directory (Chat user ID -> email/name resolution)
// ---------------------------------------------------------------------------

/** Load the domain directory once per run (impersonating the given user).
 *
 * Chat API memberships only expose users/{id}. The {id} equals the People API person ID
 * (people/{id}), so the directory listing is used to resolve IDs to emails and display names.
 */
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
// Meeting Conversation Discovery
// ---------------------------------------------------------------------------

/** Find the Chat group DM corresponding to a meeting's conversation.
 *
 * There is no API link between Calendar events and meeting conversations. Matching algorithm:
 *   1. Impersonate the organizer, list their GROUP_CHAT spaces
 *   2. Keep spaces created within [eventStart - 8 days, eventStart + 1 hour]
 *      (Google Chat creates meeting conversations up to 7 days ahead)
 *   3. Skip spaces already used by another poll
 *   4. Resolve member IDs to emails, require >= 80% of members to be event attendees
 *   5. Prefer a displayName containing the event title, then the earliest created candidate
 *      (for recurring meetings the conversation of the next instance may already exist)
 */
function findMeetingChatSpace_(context, organizerEmail, event) {
    const spacesService = UrlFetchJsonClient.createImpersonatingService(
        'ConsentPollSpaces-' + organizerEmail, context.creds, organizerEmail, SCOPE_CHAT_SPACES_READONLY);
    const spacesClient = new UrlFetchJsonClient(spacesService);

    const eventStart = new Date(event.start.dateTime || event.start.date);
    const windowStart = new Date(eventStart.getTime() - SPACE_CREATE_WINDOW_BEFORE_MS);
    const windowEnd = new Date(eventStart.getTime() + SPACE_CREATE_WINDOW_AFTER_MS);

    const candidateSpaces = [];
    let pageToken = undefined;
    do {
        const query = UrlFetchJsonClient.buildQuery({
            //filter: 'spaceType = "GROUP_CHAT"',
            pageSize: 100,
            pageToken: pageToken
        });
        const response = spacesClient.getJson(CHAT_API_BASE + '/spaces' + query) || {};

        for (const space of (response.spaces || [])) {
            if (!space.createTime || context.usedSpaces[space.name]) {
                continue;
            }
            const createTime = new Date(space.createTime);
            if (createTime >= windowStart && createTime <= windowEnd) {
                candidateSpaces.push(space);
            }
        }

        pageToken = response.nextPageToken;
    } while (pageToken);

    Logger.log('Found %s candidate chat spaces for event %s (%s)', '' + candidateSpaces.length, event.id, event.summary);

    if (candidateSpaces.length === 0) {
        return null;
    }

    const eventAttendeeEmails = new Set(
        (event.attendees || []).map(a => (a.email || '').toLowerCase()).filter(e => !!e)
    );

    const directory = getDirectory_(context, organizerEmail);
    const membersService = UrlFetchJsonClient.createImpersonatingService(
        'ConsentPollMembers-' + organizerEmail, context.creds, organizerEmail, SCOPE_CHAT_MEMBERSHIPS_READONLY);
    const membersClient = new UrlFetchJsonClient(membersService);

    const title = (event.summary || '').trim().toLowerCase();
    const matches = [];

    for (const space of candidateSpaces) {
        try {
            const memberIds = listSpaceMemberIds_(membersClient, space.name);
            if (memberIds.length === 0) {
                continue;
            }

            let matchCount = 0;
            for (const memberId of memberIds) {
                const person = directory.idToPerson[memberId];
                if (person && eventAttendeeEmails.has(person.email)) {
                    matchCount++;
                }
            }

            const overlap = matchCount / memberIds.length;
            Logger.log('Space %s: %s members, %s match event attendees (%s%% overlap)',
                space.name, '' + memberIds.length, '' + matchCount, '' + Math.round(overlap * 100));

            if (overlap >= MIN_MEMBER_OVERLAP) {
                const titleMatch = !!title && (space.displayName || '').toLowerCase().includes(title);
                matches.push({space: space, titleMatch: titleMatch, createTime: new Date(space.createTime).getTime()});
            }
        } catch (e) {
            Logger.log('Failed to check members of space %s: %s', space.name, e);
        }
    }

    if (matches.length === 0) {
        return null;
    }

    matches.sort((a, b) => (b.titleMatch - a.titleMatch) || (a.createTime - b.createTime));
    return matches[0].space;
}


/** List the Chat user IDs (users/{id} -> {id}) of all human members of a space. */
function listSpaceMemberIds_(membersClient, spaceName) {
    const ids = [];
    let pageToken = undefined;
    do {
        const query = UrlFetchJsonClient.buildQuery({
            pageSize: 100,
            pageToken: pageToken
        });
        const response = membersClient.getJson(CHAT_API_BASE + '/' + spaceName + '/members' + query) || {};

        for (const membership of (response.memberships || [])) {
            const member = membership.member || {};
            if (member.type === 'BOT') {
                continue;
            }
            const name = member.name || '';
            if (name.startsWith('users/')) {
                ids.push(name.substring('users/'.length));
            }
        }

        pageToken = response.nextPageToken;
    } while (pageToken);

    return ids;
}


/** Collect the space names already used by existing polls (avoid double posting into one conversation). */
function loadUsedSpaces_() {
    const used = {};
    const properties = getScriptProperties_().getProperties() || {};
    for (const key in properties) {
        if (!key.startsWith(STATE_KEY_PREFIX)) {
            continue;
        }
        try {
            const state = JSON.parse(properties[key]);
            if (state.spaceName) {
                used[state.spaceName] = state.eventId;
            }
        } catch (e) {
            // ignore, cleanupOldState_() removes malformed entries
        }
    }
    return used;
}


// ---------------------------------------------------------------------------
// Consent Poll Sending
// ---------------------------------------------------------------------------

/** Post the consent request into the meeting conversation as the organizer. */
function sendConsentPoll_(context, event) {
    const organizerEmail = event.organizer?.email;
    if (!organizerEmail) {
        Logger.log('No organizer email for event %s, skipping', event.id);
        return;
    }

    const space = findMeetingChatSpace_(context, organizerEmail, event);
    if (!space) {
        Logger.log('No matching meeting conversation found for event %s (%s), will retry on next trigger run',
            event.id, event.summary);
        return;
    }

    Logger.log('Found meeting conversation %s for event %s (%s)', space.name, event.id, event.summary);

    const directory = getDirectory_(context, organizerEmail);
    const attendees = (event.attendees || [])
        .map(a => (a.email || '').toLowerCase())
        .filter(e => !!e && !e.endsWith('calendar.google.com'));   // drop rooms/resources

    const names = {};
    for (const email of attendees) {
        const person = directory.emailToPerson[email];
        names[email] = person ? person.name : email;
    }

    const state = {
        eventId: event.id,
        token: Util.generateUUIDv4(),
        meetingTitle: event.summary || '(No title)',
        meetingTime: event.start.dateTime || event.start.date,
        organizerEmail: organizerEmail,
        spaceName: space.name,
        attendees: attendees,
        names: names,
        responses: {},
        pollSentAt: Date.now()
    };

    const messagesClient = createMessagesClient_(context.creds, organizerEmail);
    const message = messagesClient.postJson(CHAT_API_BASE + '/' + space.name + '/messages', {
        text: buildPollText_(state, context.webAppUrl)
    });

    state.messageName = message.name;
    saveState_(event.id, state);
    context.usedSpaces[space.name] = event.id;

    // Mark as sent (dedup)
    getScriptProperties_().setProperty(SENT_KEY_PREFIX + event.id, '' + Date.now());

    Logger.log('Posted consent request for event %s (%s) as %s in %s', event.id, event.summary, organizerEmail, space.name);
}


/** Create a Chat client impersonating the given user with the chat.messages scope. */
function createMessagesClient_(creds, userEmail) {
    const service = UrlFetchJsonClient.createImpersonatingService(
        'ConsentPollMessages-' + userEmail, creds, userEmail, SCOPE_CHAT_MESSAGES);
    return new UrlFetchJsonClient(service);
}


/** Update the poll message text in place (as the organizer, who authored it). */
function updatePollMessage_(state) {
    if (!state.messageName) {
        return;
    }
    const messagesClient = createMessagesClient_(getServiceAccountCredentials_(), state.organizerEmail);
    messagesClient.patchJson(CHAT_API_BASE + '/' + state.messageName + '?updateMask=text', {
        text: buildPollText_(state, getWebAppUrl_())
    });
}


// ---------------------------------------------------------------------------
// Message Text
// ---------------------------------------------------------------------------

/** Build the consent request text (Google Chat text formatting: *bold*, _italic_, <url|label>). */
function buildPollText_(state, webAppUrl) {
    const lines = [];

    lines.push('🎥 *Recording consent: ' + state.meetingTitle + '*');
    lines.push('Scheduled: ' + formatMeetingTime_(state.meetingTime));
    lines.push('This meeting may be recorded. Please confirm whether you consent to being recorded:');

    if (webAppUrl) {
        const consentUrl = buildResponseUrl_(webAppUrl, state, true);
        const declineUrl = buildResponseUrl_(webAppUrl, state, false);
        lines.push('✅ <' + consentUrl + '|I consent>    ❌ <' + declineUrl + '|I do not consent>');
    } else {
        lines.push('_(Web app URL not configured, responses cannot be collected.)_');
    }
    lines.push('_External guests: please state your consent verbally, the links only work for ' +
        state.organizerEmail.split('@')[1] + ' accounts._');

    const responses = state.responses || {};
    const attendees = state.attendees || [];
    const responseCount = Object.keys(responses).length;

    if (responseCount > 0) {
        const consented = [];
        const declined = [];
        const pending = [];

        for (const email of Object.keys(responses)) {
            (responses[email].consent ? consented : declined).push(displayName_(state, email));
        }
        for (const email of attendees) {
            if (!responses[email]) {
                pending.push(displayName_(state, email));
            }
        }

        // Responders that were not on the invite (e.g. forwarded invitation) count towards total
        const total = attendees.length + Object.keys(responses).filter(e => !attendees.includes(e)).length;

        lines.push('');
        lines.push(buildProgressBar_(responseCount, total) + ' *' + responseCount + '/' + total + ' responded*');
        if (consented.length > 0) {
            lines.push('✅ Consented: ' + truncateNameList_(consented));
        }
        if (declined.length > 0) {
            lines.push('❌ Declined: ' + truncateNameList_(declined));
        }
        if (pending.length > 0) {
            lines.push('⏳ Pending: ' + truncateNameList_(pending));
        }
    }

    return lines.join('\n');
}


/** Build a response link for the web app. */
function buildResponseUrl_(webAppUrl, state, consent) {
    return webAppUrl + UrlFetchJsonClient.buildQuery({
        [PARAM_EVENT]: state.eventId,
        [PARAM_TOKEN]: state.token,
        [PARAM_CONSENT]: consent ? '1' : '0'
    });
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
 * Query parameters: e=<eventId>, t=<token>, c=1|0 (omit c to just show the status).
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

    if (consentParam === undefined) {
        return renderPage_('Recording Consent', buildStatusHtml_(state, null));
    }

    if (!userEmail) {
        return renderPage_('Recording Consent',
            '<p>Could not determine your identity. Please open this link in a browser signed in with your '
            + escapeHtml_(state.organizerEmail.split('@')[1]) + ' account.</p>');
    }

    const consent = consentParam === '1';
    const updatedState = recordResponse_(eventId, userEmail, consent);
    if (!updatedState) {
        return renderPage_('Recording Consent', '<p>This consent poll is no longer active.</p>');
    }

    let notice = '';
    try {
        updatePollMessage_(updatedState);
    } catch (err) {
        Logger.log('Failed to update poll message for event %s: %s', eventId, err);
        notice = '<p class="muted">Your response was recorded, but the chat message could not be updated.</p>';
    }

    const headline = consent
        ? '<p class="ok">✅ Thank you, your consent has been recorded.</p>'
        : '<p class="no">❌ Thank you, your objection has been recorded. The organizer has been informed via the meeting chat.</p>';

    return renderPage_('Recording Consent', headline + notice + buildStatusHtml_(updatedState, userEmail));
}


/** Record a response under a script lock, returns the updated state or null if the poll is gone. */
function recordResponse_(eventId, userEmail, consent) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const state = loadState_(eventId);
        if (!state) {
            return null;
        }

        state.responses = state.responses || {};
        state.responses[userEmail] = {
            consent: consent,
            name: state.names?.[userEmail] || userEmail,
            respondedAt: Date.now()
        };

        saveState_(eventId, state);
        return state;
    } finally {
        lock.releaseLock();
    }
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

    return '<h2>' + escapeHtml_(state.meetingTitle) + '</h2>'
        + '<p class="muted">' + escapeHtml_(formatMeetingTime_(state.meetingTime)) + '</p>'
        + '<p>' + Object.keys(responses).length + '/' + all.length + ' responded</p>'
        + '<table>' + rows.join('') + '</table>';
}


/** Wrap body HTML into a minimal page. */
function renderPage_(title, bodyHtml) {
    const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
        + '<title>' + escapeHtml_(title) + '</title>'
        + '<style>body{font-family:system-ui,sans-serif;max-width:40em;margin:2em auto;padding:0 1em;color:#222}'
        + 'table{border-collapse:collapse}td{padding:.25em .75em .25em 0}.ok{color:#1b7f3b;font-weight:bold}'
        + '.no{color:#b3261e;font-weight:bold}.muted{color:#666}</style></head>'
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

/** Save consent poll state to ScriptProperties. */
function saveState_(eventId, state) {
    getScriptProperties_().setProperty(STATE_KEY_PREFIX + eventId, JSON.stringify(state));
}

/** Clean up state entries older than STATE_MAX_AGE_MS. */
function cleanupOldState_() {
    const now = Date.now();
    const properties = getScriptProperties_().getProperties() || {};

    for (const key in properties) {
        if (key.startsWith(STATE_KEY_PREFIX)) {
            try {
                const state = JSON.parse(properties[key]);
                if (state.pollSentAt && (now - state.pollSentAt) > STATE_MAX_AGE_MS) {
                    getScriptProperties_().deleteProperty(key);
                    Logger.log('Cleaned up old state: %s', key);
                }
            } catch (e) {
                // Malformed state, clean it up
                getScriptProperties_().deleteProperty(key);
            }
        } else if (key.startsWith(SENT_KEY_PREFIX)) {
            const sentAt = +properties[key];
            if (sentAt && (now - sentAt) > STATE_MAX_AGE_MS) {
                getScriptProperties_().deleteProperty(key);
                Logger.log('Cleaned up old sent marker: %s', key);
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
    const raw = (getScriptProperties_().getProperty(LOOKAHEAD_MINUTES_KEY) || '').trim();
    const minutes = Math.abs(Math.round(+raw));
    if (!minutes || Number.isNaN(minutes)) {
        return 15;
    }
    return minutes;
}

/** Get the minimum number of attendees or the default (2). */
function getMinAttendees_() {
    const raw = (getScriptProperties_().getProperty(MIN_ATTENDEES_KEY) || '').trim();
    const count = Math.abs(Math.round(+raw));
    if (!count || Number.isNaN(count)) {
        return 2;
    }
    return count;
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

/** Build a Unicode progress bar. */
function buildProgressBar_(current, total) {
    if (total <= 0) {
        return '';
    }
    const filled = Math.min(10, Math.round((current / total) * 10));
    return '█'.repeat(filled) + '░'.repeat(10 - filled);
}

/** Truncate a list of names for display, adding "and N more" if needed. */
function truncateNameList_(names) {
    if (names.length <= MAX_DISPLAY_NAMES) {
        return names.join(', ');
    }
    const shown = names.slice(0, MAX_DISPLAY_NAMES);
    const remaining = names.length - MAX_DISPLAY_NAMES;
    return shown.join(', ') + ' and ' + remaining + ' more';
}
