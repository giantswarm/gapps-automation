/**
 * Meeting Recording Consent Poll
 *
 * Asks every attendee of Google Meet meetings that may be recorded (automatic recording, transcription
 * or Gemini notes enabled) for consent before the meeting, by email with a personal link to this script's
 * web app, and documents every answer.
 *
 * Runs entirely in Apps Script:
 *   - checkUpcomingMeetings() (time trigger) scans the primary calendars of the organizers for qualifying
 *     meetings starting within the lookahead window, creates the poll state (keyed by event ID), adds a
 *     generic notice to the event description, emails every attendee (internal and external) a personal
 *     link once, sends the organizer a summary shortly before the meeting (and updates after that), and
 *     handles moved and cancelled meetings
 *   - doGet() shows the question for the attendee identified by the personal link (token), the buttons
 *     call submitConsent(), which appends the answer to the consent log spreadsheet (the documentation of
 *     consent, the answer is rejected if that fails) before it counts
 *   - On an objection the organizer is emailed right away and automatic recording/transcription/notes are
 *     switched off via the Meet REST API (spaces.patch, impersonating the organizer)
 *
 * Why not a bot in the meeting: Apps Script can not join a Meet call and posting into the meeting chat via
 * the Chat API is not reliably possible (meeting chats are system generated group DMs). Asking by email
 * before the meeting is also the better fit for consent "before recording".
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

/** Invitations are sent for meetings starting within this many hours.
 *
 * Default: 24 hours
 */
const LOOKAHEAD_HOURS_KEY = PROPERTY_PREFIX + 'lookaheadHours';

/** The organizer summary is sent this many minutes before the meeting starts (by the first run after that).
 *
 * Default: 15 minutes
 */
const SUMMARY_MINUTES_KEY = PROPERTY_PREFIX + 'summaryMinutes';

/** Minimum number of attendees for a qualifying meeting.
 *
 * Default: 2
 */
const MIN_ATTENDEES_KEY = PROPERTY_PREFIX + 'minAttendees';

/** Which meetings get a consent poll: 'recorded' or 'all'.
 *
 * 'recorded': automatic recording, transcription or notes are enabled for the Meet space (Meet REST API,
 * impersonating the organizer). 'all': every qualifying meeting.
 *
 * Default: 'recorded'
 */
const JOIN_MODE_KEY = PROPERTY_PREFIX + 'joinMode';

/** The account the emails are sent from (Gmail API, impersonated via domain-wide delegation). Required.
 *
 * Excluded from the attendees in case it is invited.
 */
const SENDER_EMAIL_KEY = PROPERTY_PREFIX + 'senderEmail';

/** Whether a generic notice is added to the event description ('true' or 'false').
 *
 * Default: true
 */
const CALENDAR_NOTICE_KEY = PROPERTY_PREFIX + 'calendarNotice';

/** Whether automatic recording/transcription/notes are switched off when an attendee objects ('true' or 'false').
 *
 * Requires the 'meetings.space.settings' scope (instead of 'meetings.space.readonly').
 *
 * Default: true
 */
const DISABLE_RECORDING_KEY = PROPERTY_PREFIX + 'disableRecordingOnObjection';

/** Optional purpose of recordings, completes "... may be recorded" (e.g. "for AI generated meeting minutes"). */
const RECORDING_PURPOSE_KEY = PROPERTY_PREFIX + 'recordingPurpose';

/** Optional link to the recording policy, shown in the emails and on the poll page. */
const POLICY_URL_KEY = PROPERTY_PREFIX + 'policyUrl';

/** Public URL of this script's web app deployment (the /exec URL).
 *
 * Default: ScriptApp.getService().getUrl()
 */
const WEB_APP_URL_KEY = PROPERTY_PREFIX + 'webAppUrl';

/** ID of the spreadsheet every response is appended to (the durable documentation of consent). Required.
 *
 * The poll state in ScriptProperties is deleted a day after the meeting, this log is kept.
 */
const CONSENT_LOG_SPREADSHEET_ID_KEY = PROPERTY_PREFIX + 'consentLogSpreadsheetId';

/** Name of the sheet in the consent log spreadsheet (created if missing). */
const CONSENT_LOG_SHEET_NAME = 'Consent Log';

/** Columns of the consent log. */
const CONSENT_LOG_HEADER = ['Responded At', 'Respondent Email', 'Respondent Name', 'Consent', 'Meeting Title',
    'Meeting Start', 'Meeting Link', 'Organizer', 'Event ID', 'Previous Answer', 'Identified By', 'Signed In As'];

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'checkUpcomingMeetings';

/** Prefix for poll state stored in ScriptProperties (keyed by event ID). */
const STATE_KEY_PREFIX = 'state.';

/** Version of the poll state format, states of other versions are deleted. */
const STATE_VERSION = 2;

/** Poll state is kept this long after the meeting ended (so the poll page can say it is closed). */
const STATE_RETENTION_AFTER_END_MS = 24 * 60 * 60 * 1000;

/** Maximum size of a serialized poll state per property (ScriptProperties allow 9 KB per value), larger states are split. */
const MAX_STATE_BYTES = 8500;

/** Prefix for the parts of large poll states ('chunk.<eventId>.<n>'). */
const CHUNK_KEY_PREFIX = 'chunk.';

/** Value of the 'state.<eventId>' property of a split state, followed by the number of parts. */
const CHUNKED_MARKER = '#chunks:';

/** Maximum number of parts of a poll state (about 100 KB, several hundred attendees). */
const MAX_STATE_CHUNKS = 12;

/** Log a warning when the script properties grow beyond this size (ScriptProperties allow 500 KB in total). */
const PROPERTIES_WARN_BYTES = 400 * 1024;

/** Start of the notice added to event descriptions, also used to detect it. */
const NOTICE_MARKER = 'Recording consent notice:';

/** Display name of the sender of all emails. */
const SENDER_NAME = 'Meeting Recording Consent';

/** OAuth scopes used via domain-wide delegation (besides the ones of the lib clients). */
const GMAIL_SEND_SCOPE = 'https://www.googleapis.com/auth/gmail.send';
const MEET_SETTINGS_SCOPE = 'https://www.googleapis.com/auth/meetings.space.settings';

/** Automatic artifacts of a Meet space (SpaceConfig.artifactConfig). */
const ARTIFACTS = [
    {config: 'recordingConfig', field: 'autoRecordingGeneration', label: 'recording'},
    {config: 'transcriptionConfig', field: 'autoTranscriptionGeneration', label: 'transcription'},
    {config: 'smartNotesConfig', field: 'autoSmartNotesGeneration', label: 'Gemini notes'}
];

/** Web app query parameter names. */
const PARAM_EVENT = 'e';
const PARAM_TOKEN = 't';


// ---------------------------------------------------------------------------
// Standard Setup (from personio-to-group/PersonioToGroup.js pattern)
// ---------------------------------------------------------------------------

/** Uninstall triggers. */
function uninstall() {
    TriggerUtil.uninstall(TRIGGER_HANDLER_FUNCTION);
}

/** Install periodic execution trigger (replacing an existing one). */
function install(delayMinutes) {
    uninstall();
    TriggerUtil.install(TRIGGER_HANDLER_FUNCTION, delayMinutes);
}

/** Allow setting properties. */
function setProperties(properties, deleteAllOthers) {
    TriggerUtil.setProperties(properties, deleteAllOthers);
}


// ---------------------------------------------------------------------------
// Core Trigger: checkUpcomingMeetings
// ---------------------------------------------------------------------------

/** Main entry point: scans organizer calendars for upcoming recorded meetings and runs their polls. */
function checkUpcomingMeetings() {
    // runs overlapping each other (a slow run) would send invitations twice
    const runLock = LockService.getUserLock();
    if (!runLock.tryLock(1000)) {
        Logger.log('Another run is still in progress, skipping this one');
        return;
    }
    try {
        checkUpcomingMeetings_();
    } finally {
        runLock.releaseLock();
    }
}


function checkUpcomingMeetings_() {
    const config = loadConfig_();
    if (!config.webAppUrl) {
        throw new Error('Web app URL unknown, deploy the web app or set ' + WEB_APP_URL_KEY);
    }
    getConsentLogSheet_();   // fail early if responses can not be documented

    const isEmailAllowed = rawEmail => {
        const email = (rawEmail || '').trim().toLowerCase();
        return (!config.emailWhiteList.length || config.emailWhiteList.includes(email))
            && config.allowedDomains.some(domain => email.endsWith('@' + domain));
    };

    Logger.log('Configured to handle accounts %s on domains %s, lookahead %s h, summary %s min before, min attendees %s, '
        + 'join mode %s, sender %s, calendar notice %s, disable recording on objection %s, web app %s',
        config.emailWhiteList.length ? config.emailWhiteList : '(all)', config.allowedDomains, '' + config.lookaheadHours,
        '' + config.summaryMinutes, '' + config.minAttendees, config.joinMode, config.senderEmail, '' + config.calendarNotice,
        '' + config.disableRecording, config.webAppUrl);

    const personioCreds = getPersonioCreds_();
    const personio = PersonioClientV1.withApiCredentials(personioCreds.clientId, personioCreds.clientSecret);
    const employees = personio.getPersonioJson('/company/employees');
    const activeEmployees = employees.filter(employee =>
        employee.attributes.status.value !== 'inactive' && isEmailAllowed(employee.attributes.email.value)
    );

    Logger.log('Processing %s active employees', '' + activeEmployees.length);

    const context = createContext_(config);
    const now = context.now;
    const timeMin = new Date(now).toISOString();
    const timeMax = new Date(now + config.lookaheadHours * 60 * 60 * 1000).toISOString();

    const seenEventIds = new Set();
    const scannedOrganizers = new Set();
    const failedOrganizers = new Set();
    const fail = (e) => context.firstError = context.firstError || e;

    for (const employee of activeEmployees) {
        const email = employee.attributes.email.value.trim().toLowerCase();
        try {
            const calendar = CalendarClient.withImpersonatingService(context.creds, email);
            // lists events ending after timeMin, so ongoing meetings are included
            const events = calendar.list('primary', {
                singleEvents: true,
                showDeleted: false,
                timeMin: timeMin,
                timeMax: timeMax
            });

            for (const event of events) {
                // Only process events where this employee is the organizer (dedup across calendars)
                if ((event.organizer?.email || '').toLowerCase() !== email) {
                    continue;
                }

                try {
                    if (isQualifyingEvent_(event, config.minAttendees)) {
                        processEvent_(context, calendar, event);
                        seenEventIds.add(event.id);
                    } else if (loadState_(event.id)) {
                        // e.g. the Meet link or attendees were removed
                        deactivatePoll_(event.id, 'no longer a qualifying meeting');
                        seenEventIds.add(event.id);
                    }
                } catch (e) {
                    Logger.log('Failed to process event %s (%s): %s', event.id, event.summary, e);
                    seenEventIds.add(event.id);   // unknown, leave the poll alone
                    fail(e);
                }
            }
            scannedOrganizers.add(email);
        } catch (e) {
            Logger.log('Failed to process calendar for user %s: %s', email, e);
            failedOrganizers.add(email);
            fail(e);
        }
    }

    Logger.log('Completed scanning %s of %s accounts', '' + scannedOrganizers.size, '' + activeEmployees.length);

    // Meetings that were cancelled, moved out of the window or whose objection was not handled yet
    try {
        reconcileUnseenPolls_(context, seenEventIds, scannedOrganizers, failedOrganizers);
    } catch (e) {
        Logger.log('Failed to reconcile polls of meetings not seen in this run: %s', e);
        fail(e);
    }

    cleanupOldState_(now);

    if (context.firstError) {
        throw context.firstError;
    }
}


/** Per-run (or per-request) context: configuration, credentials and caches. */
function createContext_(config) {
    return {
        config: config,
        creds: getServiceAccountCredentials_(),
        now: Date.now(),
        directory: undefined,   // lazily loaded: {emailToName: {}}, null if loading failed
        gmail: null,
        firstError: null
    };
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


/** The event's Meet meeting code (e.g. 'abc-mnop-xyz'). */
function getMeetingCode_(event) {
    return event.conferenceData?.conferenceId
        || (getMeetUri_(event) || '').replace(/^https:\/\/meet\.google\.com\//, '').replace(/[?#].*$/, '');
}




/** The labels of the automatic artifacts switched on for a Meet space (e.g. ['recording', 'Gemini notes']). */
function getAutoArtifacts_(space) {
    const artifactConfig = space?.config?.artifactConfig || {};
    return ARTIFACTS.filter(a => artifactConfig[a.config]?.[a.field] === 'ON').map(a => a.label);
}


/** Get the Meet space of a meeting as its organizer, or null if that fails (e.g. a Meet link owned by someone else).
 *
 * With disableRecordingOnObjection the settings scope is used, which allows reading and updating.
 */
function getSpace_(context, organizerEmail, spaceOrCode) {
    try {
        return getMeetClient_(context, organizerEmail).getSpace(spaceOrCode.startsWith('spaces/') ? spaceOrCode : 'spaces/' + spaceOrCode);
    } catch (e) {
        Logger.log('Failed to get Meet space %s as %s: %s', spaceOrCode, organizerEmail, e);
        return null;
    }
}


function getMeetClient_(context, organizerEmail) {
    return context.config.disableRecording
        ? MeetClient.withImpersonatingService(context.creds, organizerEmail, MEET_SETTINGS_SCOPE)
        : MeetClient.withImpersonatingService(context.creds, organizerEmail);
}


// ---------------------------------------------------------------------------
// Poll Processing
// ---------------------------------------------------------------------------

/** Run the poll of a qualifying event: create/update state, enforce objections, notice, invitations, summary. */
function processEvent_(context, calendar, event) {
    const config = context.config;
    const organizerEmail = event.organizer.email.toLowerCase();
    const existing = loadState_(event.id);

    // The Meet space is needed to know whether the meeting is recorded and to enforce standing objections
    const needSpace = config.joinMode === 'recorded'
        || (!!existing && config.disableRecording && hasObjection_(existing));
    const space = needSpace ? getSpace_(context, organizerEmail, existing?.space || getMeetingCode_(event)) : null;
    const autoArtifacts = getAutoArtifacts_(space);

    let recorded;
    if (config.joinMode === 'all') {
        recorded = true;
    } else if (space) {
        // recording switched off by us due to an objection keeps the poll going
        recorded = autoArtifacts.length > 0 || !!existing?.recOff;
    } else {
        recorded = !!existing?.active;   // unknown, keep as is
    }

    Logger.log('Event %s (%s) by %s: automatic %s, poll %s', event.id, event.summary, organizerEmail,
        space ? (autoArtifacts.join('/') || 'none') : '(unknown)', recorded ? 'active' : 'not needed');

    if (!recorded) {
        if (existing?.active) {
            deactivatePoll_(event.id, 'no longer recorded');
        }
        if (existing?.objPending) {
            enforceObjection_(context, event.id, space, 'objection');   // still tell the organizer
        }
        return;
    }

    upsertPoll_(context, event, space);

    // Objections not handled yet (e.g. the notification failed) or automatic recording switched on again
    const state = loadState_(event.id);
    if (state.objPending || (config.disableRecording && autoArtifacts.length > 0 && hasObjection_(state))) {
        try {
            enforceObjection_(context, event.id, space, state.objPending ? 'objection' : 'reenabled');
        } catch (e) {
            Logger.log('Failed to handle the objection for event %s (%s): %s', event.id, event.summary, e);
            context.firstError = context.firstError || e;
        }
    }

    if (config.calendarNotice && !state.noticeAt) {
        try {
            addCalendarNotice_(context, calendar, event.id);
        } catch (e) {
            Logger.log('Failed to add the notice to event %s (%s): %s', event.id, event.summary, e);
            context.firstError = context.firstError || e;
        }
    }

    sendInvitations_(context, event.id);
    sendSummaryIfDue_(context, event.id);
}


/** Create or update the poll state of an event from the event (attendees, time, title). */
function upsertPoll_(context, event, space) {
    const organizerEmail = event.organizer.email.toLowerCase();
    const senderEmail = context.config.senderEmail;

    // drop rooms/resources, people who declined the invitation and the sender
    const attendees = (event.attendees || [])
        .filter(a => !a.resource && a.responseStatus !== 'declined')
        .map(a => ({email: (a.email || '').toLowerCase(), displayName: a.displayName}))
        .filter(a => isValidEmail_(a.email) && !a.email.endsWith('calendar.google.com') && a.email !== senderEmail);

    const directory = getDirectory_(context, organizerEmail);
    const nameOf = a => ((directory?.emailToName[a.email]) || a.displayName || '').substring(0, 80);

    return updateState_(event.id, state => {
        state = state || {
            v: STATE_VERSION,
            id: event.id,
            createdAt: context.now,
            att: {}
        };

        Object.assign(state, {
            org: organizerEmail,
            title: (event.summary || '(No title)').substring(0, 200),
            start: event.start.dateTime,
            end: event.end?.dateTime || new Date(new Date(event.start.dateTime).getTime() + 60 * 60 * 1000).toISOString(),
            tz: event.start.timeZone || undefined,
            uri: getMeetUri_(event),
            code: getMeetingCode_(event),
            space: space?.name || state.space || undefined,
            active: true,
            why: undefined,
            cancelled: undefined
        });

        const current = new Set(attendees.map(a => a.email));
        for (const a of attendees) {
            const entry = state.att[a.email];
            if (!entry) {
                state.att[a.email] = {t: newToken_()};
            } else {
                delete entry.x;
            }
            const name = nameOf(a);
            if (name && name !== a.email) {
                state.att[a.email].n = name;
            }
        }
        // removed attendees are kept (marked), so they are not invited twice if added again
        for (const email of Object.keys(state.att)) {
            if (!current.has(email)) {
                state.att[email].x = 1;
            }
        }
        return state;
    });
}


/** Mark a poll inactive (no invitations or summaries), e.g. because the meeting is no longer recorded. */
function deactivatePoll_(eventId, reason, cancelled) {
    const changed = updateState_(eventId, state => {
        if (!state || (!state.active && !!state.cancelled === !!cancelled)) {
            return null;
        }
        state.active = false;
        state.why = reason;
        if (cancelled) {
            state.cancelled = Date.now();
        }
        return state;
    });
    if (changed) {
        Logger.log('Deactivated the poll of event %s: %s', eventId, reason);
    }
}


/** Whether the poll accepts answers (until the meeting ended, unless cancelled). */
function isPollOpen_(state, now) {
    return !state.cancelled && now < new Date(state.end).getTime();
}


/** Current (not removed) attendees as [email, entry] pairs. */
function currentAttendees_(state) {
    return Object.entries(state.att || {}).filter(([, entry]) => !entry.x);
}


/** Whether a current attendee does not consent. */
function hasObjection_(state) {
    return currentAttendees_(state).some(([, entry]) => entry.c === false);
}


/** Count answers of the current attendees. */
function countAnswers_(state) {
    const entries = currentAttendees_(state).map(([, entry]) => entry);
    const consented = entries.filter(entry => entry.c === true).length;
    const declined = entries.filter(entry => entry.c === false).length;
    return {
        invited: entries.length,
        consented: consented,
        declined: declined,
        pending: entries.length - consented - declined
    };
}


/** Email every current attendee who was not invited yet their personal link (each person exactly once). */
function sendInvitations_(context, eventId) {
    const state = loadState_(eventId);
    if (!state?.active || !isPollOpen_(state, context.now)) {
        return;
    }

    for (const [email, entry] of currentAttendees_(state)) {
        if (entry.i) {
            continue;
        }
        try {
            sendMail_(context, email, state.org, 'Recording consent: ' + state.title,
                buildInvitationText_(context, state, email, entry), buildInvitationHtml_(context, state, email, entry));
        } catch (e) {
            Logger.log('Failed to send the invitation for event %s (%s) to %s: %s', eventId, state.title, email, e);
            context.firstError = context.firstError || e;
            continue;
        }
        // persisted right after each mail, so a failing run does not invite anyone twice
        updateState_(eventId, current => {
            if (!current?.att?.[email]) {
                return null;
            }
            current.att[email].i = Date.now();
            return current;
        });
        Logger.log('Invited %s to the consent poll of event %s (%s)', email, eventId, state.title);
    }
}


/** Send the organizer summary shortly before the meeting (once), and an update whenever answers change after that. */
function sendSummaryIfDue_(context, eventId) {
    const state = loadState_(eventId);
    if (!state?.active || !isPollOpen_(state, context.now)) {
        return;
    }

    const signature = answersSignature_(state);
    const start = new Date(state.start).getTime();
    let kind;
    if (!state.summaryAt) {
        if (context.now < start - context.config.summaryMinutes * 60 * 1000) {
            return;
        }
        kind = 'summary';
    } else if (state.sumSig !== signature) {
        kind = 'update';
    } else {
        return;
    }

    const subject = (kind === 'summary' ? 'Recording consent summary: ' : 'Recording consent update: ') + state.title;
    sendMail_(context, state.org, null, subject, buildSummaryText_(context, state, kind), buildSummaryHtml_(context, state, kind));

    updateState_(eventId, current => {
        if (!current) {
            return null;
        }
        current.summaryAt = current.summaryAt || Date.now();
        current.sumSig = signature;
        return current;
    });
    Logger.log('Sent the %s for event %s (%s) to %s', kind, eventId, state.title, state.org);
}


/** A short signature of the current answers, to detect changes after the summary was sent. */
function answersSignature_(state) {
    const text = currentAttendees_(state)
        .map(([email, entry]) => email + '=' + (entry.c === true ? 'y' : entry.c === false ? 'n' : '-'))
        .sort()
        .join(',');
    return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text,
        Utilities.Charset.UTF_8)).substring(0, 16);
}


/** Handle an objection: switch off automatic recording/transcription/notes (if enabled) and email the organizer.
 *
 * @param context The context (config, creds).
 * @param eventId The event ID.
 * @param space The Meet space if already fetched, else null.
 * @param reason 'objection' (new objection) or 'reenabled' (automatic artifacts were switched on again despite an objection).
 */
function enforceObjection_(context, eventId, space, reason) {
    const state = loadState_(eventId);
    if (!state || state.cancelled) {
        return;
    }

    const result = {disabled: [], failed: null, attempted: false};
    if (context.config.disableRecording && isPollOpen_(state, Date.now())) {
        result.attempted = true;
        try {
            result.disabled = disableAutoArtifacts_(context, state, space);
        } catch (e) {
            Logger.log('Failed to switch off automatic recording for event %s (%s): %s', eventId, state.title, e);
            result.failed = e;
        }
    }

    if (reason === 'reenabled' && result.failed && state.offFailed) {
        // the organizer was already told to check the settings, keep trying silently on the next runs
        return;
    }

    try {
        sendMail_(context, state.org, null, 'Recording consent: objection for ' + state.title,
            buildObjectionText_(context, state, result, reason), buildObjectionHtml_(context, state, result, reason));
    } catch (e) {
        // retried by the next trigger run
        Logger.log('Failed to notify the organizer %s about an objection for event %s (%s): %s', state.org, eventId, state.title, e);
        updateState_(eventId, current => {
            if (!current) {
                return null;
            }
            current.objPending = current.objPending || Date.now();
            addRecOff_(current, result.disabled);
            return current;
        });
        throw e;
    }

    updateState_(eventId, current => {
        if (!current) {
            return null;
        }
        delete current.objPending;
        addRecOff_(current, result.disabled);
        // the organizer was told to check the settings, do not repeat that on every run
        if (result.failed) {
            current.offFailed = Date.now();
        } else {
            delete current.offFailed;
        }
        return current;
    });
    Logger.log('Handled %s for event %s (%s): notified %s, switched off %s', reason, eventId, state.title, state.org,
        result.disabled.join('/') || 'nothing');
}


/** Remember which automatic artifacts were switched off (the poll then stays active even with joinMode 'recorded'). */
function addRecOff_(state, disabled) {
    if (disabled.length) {
        state.recOff = {
            at: Date.now(),
            off: Array.from(new Set([...(state.recOff?.off || []), ...disabled]))
        };
    }
}


/** Switch off the automatic artifacts of the meeting's Meet space via spaces.patch, returns the labels switched off.
 *
 * Throws if the space can not be read or updated.
 */
function disableAutoArtifacts_(context, state, space) {
    const meet = getMeetClient_(context, state.org);
    const current = space || meet.getSpace(state.space || 'spaces/' + state.code);
    const on = ARTIFACTS.filter(a => current?.config?.artifactConfig?.[a.config]?.[a.field] === 'ON');
    if (!on.length) {
        return [];
    }

    const artifactConfig = {};
    for (const a of on) {
        artifactConfig[a.config] = {[a.field]: 'OFF'};
    }
    // spaces.patch only accepts the server generated space name, not the meeting code
    meet.patchSpace(current.name, {config: {artifactConfig: artifactConfig}},
        on.map(a => 'config.artifactConfig.' + a.config + '.' + a.field).join(','));

    Logger.log('Switched off automatic %s of %s for event %s (%s)', on.map(a => a.label).join('/'), current.name,
        state.id, state.title);
    return on.map(a => a.label);
}


/** Add a generic notice to the event description as the organizer (once, never clobbering concurrent edits). */
function addCalendarNotice_(context, calendar, eventId) {
    // read fresh and patch with If-Match, so a concurrent edit by the organizer is never overwritten
    const event = calendar.get('primary', eventId);
    const description = event.description || '';

    if (!description.includes(NOTICE_MARKER)) {
        const notice = buildCalendarNotice_(context);
        const isHtml = /<\/?[a-z][^>]*>/i.test(description);
        const updated = !description.trim() ? notice
            : isHtml ? description + '<br><br>' + escapeHtml_(notice)
                : description + '\n\n' + notice;
        calendar.patch('primary', eventId, {description: updated}, {sendUpdates: 'none'}, event.etag);
        Logger.log('Added the recording consent notice to event %s (%s)', eventId, event.summary);
    }

    updateState_(eventId, state => {
        if (!state) {
            return null;
        }
        state.noticeAt = Date.now();
        return state;
    });
}


/** Handle polls not seen in this run's scan: cancelled or moved meetings, pending objection notifications. */
function reconcileUnseenPolls_(context, seenEventIds, scannedOrganizers, failedOrganizers) {
    for (const eventId of listStateIds_()) {
        const state = loadState_(eventId);
        if (!state || seenEventIds.has(state.id) || state.cancelled || !isPollOpen_(state, context.now)) {
            continue;
        }
        const organizer = state.org;
        if (failedOrganizers.has(organizer) || !scannedOrganizers.has(organizer)) {
            continue;   // unknown (organizer not scanned), leave it alone
        }

        try {
            if (state.active) {
                reconcileUnseenPoll_(context, state);
            }
            if (loadState_(state.id)?.objPending && !loadState_(state.id)?.cancelled) {
                enforceObjection_(context, state.id, null, 'objection');
            }
        } catch (e) {
            Logger.log('Failed to reconcile the poll of event %s (%s): %s', state.id, state.title, e);
            context.firstError = context.firstError || e;
        }
    }
}


/** An active poll whose event was not in the scan: cancelled/deleted, moved out of the window or organizer changed. */
function reconcileUnseenPoll_(context, state) {
    let event = null;
    try {
        event = CalendarClient.withImpersonatingService(context.creds, state.org).get('primary', state.id);
    } catch (e) {
        const code = e.response?.getResponseCode?.();
        if (code !== 404 && code !== 410) {
            throw e;
        }
    }

    if (!event || event.status === 'cancelled') {
        deactivatePoll_(state.id, 'meeting cancelled', true);
    } else if ((event.organizer?.email || '').toLowerCase() !== state.org) {
        deactivatePoll_(state.id, 'organizer changed');
    } else if (!isQualifyingEvent_(event, context.config.minAttendees)) {
        deactivatePoll_(state.id, 'no longer a qualifying meeting');
    } else if (event.start.dateTime !== state.start || event.end?.dateTime !== state.end) {
        // moved out of the lookahead window: keep the poll (links stay valid), continue when it is due again
        updateState_(state.id, current => {
            if (!current) {
                return null;
            }
            current.start = event.start.dateTime;
            current.end = event.end?.dateTime || current.end;
            return current;
        });
        Logger.log('Event %s (%s) moved to %s', state.id, state.title, event.start.dateTime);
    }
}


// ---------------------------------------------------------------------------
// Directory (email -> name resolution)
// ---------------------------------------------------------------------------

/** Load the domain directory once per run (impersonating the given user), null if that fails. */
function getDirectory_(context, impersonatedEmail) {
    if (context.directory !== undefined) {
        return context.directory;
    }

    try {
        const directoryClient = DirectoryClient.withImpersonatingService(context.creds, impersonatedEmail);
        const people = directoryClient.listDirectoryPeople();

        const emailToName = {};
        for (const person of people) {
            const name = (person.names || []).map(n => n.displayName).find(n => !!n);
            for (const email of (person.emailAddresses || []).map(e => (e.value || '').toLowerCase()).filter(e => !!e)) {
                if (name) {
                    emailToName[email] = name;
                }
            }
        }
        Logger.log('Loaded %s directory email addresses', '' + Object.keys(emailToName).length);
        context.directory = {emailToName: emailToName};
    } catch (e) {
        Logger.log('Failed to load the directory, using calendar display names only: %s', e);
        context.directory = null;
    }
    return context.directory;
}


// ---------------------------------------------------------------------------
// Email
// ---------------------------------------------------------------------------

/** Send an email (plain text and HTML) from the sender account via the Gmail API. */
function sendMail_(context, to, replyTo, subject, text, html) {
    if (!context.gmail) {
        context.gmail = GmailClientV1.withImpersonatingService(context.creds, context.config.senderEmail, GMAIL_SEND_SCOPE);
    }
    const raw = buildMimeMessage_({
        from: context.config.senderEmail,
        fromName: SENDER_NAME,
        to: to,
        replyTo: replyTo,
        subject: subject,
        text: text,
        html: html
    });
    context.gmail.sendUserMessage(context.config.senderEmail, raw, null);
}


/** Assemble a multipart/alternative (plain text + HTML) RFC 5322 message, base64url encoded for the Gmail API. */
function buildMimeMessage_(message) {
    for (const address of [message.from, message.to, message.replyTo].filter(a => !!a)) {
        if (!isValidEmail_(address)) {
            throw new Error('Invalid email address: ' + address);
        }
    }

    const boundary = 'consent-' + Utilities.getUuid().replace(/-/g, '');
    const encodeBody = text => (Utilities.base64Encode(text, Utilities.Charset.UTF_8).match(/.{1,76}/g) || []).join('\r\n');

    const lines = [
        'From: ' + (message.fromName ? '"' + message.fromName.replace(/["\\\r\n]/g, '') + '" <' + message.from + '>' : message.from),
        'To: ' + message.to
    ];
    if (message.replyTo) {
        lines.push('Reply-To: ' + message.replyTo);
    }
    lines.push(
        'Subject: ' + encodeHeaderText_(message.subject),
        'MIME-Version: 1.0',
        'Auto-Submitted: auto-generated',
        'Content-Type: multipart/alternative;',
        ' boundary="' + boundary + '"',
        '',
        '--' + boundary,
        'Content-Type: text/plain; charset="UTF-8"',
        'Content-Transfer-Encoding: base64',
        '',
        encodeBody(message.text),
        '--' + boundary,
        'Content-Type: text/html; charset="UTF-8"',
        'Content-Transfer-Encoding: base64',
        '',
        encodeBody(message.html),
        '--' + boundary + '--',
        ''
    );

    return Utilities.base64EncodeWebSafe(Utilities.newBlob(lines.join('\r\n')).getBytes());
}


/** Encode a header value: control characters removed, RFC 2047 encoded words if it is not plain ASCII. */
function encodeHeaderText_(value) {
    const text = ('' + (value || '')).replace(/[\u0000-\u001f\u007f]+/g, ' ').trim();
    if (/^[\x20-\x7e]*$/.test(text) && text.length <= 900) {
        return text;
    }

    // lines should not exceed 78 characters: up to 42 bytes of UTF-8 become 56 base64 characters (68 per encoded word,
    // plus 'Subject: ' on the first line), characters (code points) are never split
    const chunks = [];
    let chunk = '';
    for (const c of Array.from(text)) {
        if (chunk && Util.calculateUtf8Size(chunk + c) > 42) {
            chunks.push(chunk);
            chunk = '';
        }
        chunk += c;
    }
    if (chunk) {
        chunks.push(chunk);
    }
    return chunks.map(c => '=?UTF-8?B?' + Utilities.base64Encode(c, Utilities.Charset.UTF_8) + '?=').join('\r\n ');
}


/** Check for a plain email address (nothing that could inject headers). */
function isValidEmail_(email) {
    return typeof email === 'string' && /^[^\s<>()",;:\\@]+@[a-z0-9.-]+\.[a-z]{2,}$/i.test(email);
}


/** The purpose clause, e.g. ' for AI generated meeting minutes' (or ''). */
function purposeClause_(context) {
    return context.config.recordingPurpose ? ' ' + context.config.recordingPurpose : '';
}


/** Personal poll link of an attendee. */
function buildPollUrl_(context, state, entry) {
    return context.config.webAppUrl + UrlFetchJsonClient.buildQuery({
        [PARAM_EVENT]: state.id,
        [PARAM_TOKEN]: entry.t
    });
}


function organizerLabel_(state) {
    const name = state.att?.[state.org]?.n;
    return name ? name + ' (' + state.org + ')' : state.org;
}


function buildInvitationText_(context, state, email, entry) {
    const policyUrl = context.config.policyUrl;
    return [
        'Hello' + (entry.n ? ' ' + entry.n : '') + ',',
        '',
        'you are invited to "' + state.title + '" on ' + formatMeetingTime_(state.start, state.tz)
        + ', organized by ' + organizerLabel_(state) + '.',
        '',
        'This Google Meet meeting may be recorded, transcribed or summarized with AI notes (Gemini)' + purposeClause_(context)
        + '. This only happens with the consent of everyone attending, so we ask you before the meeting:',
        '',
        'Do you consent? Please answer here (two buttons):',
        buildPollUrl_(context, state, entry),
        '',
        '- Saying no has no disadvantages for you. If anyone objects, the organizer is notified and the meeting must not be '
        + 'recorded' + (context.config.disableRecording ? ' (automatic recording is switched off)' : '') + '.',
        '- You can change or withdraw your answer on that page until the meeting ends. Afterwards, please contact the organizer.',
        '- The link is personal: it records the answer for ' + email + '. Please do not forward it.',
        '- Your answer is documented (time, your email address, the answer) as proof of consent.',
        ...(policyUrl ? ['- Recording policy: ' + policyUrl] : []),
        '',
        'Replies to this email go to the organizer.',
        '',
        '--',
        'Sent automatically on behalf of the organizer by the meeting recording consent tool.'
    ].join('\n');
}


function buildInvitationHtml_(context, state, email, entry) {
    const policyUrl = context.config.policyUrl;
    const url = buildPollUrl_(context, state, entry);
    return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#222">'
        + '<p>Hello' + (entry.n ? ' ' + escapeHtml_(entry.n) : '') + ',</p>'
        + '<p>you are invited to <b>' + escapeHtml_(state.title) + '</b> on ' + escapeHtml_(formatMeetingTime_(state.start, state.tz))
        + ', organized by ' + escapeHtml_(organizerLabel_(state)) + '.</p>'
        + '<p>This Google Meet meeting may be recorded, transcribed or summarized with AI notes (Gemini)'
        + escapeHtml_(purposeClause_(context)) + '. This only happens with the consent of everyone attending, so we ask you '
        + 'before the meeting:</p>'
        + '<p><a href="' + escapeHtml_(url) + '" style="display:inline-block;padding:.6em 1.2em;border-radius:.4em;'
        + 'background:#1a73e8;color:#fff;text-decoration:none;font-weight:bold">Answer: do you consent?</a></p>'
        + '<ul>'
        + '<li>Saying no has no disadvantages for you. If anyone objects, the organizer is notified and the meeting must not be '
        + 'recorded' + (context.config.disableRecording ? ' (automatic recording is switched off)' : '') + '.</li>'
        + '<li>You can change or withdraw your answer on that page until the meeting ends. Afterwards, please contact the organizer.</li>'
        + '<li>The link is personal: it records the answer for ' + escapeHtml_(email) + '. Please do not forward it.</li>'
        + '<li>Your answer is documented (time, your email address, the answer) as proof of consent.</li>'
        + (policyUrl ? '<li><a href="' + escapeHtml_(policyUrl) + '">Recording policy</a></li>' : '')
        + '</ul>'
        + '<p>Replies to this email go to the organizer.</p>'
        + '<p style="color:#666">Sent automatically on behalf of the organizer by the meeting recording consent tool.</p>'
        + '</body></html>';
}


/** Summary lists: [{label, people: ['Name <email>', ...]}] */
function summaryLists_(state) {
    const label = (email, entry) => entry.n ? entry.n + ' <' + email + '>' : email;
    const current = currentAttendees_(state);
    return [
        {label: 'Do NOT consent', people: current.filter(([, e]) => e.c === false).map(([m, e]) => label(m, e))},
        {label: 'Consented', people: current.filter(([, e]) => e.c === true).map(([m, e]) => label(m, e))},
        {label: 'No answer yet', people: current.filter(([, e]) => e.c !== true && e.c !== false).map(([m, e]) => label(m, e))}
    ];
}


/** The verdict line of a summary. */
function summaryVerdict_(state) {
    const counts = countAnswers_(state);
    if (counts.declined > 0) {
        return 'At least one attendee does NOT consent: do not record, transcribe or take AI notes of this meeting.';
    }
    if (counts.pending > 0) {
        return 'Not everyone answered yet: only record if everyone present consented (ask in the meeting and make sure '
            + 'they answer via their link).';
    }
    return 'All invited attendees consented.';
}


function recOffLine_(state) {
    return state.recOff ? 'Automatic ' + state.recOff.off.join(', ') + ' was switched off on '
        + formatMeetingTime_(new Date(state.recOff.at).toISOString(), state.tz) + ' because an attendee objected. '
        + 'Do not switch it on again.' : '';
}


function buildSummaryText_(context, state, kind) {
    const counts = countAnswers_(state);
    const lines = [
        (kind === 'summary' ? 'Recording consent summary' : 'Updated recording consent') + ' for "' + state.title + '" on '
        + formatMeetingTime_(state.start, state.tz) + ':',
        '',
        counts.invited + ' attendees asked: ' + counts.consented + ' consented, ' + counts.declined + ' do not consent, '
        + counts.pending + ' did not answer.',
        summaryVerdict_(state)
    ];
    if (state.recOff) {
        lines.push(recOffLine_(state));
    }
    for (const list of summaryLists_(state)) {
        lines.push('', list.label + ' (' + list.people.length + '):', ...list.people.map(p => '- ' + p));
    }
    lines.push('',
        'Guests joining without being on the invitation were not asked: ask them before recording, or add them to the event '
        + '(they get their own link with the next run).',
        'Please keep this list confidential: attendees were told that saying no has no disadvantages for them.',
        'Every answer is documented in the consent log. You get an update email when answers change before the meeting ends.');
    return lines.join('\n');
}


function buildSummaryHtml_(context, state, kind) {
    const counts = countAnswers_(state);
    let html = '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#222">'
        + '<p>' + (kind === 'summary' ? 'Recording consent summary' : 'Updated recording consent') + ' for <b>'
        + escapeHtml_(state.title) + '</b> on ' + escapeHtml_(formatMeetingTime_(state.start, state.tz)) + ':</p>'
        + '<p>' + counts.invited + ' attendees asked: ' + counts.consented + ' consented, ' + counts.declined
        + ' do not consent, ' + counts.pending + ' did not answer.</p>'
        + '<p><b>' + escapeHtml_(summaryVerdict_(state)) + '</b></p>'
        + (state.recOff ? '<p>' + escapeHtml_(recOffLine_(state)) + '</p>' : '');
    for (const list of summaryLists_(state)) {
        html += '<p>' + escapeHtml_(list.label) + ' (' + list.people.length + '):</p>'
            + (list.people.length ? '<ul>' + list.people.map(p => '<li>' + escapeHtml_(p) + '</li>').join('') + '</ul>' : '');
    }
    return html + '<p style="color:#666">Guests joining without being on the invitation were not asked: ask them before '
        + 'recording, or add them to the event. Please keep this list confidential: attendees were told that saying no '
        + 'has no disadvantages for them. Every answer is documented in the consent log. You get an update email when answers '
        + 'change before the meeting ends.</p></body></html>';
}


/** Objection notice lines (without the objector's name, the summary/update lists names). */
function objectionLines_(state, result, reason) {
    const lines = [];
    lines.push(reason === 'reenabled'
        ? 'Automatic recording, transcription or notes were switched on again for "' + state.title + '" on '
        + formatMeetingTime_(state.start, state.tz) + ', although an attendee does not consent to being recorded.'
        : 'An attendee of "' + state.title + '" on ' + formatMeetingTime_(state.start, state.tz)
        + ' does not consent to being recorded.');
    lines.push('Do not record, transcribe or take AI notes of this meeting.');
    if (result.disabled.length) {
        lines.push('Automatic ' + result.disabled.join(', ') + ' of the Meet space was switched off. Do not switch it on again.');
    } else if (result.failed) {
        lines.push('Automatic recording/transcription/notes could NOT be switched off automatically. Please check the meeting '
            + 'settings yourself (Google Calendar > event > Meet video conferencing settings).');
    } else if (result.attempted) {
        lines.push('Automatic recording, transcription and notes are not switched on for the Meet space.');
    }
    if (state.recOff && !result.disabled.length) {
        lines.push(recOffLine_(state));
    }
    lines.push('Who answered what is listed in the summary sent before the meeting (and its updates) and documented in the consent log.');
    return lines;
}


function buildObjectionText_(context, state, result, reason) {
    return objectionLines_(state, result, reason).join('\n\n');
}


function buildObjectionHtml_(context, state, result, reason) {
    return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#222">'
        + objectionLines_(state, result, reason).map(line => '<p>' + escapeHtml_(line) + '</p>').join('')
        + '</body></html>';
}


/** The generic (non-personal) notice for the event description. */
function buildCalendarNotice_(context) {
    return NOTICE_MARKER + ' this meeting may be recorded, transcribed or summarized with AI notes'
        + purposeClause_(context) + '. Every attendee gets a personal email from ' + context.config.senderEmail
        + ' asking for consent. Saying no has no disadvantages; if anyone objects, the meeting must not be recorded.'
        + (context.config.policyUrl ? ' Recording policy: ' + context.config.policyUrl : '');
}


// ---------------------------------------------------------------------------
// Web App: Response Handling
// ---------------------------------------------------------------------------

/** Web app entry point: shows the question, the answer buttons and the respondent's answer.
 *
 * Deployment: execute as the deploying user, access for anyone (including anonymous users), since external
 * attendees have no company account. The respondent is identified by the personal link (event ID + token)
 * that was emailed to their address only.
 *
 * Query parameters: e=<eventId>, t=<token>. Answers are submitted with the buttons (submitConsent()),
 * never by just opening a link, so link previews or link scanners can not record an answer.
 */
function doGet(e) {
    const params = (e && e.parameter) || {};
    const eventId = '' + (params[PARAM_EVENT] || '');
    const token = '' + (params[PARAM_TOKEN] || '');

    const found = findRespondent_(eventId, token);
    if (!found) {
        return renderPage_('<p>This consent poll link is invalid or the poll is no longer available. '
            + 'If you have questions, please contact the meeting organizer.</p>');
    }

    const state = found.state;
    const open = isPollOpen_(state, Date.now());
    return renderPage_(buildQuestionHtml_(state, open) + '<div id="result">' + buildStatusHtml_(state, found.email) + '</div>',
        open ? {eventId: eventId, token: token} : null);
}


/** Called from the web app page (google.script.run): records the answer of the link's owner, returns HTML. */
function submitConsent(eventId, token, consent) {
    if (typeof eventId !== 'string' || typeof token !== 'string' || typeof consent !== 'boolean'
        || eventId.length > 1024 || token.length > 100) {
        throw new Error('Invalid request');
    }

    let result;
    try {
        result = recordResponse_(eventId, token, consent, getActiveUserEmail_());
    } catch (e) {
        Logger.log('Failed to record an answer for event %s: %s', eventId, e);
        throw new Error('Your answer could not be documented, so it was not recorded. Please try again later or contact '
            + 'the organizer.');
    }

    if (result.status === 'invalid') {
        return '<p class="no">This consent poll link is invalid or the poll is no longer available.</p>';
    }
    if (result.status === 'closed') {
        return '<p class="no">This poll is closed. To change or withdraw your answer, please contact the organizer '
            + escapeHtml_(result.state.org) + '.</p>';
    }

    let notice = '';
    if (result.newObjection) {
        try {
            enforceObjection_(createContext_(loadConfig_()), eventId, null, 'objection');
        } catch (e) {
            Logger.log('Failed to handle the objection for event %s right away (retried later): %s', eventId, e);
            notice = '<p class="muted">The organizer will be notified shortly.</p>';
        }
    }

    const headline = consent
        ? '<p class="ok">Thank you, your consent has been recorded. You can withdraw it on this page until the meeting ends.</p>'
        : '<p class="no">Thank you, your objection has been recorded. The organizer is notified and the meeting must not be '
        + 'recorded.</p>';

    return headline + notice + buildStatusHtml_(result.state, result.email);
}


/** Find the poll and the attendee a personal link belongs to, null if invalid. */
function findRespondent_(eventId, token) {
    if (!eventId || !token) {
        return null;
    }
    const state = loadState_(eventId);
    const email = findAttendeeByToken_(state, token);
    return email ? {state: state, email: email} : null;
}


function findAttendeeByToken_(state, token) {
    if (!state || state.v !== STATE_VERSION || !token) {
        return null;
    }
    return Object.keys(state.att || {}).find(email => state.att[email].t === token) || null;
}


/** The email of the signed in person using the web app (lower case), or '' if unknown (anonymous/external). */
function getActiveUserEmail_() {
    try {
        return (Session.getActiveUser().getEmail() || '').toLowerCase();
    } catch (err) {
        return '';
    }
}


/** Record an answer under the script lock, documenting it first.
 *
 * @return {{status: string, state: Object, email: string, newObjection: boolean}} status is 'ok', 'closed' or 'invalid'
 */
function recordResponse_(eventId, token, consent, signedInAs) {
    const result = {status: 'invalid', state: null, email: null, newObjection: false};
    updateState_(eventId, state => {
        const email = findAttendeeByToken_(state, token);
        if (!email) {
            return null;
        }
        result.email = email;
        result.state = state;
        if (!isPollOpen_(state, Date.now())) {
            result.status = 'closed';
            return null;
        }

        const entry = state.att[email];
        const previous = entry.c;
        const now = Date.now();

        // Document first: an answer that can not be documented is not recorded either (the user sees an error)
        appendConsentLog_(state, email, entry, consent, previous, now, signedInAs);

        entry.c = consent;
        entry.a = now;
        if (!consent && previous !== false) {
            state.objPending = now;
            result.newObjection = true;
        }
        result.status = 'ok';
        return state;
    });
    return result;
}


/** Append an answer to the consent log spreadsheet (the durable documentation, kept after the poll state is gone). */
function appendConsentLog_(state, email, entry, consent, previous, respondedAt, signedInAs) {
    const answer = c => c === true ? 'consent' : c === false ? 'no consent' : '';
    getConsentLogSheet_().appendRow([
        new Date(respondedAt),
        email,
        entry.n || '',
        answer(consent),
        state.title,
        new Date(state.start),
        state.uri || '',
        state.org,
        state.id,
        answer(previous),
        'personal link emailed to the respondent' + (entry.i ? ' on ' + new Date(entry.i).toISOString() : ''),
        signedInAs || ''
    ].map(value => typeof value === 'string' && /^[=+\-@\t\r]/.test(value) ? "'" + value : value));   // no formula injection
}


/** The consent log sheet (created with a header row if missing), throws if not configured. */
function getConsentLogSheet_() {
    const spreadsheetId = (getScriptProperties_().getProperty(CONSENT_LOG_SPREADSHEET_ID_KEY) || '').trim();
    if (!spreadsheetId) {
        throw new Error('Consent log not configured, set ' + CONSENT_LOG_SPREADSHEET_ID_KEY);
    }
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(CONSENT_LOG_SHEET_NAME);
    if (!sheet) {
        sheet = spreadsheet.insertSheet(CONSENT_LOG_SHEET_NAME);
        sheet.appendRow(CONSENT_LOG_HEADER);
        sheet.setFrozenRows(1);
    }
    return sheet;
}


/** Render the question (with the answer buttons while the poll is open). */
function buildQuestionHtml_(state, open) {
    const purpose = (getScriptProperties_().getProperty(RECORDING_PURPOSE_KEY) || '').trim();
    const policyUrl = (getScriptProperties_().getProperty(POLICY_URL_KEY) || '').trim();

    // HtmlService pages run in an iframe, links must navigate the top window
    let html = '<h2>' + escapeHtml_(state.title) + '</h2>'
        + '<p class="muted">' + escapeHtml_(formatMeetingTime_(state.start, state.tz)) + ', organized by '
        + escapeHtml_(state.org) + '</p>'
        + '<p>This meeting may be recorded, transcribed or summarized with AI notes' + (purpose ? ' ' + escapeHtml_(purpose) : '')
        + '. Do you consent? Saying no has no disadvantages for you. If anyone objects, the meeting must not be recorded.'
        + (policyUrl ? ' See the <a target="_top" href="' + escapeHtml_(policyUrl) + '">recording policy</a>.' : '')
        + '</p>';

    if (state.cancelled) {
        return html + '<p class="no">This meeting was cancelled, the poll is closed.</p>';
    }
    if (!open) {
        return html + '<p class="no">This poll is closed. To change or withdraw your answer, please contact the organizer '
            + escapeHtml_(state.org) + '.</p>';
    }
    return html + '<p class="muted">You can change or withdraw your answer on this page until the meeting ends.</p><p>'
        + '<button class="button ok" onclick="answer(true)">I consent</button> '
        + '<button class="button no" onclick="answer(false)">I do not consent</button>'
        + '</p>';
}


/** Render the respondent's own answer and the aggregate counts (never other attendees' names or answers). */
function buildStatusHtml_(state, email) {
    const entry = state.att[email] || {};
    const counts = countAnswers_(state);
    const own = entry.c === true ? '<span class="ok">you consent</span>'
        : entry.c === false ? '<span class="no">you do not consent</span>' : 'not answered yet';
    return '<p>Your answer (' + escapeHtml_(email) + '): ' + own
        + (entry.a ? ' <span class="muted">(' + escapeHtml_(formatMeetingTime_(new Date(entry.a).toISOString(), state.tz))
            + ')</span>' : '') + '</p>'
        + '<p class="muted">All attendees: ' + (counts.consented + counts.declined) + ' of ' + counts.invited + ' answered, '
        + counts.consented + ' consent, ' + counts.declined + ' do not consent.</p>';
}


/** Wrap body HTML into a minimal page, with poll = {eventId, token} including the answer script. */
function renderPage_(bodyHtml, poll) {
    const title = 'Recording Consent';
    // answer() submits via google.script.run (a POST) and shows the result
    const script = !poll ? '' : '<script>const poll = ' + JSON.stringify(poll).replace(/</g, '\\u003c') + ';'
        + 'function answer(consent) {'
        + '  const result = document.getElementById("result");'
        + '  document.querySelectorAll("button").forEach(b => b.disabled = true);'
        + '  result.innerHTML = "<p class=\\"muted\\">Saving ...</p>";'
        + '  google.script.run'
        + '    .withSuccessHandler(html => { result.innerHTML = html; document.querySelectorAll("button").forEach(b => b.disabled = false); })'
        + '    .withFailureHandler(err => { result.textContent = "Your answer could not be saved, please try again: " + err.message;'
        + '      document.querySelectorAll("button").forEach(b => b.disabled = false); })'
        + '    .submitConsent(poll.eventId, poll.token, consent);'
        + '}</script>';
    const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
        + '<title>' + title + '</title>'
        + '<style>body{font-family:system-ui,sans-serif;max-width:40em;margin:2em auto;padding:0 1em;color:#222}'
        + '.ok{color:#1b7f3b;font-weight:bold}.no{color:#b3261e;font-weight:bold}.muted{color:#666}'
        + '.button{display:inline-block;padding:.5em 1em;margin:0 .5em .5em 0;border:2px solid;border-radius:.5em;'
        + 'background:#fff;font:inherit;cursor:pointer}.button:disabled{opacity:.5;cursor:wait}</style>' + script + '</head>'
        + '<body><h1>' + title + '</h1>' + bodyHtml + '</body></html>';
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
//
// One ScriptProperties entry per event ('state.<eventId>'), compact JSON (split into 'chunk.<eventId>.<n>' if large):
//   {v, id, org, title, start, end, tz, uri, code, space, createdAt, active, why, cancelled,
//    att: {<email>: {t: token, n: name, i: invitedAt, c: consent (true/false), a: answeredAt, x: 1 if removed}},
//    noticeAt, summaryAt, sumSig, objPending, offFailed, recOff: {at, off: [labels]}}

/** A new unguessable personal token (122 random bits from a secure random UUID). */
function newToken_() {
    return Utilities.getUuid().replace(/-/g, '');
}


/** Parse a serialized state, null if malformed or of another version. */
function parseState_(raw) {
    try {
        const state = raw ? JSON.parse(raw) : null;
        return state && state.v === STATE_VERSION ? state : null;
    } catch (e) {
        return null;
    }
}


/** The event IDs of all stored poll states. */
function listStateIds_() {
    return Object.keys(getScriptProperties_().getProperties() || {})
        .filter(key => key.startsWith(STATE_KEY_PREFIX))
        .map(key => key.substring(STATE_KEY_PREFIX.length));
}


function chunkKey_(eventId, index) {
    return CHUNK_KEY_PREFIX + eventId + '.' + index;
}


/** The number of parts of a stored state value (0 if not split). */
function chunkCount_(value) {
    return (value || '').startsWith(CHUNKED_MARKER) ? (+value.substring(CHUNKED_MARKER.length) || 0) : 0;
}


/** Load poll state from ScriptProperties (joining the parts of a split state), null if missing or malformed. */
function loadState_(eventId) {
    const properties = getScriptProperties_();
    const value = properties.getProperty(STATE_KEY_PREFIX + eventId);
    const chunks = chunkCount_(value);
    if (!chunks) {
        return parseState_(value);
    }
    const parts = [];
    for (let i = 0; i < chunks; ++i) {
        parts.push(properties.getProperty(chunkKey_(eventId, i)) || '');
    }
    return parseState_(parts.join(''));
}


/** Update poll state under a script lock: updater(state or null) returns the new state (null: no change). */
function updateState_(eventId, updater) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const state = updater(loadState_(eventId));
        if (state) {
            writeState_(eventId, state);
        }
        return state;
    } finally {
        lock.releaseLock();
    }
}


/** Store poll state, split into parts if it exceeds the ScriptProperties value size limit (call with the script lock). */
function writeState_(eventId, state) {
    const properties = getScriptProperties_();
    const json = JSON.stringify(state);
    const previousChunks = chunkCount_(properties.getProperty(STATE_KEY_PREFIX + eventId));

    const values = {};
    let chunks = 0;
    if (Util.calculateUtf8Size(json) <= MAX_STATE_BYTES) {
        values[STATE_KEY_PREFIX + eventId] = json;
    } else {
        const parts = splitUtf8_(json, MAX_STATE_BYTES);
        if (parts.length > MAX_STATE_CHUNKS) {
            throw new Error('Poll state of event ' + eventId + ' (' + state.title + ') too large ('
                + Object.keys(state.att || {}).length + ' attendees)');
        }
        parts.forEach((part, i) => values[chunkKey_(eventId, i)] = part);
        values[STATE_KEY_PREFIX + eventId] = CHUNKED_MARKER + parts.length;
        chunks = parts.length;
    }

    properties.setProperties(values, false);
    for (let i = chunks; i < previousChunks; ++i) {
        properties.deleteProperty(chunkKey_(eventId, i));
    }
}


/** Delete a poll state (with all its parts). */
function deleteState_(eventId) {
    const properties = getScriptProperties_();
    const chunks = chunkCount_(properties.getProperty(STATE_KEY_PREFIX + eventId));
    properties.deleteProperty(STATE_KEY_PREFIX + eventId);
    for (let i = 0; i < chunks; ++i) {
        properties.deleteProperty(chunkKey_(eventId, i));
    }
}


/** Split text into parts of at most maxBytes UTF-8 bytes (never splitting a character). */
function splitUtf8_(text, maxBytes) {
    const parts = [];
    let part = '';
    let size = 0;
    for (const c of text) {
        const charSize = Util.calculateUtf8Size(c);
        if (size + charSize > maxBytes) {
            parts.push(part);
            part = '';
            size = 0;
        }
        part += c;
        size += charSize;
    }
    if (part) {
        parts.push(part);
    }
    return parts;
}


/** Delete poll state a day after the meeting ended (and malformed/outdated entries), warn about the size limit. */
function cleanupOldState_(now) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        for (const eventId of listStateIds_()) {
            const state = loadState_(eventId);
            const end = state ? new Date(state.end).getTime() : NaN;
            if (!state || !(now <= end + STATE_RETENTION_AFTER_END_MS)) {
                deleteState_(eventId);
                Logger.log('Cleaned up poll state of event %s', eventId);
            }
        }

        // parts left behind (e.g. a failed write)
        const properties = getScriptProperties_().getProperties() || {};
        for (const key of Object.keys(properties)) {
            const match = key.startsWith(CHUNK_KEY_PREFIX) && key.substring(CHUNK_KEY_PREFIX.length).match(/^(.*)\.(\d+)$/);
            if (match && !(+match[2] < chunkCount_(properties[STATE_KEY_PREFIX + match[1]]))) {
                getScriptProperties_().deleteProperty(key);
            }
        }
    } finally {
        lock.releaseLock();
    }

    const properties = getScriptProperties_().getProperties() || {};
    const totalBytes = Object.keys(properties)
        .reduce((sum, key) => sum + Util.calculateUtf8Size(key) + Util.calculateUtf8Size(properties[key] || ''), 0);
    if (totalBytes > PROPERTIES_WARN_BYTES) {
        Logger.log('WARNING: script properties use %s bytes, the limit is 500 KB', '' + totalBytes);
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

/** Load and validate the configuration from the script properties. */
function loadConfig_() {
    const properties = getScriptProperties_();
    const get = key => (properties.getProperty(key) || '').trim();
    const list = key => get(key).split(',').map(v => v.trim()).filter(v => !!v);

    const senderEmail = get(SENDER_EMAIL_KEY).toLowerCase();
    if (!isValidEmail_(senderEmail)) {
        throw new Error('Sender account not configured, set ' + SENDER_EMAIL_KEY);
    }

    return {
        allowedDomains: list(ALLOWED_DOMAINS_KEY).map(v => v.toLowerCase()),
        emailWhiteList: list(EMAIL_WHITELIST_KEY).map(v => v.toLowerCase()),
        lookaheadHours: getPositiveIntProperty_(LOOKAHEAD_HOURS_KEY, 24),
        summaryMinutes: getPositiveIntProperty_(SUMMARY_MINUTES_KEY, 15),
        minAttendees: getPositiveIntProperty_(MIN_ATTENDEES_KEY, 2),
        joinMode: get(JOIN_MODE_KEY).toLowerCase() === 'all' ? 'all' : 'recorded',
        senderEmail: senderEmail,
        calendarNotice: getBooleanProperty_(CALENDAR_NOTICE_KEY, true),
        disableRecording: getBooleanProperty_(DISABLE_RECORDING_KEY, true),
        recordingPurpose: get(RECORDING_PURPOSE_KEY),
        policyUrl: get(POLICY_URL_KEY),
        webAppUrl: getWebAppUrl_()
    };
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

/** Read a positive integer property, else the default. */
function getPositiveIntProperty_(key, defaultValue) {
    const raw = (getScriptProperties_().getProperty(key) || '').trim();
    const value = Math.abs(Math.round(+raw));
    if (!value || Number.isNaN(value)) {
        return defaultValue;
    }
    return value;
}

/** Read a boolean property ('true'/'false', 'yes'/'no', '1'/'0', 'on'/'off'), else the default. */
function getBooleanProperty_(key, defaultValue) {
    const raw = (getScriptProperties_().getProperty(key) || '').trim().toLowerCase();
    if (['true', 'yes', '1', 'on'].includes(raw)) {
        return true;
    }
    if (['false', 'no', '0', 'off'].includes(raw)) {
        return false;
    }
    return defaultValue;
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

/** Format a meeting time in the event's time zone (else the script's) and UTC, e.g. 'Tue, 17 Nov 2026 10:00 Europe/Berlin (09:00 UTC)'. */
function formatMeetingTime_(isoTime, timeZone) {
    if (!isoTime) {
        return '(Unknown time)';
    }
    try {
        const date = new Date(isoTime);
        const tz = timeZone || Session.getScriptTimeZone();
        const local = Utilities.formatDate(date, tz, 'EEE, d MMM yyyy HH:mm') + ' ' + tz;
        return tz === 'UTC' || tz === 'Etc/UTC' ? local : local + ' (' + Utilities.formatDate(date, 'UTC', 'HH:mm') + ' UTC)';
    } catch (e) {
        return isoTime;
    }
}
