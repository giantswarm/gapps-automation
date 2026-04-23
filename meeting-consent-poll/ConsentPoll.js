/**
 * Meeting Recording Consent Poll
 *
 * Posts an interactive consent poll card into Google Meet in-meeting chat
 * for GDPR-compliant recording consent collection.
 *
 * Architecture:
 *   - Time-based trigger scans employee calendars for upcoming meetings with Meet links
 *   - Finds the meeting's Chat space by impersonating the organizer
 *   - Installs the Chat App into the meeting Chat space
 *   - Posts a consent poll card visible to all participants
 *   - onCardClick handles consent/decline button clicks, updates card in-place
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

/** White-list to restrict operation to a few tester email accounts.
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

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'checkUpcomingMeetings';

/** Prefix for consent state stored in ScriptProperties (keyed by event ID). */
const STATE_KEY_PREFIX = 'state.';

/** Prefix for sent dedup markers stored in ScriptProperties (keyed by event ID). */
const SENT_KEY_PREFIX = 'sent.';

/** Card action names. */
const ACTION_CONSENT = 'handleConsent';
const ACTION_DECLINE = 'handleDecline';
const ACTION_REFRESH = 'handleRefresh';

/** Maximum age for state entries before cleanup (48 hours). */
const STATE_MAX_AGE_MS = 48 * 60 * 60 * 1000;

/** Chat API base URL. */
const CHAT_API_BASE = 'https://chat.googleapis.com/v1';

/** Maximum number of attendee names to show in the card before truncating. */
const MAX_DISPLAY_NAMES = 20;


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
// Chat App Event Handlers
// ---------------------------------------------------------------------------

/** Called when the Chat App is added to a space. */
function onAddToSpace(event) {
    return {text: 'Meeting Recording Consent Poll is active. Consent cards will be posted automatically for upcoming meetings.'};
}

/** Called when a card button is clicked. */
function onCardClick(event) {
    const action = event.action || {};
    const invokedFunction = action.actionMethodName || '';
    const params = {};
    for (const param of (action.parameters || [])) {
        params[param.key] = param.value;
    }

    const eventId = params.eventId;
    if (!eventId) {
        return {text: 'Error: Missing event ID in action parameters.'};
    }

    const user = event.user || {};
    const userEmail = user.email || '';
    const userName = user.displayName || userEmail;

    if (invokedFunction === ACTION_CONSENT) {
        return handleConsentResponse_(eventId, userEmail, userName, true);
    } else if (invokedFunction === ACTION_DECLINE) {
        return handleConsentResponse_(eventId, userEmail, userName, false);
    } else if (invokedFunction === ACTION_REFRESH) {
        return handleRefreshCard_(eventId);
    }

    return {text: 'Unknown action: ' + invokedFunction};
}


// ---------------------------------------------------------------------------
// Core Trigger: checkUpcomingMeetings
// ---------------------------------------------------------------------------

/** Main entry point — scans employee calendars for upcoming meetings and posts consent polls. */
async function checkUpcomingMeetings() {

    const allowedDomains = (getScriptProperties_().getProperty(ALLOWED_DOMAINS_KEY) || '')
        .split(',')
        .map(d => d.trim())
        .filter(d => !!d);

    const emailWhiteList = getEmailWhiteList_();
    const isEmailAllowed = email => (!emailWhiteList.length || emailWhiteList.includes(email))
        && allowedDomains.some(domain => email.endsWith('@' + domain));

    const lookaheadMinutes = getLookaheadMinutes_();
    const minAttendees = getMinAttendees_();

    Logger.log('Configured to handle accounts %s on domains %s, lookahead %s min, min attendees %s',
        emailWhiteList.length ? emailWhiteList : '(all)', allowedDomains, lookaheadMinutes, minAttendees);

    const personioCreds = getPersonioCreds_();
    const personio = PersonioClientV1.withApiCredentials(personioCreds.clientId, personioCreds.clientSecret);
    const employees = await personio.getPersonioJson('/company/employees');
    const activeEmployees = employees.filter(employee =>
        employee.attributes.status.value !== 'inactive' && isEmailAllowed(employee.attributes.email.value)
    );

    Logger.log('Processing %s active employees', '' + activeEmployees.length);

    const now = new Date();
    const timeMin = now.toISOString();
    const timeMax = Util.addDateMillies(new Date(now), lookaheadMinutes * 60 * 1000).toISOString();

    const employeeByEmail = {};
    for (const emp of activeEmployees) {
        employeeByEmail[emp.attributes.email.value] = emp;
    }

    let firstError = null;
    let processedCount = 0;

    for (const employee of activeEmployees) {
        const email = employee.attributes.email.value;

        try {
            const calendar = await CalendarClient.withImpersonatingService(getServiceAccountCredentials_(), email);
            const events = await calendar.list('primary', {
                singleEvents: true,
                showDeleted: false,
                timeMin: timeMin,
                timeMax: timeMax
            });

            for (const event of events) {
                // Only process events where this employee is the organizer (dedup across calendars)
                if (event.organizer?.email !== email) {
                    continue;
                }

                if (!isQualifyingEvent_(event, minAttendees)) {
                    continue;
                }

                // Skip if already sent
                const sentKey = SENT_KEY_PREFIX + event.id;
                if (getScriptProperties_().getProperty(sentKey)) {
                    continue;
                }

                try {
                    await sendConsentPoll_(event, employeeByEmail);
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
    const hasVideoEntrypoint = (event.conferenceData?.entryPoints || [])
        .some(ep => ep.entryPointType === 'video' && (ep.uri || '').includes('meet.google.com'));

    return hasVideoEntrypoint;
}


// ---------------------------------------------------------------------------
// Meeting Chat Space Discovery
// ---------------------------------------------------------------------------

/** Find the Chat space corresponding to a meeting's in-meeting chat.
 *
 * There is no direct API link between Calendar events and Chat spaces. The matching algorithm:
 *   1. Impersonate the organizer, list their GROUP_CHAT spaces
 *   2. Filter by createTime within [eventStart - 10min, eventStart + 30min]
 *   3. For candidates, list members and compare with calendar attendees
 *   4. Return the space with >= 80% member overlap, or null
 */
async function findMeetingChatSpace_(organizerEmail, event) {
    const creds = getServiceAccountCredentials_();
    const spacesService = await UrlFetchJsonClient.createImpersonatingService(
        'ConsentPollSpaces-' + organizerEmail, creds, organizerEmail,
        'https://www.googleapis.com/auth/chat.spaces.readonly'
    );
    const spacesClient = new UrlFetchJsonClient(spacesService);

    const eventStart = new Date(event.start.dateTime || event.start.date);
    const windowStart = new Date(eventStart.getTime() - 10 * 60 * 1000);
    const windowEnd = new Date(eventStart.getTime() + 30 * 60 * 1000);

    // List GROUP_CHAT spaces with pagination
    const candidateSpaces = [];
    let pageToken = undefined;
    do {
        const query = UrlFetchJsonClient.buildQuery({
            filter: 'spaceType = "GROUP_CHAT"',
            pageSize: 100,
            pageToken: pageToken
        });
        const response = await spacesClient.getJson(CHAT_API_BASE + '/spaces' + query);
        const spaces = response.spaces || [];

        for (const space of spaces) {
            if (space.createTime) {
                const createTime = new Date(space.createTime);
                if (createTime >= windowStart && createTime <= windowEnd) {
                    candidateSpaces.push(space);
                }
            }
        }

        pageToken = response.nextPageToken;
    } while (pageToken);

    Logger.log('Found %s candidate chat spaces for event %s (%s)', candidateSpaces.length, event.id, event.summary);

    if (candidateSpaces.length === 0) {
        return null;
    }

    // Build set of event attendee emails
    const eventAttendeeEmails = new Set(
        (event.attendees || []).map(a => (a.email || '').toLowerCase()).filter(e => !!e)
    );

    // Check each candidate for member overlap
    const membersService = await UrlFetchJsonClient.createImpersonatingService(
        'ConsentPollMembers-' + organizerEmail, creds, organizerEmail,
        'https://www.googleapis.com/auth/chat.memberships.readonly'
    );
    const membersClient = new UrlFetchJsonClient(membersService);

    for (const space of candidateSpaces) {
        try {
            const memberEmails = await listSpaceMemberEmails_(membersClient, space.name);

            // Calculate overlap: what fraction of chat members are also event attendees
            if (memberEmails.length === 0) {
                continue;
            }

            let matchCount = 0;
            for (const memberEmail of memberEmails) {
                if (eventAttendeeEmails.has(memberEmail.toLowerCase())) {
                    matchCount++;
                }
            }

            const overlap = matchCount / memberEmails.length;
            Logger.log('Space %s: %s members, %s match event attendees (%.0f%% overlap)',
                space.name, memberEmails.length, matchCount, overlap * 100);

            if (overlap >= 0.8) {
                return space;
            }
        } catch (e) {
            Logger.log('Failed to check members of space %s: %s', space.name, e);
        }
    }

    return null;
}


/** List all member emails of a Chat space. */
async function listSpaceMemberEmails_(membersClient, spaceName) {
    const emails = [];
    let pageToken = undefined;
    do {
        const query = UrlFetchJsonClient.buildQuery({
            pageSize: 100,
            pageToken: pageToken
        });
        const response = await membersClient.getJson(CHAT_API_BASE + '/' + spaceName + '/members' + query);
        const memberships = response.memberships || [];

        for (const membership of memberships) {
            // Skip bot members
            if (membership.member?.type === 'BOT') {
                continue;
            }
            const email = membership.member?.name;
            if (email && email.startsWith('users/')) {
                // member.name is "users/<email>" for human members when listed via user auth
                const memberEmail = membership.member?.email || '';
                if (memberEmail) {
                    emails.push(memberEmail);
                }
            }
        }

        pageToken = response.nextPageToken;
    } while (pageToken);

    return emails;
}


// ---------------------------------------------------------------------------
// Chat App Installation
// ---------------------------------------------------------------------------

/** Install the Chat App into a Chat space by adding it as a BOT member. */
async function installAppInSpace_(organizerEmail, spaceName) {
    const creds = getServiceAccountCredentials_();
    const appService = await UrlFetchJsonClient.createImpersonatingService(
        'ConsentPollAppInstall-' + organizerEmail, creds, organizerEmail,
        'https://www.googleapis.com/auth/chat.memberships.app'
    );
    const appClient = new UrlFetchJsonClient(appService);

    try {
        await appClient.postJson(CHAT_API_BASE + '/' + spaceName + '/members', {
            member: {
                name: 'users/app',
                type: 'BOT'
            }
        });
        Logger.log('Installed Chat App into space %s', spaceName);
    } catch (e) {
        // Ignore "already member" errors
        if (e.message && e.message.includes('ALREADY_EXISTS')) {
            Logger.log('Chat App already installed in space %s', spaceName);
        } else {
            throw e;
        }
    }
}


// ---------------------------------------------------------------------------
// Consent Poll Sending
// ---------------------------------------------------------------------------

/** Send a consent poll card to the meeting's chat space. */
async function sendConsentPoll_(event, employeeByEmail) {
    const organizerEmail = event.organizer?.email;
    if (!organizerEmail) {
        Logger.log('No organizer email for event %s, skipping', event.id);
        return;
    }

    // Initialize state
    const attendees = (event.attendees || [])
        .map(a => a.email)
        .filter(e => !!e);

    const state = {
        meetingTitle: event.summary || '(No title)',
        meetingTime: event.start.dateTime || event.start.date,
        eventId: event.id,
        organizerEmail: organizerEmail,
        attendees: attendees,
        responses: {},
        pollSentAt: Date.now()
    };

    // Find the meeting's Chat space
    const space = await findMeetingChatSpace_(organizerEmail, event);
    if (!space) {
        Logger.log('No matching chat space found for event %s (%s), will retry on next trigger run',
            event.id, event.summary);
        return;
    }

    state.spaceName = space.name;
    Logger.log('Found chat space %s for event %s (%s)', space.name, event.id, event.summary);

    // Install the Chat App into the space
    await installAppInSpace_(organizerEmail, space.name);

    // Post the consent card as the Chat App (bot auth)
    const creds = getServiceAccountCredentials_();
    const botService = await UrlFetchJsonClient.createImpersonatingService(
        'ConsentPollBot', creds, creds.client_email,
        'https://chat.googleapis.com/auth/chat.bot'
    );
    const botClient = new UrlFetchJsonClient(botService);

    const card = buildConsentCard_(state);
    const cardId = 'consent-poll-' + event.id;

    const message = await botClient.postJson(CHAT_API_BASE + '/' + space.name + '/messages', {
        cardsV2: [{
            cardId: cardId,
            card: card
        }]
    });

    state.messageName = message.name;
    saveState_(event.id, state);

    // Mark as sent (dedup)
    getScriptProperties_().setProperty(SENT_KEY_PREFIX + event.id, '' + Date.now());

    Logger.log('Posted consent poll card for event %s (%s) in space %s', event.id, event.summary, space.name);
}


// ---------------------------------------------------------------------------
// Card Building
// ---------------------------------------------------------------------------

/** Build the consent poll card. */
function buildConsentCard_(state) {
    const formattedTime = formatMeetingTime_(state.meetingTime);

    const sections = [];

    // Section 1: Meeting info
    sections.push({
        widgets: [
            {
                decoratedText: {
                    topLabel: 'Meeting',
                    text: state.meetingTitle,
                    startIcon: {knownIcon: 'INVITE'}
                }
            },
            {
                decoratedText: {
                    topLabel: 'Scheduled',
                    text: formattedTime,
                    startIcon: {knownIcon: 'CLOCK'}
                }
            },
            {
                textParagraph: {
                    text: '<b>This meeting may be recorded. Please indicate your consent below.</b>'
                }
            }
        ]
    });

    // Section 2: Consent buttons
    sections.push({
        widgets: [{
            buttonList: {
                buttons: [
                    {
                        text: 'I Consent',
                        color: {red: 0.2, green: 0.66, blue: 0.33, alpha: 1},
                        onClick: {
                            action: {
                                actionMethodName: ACTION_CONSENT,
                                parameters: [{key: 'eventId', value: state.eventId}]
                            }
                        }
                    },
                    {
                        text: 'I Do Not Consent',
                        color: {red: 0.84, green: 0.18, blue: 0.18, alpha: 1},
                        onClick: {
                            action: {
                                actionMethodName: ACTION_DECLINE,
                                parameters: [{key: 'eventId', value: state.eventId}]
                            }
                        }
                    }
                ]
            }
        }]
    });

    // Section 3: Status (only shown after first response)
    const responses = state.responses || {};
    const responseCount = Object.keys(responses).length;
    const totalAttendees = (state.attendees || []).length;

    if (responseCount > 0) {
        const consented = [];
        const declined = [];
        const pending = [];

        const respondedEmails = new Set(Object.keys(responses));

        for (const email of Object.keys(responses)) {
            const resp = responses[email];
            const displayName = resp.name || email;
            if (resp.consent) {
                consented.push(displayName);
            } else {
                declined.push(displayName);
            }
        }

        for (const email of (state.attendees || [])) {
            if (!respondedEmails.has(email)) {
                pending.push(email);
            }
        }

        const progressBar = buildProgressBar_(responseCount, totalAttendees);
        const statusWidgets = [];

        statusWidgets.push({
            textParagraph: {
                text: progressBar + ' ' + responseCount + '/' + totalAttendees + ' responded'
            }
        });

        if (consented.length > 0) {
            statusWidgets.push({
                textParagraph: {
                    text: '\u2705 <b>Consented:</b> ' + truncateNameList_(consented)
                }
            });
        }

        if (declined.length > 0) {
            statusWidgets.push({
                textParagraph: {
                    text: '\u274C <b>Declined:</b> ' + truncateNameList_(declined)
                }
            });
        }

        if (pending.length > 0) {
            statusWidgets.push({
                textParagraph: {
                    text: '\u23F3 <b>Pending:</b> ' + truncateNameList_(pending)
                }
            });
        }

        sections.push({widgets: statusWidgets});
    }

    // Section 4: Refresh button
    sections.push({
        widgets: [{
            buttonList: {
                buttons: [{
                    text: 'Refresh',
                    onClick: {
                        action: {
                            actionMethodName: ACTION_REFRESH,
                            parameters: [{key: 'eventId', value: state.eventId}]
                        }
                    }
                }]
            }
        }]
    });

    return {
        header: {
            title: 'Recording Consent',
            subtitle: state.meetingTitle,
            imageUrl: 'https://fonts.gstatic.com/s/i/googlematerialicons/videocam/v6/gm_grey-24dp/2x/gm_videocam_gm_grey_24dp.png',
            imageType: 'CIRCLE'
        },
        sections: sections
    };
}


// ---------------------------------------------------------------------------
// Response Handling
// ---------------------------------------------------------------------------

/** Handle a consent/decline button click. */
function handleConsentResponse_(eventId, userEmail, userName, consented) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return {text: 'The poll is being updated, please try again in a moment.'};
    }

    try {
        const state = loadState_(eventId);
        if (!state) {
            return {text: 'This consent poll is no longer active.'};
        }

        state.responses[userEmail] = {
            consent: consented,
            name: userName,
            respondedAt: Date.now()
        };

        saveState_(eventId, state);

        const card = buildConsentCard_(state);
        return {
            actionResponse: {type: 'UPDATE_MESSAGE'},
            cardsV2: [{
                cardId: 'consent-poll-' + eventId,
                card: card
            }]
        };
    } finally {
        lock.releaseLock();
    }
}

/** Handle a refresh button click. */
function handleRefreshCard_(eventId) {
    const state = loadState_(eventId);
    if (!state) {
        return {text: 'This consent poll is no longer active.'};
    }

    const card = buildConsentCard_(state);
    return {
        actionResponse: {type: 'UPDATE_MESSAGE'},
        cardsV2: [{
            cardId: 'consent-poll-' + eventId,
            card: card
        }]
    };
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
    const filled = Math.round((current / total) * 10);
    const empty = 10 - filled;
    return '\u2588'.repeat(filled) + '\u2591'.repeat(empty);
}

/** Truncate a list of names for card display, adding "and N more" if needed. */
function truncateNameList_(names) {
    if (names.length <= MAX_DISPLAY_NAMES) {
        return names.join(', ');
    }
    const shown = names.slice(0, MAX_DISPLAY_NAMES);
    const remaining = names.length - MAX_DISPLAY_NAMES;
    return shown.join(', ') + ' and ' + remaining + ' more';
}
