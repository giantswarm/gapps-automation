/** Script to generate various meeting reports and share team meeting artifacts.
 *
 * NOTE: You may install periodic triggers to keep these reports up-to-date in the output sheets.
 */

/** The prefix for properties specific to this script in the project. */
const PROPERTY_PREFIX = 'Meetings.';

/** Personio clientId and clientSecret, separated by '|'. */
const PERSONIO_TOKEN_KEY = PROPERTY_PREFIX + 'personioToken';

/** Service account credentials (in JSON format, as downloaded from Google Management Console). */
const SERVICE_ACCOUNT_CREDENTIALS_KEY = PROPERTY_PREFIX + 'serviceAccountCredentials';

/** Filter for allowed domains (to avoid working and failing on users present on foreign domains). */
const ALLOWED_DOMAINS_KEY = PROPERTY_PREFIX + 'allowedDomains';

/** White-list to restrict synchronization to a few tester email accounts.
 *
 * Must be one email or a comma separated list.
 *
 * Default: null or empty
 */
const EMAIL_WHITELIST_KEY = PROPERTY_PREFIX + 'emailWhiteList';

/** Lookahead days for event/time-off synchronization.
 *
 * Default: 6 * 30 days, should scale up to ~18 months
 */
const LOOKAHEAD_DAYS_KEY = PROPERTY_PREFIX + 'lookaheadDays';

/** Lookback days for event/time-off synchronization.
 *
 * Default: 30 days, should scale up to ~6 months, avoid changing events too far back
 */
const LOOKBACK_DAYS_KEY = PROPERTY_PREFIX + 'lookbackDays';

/** ID of sheet to write reports to. */
const REPORT_SHEET_KEY = PROPERTY_PREFIX + 'reportSheet';

/** ID of shared team meetings calendar to restrict automatic operations (optional). */
const TEAM_MEETINGS_CALENDAR = PROPERTY_PREFIX + 'teamMeetingsCalendar';

/** Slack App bot token for sharing artifacts with the whole organization (posts will be to each channel where the app is invited).
 *  Required scopes: chat:write, channels:read, groups:read, users:read, users:read.email
 */
const SLACK_BOT_TOKEN = PROPERTY_PREFIX + 'slackBotToken';

/** Google Gemini API key to summarize meeting content. */
const GEMINI_API_KEY = PROPERTY_PREFIX + 'geminiApiKey';

/** Share domain (domain to share meeting recordings/summaries with). */
const SHARE_DOMAIN_NAME = PROPERTY_PREFIX + 'shareDomain';

/** Shared Drive folder ID to move meeting artifacts into (optional). */
const ARTIFACTS_FOLDER_ID = PROPERTY_PREFIX + 'artifactsFolderId';

/** URL to the company glossary for correcting terminology in summaries. */
const GLOSSARY_URL = 'https://raw.githubusercontent.com/giantswarm/handbook/main/content/docs/glossary/_index.md';


/** Build meetings statistic per employee and summarized. */
async function listMeetingStatistic() {

    const allowedDomains = (getScriptProperties_().getProperty(ALLOWED_DOMAINS_KEY) || '')
        .split(',')
        .map(d => d.trim());

    const weeks = 4;   // 4 weeks back from now

    const now = new Date();
    const weekBounds = [];
    for (let week = weeks - 1; week >= 0; --week) {
        weekBounds.push({
            start: new Date(new Date().setDate(now.getDate() - ((week + 1) * 7))),
            end: new Date(new Date().setDate(now.getDate() - (week * 7)))
        });
    }

    const getWeekIndex = (startDate, endDate) => {
        for (let i = 0; i < weekBounds.length; ++i) {
            if (+startDate >= +weekBounds[i].start && +startDate <= +weekBounds[i].end) {
                return i;
            }
        }
        for (let i = 0; i < weekBounds.length; ++i) {
            if (+endDate >= +weekBounds[i].start && +endDate <= +weekBounds[i].end) {
                return i;
            }
        }
        return null;
    };

    const employeeStats = {};
    try {
        let haveEmployees = false;
        // visitEvents_() will call our visitor function for each employee and calendar event combination
        await visitEvents_(async (event, employee, employees, calendar, personio) => {

            if (!haveEmployees) {
                for (const e of employees) {
                    const email = e.attributes.email.value;

                    const stats = {email: email, count: 0, duration1on1: 0, durationSig: 0, weeks: []};
                    for (let i = 0; i < weeks; ++i) {
                        stats.weeks.push(0.0);
                    }

                    employeeStats[email] = stats;
                }
                haveEmployees = true;
            }

            const email = employee.attributes.email.value;
            const stats = employeeStats[email];

            if (!event.extendedProperties?.private?.timeOffId
                && event.eventType !== 'outOfOffice'
                && getGiantSwarmAttendees_(event, allowedDomains).length === (event?.attendees?.length || 0)) {
                const start = new Date(event.start.dateTime);
                const end = new Date(event.end.dateTime);
                const durationHours = (end - start) / (60.0 * 60 * 1000);
                if (!isNaN(durationHours)) {
                    const week = getWeekIndex(start, end);
                    if (week != null) {

                        let validMeeting = (event?.attendees?.length) > 1;
                        if (isSigOrChapterOrSyncEvent_(event)) {
                            validMeeting = true;
                            stats.durationSig += durationHours;
                        } else if (isOneOnOne_(event, employees, email)) {
                            validMeeting = true;
                            stats.duration1on1 += durationHours;
                        }

                        if (validMeeting) {
                            stats.count += 1;
                            stats.weeks[week] += durationHours;
                        }
                    } else {
                        console.log(`Event for ${email} doesn't fit any week: ${JSON.stringify(event)}`);
                    }
                }
            }

            return true;
        }, {
            singleEvents: true, // return recurring events rolled out into individual event instances
            timeMin: new Date(new Date().setDate(now.getDate() - (weeks * 7))).toISOString(),
            timeMax: now.toISOString()
        });
    } catch (e) {
        Logger.log("First error while visiting calendar events: " + e);
    }

    const statistics = Object.values(employeeStats);
    statistics.sort((a, b) => Util.median(b.weeks) - Util.median(a.weeks));

    const rows = statistics.map(stats => {
        const row = [stats.email, stats.count, +stats.duration1on1.toFixed(2), +stats.durationSig.toFixed(2)];

        for (const weekHours of stats.weeks) {
            row.push(+weekHours.toFixed(2));
        }

        return row;
    });

    const header = ["Email", "Meeting Count", "Total 1on1 (h)", "SIG/Chapter (h)"];
    for (let i = 0; i < weekBounds.length; ++i) {
        header.push('' + weekBounds[i].start.toDateString() + ' +7d (h)');
    }
    rows.unshift(header);

    const spreadsheet = SpreadsheetApp.openById(getReportSheetId_());
    const sheet = SheetUtil.ensureSheet(spreadsheet, "Meeting_Hours");
    sheet.getRange(1, 1, sheet.getMaxRows(), header.length).clearContent();
    sheet.getRange(1, 1, rows.length, header.length).setValues(rows);
}


/** Extract recurring 1on1s from all personal calendars. */
async function listMeetings() {

    const allowedDomains = (getScriptProperties_().getProperty(ALLOWED_DOMAINS_KEY) || '')
        .split(',')
        .map(d => d.trim());

    const recurring1on1Ids = {};
    const rows = [];
    try {
        // visitEvents_() will call our visitor function for each employee and calendar event combination
        await visitEvents_(async (event, employee, employees, calendar, personio) => {
            const email = employee.attributes.email.value;

            // recurring event, organized by this employee with 2 attendees means it's a 1on1
            if (isOneOnOne_(event, employees, email)) {
                const otherAttendeeEmail = event.attendees.find(attendee => attendee.email !== email)?.email;

                if (!recurring1on1Ids[event.iCalUID]) {
                    recurring1on1Ids[event.iCalUID] = event;
                    rows.push([email, otherAttendeeEmail, event.summary, event.start.dateTime, event.recurrence.join('\n')]);
                }
            }

            return true;
        });
    } catch (e) {
        Logger.log("First error while visiting calendar events: " + e);
    }

    Logger.log(`Found ${rows.length} recurring 1on1s`);

    const header = ["Organizer/Owner", "Other Attendee", "Summary", "First Start DateTime", "Recurring Rules"];
    rows.unshift(header);

    const spreadsheet = SpreadsheetApp.openById(getReportSheetId_());
    const sheet = SheetUtil.ensureSheet(spreadsheet, "1on1recurring");
    sheet.getRange(1, 1, sheet.getMaxRows(), header.length).clearContent();
    sheet.getRange(1, 1, rows.length, header.length).setValues(rows);
}


/** Debug helper for shareTeamMeetingArtifacts that collects event filtering information. */
async function debugShareTeamMeetingArtifacts() {
    const shareDomainName = getScriptProperties_().getProperty(SHARE_DOMAIN_NAME) || '';

    const fetchTimeMin = Util.addDateMillies(new Date(), -14 * 24 * 60 * 60 * 1000);
    fetchTimeMin.setUTCHours(24, 0, 0, 0);

    const fetchTimeMax = Util.addDateMillies(new Date(), 24 * 60 * 60 * 1000);
    fetchTimeMax.setUTCHours(24, 0, 0, 0);
    const listEventParams = {
        timeMin: fetchTimeMin.toISOString(),
        timeMax: fetchTimeMax.toISOString(),
    };

    const teamMeetingsCalendarId = getScriptProperties_().getProperty(TEAM_MEETINGS_CALENDAR) || '';

    // Map of event.summary to {event, reason}
    const eventMap = {};

    // Track events by unique ID to see if creator ever matches a calendar owner
    const eventCreatorTracking = {}; // iCalUID -> { summary, creator, seenAsOwner: boolean, start }

    try {
        await visitEvents_(async (event, employee, employees, calendar, personio) => {
            const email = employee.attributes.email.value;
            const summary = event.summary || 'Untitled Event';
            const eventId = event.iCalUID || event.id;

            if (!shareDomainName) {
                return true;
            }

            if (teamMeetingsCalendarId && event?.organizer?.email !== teamMeetingsCalendarId) {
                return true;
            }

            const eventEnd = new Date(event.end?.dateTime || event.end?.date);
            if (isNaN(eventEnd) || eventEnd > new Date()) {
                return true;
            }

            const eventStart = new Date(event.start?.dateTime || event.start?.date);
            if (isNaN(eventStart) || eventStart < fetchTimeMin || eventStart > fetchTimeMax) {
                return true;
            }

            // Track this event across all visits
            if (!eventCreatorTracking[eventId]) {
                eventCreatorTracking[eventId] = {
                    summary: summary,
                    creator: event?.creator?.email,
                    organizer: event?.organizer?.email,
                    start: event.start?.dateTime || event.start?.date,
                    seenAsOwner: false
                };
            }

            if (event?.creator?.email !== email) {
                return true;
            }

            // Mark that we've seen this event where the creator is the calendar owner
            eventCreatorTracking[eventId].seenAsOwner = true;

            // Check each filter condition in order
            if (!isSigOrChapterOrSyncEvent_(event)) {
                // Provide detailed breakdown matching the updated isSigOrChapterOrSyncEvent_ logic
                const summaryMatchSigChapterWg = /(^|\s)(SIG|chapter|WG|Jour Fixe|Weekly)(\s|$)/i.test(event.summary);
                const summaryMatchSync = /(^|\s)(Sync)(\s|$)/i.test(event.summary);
                const hasGroup = (event.attendees || []).find(a =>
                    a.email.startsWith('sig-') ||
                    a.email.startsWith('wg-') ||
                    a.email.startsWith('chapter-') ||
                    a.email.startsWith('all@') ||
                    a.email.startsWith('giantswarm.io@')
                );
                const hasMany = (event.attendees || []).length > 2;
                const attendeeEmails = (event.attendees || []).map(a => a.email).join(', ');

                const details = [];
                details.push(`summary: "${event.summary}"`);
                details.push(`summaryMatchSigChapterWg: ${summaryMatchSigChapterWg}`);
                details.push(`summaryMatchSync: ${summaryMatchSync}`);
                details.push(`hasGroup: ${!!hasGroup}${hasGroup ? ` (${hasGroup.email})` : ''}`);
                details.push(`hasMany: ${hasMany} (${(event.attendees || []).length} attendees)`);
                details.push(`attendees: ${attendeeEmails || 'none'}`);

                updateEventMap_(eventMap, summary, event, `Excluded: Not a SIG/Chapter/WG event - ${details.join('; ')}`);
                return true;
            }

            const publishedAt = +event.extendedProperties?.private?.attachmentsPublishedAt;
            const eventStartPlusOneHour = new Date(event.start.dateTime).getTime() + 60 * 60 * 1000;
            if (publishedAt && publishedAt >= eventStartPlusOneHour) {
                updateEventMap_(eventMap, summary, event, `Excluded: Already published at ${new Date(publishedAt).toISOString()}`);
                return true;
            }

            const geminiNotes = event.attachments?.find(a => a.mimeType === 'application/vnd.google-apps.document'
                && a.title?.includes('Notes by Gemini')) || event.attachments?.find(a => a.mimeType === 'application/vnd.google-apps.document'
                && a.title?.includes('Notes'));
            const recording = event.attachments?.find(a => a.mimeType === 'video/mp4'
                && a.title?.includes('Recording')
                && a.title?.includes(event.summary));

            if (!geminiNotes && !recording) {
                updateEventMap_(eventMap, summary, event, 'Excluded: No Gemini notes or recording found');
                return true;
            } else if (!geminiNotes) {
                updateEventMap_(eventMap, summary, event, `Excluded: No Gemini notes found (recording: ${recording?.title || 'none'})`);
                return true;
            } else if (!recording) {
                updateEventMap_(eventMap, summary, event, `Excluded: No recording found (notes: ${geminiNotes?.title || 'none'})`);
                return true;
            }

            // All checks passed - would be shared
            updateEventMap_(eventMap, summary, event, `Included: Would share artifacts (notes: ${geminiNotes.title}, recording: ${recording.title})`);

            return true;
        }, listEventParams);
    } catch (e) {
        Logger.log("Error while analyzing events: " + e);
    }

    // Find events that passed SIG/Chapter filter but creator never matched any calendar owner
    const neverOwnedByCreator = Object.entries(eventCreatorTracking)
        .filter(([id, info]) => !info.seenAsOwner)
        .map(([id, info]) => ({
            summary: info.summary,
            creator: info.creator,
            organizer: info.organizer,
            start: info.start
        }));

    // Log each event individually (one line per event)
    // Filter out events outside the query time range
    Logger.log("=== Events by Name (most recent occurrence) ===");
    for (const [summary, data] of Object.entries(eventMap)) {
        const eventStart = new Date(data.event.start);
        if (eventStart >= fetchTimeMin && eventStart <= fetchTimeMax) {
            Logger.log(JSON.stringify({ summary, event: data.event, reason: data.reason }));
        }
    }

    Logger.log("=== Events Without Creator ===");
    for (const event of neverOwnedByCreator) {
        const eventStart = new Date(event.start);
        if (eventStart >= fetchTimeMin && eventStart <= fetchTimeMax) {
            Logger.log(JSON.stringify(event));
        }
    }

    Logger.log("=== Stats ===");
    Logger.log(JSON.stringify({
        totalUniqueEvents: Object.keys(eventCreatorTracking).length,
        eventsWithoutCreator: neverOwnedByCreator.length
    }));
}


/** Helper to update event map with most recent event for each summary. */
function updateEventMap_(eventMap, summary, event, reason) {
    const eventStart = new Date(event.start?.dateTime || event.start?.date);

    if (!eventMap[summary]) {
        eventMap[summary] = {
            event: {
                start: event.start?.dateTime || event.start?.date,
                end: event.end?.dateTime || event.end?.date,
                creator: event.creator?.email,
                organizer: event.organizer?.email,
                attachments: event.attachments?.length || 0
            },
            reason: reason
        };
    } else {
        // Update if this event is more recent
        const existingStart = new Date(eventMap[summary].event.start);
        if (eventStart > existingStart) {
            eventMap[summary] = {
                event: {
                    start: event.start?.dateTime || event.start?.date,
                    end: event.end?.dateTime || event.end?.date,
                    creator: event.creator?.email,
                    organizer: event.organizer?.email,
                    attachments: event.attachments?.length || 0
                },
                reason: reason
            };
        }
    }
}


/** Share team meeting artifacts with the organization. */
async function shareTeamMeetingArtifacts() {

    const shareDomainName = getScriptProperties_().getProperty(SHARE_DOMAIN_NAME) || '';
    const artifactsFolderId = getScriptProperties_().getProperty(ARTIFACTS_FOLDER_ID) || '';
    if (!shareDomainName) {
        Logger.log('No share domain configured, skipping sharing team meeting artifacts');
        return;
    }

    // limit to 2 weeks in the past (performance, guard against reposting old meets on filter change)
    const fetchTimeMin = Util.addDateMillies(new Date(), -14 * 24 * 60 * 60 * 1000);
    fetchTimeMin.setUTCHours(24, 0, 0, 0); // round up to end of day

    // limit Google Calendar events fetch up until today (performance)
    const fetchTimeMax = Util.addDateMillies(new Date(), 24 * 60 * 60 * 1000);
    fetchTimeMax.setUTCHours(24, 0, 0, 0); // round up to end of day
    const listEventParams = {
        timeMin: fetchTimeMin.toISOString(),
        timeMax: fetchTimeMax.toISOString(),
        conferenceDataVersion: 1,
    };

    // Fetch company glossary once for correcting terminology in summaries
    let glossary = '';
    try {
        const glossaryResponse = UrlFetchApp.fetch(GLOSSARY_URL, { muteHttpExceptions: true });
        if (glossaryResponse.getResponseCode() === 200) {
            glossary = glossaryResponse.getContentText();
        }
    } catch (e) {
        Logger.log(`Failed to fetch glossary: ${e}`);
    }

    let hits = 0;
    let hits_published = 0;
    const teamMeetingsCalendarId = getScriptProperties_().getProperty(TEAM_MEETINGS_CALENDAR) || '';
    try {
        // visitEvents_() will call the visitor function for each employee and calendar event combination
        await visitEvents_(async (event, employee, employees, calendar, personio) => {
            const email = employee.attributes.email.value;

            // Only process SIG/Chapter/WG events
            if (!isSigOrChapterOrSyncEvent_(event)) {
                return true;
            }

            // If team meetings calendar is configured, only process events from that calendar
            if (teamMeetingsCalendarId && event?.organizer?.email !== teamMeetingsCalendarId) {
                return true;
            }

            // Only the creator initially owns the meeting attributes and can share them
            if (event?.creator?.email !== email) {
                return true;
            }

            // Skip events that haven't ended yet - future events can't have legitimate
            // notes/recordings, and any attachments found would be from a previous occurrence
            const eventEnd = new Date(event.end?.dateTime || event.end?.date);
            if (isNaN(eventEnd) || eventEnd > new Date()) {
                return true;
            }

            // Skip events whose start timestamp is outside the intended query range
            // (handles recurring events that return with original start date from years ago)
            const eventStart = new Date(event.start?.dateTime || event.start?.date);
            if (isNaN(eventStart) || eventStart < fetchTimeMin || eventStart > fetchTimeMax) {
                return true;
            }

            const publishedAt = +event.extendedProperties?.private?.attachmentsPublishedAt;
            const eventStartPlusOneHour = new Date(event.start.dateTime).getTime() + 60 * 60 * 1000;
            if (publishedAt && publishedAt >= eventStartPlusOneHour) {
                Logger.log("event " + event.summary + " at " + event.start.dateTime + " already published at " + publishedAt);
                return true;
            }

            const geminiNotes = event.attachments?.find(a => a.mimeType === 'application/vnd.google-apps.document'
                && a.title?.includes('Notes by Gemini')) || event.attachments?.find(a => a.mimeType === 'application/vnd.google-apps.document'
                && a.title?.includes('Notes'));
            const recording = event.attachments?.find(a => a.mimeType === 'video/mp4'
                && a.title?.includes('Recording')
                && a.title?.includes(event.summary));

            console.log('meet: ' + event.summary + ' ' + event.start.dateTime + ' notes=' + geminiNotes + ' rec=' + recording);

            if (geminiNotes && recording) {

                hits++;

                // Share with organization
                try {
                    const drive = await DriveClientV1.withImpersonatingService(getServiceAccountCredentials_(), email);

                    // Move artifacts to shared drive folder (inherits shared drive permissions)
                    // or share with domain if no folder is configured
                    if (artifactsFolderId) {
                        try {
                            Logger.log(`Moving notes ${geminiNotes.fileId} to folder ${artifactsFolderId}`);
                            await drive.moveToFolder(geminiNotes.fileId, artifactsFolderId);
                        } catch (e) {
                            Logger.log(`Failed to move notes ${geminiNotes.fileId}: ${e}`);
                            return true;
                        }
                        try {
                            Logger.log(`Moving recording ${recording.fileId} to folder ${artifactsFolderId}`);
                            await drive.moveToFolder(recording.fileId, artifactsFolderId);
                        } catch (e) {
                            Logger.log(`Failed to move recording ${recording.fileId}: ${e}`);
                            return true;
                        }
                    } else {
                        Logger.log(`Sharing notes ${geminiNotes.fileId} for event ${event.summary}`);
                        try {
                            await drive.shareWith(geminiNotes.fileId, 'domain', shareDomainName,'reader', true);
                        } catch (e) {
                            Logger.log(`Failed to share notes ${geminiNotes.fileId}: ${e}`);
                            return true;
                        }
                        Logger.log(`Sharing recording ${recording.fileId} for event ${event.summary}`);
                        try {
                            await drive.shareWith(recording.fileId, 'domain', shareDomainName,'reader', true);
                        } catch (e) {
                            Logger.log(`Failed to share recording ${recording.fileId}: ${e}`);
                            return true;
                        }
                    }

                    // Summarize and post to Slack
                    await summarizeAndPostToSlack_(event, geminiNotes, recording, drive, employees, glossary);

                    // Mark as published
                    hits_published++;
                    setEventPrivateProperty_(event, 'attachmentsPublishedAt', Date.now());
                    await calendar.update('primary', event.id, event);

                    Logger.log(`Successfully shared artifacts for ${event.summary} at ${event.start.dateTime}`);
                } catch (e) {
                    Logger.log(`Failed to share artifacts for ${event.summary}: ${e}`);
                }
            }

            return true;
        }, listEventParams);
    } catch (e) {
        Logger.log("First error while visiting calendar events: " + e);
    }

    Logger.log(`Found ${hits} recorded meets, published ${hits_published} new artifacts`);
}


function getGiantSwarmAttendees_(event, allowedDomains) {
    return (event.attendees || []).filter(attendee => allowedDomains.filter(domain => (attendee.email || '').includes(domain)).length);
}


function isTeamEvent_(event) {
    return (event.attendees || []).find(attendee => attendee.email.startsWith('team-'));
}


function isSigOrChapterOrSyncEvent_(event) {
    const emailPrefixes = ['sig-', 'wg-', 'chapter-', 'all@', 'giantswarm.io@'];
    const hasGroup = event?.attendees?.length && emailPrefixes.some(prefix => event.attendees.some(a => a.email.startsWith(prefix)));
    const hasMany = (event.attendees || []).length > 2;
    return /(^|\s)(SIG|chapter|WG|Jour Fixe|Weekly)(\s|$)/i.test(event.summary)
        || (/(^|\s)(Sync)(\s|$)/i.test(event.summary) && hasMany)
        || hasGroup;
}


function isOneOnOne_(event, employees, ownerEmail) {
    return Array.isArray(event.attendees) && event.attendees.length === 2
        && ((Array.isArray(event.recurrence) && event.recurrence.length) || event.recurringEventId)
        && !isSigOrChapterOrSyncEvent_(event)
        && !isTeamEvent_(event)
        && event.attendees.find(attendee => attendee.email === ownerEmail)
        && event.attendees.filter(attendee => employees.find(employee => employee.attributes.email.value === attendee.email)).length === event.attendees.length;
}


/** Summarize meeting notes and post to Slack.
 *
 * @param event The calendar event
 * @param geminiNotes The Gemini notes attachment object
 * @param recording The recording attachment object
 * @param drive The DriveClientV1 instance with access to the documents
 * @param employees The list of active employees from Personio
 * @param glossary Company glossary markdown for correcting terminology (optional)
 */
async function summarizeAndPostToSlack_(event, geminiNotes, recording, drive, employees, glossary) {
    try {
        const geminiApiKey = getGeminiApiKey_();
        const slackBotToken = getSlackBotToken_();

        if (!geminiApiKey) {
            Logger.log('No Gemini API key configured, skipping summarization');
            return;
        }

        if (!slackBotToken) {
            Logger.log('No Slack bot token configured, skipping Slack posting');
            return;
        }

        // Initialize clients
        const gemini = new GeminiRestClient(geminiApiKey);
        const slack = new SlackWebClient(slackBotToken);

        // Extract file ID from the notes URL
        const notesFileId = geminiNotes.fileId || DriveClientV1.extractFileId(geminiNotes.fileUrl);
        if (!notesFileId) {
            throw new Error('Could not extract file ID from Gemini notes');
        }

        Logger.log(`Summarizing notes ${notesFileId} for event ${event.summary}`);

        // Format the date
        const eventDate = new Date(event.start.dateTime || event.start.date);
        const formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

        // Limit mentions to actual Meet attendees to avoid false name matches against unrelated employees
        let meetAttendeeEmails = null;
        try {
            meetAttendeeEmails = await getMeetAttendeeEmails_(event);
            if (meetAttendeeEmails) {
                Logger.log(`Found ${meetAttendeeEmails.size} Meet attendees for ${event.summary}`);
            }
        } catch (e) {
            Logger.log(`Failed to get Meet attendees for ${event.summary}, falling back to all employees: ${e}`);
        }

        const attendeeLabel = meetAttendeeEmails ? 'meeting attendees' : 'all employees';
        const nameToSlackMentionMapping = async () => {
            const relevantEmployees = meetAttendeeEmails
                ? employees.filter(emp => meetAttendeeEmails.has(emp.attributes.email.value))
                : employees;
            const employeeEmails = relevantEmployees.map(emp => emp.attributes.email.value).filter(Boolean);
            if (employeeEmails.length === 0) {
                return '';
            }

            const mappingLines = [];
            for (const email of employeeEmails) {
                try {
                    const slackUser = await getSlackUserByEmailCached_(slack, email);
                    if (slackUser) {
                        // Get the employee's full name from Personio
                        const employee = employees.find(emp => emp.attributes.email.value === email);
                        const firstName = employee?.attributes?.first_name?.value || '';
                        const lastName = employee?.attributes?.last_name?.value || '';
                        const fullName = `${firstName} ${lastName}`.trim();

                        const slackMention = `<@${slackUser.id}>`;
                        const displayName = fullName || slackUser.profile?.display_name || slackUser.real_name || slackUser.name;
                        mappingLines.push(`${displayName} -> ${slackMention}`);
                    }
                } catch (e) {
                    Logger.log(`Failed to lookup Slack user for email ${email}: ${e}`);
                }
            }

            return mappingLines.join('\n');
        };

        // Create the summarization prompt
        const prompt = `You are summarizing meeting notes for a Slack post. Generate ONLY the content for the takeaways sections as described below. Do NOT include the header, meeting name, date, or document links - those will be added separately.

Generate the following sections:

🧠 *Top Takeaways*
Use plain bullet points (•) for each takeaway. Do NOT use emoji prefixes on individual items.
Keep each takeaway to 1-2 sentences maximum. Focus on outcomes, not discussions.

🙋 *Key Contributors*

Note: Only include this section if any of the following apply:
- Someone is assigned an action or deliverable
- A decision was clearly advocated for or blocked by a participant
- Someone flagged a risk, dependency, or introduced a new direction

Skip this section entirely if there's nothing impactful to highlight - this isn't about logging everyone's contributions.

If you do include this section, use the correct mention using the Slack user ID (see lookup table below) and format each entry as:
• <@{{$SLACK_MENTION}}> – {{Brief description of their assigned task, key decision, or flagged blocker}}

Keep it short and outcome-focused. Only mention people with specific, actionable contributions.

Guidelines:
- Be concise and direct
- Focus on outcomes and decisions, not process
- Use clear, simple language
- Only output the sections above, nothing else

Name to Slack user ID lookup table (format "$DISPLAY_NAME -> $SLACK_MENTION") for ${attendeeLabel}:
<slack_mention_lookup_begin>
${await nameToSlackMentionMapping()}
</slack_mention_lookup_end>

${glossary ? `
Company terminology glossary — use this to correct any misspelled or misunderstood terms (e.g. from speech-to-text errors) in your summary:
<glossary_end>
${glossary}
</glossary_end>` : ''}`;

        // Summarize the document
        const summaryContent = await gemini.summarizeGoogleDoc(notesFileId, prompt, drive);
        Logger.log(`Generated summary: ${summaryContent}`);

        // Broadcast to all channels where the bot is a member
        Logger.log(`Broadcasting summary to Slack channels`);
        const channels = await slack.getUserChannels();
        for (const channel of channels) {
            await postMeetingSummary_(slack, channel.id, event.summary, formattedDate, summaryContent, geminiNotes.fileUrl, recording.fileUrl);
        }

        Logger.log(`Successfully posted summary for ${event.summary} to Slack`);
    } catch (e) {
        Logger.log(`Failed to summarize and post to Slack: ${e}`);
        throw e;
    }
}


/** Post a formatted meeting summary to a Slack channel.
 *
 * @param slackClient The SlackWebClient instance
 * @param channelId The channel to post to
 * @param meetingName The name of the meeting
 * @param date The formatted date string
 * @param summaryContent The AI-generated summary content (takeaways and contributors)
 * @param docLink Link to the Google Doc
 * @param recordingLink Link to the recording
 */
async function postMeetingSummary_(slackClient, channelId, meetingName, date, summaryContent, docLink, recordingLink) {
    const blocks = [
        {
            type: 'section',
            text: {
                type: 'mrkdwn',
                text: `*${meetingName}*  ·  ${date}`
            }
        },
        {
            type: 'divider'
        },
        // Split summary into multiple section blocks to stay within Slack's
        // 3000-character limit per section text field.
        ...splitForSlackBlocks_(summaryContent),
        {
            type: 'divider'
        },
        {
            type: 'context',
            elements: [
                {
                    type: 'mrkdwn',
                    text: `<${docLink}|Full Summary + Transcript>  ·  <${recordingLink}|Recording>`
                }
            ]
        }
    ];

    // Use plain text as fallback for notifications
    const fallbackText = `Meeting Summary: ${meetingName} (${date})`;

    await slackClient.postMessage(channelId, fallbackText, blocks);
}

/** Max characters allowed in a single Slack section block text field. */
const SLACK_SECTION_TEXT_LIMIT = 3000;

/** Split text into Slack section blocks, each within the character limit.
 *
 *  Splits on paragraph boundaries (\n\n), merges small paragraphs into
 *  chunks that fit, and as a last resort splits on single newlines.
 */
function splitForSlackBlocks_(text) {
    if (text.length <= SLACK_SECTION_TEXT_LIMIT) {
        return [{ type: 'section', text: { type: 'mrkdwn', text: text } }];
    }

    // Split on paragraph boundaries and merge small paragraphs into chunks that fit
    const paragraphs = text.split(/\n\n+/).filter(function(s) { return s.length > 0; });
    const chunks = [];
    let buffer = '';
    for (const para of paragraphs) {
        const separator = buffer ? '\n\n' : '';
        if (buffer.length + separator.length + para.length <= SLACK_SECTION_TEXT_LIMIT) {
            buffer += separator + para;
        } else {
            if (buffer) chunks.push(buffer);
            buffer = para;
        }
    }
    if (buffer) chunks.push(buffer);

    // Further split any chunk that still exceeds the limit (single huge paragraph)
    const finalChunks = [];
    for (const chunk of chunks) {
        if (chunk.length <= SLACK_SECTION_TEXT_LIMIT) {
            finalChunks.push(chunk);
        } else {
            finalChunks.push(...splitOnBoundary_(chunk, SLACK_SECTION_TEXT_LIMIT));
        }
    }

    return finalChunks.map(function(c) {
        return { type: 'section', text: { type: 'mrkdwn', text: c.trim() } };
    }).filter(function(b) { return b.text.text.length > 0; });
}

/** Split a single oversized chunk on paragraph or line boundaries. */
function splitOnBoundary_(text, limit) {
    const results = [];
    let remaining = text;

    while (remaining.length > limit) {
        // Try to break at a double-newline (paragraph boundary)
        let splitIdx = remaining.lastIndexOf('\n\n', limit);
        // Fall back to single newline
        if (splitIdx <= 0) splitIdx = remaining.lastIndexOf('\n', limit);
        // Last resort: hard break at the limit, avoiding surrogate pair splits
        if (splitIdx <= 0) {
            splitIdx = limit;
            const code = remaining.charCodeAt(splitIdx - 1);
            if (code >= 0xD800 && code <= 0xDBFF) splitIdx--;
        }

        results.push(remaining.slice(0, splitIdx));
        remaining = remaining.slice(splitIdx).replace(/^\n+/, '');
    }

    if (remaining) results.push(remaining);
    return results;
}


/** Set a private property on an event. */
function setEventPrivateProperty_(event, key, value) {
    const props = event.extendedProperties ? event.extendedProperties : event.extendedProperties = {};
    const privateProps = props.private ? props.private : props.private = {};
    privateProps[key] = value;
}


/** Utility function to visit and handle all personal employee Gcal events.
 *
 * @param visitor The visitor function which receives (event, employee, employees, calendar, personio) as arguments and may return false to stop iteration.
 * @param listParams Additional parameter overrides for calendar.list().
 */
async function visitEvents_(visitor, listParams) {

    const allowedDomains = (getScriptProperties_().getProperty(ALLOWED_DOMAINS_KEY) || '')
        .split(',')
        .map(d => d.trim());

    const emailWhiteList = getEmailWhiteList_();
    const isEmailAllowed = email => (!emailWhiteList.length || emailWhiteList.includes(email))
        && allowedDomains.includes(email.substring(email.lastIndexOf('@') + 1));

    Logger.log('Configured to handle accounts %s on domains %s', emailWhiteList.length ? emailWhiteList : '', allowedDomains);

    // all timing related activities are relative to this EPOCH
    const epoch = new Date();

    // how far back to sync events/time-offs
    const lookbackMillies = -Math.round(getLookbackDays_() * 24 * 60 * 60 * 1000);
    // how far into the future to sync events/time-offs
    const lookaheadMillies = Math.round(getLookaheadDays_() * 24 * 60 * 60 * 1000);

    const fetchTimeMin = Util.addDateMillies(new Date(epoch), lookbackMillies);
    fetchTimeMin.setUTCHours(0, 0, 0, 0); // round down to start of day
    const fetchTimeMax = Util.addDateMillies(new Date(epoch), lookaheadMillies);
    fetchTimeMax.setUTCHours(24, 0, 0, 0); // round up to end of day

    const personioCreds = getPersonioCreds_();
    const personio = PersonioClientV1.withApiCredentials(personioCreds.clientId, personioCreds.clientSecret);

    // load and prepare list of employees to process
    const employees = await personio.getPersonioJson('/company/employees').filter(employee =>
        employee.attributes.status.value !== 'inactive' && isEmailAllowed(employee.attributes.email.value)
    );

    Logger.log('Visiting events between %s and %s for %s accounts', listParams?.timeMin || fetchTimeMin.toISOString(), listParams?.timeMax || fetchTimeMax.toISOString(), '' + employees.length);

    let firstError = null;
    let processedCount = 0;
    let done = false;
    for (const employee of employees) {

        const email = employee.attributes.email.value;

        // we keep operating if handling calendar of a single user fails
        try {
            const calendar = await CalendarClient.withImpersonatingService(getServiceAccountCredentials_(), email);
            const allEvents = await calendar.list('primary', {
                singleEvents: false, // return original recurring events, not the individual, rolled out instances
                showDeleted: false, // no cancelled/deleted events
                timeMin: fetchTimeMin.toISOString(),
                timeMax: fetchTimeMax.toISOString(),
                ...listParams
            });

            for (const event of allEvents) {
                done = !await visitor(event, employee, employees, calendar, personio);
                if (done) {
                    break;
                }
            }
        } catch (e) {
            Logger.log('Failed to visit events of user %s: %s', email, e);
            firstError = firstError || e;
        }
        ++processedCount;
        if (done) {
            break;
        }
    }

    Logger.log('Completed visiting events for %s of %s accounts', '' + processedCount, '' + employees.length);

    if (firstError) {
        throw firstError;
    }
}


/** Get emails of people who actually attended the Google Meet for a calendar event.
 *
 * Uses the Google Meet REST API to retrieve conference participants, then resolves
 * their display names to primary emails via the Google Directory (People API).
 * Both data sources use Google's identity system, so display names match reliably.
 *
 * Requires the following scopes to be authorized for the service account
 * in Workspace Admin Console (domain-wide delegation):
 * - meetings.space.readonly (Google Meet conference records)
 * - directory.readonly (Google Directory people listing)
 *
 * @param event The calendar event (must include conferenceData, requires conferenceDataVersion=1 in list params)
 * @returns {Promise<Set<string>|null>} Set of attendee email addresses, or null if attendance could not be determined
 */
async function getMeetAttendeeEmails_(event) {
    const meetingCode = event.conferenceData?.conferenceId
        || extractMeetingCode_(event);

    if (!meetingCode) {
        return null;
    }

    const email = event.creator?.email;
    if (!email) {
        return null;
    }

    const creds = getServiceAccountCredentials_();
    const meet = await MeetClient.withImpersonatingService(creds, email);

    // Search for conference records within a window around the event
    const eventStart = new Date(event.start.dateTime || event.start.date);
    const startFilter = new Date(eventStart.getTime() - 30 * 60 * 1000).toISOString();
    const endFilter = new Date(eventStart.getTime() + 30 * 60 * 1000).toISOString();

    const filter = `space.meeting_code="${meetingCode}" AND start_time>="${startFilter}" AND start_time<="${endFilter}"`;
    const records = await meet.listConferenceRecords(filter);

    if (!records.length) {
        return null;
    }

    const participants = await meet.listParticipants(records[0].name);

    // Resolve participant display names to emails via Google Directory
    // (both Meet and Directory use Google identity, so names match)
    const directoryNameToEmail = await getDirectoryNameToEmail_(creds, email);

    const attendeeEmails = new Set();
    for (const participant of participants) {
        const displayName = (participant.signedinUser?.displayName || '').trim().toLowerCase();
        if (displayName && directoryNameToEmail[displayName]) {
            attendeeEmails.add(directoryNameToEmail[displayName]);
        }
    }

    return attendeeEmails.size > 0 ? attendeeEmails : null;
}


/** Build a mapping from Google Directory display names (lowercase) to primary emails.
 *
 * @param serviceAccountCredentials Service account credentials JSON
 * @param impersonateEmail Email of domain user to impersonate
 * @returns {Promise<Object>} Map of lowercase display name to primary email
 */
async function getDirectoryNameToEmail_(serviceAccountCredentials, impersonateEmail) {
    const directory = await DirectoryClient.withImpersonatingService(serviceAccountCredentials, impersonateEmail);
    const people = await directory.listDirectoryPeople();

    const nameToEmail = {};
    for (const person of people) {
        const name = person.names?.[0]?.displayName;
        const email = person.emailAddresses?.[0]?.value;
        if (name && email) {
            nameToEmail[name.trim().toLowerCase()] = email;
        }
    }

    return nameToEmail;
}


/** Extract Google Meet meeting code from calendar event entry points. */
function extractMeetingCode_(event) {
    const videoEntryPoint = event.conferenceData?.entryPoints?.find(ep => ep.entryPointType === 'video');
    if (videoEntryPoint?.uri) {
        const match = videoEntryPoint.uri.match(/meet\.google\.com\/([a-z]+-[a-z]+-[a-z]+)/);
        if (match) {
            return match[1];
        }
    }
    return null;
}


/** Get script properties. */
function getScriptProperties_() {
    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties) {
        throw new Error('ScriptProperties not accessible');
    }

    return scriptProperties;
}


/** Get the report target sheet ID. */
function getReportSheetId_() {
    const sheetId = getScriptProperties_().getProperty(REPORT_SHEET_KEY);
    if (!sheetId) {
        throw new Error("No report target sheet ID specified at script property " + REPORT_SHEET_KEY);
    }

    return sheetId;
}


/** Get the service account credentials. */
function getServiceAccountCredentials_() {
    const creds = getScriptProperties_().getProperty(SERVICE_ACCOUNT_CREDENTIALS_KEY);
    if (!creds) {
        throw new Error("No service account credentials at script property " + SERVICE_ACCOUNT_CREDENTIALS_KEY);
    }

    return JSON.parse(creds);
}


/** Get the email account white-list (optional, leave empty to sync all suitable accounts). */
function getEmailWhiteList_() {
    return (getScriptProperties_().getProperty(EMAIL_WHITELIST_KEY) || '').trim()
        .split(',').map(email => email.trim()).filter(email => !!email);
}


/** Get the number of lookahead days (positive integer) or the default (6 * 30 days). */
function getLookaheadDays_() {
    const creds = (getScriptProperties_().getProperty(LOOKAHEAD_DAYS_KEY) || '').trim();
    const lookaheadDays = Math.abs(Math.round(+creds));
    if (!lookaheadDays || Number.isNaN(lookaheadDays)) {
        // use default: 6 months
        return 6 * 30;
    }

    return lookaheadDays;
}


/** Get the number of lookback days (positive integer) or the default (30 days). */
function getLookbackDays_() {
    const creds = (getScriptProperties_().getProperty(LOOKBACK_DAYS_KEY) || '').trim();
    const lookbackDays = Math.abs(Math.round(+creds));
    if (!lookbackDays || Number.isNaN(lookbackDays)) {
        // use default: 30 days
        return 30;
    }

    return lookbackDays;
}


/** Get the Personio token. */
function getPersonioCreds_() {
    const credentialFields = (getScriptProperties_().getProperty(PERSONIO_TOKEN_KEY) || '|')
        .split('|')
        .map(field => field.trim());

    return {clientId: credentialFields[0], clientSecret: credentialFields[1]};
}


/** Get Slack user by email with caching to avoid rate limits.
 *
 * @param slack The SlackWebClient instance
 * @param email The email address to look up
 * @returns The Slack user object or null if not found
 */
async function getSlackUserByEmailCached_(slack, email) {
    const CACHE_TTL_MIN_MS = 3.5 * 24 * 60 * 60 * 1000; // 0.5 weeks
    const CACHE_TTL_MAX_MS = 14 * 24 * 60 * 60 * 1000;  // 2 weeks
    const cacheKey = `slackUser.${email}`;
    const scriptProps = getScriptProperties_();

    // Try to get from cache first (timestamp and TTL stored with value to minimize property accesses)
    const cachedData = scriptProps.getProperty(cacheKey);
    if (cachedData) {
        const { ts, ttl, user } = JSON.parse(cachedData);
        if (ts && ttl && (Date.now() - ts) < ttl) {
            return user;
        }
    }

    // Fetch from Slack API if not cached or expired
    const slackUser = await slack.lookupByEmail(email);
    if (slackUser) {
        // Randomize TTL to avoid mass expiry causing rate limits
        const ttl = CACHE_TTL_MIN_MS + Math.random() * (CACHE_TTL_MAX_MS - CACHE_TTL_MIN_MS);
        scriptProps.setProperty(cacheKey, JSON.stringify({ ts: Date.now(), ttl, user: slackUser }));
    }

    return slackUser;
}


/** Get the Slack bot token for posting messages. */
function getSlackBotToken_() {
    return getScriptProperties_().getProperty(SLACK_BOT_TOKEN) || null;
}


/** Get the Gemini API key for summarization. */
function getGeminiApiKey_() {
    return getScriptProperties_().getProperty(GEMINI_API_KEY) || null;
}


/** Set script properties.
 *
 * Usage: clasp run 'setProperties' --params '[{"Meetings.artifactsFolderId": "FOLDER_ID"}, false]'
 */
function setProperties(properties, deleteAllOthers) {
    TriggerUtil.setProperties(properties, deleteAllOthers);
}
