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
                        if (isSigOrChapter_(event)) {
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


/** Share team meeting artifacts with the organization. */
async function shareTeamMeetingArtifacts() {

    const shareDomainName = getScriptProperties_().getProperty(SHARE_DOMAIN_NAME) || '';
    if (!shareDomainName) {
        Logger.log('No share domain configured, skipping sharing team meeting artifacts');
        return;
    }

    let hits = 0;
    let hits_published = 0;
    const teamMeetingsCalendarId = getScriptProperties_().getProperty(TEAM_MEETINGS_CALENDAR) || '';
    try {
        // visitEvents_() will call the visitor function for each employee and calendar event combination
        await visitEvents_(async (event, employee, employees, calendar, personio) => {
            const email = employee.attributes.email.value;

            // Only process SIG/Chapter/WG events
            if (!isSigOrChapter_(event)) {
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

            const publishedAt = +event.extendedProperties?.private?.attachmentsPublishedAt;
            if (publishedAt) {
                Logger.log("event " + event.summary + " at " + event.start.dateTime + " already published at " + publishedAt);
                return true;
            }

            const geminiNotes = event.attachments?.find(a => a.mimeType === 'application/vnd.google-apps.document'
                && a.fileUrl.includes('usp=meet_tnfm_calendar')
                && a.title.includes('Notes'));
            const recording = event.attachments?.find(a => a.mimeType === 'video/mp4'
                && a.title.includes('Recording'));

            console.log('meet: ' + event.summary + ' ' + event.start.dateTime + ' notes=' + geminiNotes + ' rec=' + recording);

            if (geminiNotes && recording) {

                hits++;

                // Share with organization
                try {
                    const drive = await DriveClientV1.withImpersonatingService(getServiceAccountCredentials_(), email);

                    // Share the Gemini notes document
                    Logger.log(`Sharing notes ${geminiNotes.fileId} for event ${event.summary}`);
                    try {
                        await drive.shareWith(geminiNotes.fileId, 'domain', shareDomainName,'reader', true);
                    } catch (e) {
                        Logger.log(`Failed to share notes ${geminiNotes.fileId} (may have been moved already): ${e}`);
                    }

                    // Share the recording video
                    Logger.log(`Sharing recording ${recording.fileId} for event ${event.summary}`);
                    try {
                        await drive.shareWith(recording.fileId, 'domain', shareDomainName,'reader', true);
                    } catch (e) {
                        Logger.log(`Failed to share recording ${recording.fileId} (may have been moved already): ${e}`);
                    }

                    // Summarize and post to Slack
                    await summarizeAndPostToSlack_(event, geminiNotes, recording, drive, employees);

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
        });
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


function isSigOrChapter_(event) {
    return /(^|\s)(SIG|chapter|WG)(\s|$)/i.test(event.summary) || (event.attendees || []).find(a => a.email.startsWith('sig-') || a.email.startsWith('wg-') || a.email.startsWith('all@'));
}


function isOneOnOne_(event, employees, ownerEmail) {
    return Array.isArray(event.attendees) && event.attendees.length === 2
        && ((Array.isArray(event.recurrence) && event.recurrence.length) || event.recurringEventId)
        && !isSigOrChapter_(event)
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
 */
async function summarizeAndPostToSlack_(event, geminiNotes, recording, drive, employees) {
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
        const formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'dd/MM/yy');

        const nameToSlackHandleMapping = async () => {
            // Use all active employees instead of event.attendees (which may be incomplete)
            const employeeEmails = employees.map(emp => emp.attributes.email.value).filter(Boolean);
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

                        const slackHandle = slackUser.name; // Slack username (without @)
                        const displayName = fullName || slackUser.profile?.display_name || slackUser.real_name || slackUser.name;
                        mappingLines.push(`${displayName} -> @${slackHandle}`);
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

🧠 Top Takeaways
Format each takeaway with an appropriate emoji prefix:
- Use ✅ for decisions made
- Use 🚧 for blockers or challenges
- Use 🔜 for next steps or action items
- Use 💡 for key insights or ideas
- Use ⚠️ for risks or concerns

Keep each takeaway to 1-2 sentences maximum. Focus on outcomes, not discussions.

🙋 Key Contributors & Takeaways

Note: Only include this section if any of the following apply:
- Someone is assigned an action or deliverable
- A decision was clearly advocated for or blocked by a participant
- Someone flagged a risk, dependency, or introduced a new direction

🎯 Skip this section entirely if there's nothing impactful to highlight - this isn't about logging everyone's contributions.

If you do include this section, format each entry as:
• @{{SlackHandle}} – {{Brief description of their assigned task, key decision, or flagged blocker}}

Keep it short and outcome-focused. Only mention people with specific, actionable contributions.

Guidelines:
- Be concise and direct
- Focus on outcomes and decisions, not process
- Use clear, simple language
- Only output the sections above, nothing else

Name to SlackHandle mapping in the format "{{Display Name}} -> @{{SlackHandle}}" for employees, for replacing names with handles:
${await nameToSlackHandleMapping()}`;

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
            fields: [
                {
                    type: 'mrkdwn',
                    text: `*📌 Meeting:*\n${meetingName}`
                },
                {
                    type: 'mrkdwn',
                    text: `*📅 Date:*\n${date}`
                }
            ]
        },
        {
            type: 'divider'
        },
        {
            type: 'section',
            text: {
                type: 'mrkdwn',
                text: summaryContent
            }
        },
        {
            type: 'divider'
        },
        {
            type: 'section',
            text: {
                type: 'mrkdwn',
                text: `📄 <${docLink}|Full Summary + Transcript>\n🎥 <${recordingLink}|Recording>`
            }
        }
    ];

    // Use plain text as fallback for notifications
    const fallbackText = `Meeting Summary: ${meetingName} (${date})`;

    await slackClient.postMessage(channelId, fallbackText, blocks);
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

    Logger.log('Visiting events between %s and %s for %s accounts', fetchTimeMin.toISOString(), fetchTimeMax.toISOString(), '' + employees.length);

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
