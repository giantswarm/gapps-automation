/** Script to generate various meeting reports.
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


/** Build meetings statistic per employee and summarized. */
async function listMeetingStatistic() {

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

    const allowedDomains = (getScriptProperties_().getProperty(ALLOWED_DOMAINS_KEY) || '')
        .split(',')
        .map(d => d.trim());

    const employeeStats = {};
    try {
        let haveEmployees = false;
        // visitEvents_() will call our visitor function for each employee and calendar event combination
        await visitEvents_((event, employee, employees, calendar, personio) => {

            const getGiantSwarmAttendees = attendees =>
                attendees.filter(attendee => allowedDomains.filter(domain => (attendee.email || '').includes(domain)).length);

            if (!haveEmployees) {
                for (const e of employees) {
                    const email = e.attributes.email.value;

                    const stats = {email: email, count: 0, weeks: []};
                    for (let i = 0; i < weeks; ++i) {
                        stats.weeks.push(0.0);
                    }

                    employeeStats[email] = stats;
                }
                haveEmployees = true;
            }

            const stats = employeeStats[employee.attributes.email.value];

            if (Array.isArray(event.attendees) && event.attendees.length > 1
                && !event.extendedProperties?.private?.timeOffId
                && event.eventType !== 'outOfOffice'
                && getGiantSwarmAttendees(event.attendees).length === event.attendees.length) {
                const start = new Date(event.start.dateTime);
                const end = new Date(event.end.dateTime);
                const durationHours = (end - start) / (60.0 * 60 * 1000);
                if (!isNaN(durationHours)) {
                    const week = getWeekIndex(start, end);
                    if (week != null) {
                        stats.count += 1;
                        stats.weeks[week] += durationHours;
                    } else {
                        console.log(`Event for ${employee.attributes.email.value} doesn't fit any week: ${JSON.stringify(event)}`);
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
        const row = [stats.email, stats.count];

        for (const weekHours of stats.weeks) {
            row.push(+weekHours.toFixed(2));
        }

        return row;
    });

    const header = ["Email", "Meeting Count"];
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

    const recurring1on1Ids = {};
    const rows = [];
    try {
        // visitEvents_() will call our visitor function for each employee and calendar event combination
        await visitEvents_((event, employee, employees, calendar, personio) => {
            const email = employee.attributes.email.value;

            const hasNonEmployeeAttendees = attendees => {
                for (const attendee of attendees) {
                    if (!employees.find(e => e.attributes.email.value === attendee.email)) {
                        return true;
                    }
                }

                return false;
            };

            // recurring event, organized by this employee with 2 attendees means it's a 1on1
            if (Array.isArray(event.attendees) && event.attendees.length === 2
                && Array.isArray(event.recurrence) && event.recurrence.length
                && event.attendees.find(attendee => attendee.email === email)
                && !hasNonEmployeeAttendees(event.attendees)) {

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
                done = !visitor(event, employee, employees, calendar, personio);
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
