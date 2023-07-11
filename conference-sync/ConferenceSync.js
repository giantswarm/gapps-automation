/**
 * Functions for syncing events from forms/spreadsheets to calendar entries.
 *
 * Original author: marcus@giantswarm.io
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */

/** The prefix for properties specific to this script in the project. */
const PROPERTY_PREFIX = 'ConferenceSync.';

/** The Google Calendar to create conference events in. */
const CALENDAR_ID_KEY = PROPERTY_PREFIX + 'calendarId';

/** The sheet name inside the Google Sheets spreadsheet which serves as a container for this script project. */
const SHEET_NAME_KEY = PROPERTY_PREFIX + 'sheetName';

/** The URL of some kind of table or overview page or document. */
const OVERVIEW_URL_KEY = PROPERTY_PREFIX + 'overviewUrl';

/** A Slack app bot token.
 *
 * The connected Slack app must:
 *   - have at least the scopes channels:read, groups:read, chat:write and chat:write.public
 *   - be added to a private or public channel to which notifications will be posted
 */
const SLACK_BOT_TOKEN_KEY = PROPERTY_PREFIX + 'slackBotToken';

/** Call for papers deadline warning threshold in days (default: 7). */
const DEADLINE_WARNING_DAYS = PROPERTY_PREFIX + 'deadlineWarnDays';

/** Register this function with a "On form submit" trigger in Sheet context.
 *
 * USAGE:
 *
 *  1. Deploy this script into a spreadsheet with the right columns (conference form)
 *  2. Ensure the calendar is shared with and subscribed by the account running the script
 *  3. Set the property ConferenceSync.calendarId to the calendar ID.
 *  4. Setup "On form submit" trigger to call onFormSubmit()
 */
function onFormSubmit() {
    const calendarId = getScriptProperties_().getProperty(CALENDAR_ID_KEY);
    if (!calendarId) {
        throw new Error('Script property ConferenceSync.calendarId not configured, please set to target calendar ID');
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
        throw new Error('No active spreadsheet, please deploy this script into a spreadsheet context');
    }

    const sheetName = getScriptProperties_().getProperty(SHEET_NAME_KEY);
    syncToCalendar_(spreadsheet.getId(), sheetName, calendarId);
}

/** Synchronize rows from a sheet with well-known columns from a conference form to a calendar. */
function syncToCalendar_(spreadsheetId, sheetName, calendarId) {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
        throw new Error(`Failed to access calendar ${calendarId}, please check permissions`);
    }
    const sheet = spreadsheet.getSheetByName(sheetName);
    SheetUtil.mapRows(sheet.getSheetValues(1, 1, -1, -1), null).forEach((row, i) => {

        const untilDeadlineMs = new Date(row.deadline) - new Date();
        if (!row.notified_deadline_date && untilDeadlineMs) {
            const deadlineWarnThresholdMs = +(getScriptProperties_().getProperty(DEADLINE_WARNING_DAYS) || 7) * 24 * 60 * 60 * 1000;
            if (untilDeadlineMs > 0 && untilDeadlineMs < deadlineWarnThresholdMs) {
                if (notifySlack_(row)) {
                    sheet.getRange(i + 2, row.$columnIndex.notified_deadline_date + 1).setValue(new Date());
                }
            }
        }

        if (row.calender_url) {
            // skip entries which have a calendar entry already
            return;
        }

        const event = calendar.createAllDayEvent(`${row.conference_name} - CfP Deadline`, row.deadline);
        event.setLocation(row.submission_website)

        let description = `
<b>${row.conference_name}</b>
<u>CfP Closes:</u> ${formatDate_(row.deadline)}
<u>CfP Submission:</u> ${row.submission_website}`;

        if (row.max_number_of_submissions) {
            description += `\n<u>Max Submissions:</u> ${row.max_number_of_submissions}`
        }

        description += "\n"

        if (row.conference_website) {
            description += `\n<u>Event URL:</u> ${row.conference_website}`
        }
        if (row.conference_date_start && row.conference_date_end) {
            description += `\n<u>Event Dates:</u> ${formatDate_(row.conference_date_start)} - ${formatDate_(row.conference_date_end)}`
        } else if (row.conference_date_start) {
            description += `\n<u>Event Date:</u> ${formatDate_(row.conference_date_start)}`
        }
        if (row.city || row.country) {
            description += `\n<u>Location:</u> ${[row.city, row.country].filter(a => !!a).join(', ')}`
        }
        if (row.conference_theme) {
            description += `\n<u>Theme:</u> ${row.conference_theme}`
        }

        event.setDescription(description);

        const eventId = Utilities.base64Encode(event.getId().split('@')[0] + calendarId).replace('=', '');
        const url = `https://calendar.google.com/calendar/event?eid=${eventId}`;
        sheet.getRange(i + 2, row.$columnIndex.calender_url + 1).setValue(url);
    });
}

function formatDate_(d) {
    const d2 = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0));
    return Utilities.formatDate(d2, "UTC", "d MMM, yyyy");
}

/** Notifies Slack channels about the approaching CfP deadline. */
function notifySlack_(cfpRecord) {
    const botToken = getScriptProperties_().getProperty(SLACK_BOT_TOKEN_KEY);
    if (!botToken) {
        return false;
    }
    const slack = new SlackWebClient(botToken);
    const untilDeadlineDays = Math.round((new Date(cfpRecord.deadline) - new Date()) / (24 * 60 * 60 * 1000));
    slack.broadcastMessage(`:star: Dear Researchers and Rockstars! :star2:

The *call for papers* deadline for *${cfpRecord.conference_name}* is only *${untilDeadlineDays} ${untilDeadlineDays > 1 ? 'days' : 'day'}* away!

:loudspeaker:  Please hurry, we'd all love you hear your talk!

All the details here: ${getScriptProperties_().getProperty(OVERVIEW_URL_KEY)}`);

    return true;
}

/** Get script properties. */
function getScriptProperties_() {
    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties) {
        throw new Error('ScriptProperties not accessible');
    }

    return scriptProperties;
}
