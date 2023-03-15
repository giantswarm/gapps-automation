/**
 * Functions for syncing events from forms/spreadsheets to calendar entries.
 *
 * Original author: marcus@giantswarm.io
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */

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

    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties) {
        throw new Error('ScriptProperties not accessible');
    }

    const calendarId = scriptProperties.getProperty('ConferenceSync.calendarId')
    if (!calendarId) {
        throw new Error('Script property ConferenceSync.calendarId not configured, please set to target calendar ID');
    }

    const sheetName = scriptProperties.getProperty('ConferenceSync.sheetName')

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
        throw new Error('No active spreadsheet, please deploy this script into a spreadsheet context');
    }

    syncToCalendar_(spreadsheet.getId(), sheetName, calendarId);
}

/** Synchronize rows from a sheet with well-known columns from a conference form to a calendar. */
function syncToCalendar_(spreadsheetId, sheetName, calendarId) {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : spreadsheet.getSheets()[0];
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
        throw new Error(`Failed to access calendar ${calendarId}, please check permissions`);
    }

    const data = sheet.getDataRange().getValues();
    const columnIndices = getColumnIndices_(data[0]);

    data.forEach((row, i) => {

        const calendarUrl = row[columnIndices.CalURL];
        if (i === 0 || calendarUrl) {
            // Skip the header and entries which have a calendar entry already
            return;
        }

        const name = row[columnIndices.Name];
        const cfpDeadline = row[columnIndices.CfpDeadline];
        const cfpURL = row[columnIndices.CfpURL];
        const eventURL = row[columnIndices.ConferenceURL];
        const eventStart = row[columnIndices.ConferenceStart];
        const eventEnd = row[columnIndices.ConferenceEnd];
        const city = row[columnIndices.City];
        const country = row[columnIndices.Country];
        const theme = row[columnIndices.Theme];
        const maxSubmissions = row[columnIndices.MaxSubmissions];

        const formattedName = name + " - CfP Deadline"
        const event = calendar.createAllDayEvent(formattedName, cfpDeadline);
        event.setLocation(cfpURL)

        let description = `
<b>${name}</b>
<u>CfP Closes:</u> ${formatDate_(cfpDeadline)}
<u>CfP Submission:</u> ${cfpURL}`;
        if (maxSubmissions) {
            description += `\n<u>Max Submissions:</u> ${maxSubmissions}`
        }

        description += "\n"

        if (eventURL) {
            description += `\n<u>Event URL:</u> ${eventURL}`
        }
        if (eventStart && eventEnd) {
            description += `\n<u>Event Dates:</u> ${formatDate_(eventStart)} - ${formatDate_(eventEnd)}`
        } else if (eventStart) {
            description += `\n<u>Event Date:</u> ${formatDate_(eventStart)}`
        }
        if (city || country) {
            description += `\n<u>Location:</u> ${[city, country].filter(a => !!a).join(', ')}`
        }
        if (theme) {
            description += `\n<u>Theme:</u> ${theme}`
        }

        event.setDescription(description);

        // Reminder 7 days before
        event.addEmailReminder(10080)

        const eventId = Utilities.base64Encode(event.getId().split('@')[0] + calendarId).replace('=', '');
        const url = `https://calendar.google.com/calendar/event?eid=${eventId}`;
        sheet.getRange(i + 1, columnIndices.CalURL + 1).setValue(url);
    });
}

function formatDate_(d) {
    const d2 = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0));
    return Utilities.formatDate(d2, "UTC", "d MMM, yyyy");
}

function getColumnIndices_(headerRow) {
    const indexOfColumn = columnName => {
        const index = headerRow.indexOf(columnName);
        if (index < 0) {
            throw new Error(`Required column "${columnName}" not found in sheet`);
        }
        return index;
    }

    return {
        Timestamp: indexOfColumn("Timestamp"),
        Name: indexOfColumn("Conference Name"),
        ConferenceURL: indexOfColumn("Conference Website"),
        CfpURL: indexOfColumn("Submission Website"),
        ConferenceStart: indexOfColumn("Conference Date Start"),
        ConferenceEnd: indexOfColumn("Conference Date End"),
        CfpDeadline: indexOfColumn("Deadline"),
        Country: indexOfColumn("Country"),
        City: indexOfColumn("City"),
        Theme: indexOfColumn("Conference Theme"),
        MaxSubmissions: indexOfColumn("Max number of submissions"),
        CalURL: indexOfColumn("Calender URL")
    };
}
