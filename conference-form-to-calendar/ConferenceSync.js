/**
 * Functions for syncing events from forms/spreadsheets to calendar entries.
 *
 * Original author: marcus@giantswarm.io
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */

/** Register this function with a "On form submit" trigger in Sheet context. */
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

    const data = sheet.getDataRange().getValues();
    const columnIndexes = getColumnIndexes_(data[0]);

    data.forEach((row, i) => {

        const calendarUrl = row[columnIndexes.CalURL];
        if (i === 0 || calendarUrl) {
            // Skip the header and entries which have a calendar entry already
            return;
        }

        const name = row[columnIndexes.Name];
        const cfpDeadline = row[columnIndexes.CfpDeadline];
        const cfpURL = row[columnIndexes.CfpURL];
        const eventURL = row[columnIndexes.ConferenceURL];
        const eventStart = row[columnIndexes.ConferenceStart];
        const eventEnd = row[columnIndexes.ConferenceEnd];
        const city = row[columnIndexes.City];
        const country = row[columnIndexes.Country];
        const theme = row[columnIndexes.Theme];
        const maxSubmissions = row[columnIndexes.MaxSubmissions];

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
        sheet.getRange(i + 1, columnIndexes.CalURL + 1).setValue(url);
    });
}

function formatDate_(d) {
    const d2 = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0));
    return Utilities.formatDate(d2, "UTC", "d MMM, yyyy");
}

function getColumnIndexes_(headerRow) {
    const columns = {};

    headerRow.forEach((cell, i) => {
        switch (cell) {
            case "Timestamp":
                columns.Timestamp = i;
                break
            case "Conference Name":
                columns.Name = i;
                break
            case "Conference Website":
                columns.ConferenceURL = i;
                break
            case "Submission Website":
                columns.CfpURL = i;
                break
            case "Conference Date Start":
                columns.ConferenceStart = i;
                break
            case "Conference Date End":
                columns.ConferenceEnd = i;
                break
            case "Deadline":
                columns.CfpDeadline = i;
                break
            case "Country":
                columns.Country = i;
                break
            case "City":
                columns.City = i;
                break
            case "Conference Theme":
                columns.Theme = i;
                break
            case "Max number of submissions":
                columns.MaxSubmissions = i;
                break
            case "Calender URL":
                columns.CalURL = i;
                break
            default:
                break
        }
    })
    return columns
}