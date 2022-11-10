/**
 * Ensure new joiners see all required Google Calendars.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */


/** The prefix for properties specific to this script in the project. */
const PROPERTY_PREFIX = 'AddCalendars.';

/** Comma-separated list of calendar IDs required for new joiners. */
const REQUIRED_CALENDARS_KEY = PROPERTY_PREFIX + 'required';

/** Personio clientId and clientSecret, separated by '|'. */
const PERSONIO_TOKEN_KEY = PROPERTY_PREFIX + 'personioToken';

/** Service account credentials (in JSON format, as downloaded from Google Management Console. */
const SERVICE_ACCOUNT_CREDENTIALS_KEY = PROPERTY_PREFIX + 'serviceAccountCredentials';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'addCalendarsNewJoiners';


/** Main entry point.
 *
 * Add required calendars to all GApps organization users tagged as "onboarding" in Personio.
 *
 * This requires a configured Personio access token and the user impersonation API enabled for the cloud project:
 *  https://cloud.google.com/iam/docs/impersonating-service-accounts
 */
function addCalendarsToNewJoiners() {

    const calendarIds = getRequiredCalendars();
    if (!calendarIds || calendarIds.length === 0) {
        Logger.log("No calendars configured, exiting");
        return;
    }

    const newJoinerEmails = getPersonioEmployeeEmailsByStatus_('onboarding');

    let firstError = null;

    for (const primaryEmail of newJoinerEmails) {

        // we keep operating if handling calendars of a single user fails
        try {
            const calendarListService = new CalendarListService(getServiceAccountCredentials_(), primaryEmail);

            const existingCalendars = calendarListService.list();
            for (const calendarId of calendarIds) {
                if (!existingCalendars.some(calendar => calendar.id === calendarId)) {
                    // continue operating if adding a single calendar fails
                    try {
                        const newItem = calendarListService.insert({id: calendarId});
                        existingCalendars.push(newItem); // to handle duplicates in calendarIds
                        Logger.log('Added calendar %s to user %s', calendarId, primaryEmail);
                    } catch (e) {
                        Logger.log('Failed to add calendar %s to user %s: %s', calendarId, primaryEmail, e.message);
                    }
                }
            }
        } catch (e) {
            Logger.log('Failed to access calendars of user %s: %s', primaryEmail, e.message);
            firstError = firstError || e;
        }
    }

    if (firstError) {
        throw firstError;
    }
}


/** Uninstall time based execution trigger for this script. */
function uninstall() {
    // Remove pre-existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === TRIGGER_HANDLER_FUNCTION) {
            ScriptApp.deleteTrigger(trigger);
            Logger.log("Uninstalled time based trigger for %s", TRIGGER_HANDLER_FUNCTION);
        }
    }
}


/** Setup for periodic execution and do some checks.
 * Supported values for delayMinutes: 1, 5, 10, 15 or 30
 */
function install(delayMinutes) {
    uninstall();

    const delay = sanitizeDelayMinutes_(delayMinutes);
    Logger.log("Installing time based trigger (every %s minutes)", delayMinutes);

    ScriptApp.newTrigger(TRIGGER_HANDLER_FUNCTION)
        .timeBased()
        .everyMinutes(delay)
        .create();

    Logger.log("Installed time based trigger for %s every %s minutes", TRIGGER_HANDLER_FUNCTION, delay);
}


/** Helper function to configure the required script properties.
 *
 * USAGE EXAMPLE:
 *   clasp run 'setProperties' --params '[{"persionio-dump.SHEET_ID": "SOME_PERSONIO_URL|CLIENT_ID|CLIENT_SECRET"}, false]'
 *
 * Warning: passing argument true for parameter deleteAllOthers will also cause the schema to be reset!
 */
function setProperties(properties, deleteAllOthers) {
    PropertiesService.getScriptProperties().setProperties(properties, deleteAllOthers);
}


/** Sanitize delay minutes input.
 *
 * ClockTriggerBuilder supported only a limited number of values, see:
 * https://developers.google.com/apps-script/reference/script/clock-trigger-builder#everyMinutes(Integer)
 */
function sanitizeDelayMinutes_(delayMinutes) {
    return [1, 5, 10, 15, 30].reduceRight((v, prev) =>
        typeof +delayMinutes === 'number' && v <= +delayMinutes ? v : prev);
}


/** Get script properties. */
function getScriptProperties_() {
    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties) {
        Logger.log('ScriptProperties not accessible');
        return throw new Error("ScriptProperties not accessible");
    }

    return scriptProperties;
}


/** Get the list of required calendars. */
function getRequiredCalendars_() {
    const requiredCalendarsList = getScriptProperties_().getProperty(REQUIRED_CALENDARS_KEY) || '';

    return requiredCalendarsList.split(',').map(calId => calId.trim());
}


/** Get the service account credentials. */
function getServiceAccountCredentials_() {
    const creds = getScriptProperties_().getProperty(SERVICE_ACCOUNT_CREDENTIALS_KEY);
    if (!creds) {
        throw new Error("No service account credentials at script property " + SERVICE_ACCOUNT_CREDENTIALS_KEY);
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


/** Get the list of new joiner primary organization emails.
 *
 * @param status One of: onboarding, active, offboarding, inactive
 */
function getPersonioEmployeeEmailsByStatus_(status) {

    const personioCreds = getPersonioCreds_();
    const personio = new PersonioClientV1(personioCreds.clientId, personioCreds.clientSecret);

    const data = personio.fetch('/company/employees.json');
    const emails = [];
    for (const item of data) {
        const attributes = item?.attributes;

        if (attributes?.status?.value === status) {
            const email = attributes?.email?.value;
            if (email) {
                emails.push(email);
            }
        }
    }

    return emails;
}
