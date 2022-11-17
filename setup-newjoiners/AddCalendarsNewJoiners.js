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

/** Filter for allowed domains (to avoid working and failing on users present on foreign domains). */
const ALLOWED_DOMAINS_KEY = PROPERTY_PREFIX + 'allowedDomains';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'addCalendarsNewJoiners';


/** Main entry point.
 *
 * Add required calendars to all GApps organization users tagged as "onboarding" in Personio.
 *
 * This requires a configured Personio access token and the user impersonation API enabled for the cloud project:
 *  https://cloud.google.com/iam/docs/impersonating-service-accounts
 *
 * Requires the following script properties for operation:
 *
 *   AddCalendars.personioToken              CLIENT_ID|CLIENT_SECRET
 *   AddCalendars.serviceAccountCredentials  {...SERVICE_ACCOUNT_CREDENTIALS...}
 *   AddCalendars.allowedDomains             giantswarm.io,giantswarm.com
 *
 * One can use the following command line to compress the service account creds into one line:
 *   $ cat credentials.json | tr -d '\n '
 *
 * The service account must be configured correctly and have at least permission for these scopes:
 *   https://www.googleapis.com/auth/calendar
 */
function addCalendarsToNewJoiners() {

    const calendarIds = getRequiredCalendars_();
    if (!calendarIds || calendarIds.length === 0) {
        Logger.log("No calendars configured, exiting");
        return;
    }
    Logger.log('Configured to ensure calendars: %s', calendarIds);

    const allowedDomains = (getScriptProperties_().getProperty(ALLOWED_DOMAINS_KEY) || '')
        .split(',')
        .map(d => d.trim());
    const isEmailDomainAllowed = email => allowedDomains.includes(email.substring(email.lastIndexOf('@') + 1));

    Logger.log('Configured to handle users on domains: %s', allowedDomains);

    const newJoinerEmails = getPersonioEmployeeEmailsByStatus_('onboarding').filter(email => isEmailDomainAllowed(email));

    let firstError = null;

    for (const primaryEmail of newJoinerEmails) {
        // we keep operating if handling calendars of a single user fails
        try {
            subscribeUserCalendars_(getServiceAccountCredentials_(), primaryEmail, calendarIds);
        } catch (e) {
            Logger.log('Failed to ensure calendars of user %s: %s', primaryEmail, e.message);
            firstError = firstError || e;
        }
    }

    if (firstError) {
        throw firstError;
    }
}


/** Subscribe a single account to all the specified calendars. */
function subscribeUserCalendars_(serviceAccountCredentials, primaryEmail, calendarIds) {

    const calendarList = CalendarListClient.withImpersonatingService(serviceAccountCredentials, primaryEmail);

    const existingCalendars = calendarList.list();
    for (const calendarId of calendarIds) {
        if (!existingCalendars.some(calendar => calendar.id === calendarId)) {
            // continue operating if adding a single calendar fails
            try {
                const newItem = calendarList.insert({id: calendarId});
                existingCalendars.push(newItem); // to handle duplicates in calendarIds
                Logger.log('Added calendar %s to user %s', calendarId, primaryEmail);
            } catch (e) {
                Logger.log('Failed to add calendar %s to user %s: %s', calendarId, primaryEmail, e.message);
            }
        }
    }
}


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


/** Get script properties. */
function getScriptProperties_() {
    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties) {
        throw new Error('ScriptProperties not accessible');
    }

    return scriptProperties;
}


/** Get the list of required calendars. */
function getRequiredCalendars_() {
    const requiredCalendarsList = getScriptProperties_().getProperty(REQUIRED_CALENDARS_KEY) || '';

    return requiredCalendarsList.split(',')
        .map(calId => calId.trim())
        .map(calId => Utilities.newBlob(Utilities.base64Decode(calId)).getDataAsString());
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
    const personio = PersonioClientV1.withApiCredentials(personioCreds.clientId, personioCreds.clientSecret);

    const data = personio.getPersonioJson('/company/employees');
    const emails = [];
    for (const item of data) {
        const attributes = item?.attributes;

        if (attributes?.status?.value === status) {
            const email = (attributes?.email?.value || '').trim();
            if (email) {
                emails.push(email);
            }
        }
    }

    return emails;
}
