/**
 * Truncate mailboxes of certain (automated) users or to save space.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */


/** The prefix for properties specific to this script in the project. */
const PROPERTY_PREFIX = 'TruncateMailboxes.';

/** Service account credentials. */
const SERVICE_ACCOUNT_CREDENTIALS_KEY = PROPERTY_PREFIX + 'serviceAccountCredentials';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'truncateMailboxes';


/** Main entry point.
 *
 * Truncate mailboxes of users listed in the Script Properties by search filters.
 *
 * Requires the following script properties for operation:
 *
 *   TruncateMailboxes.serviceAccountCredentials  {...SERVICE_ACCOUNT_CREDENTIALS...}
 *   TruncateMailboxes.EMAIL_1                    CLEAN_FILTER
 *   TruncateMailboxes.EMAIL_2                    CLEAN_FILTER
 *   ...
 *
 *   EMAIL should be a primary email within this organization.
 *
 *   CLEAN_FILTER is a gmail search expression, adhering to: https://support.google.com/mail/answer/7190?hl=en
 *      Examples:
 *        older_than:30d
 *        size:1000000
 *
 * Requires service account credentials (with domain-wide delegation enabled) to be set, for example via:
 *
 *   $ clasp run 'setProperties' --params '[{"TruncateMailboxes.serviceAccountCredentials": "{...ESCAPED_JSON...}"}, false]'
 *
 * One may use the following command line to compress the service account creds into one line:
 *
 *   $ cat credentials.json | tr -d '\n '
 *
 * The service account must be configured correctly and have at least permission for these scopes:
 *   https://www.googleapis.com/auth/calendar
 *
 * Ensure that this methods ExecutionApi scope is limited to MYSELF in the manifest,
 * to prevent other unauthorized domain users from using its features.
 */
function truncateMailboxes() {

    let firstError = null;

    // we keep operating if a single sync task fails
    for (const task of getTasks_()) {

        try {
            cleanUserMailbox_(getServiceAccountCredentials_(), task.primaryEmail, task.cleanFilter);
        } catch (e) {
            Logger.log('Failed to clean mailbox %s: %s', task.primaryEmail, e.message);
            firstError = firstError || e;
        }
    }

    if (firstError) {
        throw firstError;
    }
}


/** Clean the specified user's mailbox based on search expression.
 *
 * WARNING: This is intended to be used for emergencies, adhering to regulatory requirements or to
 *          keep automated account's inboxes clean. Use with utmost care!
 */
function cleanUserMailbox_(serviceAccountCredentials, primaryEmail, searchExpression) {

    const gmail = GmailClientV1.withImpersonatingService(serviceAccountCredentials, primaryEmail);

    const cleanMessages = (messages) => {
        gmail.deleteUserMessages(primaryEmail, messages.map(message => message.id));
        return [messages.length]; // accumulate only stats
    };

    const batches = gmail.listUserMessages(primaryEmail, searchExpression, cleanMessages);

    const removalCount = batches.reduce((s1, s2) => s1 + s2);
    Logger.log("mailbox %s: deleted %s messages searching: %s", primaryEmail, ""+removalCount, searchExpression);
}


function getTasks_() {

    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties) {
        Logger.log('ScriptProperties not accessible');
        return [];
    }

    const properties = PropertiesService.getScriptProperties().getProperties() || {};
    const tasks = [];
    for (const key in properties) {

        const safeKey = key.trim();
        if (!safeKey.startsWith(PROPERTY_PREFIX))
            continue;

        const primaryEmail = safeKey.replace(PROPERTY_PREFIX, '');
        // an email should at least look like x@y.z
        if (!primaryEmail || !primaryEmail.includes('@') || primaryEmail.length < 5) {
            continue;
        }

        try {
            const cleanFilter = properties[key] || '';
            if (cleanFilter) {
                tasks.push({primaryEmail: primaryEmail, cleanFilter: cleanFilter});
            } else {
                Logger.log("Skipped task: key %s: value (CLEAN_FILTER) must be a gmail search expression", key);
            }
        } catch (e) {
            Logger.log('Skipped task: Incorrect config for property key %s: %s', key, e.message);
        }
    }

    return tasks;
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


/** Get the service account credentials. */
function getServiceAccountCredentials_() {
    const creds = getScriptProperties_().getProperty(SERVICE_ACCOUNT_CREDENTIALS_KEY);
    if (!creds) {
        throw new Error("No service account credentials at script property " + SERVICE_ACCOUNT_CREDENTIALS_KEY);
    }

    return JSON.parse(creds);
}
