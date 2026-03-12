/**
 * Sync Personio employees to Google Workspace group membership
 *
 *   - The source URL, tokens, target group and filter criteria for sync process
 *     must be specified as ScriptProperties (see setProperties() helper function).
 *
 *     FORMAT:
 *          Script Property Key: PersonioToGroup.<GROUP_EMAIL>
 *          Script Property Value: FULL_PERSONIO_API_URL|PERSONIO_CLIENT_ID|PERSONIO_CLIENT_SECRET|FILTER
 *
 *     EXAMPLE:
 *         "PersonioToGroup.engineers@example.com": "/company/employees|abc123|secret456|status=active,employment_type=internal"
 *
 *     FILTER is a comma-separated list of attribute=value pairs (AND logic).
 *     Attribute names correspond to Personio API v1 employee attribute keys.
 *     Values are matched case-insensitively against the unboxed attribute value.
 *
 *   - The executing account must have admin privileges to manage the target group(s).
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */

/** The prefix for properties specific to this script in the project. */
const PROPERTY_PREFIX = 'PersonioToGroup.';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'syncPersonioToGroup';


/** Main entry point.
 *
 * Take configuration from ScriptProperties and perform synchronization.
 */
async function syncPersonioToGroup() {

    let firstError = null;

    // we keep operating if a single sync task fails
    for (const task of getTasks_()) {

        const personio = PersonioClientV1.withApiCredentials(task.source.clientId, task.source.clientSecret);

        let employees = null;
        try {
            employees = await personio.getPersonioJson(task.source.url);
        } catch (e) {
            Logger.log('Failed to fetch Personio data for group %s: %s', task.groupEmail, e.message);
            firstError = firstError || e;
            continue;
        }

        try {
            const matching = filterEmployees_(employees, task.filter);
            const targetEmails = matching
                .map(emp => getEmployeeEmail_(emp))
                .filter(email => email != null);

            Logger.log('Group %s: %s employees match filter (%s with email)', task.groupEmail, matching.length, targetEmails.length);

            const result = await reconcileGroupMembers_(task.groupEmail, targetEmails);
            Logger.log('Group %s: added %s, removed %s', task.groupEmail, result.added, result.removed);
        } catch (e) {
            Logger.log('Failed to reconcile group %s: %s', task.groupEmail, e.message);
            firstError = firstError || e;
        }
    }

    if (firstError) {
        throw firstError;
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

        const groupEmail = safeKey.replace(PROPERTY_PREFIX, '');
        if (!groupEmail) {
            continue;
        }

        try {
            const rawProperty = properties[key] || '';
            const parts = rawProperty.trim().split('|');

            if (parts.length === 4) {
                const sourceSpec = {
                    url: parts[0].trim(),
                    clientId: parts[1].trim(),
                    clientSecret: parts[2].trim()
                };

                const filterSpec = parts[3].trim();
                const filter = filterSpec ? filterSpec.split(',').map(pair => {
                    const [filterKey, ...rest] = pair.trim().split('=');
                    return {key: filterKey.trim(), value: rest.join('=').trim()};
                }) : [];

                if (sourceSpec.url && sourceSpec.clientId && sourceSpec.clientSecret) {
                    tasks.push({groupEmail: groupEmail, source: sourceSpec, filter: filter});
                } else {
                    Logger.log("Skipped task: Empty fields in property value for key %s: %s", key, rawProperty);
                }
            } else {
                Logger.log("Skipped task: Expected 4 fields (URL, CLIENT_ID, CLIENT_SECRET, FILTER) in property value for key %s: %s", key, rawProperty);
            }
        } catch (e) {
            Logger.log('Skipped task: Incorrect config for property key %s: %s', key, e.message);
        }
    }

    return tasks;
}


/** Filter Personio employee objects by the given filter criteria (AND logic).
 *
 * Each filter criterion is matched case-insensitively against the unboxed attribute value.
 */
function filterEmployees_(employees, filter) {
    if (!filter || filter.length === 0) {
        return employees;
    }

    return employees.filter(employee => {
        const attributes = employee?.attributes;
        if (!attributes) return false;

        return filter.every(criterion => {
            const attribute = attributes[criterion.key];
            const value = unboxValue_(attribute);
            if (value == null) return false;

            return String(value).toLowerCase() === criterion.value.toLowerCase();
        });
    });
}


/** Unbox a Personio attribute value (handle {label, value} boxing). */
function unboxValue_(boxed) {
    if (boxed?.value !== undefined) {
        return boxed.value;
    }
    return boxed;
}


/** Extract email from a Personio employee object. */
function getEmployeeEmail_(employee) {
    const email = unboxValue_(employee?.attributes?.email);
    return email || null;
}


/** Reconcile group membership: add missing members, remove non-matching ones.
 *
 * @param {string} groupEmail The target group email address.
 * @param {string[]} targetEmails The desired member email addresses.
 * @return {{added: number, removed: number}} Summary of changes made.
 */
async function reconcileGroupMembers_(groupEmail, targetEmails) {
    // Resolve group email to internal ID (more reliable than email as groupKey)
    const group = AdminDirectory.Groups.get(groupEmail);
    const groupKey = group.id;

    // List current members with pagination
    const currentEmails = new Set();
    const options = {maxResults: 200};
    do {
        const response = AdminDirectory.Members.list(groupKey, options);
        const members = response.members || [];
        for (const member of members) {
            if (member.email) {
                currentEmails.add(member.email.toLowerCase());
            }
        }
        options.pageToken = response.nextPageToken;
    } while (options.pageToken);

    const targetSet = new Set(targetEmails.map(e => e.toLowerCase()));

    // Compute diff
    const toAdd = [...targetSet].filter(email => !currentEmails.has(email));
    const toRemove = [...currentEmails].filter(email => !targetSet.has(email));

    let added = 0;
    for (const email of toAdd) {
        Logger.log('Group %s: adding %s', groupEmail, email);
        AdminDirectory.Members.insert({email: email, role: 'MEMBER'}, groupKey);
        added++;
    }

    let removed = 0;
    for (const email of toRemove) {
        Logger.log('Group %s: removing %s', groupEmail, email);
        AdminDirectory.Members.remove(groupKey, email);
        removed++;
    }

    return {added: added, removed: removed};
}
