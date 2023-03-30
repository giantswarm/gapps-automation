/**
 * Edit Personio Employee records (privileged access only).
 *
 * IMPORTANT:
 *   For security reasons, this script MUST be deployed with executionApi.access = "MYSELF" in appsscript.json (the manifest).
 *
 *   See: https://developers.google.com/apps-script/manifest/web-app-api-executable#executionapi
 *
 * Only a handful of employee properties can be modified via API, see:
 *   https://developer.personio.de/reference/patch_company-employees-employee-id
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */

/** The prefix for properties specific to this script in the project.
 *
 * Making this dynamic didn't work as DriveApp.getFileById() tends to throw 500s and is otherwise quite slow.
 */
const PROPERTY_PREFIX = 'PersonioEdit.';

/** Key for the pipe separated Personio token (CLIENT_ID|CLIENT_SECRET). */
const PERSONIO_TOKEN_KEY = PROPERTY_PREFIX + 'personioToken'

/** Patch one or multiple employees.
 *
 * The employee objects must contain the field "id" and be otherwise structured as documented on:
 *   https://developer.personio.de/reference/patch_company-employees-employee-id
 *
 * Execution stops at the first error, and error message containing the failed employee is returned.
 *
 * USAGE:
 *
 *   Only via script API or debugger, clasp example:
 *
 *   $ clasp run 'patchEmployees' --params '[{"id": 1540283, "last_name": "Ajmera"}, {"id": 8919186, "last_name": "Wu"}]'
 *
 */
function patchEmployees(...employees) {

    const creds = getPersonioCreds_();
    const personio = PersonioClientV1.withApiCredentials(creds.clientId, creds.clientSecret);

    // we keep operating if a single sync task fails
    for (const employee of employees) {

        if (!Util.isObject(employee) || employee.id == null) {
            throw new Error('Invalid argument specified (not a partial employee with "id"): ' + JSON.stringify(employee));
        }

        const {id, ...employeeData} = employee;
        try {
            personio.fetchJson(`/company/employees/${id}`,
                {
                    "method": "patch",
                    "contentType": "application/json",
                    "payload": JSON.stringify(employeeData)
                });
            Logger.log('Patched employee %s: %s', id, JSON.stringify(employeeData));
        } catch (e) {
            throw new Error(`Failed to patch employee ${id}: ${e.message}`);
        }
    }
}


/** Allow setting properties. */
function setProperties(properties, deleteAllOthers) {
    TriggerUtil.setProperties(properties, deleteAllOthers);
}


function getPersonioCreds_() {

    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties) {
        throw new Error('ScriptProperties not accessible');
    }

    const key = PERSONIO_TOKEN_KEY;
    const rawProperty = scriptProperties.getProperty(key) || '';
    const apiParts = rawProperty.trim().split('|');

    if (apiParts.length === 2) {
        const sourceSpec = {
            clientId: apiParts[0].trim(),
            clientSecret: apiParts[1].trim()
        };

        if (sourceSpec.url && sourceSpec.clientId && sourceSpec.clientSecret) {
            return sourceSpec;
        } else {
            throw new Error('Empty fields in property value for key ' + key + ': ' + rawProperty);
        }
    } else {
        throw new Error('Expected 2 fields (CLIENT_ID, CLIENT_SECRET) in property value for key ' + key + ': ' + rawProperty);
    }
}
