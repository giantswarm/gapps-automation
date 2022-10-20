/**
 * Dump Personio Context into a Sheet
 *
 *   - The source URL, token and target sheets for sync process
 *     must be specified as ScriptProperties (also see setProperties()).
 *
 *     FORMAT:
 *          Script Property Key: SHEET_ID
 *          Script Property Value: FULL_PERSONIO_API_URL|PERSONIO_CLIENT_ID|PERSONIO_CLIENT_SECRET
 *
 *     EXAMPLE:
 *         "personio-dump.1wX-VnjLVkBL74SC-8qMt_oeib4VGMlDpZuJzrd_NZUE": "/company/employees|lkjklasdj|lkjakasd|"
 *
 *   - The target sheet must be accessible (shared with) the account running the script, for example:
 *
 *     personio-sync@gapps-automation-2022.iam.gserviceaccount.com
 *
 * This script uses the Advanced Sheets Service (make sure to enable it in the Google Cloud Project).
 */


const VALUE_INPUT_OPTIONS = {RAW: 'RAW', USER_ENTERED: 'USER_ENTERED'};

/** How to store values in sheet fields (with or without parsing). */
const DEFAULT_VALUE_INPUT_OPTION = VALUE_INPUT_OPTIONS.RAW;

/** The prefix for properties specific to this script in the project.
 *
 * Making this dynamic didn't work as DriveApp.getFileById() tends to throw 500s and is otherwise quite slow.
 */
const PROPERTY_PREFIX = 'personio-dump.';


/** Main entry point.
 *
 * Take configuration from ScriptProperties and perform synchronization.
 */
function sync() {

    let firstError = null;

    // we keep operating if a single sync task fails
    for (const task of getTasks_()) {

        const personio = new PersonioClientV1(task.source.clientId, task.source.clientSecret);

        let data = null;
        try {
            data = personio.fetch(task.source.url);
        } catch (e) {
            Logger.log('Failed to fetch Personio data for sheet %s: %s', task.spreadsheetId, e.message);
            firstError = firstError || e;
        }

        // TODO Handle some special data/formats (for example lists, images or other blobs)?
        let relations = null;
        try {
            relations = transformPersonioDataToRelations_(data);
        } catch (e) {
            Logger.log('Failed to write rows to sheet %s: %s', task.spreadsheetId, e.message);
            firstError = firstError || e;
        }

        try {
            writeRelationsToSheet_(task.spreadsheetId, relations, DEFAULT_VALUE_INPUT_OPTION);
        } catch (e) {
            Logger.log('Failed to write rows to sheet %s: %s', task.spreadsheetId, e.message);
            firstError = firstError || e;
        }
    }

    if (firstError) {
        throw firstError;
    }
}


/** Helper function to configure the required script properties.
 *
 * USAGE EXAMPLE:
 *   clasp run 'setProperties' --params '[{"persionio-dump.SHEET_ID": "SOME_PERSONIO_URL|CLIENT_ID|CLIENT_SECRET"}, true]'
 */
function setProperties(properties, deleteAllOthers) {
    PropertiesService.getScriptProperties().setProperties(properties, deleteAllOthers);
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

        // TODO Extract sheet ID from sheets URL if such is present
        const spreadsheetId = safeKey.replace(PROPERTY_PREFIX, '');
        if (!spreadsheetId) {
            continue;
        }

        try {
            // TODO Validate object and output proper log messsage
            const apiParts = properties[key].trim().split('|');
            const sourceSpec = {
                url: apiParts[0].trim(),
                clientId: apiParts[1].trim(),
                clientSecret: apiParts[2].trim()
            };

            tasks.push({spreadsheetId: spreadsheetId, source: sourceSpec});
        } catch (e) {
            Logger.log('Incorrect API config for property key %s: %s', key, e.message);
        }
    }

    return tasks;
}


/** Transform objects contained in Persionio API response data.
 *
 *  Output Example:
 *
 *  {
 *      Employee: { headers: { first_name: "First Name"}, rows: [["First Name"], ["Jonas"]]},
 *      TimeOff: { ... }
 *  }
 *
 */
function transformPersonioDataToRelations_(data) {

    const relations = {};  // ie. types

    for (const item of data) {

        // make multiple relations (by field type) support optional?
        if (!item || !item.type)
            continue;

        // TODO Implement more fields here?
        // map attribute labels to columns
        const attrs = item.attributes;
        if (!attrs)
            continue;

        if (relations[item.type] === undefined) {
            relations[item.type] = {headers: {}, rows: []};
        }

        const relation = relations[item.type];

        for (const id in attrs) {
            if (!relation.headers[id]) {
                // track columns
                // TODO Can we guarantee they are always the same for each object?
                const attr = attrs[id];
                const label = (attr.label || '').trim();
                const uniform_id = (attr.uniform_id || '').trim();
                relation.headers[id] = label || uniform_id || id.trim();
            }
        }
    }

    // first rows for each relation (headers)
    for (const relType in relations) {
        const relation = relations[relType];
        relation.rows.push(Object.values(relation.headers));
    }

    // TODO Are there more field values (except attributes)?
    for (const item of data) {

        if (!item || !item.type)
            continue;

        const relation = relations[item.type];
        if (!relation)
            continue;

        const row = [];
        for (const header in relation.headers) {
            row.push(item.attributes[header]?.value);
        }
        relation.rows.push(row);
    }

    return relations;
}


/** Creates a "custom" sheet with the specified name (if it doesn't exist) and returns the sheetId. */
function addOrGetSheet_(spreadsheetId, sheetTitle) {
    try {
        const requests = [{
            'addSheet': {
                'properties': {
                    'title': sheetTitle
                }
            }
        }];

        const response = Sheets.Spreadsheets.batchUpdate({'requests': requests}, spreadsheetId);
        return response.replies[0].addSheet.properties.sheetId;
    } catch (e) {
        if (e.details.code === 400) {
            try {
                const existingSheet = Sheets.Spreadsheets.get(spreadsheetId).sheets.find(sheet => sheet.properties.title === sheetTitle);
                if (existingSheet) {
                    return existingSheet.sheetId;
                }
            } catch (e1) {
                Logger.log('Failed to lookup existing sheet: %s', e.message);
            }
        }
        Logger.log('Failed to create new sheet: %s', e.message);
        throw e;
    }
}


function writeRelationsToSheet_(spreadsheetId, relations, valueInputOption) {

    // Store each relation in a corresponding sheet
    const request = {
        'valueInputOption': valueInputOption,
        'data': Object.entries(relations).map(([relType, relation]) => {
            // add sheet or use existing
            addOrGetSheet_(spreadsheetId, relType);
            return {
                range: "'" + relType + "'",
                majorDimension: 'ROWS',
                values: relation.rows
            }
        })
    };

    // TODO Do something with the reponse? Log number of changed cells?
    Sheets.Spreadsheets.Values.batchUpdate(request, spreadsheetId);
}
