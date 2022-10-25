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
 *         "PersonioDump.1wX-VnjLVkBL74SC-8qMt_oeib4VGMlDpZuJzrd_NZUE": "/company/employees.json|lkjklasdj|lkjakasd|"
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
const PROPERTY_PREFIX = 'PersonioDump.';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'dumpPersonio';


/** Main entry point.
 *
 * Take configuration from ScriptProperties and perform synchronization.
 */
function dumpPersonio() {

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


/** Setup for periodic execution and do some checks. */
function install(delayMinutes) {
    uninstall();

    Logger.log("Installing time based trigger: %s", delayMinutes);
    const delay = delayMinutes || 30;

    ScriptApp.newTrigger(TRIGGER_HANDLER_FUNCTION)
        .timeBased()
        .everyMinutes(delay)
        //.everyHours(1)
        .create();

    Logger.log("Installed time based trigger for %s every %s minutes", TRIGGER_HANDLER_FUNCTION, delay);
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

    const relations = {};  // ie. tables in a relational schema

    const hasParentOfType = (parents, relType) => !!parents.find(parent => parent.type === relType);

    const getItemType = (item) => item?.type;

    const getItemId = (item) => {
        if (item?.attributes) {
            const idOrObject = item.attributes['id'];
            return Util.isObject(idOrObject) ? idOrObject.value : idOrObject;
        }
        return '';
    };

    // scan an object, recursively, returns true if object was handled, false otherwise
    // TODO Are there more field values (except attributes)?
    const scanObject = (item, parents) => {

        // map attribute labels to columns
        const attrs = item?.attributes;
        const itemType = getItemType(item);
        if (!Util.isObject(attrs) || !itemType)
            return undefined;  // can't handle in a meaningful way (has no attributes or type)

        if (hasParentOfType(parents, itemType))
            return itemType;

        if (!relations[itemType]) {
            relations[itemType] = {headers: {}, rows: [], ids: {'': false}};
        }

        const relation = relations[itemType];

        for (const id in attrs) {
            if (!relation.headers[id]) {
                const attr = attrs[id];
                // track nested types
                const value = attr && attr.value !== undefined ? attr.value : attr;

                const values = Array.isArray(value) ? value : [value];
                for (const unpackedValue of values)
                {
                    if (unpackedValue === null || unpackedValue === undefined) {
                        // skip or register later
                        // TODO This will drop columns/relations with all values set to NULL
                        // TODO This could lead to schema inconsistencies without schema LOAD/SAVE support
                        continue;
                    }

                    // TODO We support only uniform values in arrays
                    let foreignType = null;
                    if (Util.isObject(unpackedValue)) {
                        foreignType = scanObject(unpackedValue, parents.concat([item]));
                        if (!foreignType) {
                            continue;
                        }

                        // nested (possibly circular object)
                    }

                    // track columns
                    const label = (attr?.label || '').trim();
                    const uniform_id = (attr?.uniform_id || '').trim();
                    relation.headers[id] = {
                        title: (label || uniform_id || id.trim()) + (foreignType ? '_' + foreignType + '_id' : ''),
                        foreignType: foreignType
                    };
                }
            }
        }

        return itemType;
    };

    // convert values recursively, returns true if item was handled, false if not
    // TODO Are there more field values (except attributes)?
    const convertObject = (item, parents) => {

        const itemType = getItemType(item);
        const relation = relations[itemType];
        if (!relation)
            return undefined;

        if (hasParentOfType(parents, itemType))
            return itemType;

        // filter out duplicate rows by ID
        const itemId = getItemId(item);
        if (relation.ids[itemId]) {
            return itemType;
        } else {
            relation.ids[itemId] = true;
        }

        const row = [];
        for (const column in relation.headers) {
            const header = relation.headers[column];
            const attr = item.attributes[column];
            const value = attr && attr.value !== undefined ? attr.value : attr;

            const values = Array.isArray(value) ? value : [value];
            let field = [];
            for (const unpackedValue of values) {
                if (header.foreignType) {
                    // TODO Should we push "name" or smth else or just the foreign key?
                    field.push(getItemId(unpackedValue));

                    // nested type, in case of loop just store id
                    convertObject(unpackedValue, parents.concat([item]));
                } else {
                    // plain value, null/undefined -> ''
                    field.push(unpackedValue === null || unpackedValue === undefined ? '' : unpackedValue);
                }
            }

            row.push(field.join(','));
        }

        relation.rows.push(row);

        return itemType;
    };

    // #1 Build schema (scan data returned from API)
    // TODO LOAD schema support in ScriptProperty? (To be able to automatically version control a stable schema)
    for (const item of data) {
        scanObject(item, []);
    }
    // TODO STORE schema support in ScriptProperty? (To be able to automatically version control a stable schema)

    // #2 Set headers (first rows for each relation)
    for (const relType in relations) {
        const relation = relations[relType];
        relation.rows.push(Object.values(relation.headers).map(header => header.title));
    }

    // #3 Convert data (rows)
    for (const item of data) {
        convertObject(item, []);
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
