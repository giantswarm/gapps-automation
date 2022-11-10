/**
 * Dump Personio Context into a Sheet
 *
 *   - The source URL, tokens and target sheets for sync process
 *     must be specified as ScriptProperties (see setProperties() helper function).
 *
 *     FORMAT:
 *          Script Property Key: PersonioDump.SHEET_ID
 *          Script Property Value: FULL_PERSONIO_API_URL|PERSONIO_CLIENT_ID|PERSONIO_CLIENT_SECRET
 *
 *     EXAMPLE:
 *         "PersonioDump.1wX-VnjLVkBL74SC-8qMt_oeib4VGMlDpZuJzrd_NZUE": "/company/employees.json|lkjklasdj|lkjakasd"
 *
 *   - The target sheet must be accessible by the account running the script, for example:
 *
 *     automation@giantswarm.io
 *
 * This script uses the Advanced Sheets Service, make sure to enable it in the Google Cloud Project.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
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
            continue;
        }

        let relations = null;
        try {
            relations = transformPersonioDataToRelations_(data);
        } catch (e) {
            Logger.log('Failed to transform Personio data for sheet %s: %s', task.spreadsheetId, e.message);
            firstError = firstError || e;
            continue;
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


/** Sanitize delay minutes input.
 *
 * ClockTriggerBuilder supported only a limited number of values, see:
 * https://developers.google.com/apps-script/reference/script/clock-trigger-builder#everyMinutes(Integer)
 */
function sanitizeDelayMinutes_(delayMinutes) {
    return [1, 5, 10, 15, 30].reduceRight((v, prev) =>
        typeof +delayMinutes === 'number' && v <= +delayMinutes ? v : prev);
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
 *   clasp run 'setProperties' --params '[{"PersonioDump.SHEET_ID": "SOME_PERSONIO_URL|CLIENT_ID|CLIENT_SECRET"}, false]'
 *
 * Warning: passing argument true for parameter deleteAllOthers will also cause the schema to be reset!
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
        if (!safeKey.startsWith(PROPERTY_PREFIX) || safeKey === PROPERTY_PREFIX + 'schema')
            continue;

        const spreadsheetId = safeKey.replace(PROPERTY_PREFIX, '');
        if (!spreadsheetId) {
            continue;
        }

        try {
            const rawProperty = properties[key] || '';
            const apiParts = rawProperty.trim().split('|');

            if (apiParts.length === 3) {
                const sourceSpec = {
                    url: apiParts[0].trim(),
                    clientId: apiParts[1].trim(),
                    clientSecret: apiParts[2].trim()
                };

                if (sourceSpec.url && sourceSpec.clientId && sourceSpec.clientSecret) {
                    tasks.push({spreadsheetId: spreadsheetId, source: sourceSpec});
                } else {
                    Logger.log("Skipped task: Empty fields in property value for key %s: %s", key, rawProperty);
                }
            } else {
                Logger.log("Skipped task: Expected 3 fields (URL, CLIENT_ID, CLIENT_SECRET) in property value for key %s: %s", key, rawProperty);
            }
        } catch (e) {
            Logger.log('Skipped task: Incorrect API config for property key %s: %s', key, e.message);
        }
    }

    return tasks;
}


/** Transform objects from Persionio API v1 response data to normalized relations.
 *
 *  Output Example:
 *
 *  {
 *      Employee: { headers: { first_name: {title: "First Name"}}, rows: [["First Name"], ["Jonas"]]},
 *      Department: { ... }
 *  }
 *
 */
function transformPersonioDataToRelations_(data) {

    const relations = {version: 0};  // ie. tables in a relational schema

    const calculateUtf8Size = str => {
        // returns the byte length of an utf8 string
        let s = str.length;
        for (let i = str.length - 1; i >= 0; i--) {
            let code = str.charCodeAt(i);
            if (code > 0x7f && code <= 0x7ff) s++;
            else if (code > 0x7ff && code <= 0xffff) s += 2;
            if (code >= 0xDC00 && code <= 0xDFFF) i--; //trail surrogate
        }
        return s;
    };

    const loadSchema = () => {
        const existingSchema = PropertiesService.getScriptProperties().getProperty(PROPERTY_PREFIX + 'schema');
        if (existingSchema) {
            const schema = JSON.parse(existingSchema);
            // copy properties
            for (const relType in schema) {
                relations[relType] = schema[relType];
            }

            Logger.log("Loaded schema: version=%s, size=%s (max 9216)", relations.version, calculateUtf8Size(existingSchema));
        } else {
            Logger.log("No existing schema found");
        }
    };

    const saveSchema = () => {
        // Persist updated schema version
        const schema = {version: relations.version};
        for (const relType in schema) {
            schema[relType] = {headers: relations[relType].headers};
        }

        const updatedSchema = JSON.stringify(relations);
        Logger.log("Saving schema: version=%s, size=%s (max 9216)", relations.version, calculateUtf8Size(updatedSchema));
        PropertiesService.getScriptProperties().setProperty(PROPERTY_PREFIX + 'schema', updatedSchema);
    };

    const hasParentOfType = (parents, relType) => !!parents.find(parent => parent.type === relType);

    const unboxValue = boxedValue => boxedValue?.value !== undefined ? boxedValue.value : boxedValue;

    const getItemId = (item) => {
        if (item?.attributes) {
            return unboxValue(item.attributes['id']);
        }
        // null/undefined: this item has no ID (ie. item == null)
        return null;
    };

    // scan an object, recursively
    // returns the item type (relation) if the object was handled, null otherwise
    const scanObject = (item, parents) => {

        // map attribute labels to columns
        const attributes = item?.attributes;
        const itemType = item?.type;
        if (!Util.isObject(attributes) || !itemType)
            throw new Error(`Unknown object without type or attributes: ${JSON.stringify(item)}`);

        if (hasParentOfType(parents, itemType))
            return itemType;

        if (!relations[itemType]) {
            relations[itemType] = {headers: {}, rows: [], ids: {}};
        }

        const relation = relations[itemType];

        for (const id in attributes) {

            let header = relation.headers[id];
            if (header && header.v === relations.version) {
                // header already scanned
                continue;
            }

            const attribute = attributes[id];
            const value = unboxValue(attribute);

            const values = Array.isArray(value) ? value : [value];
            for (const unpackedValue of values) {
                if (unpackedValue === null || unpackedValue === undefined) {
                    // skip or register later
                    // NOTE: This will drop columns/relations with all values set to NULL
                    continue;
                }

                // We support varying values in arrays, but that shouldn't occur
                let foreignType = null;
                if (Util.isObject(unpackedValue)) {
                    foreignType = scanObject(unpackedValue, parents.concat([item]));
                    if (!foreignType) {
                        continue;
                    }
                    // nested, possibly circular type was registered
                }

                // add column header
                const label = (attribute?.label || '').trim();
                const uniform_id = (attribute?.uniform_id || '').trim();
                if (!header) {
                    // new column found
                    relation.headers[id] = header = {};
                }
                header.title = (label || uniform_id || id.trim()) + (foreignType ? '_' + foreignType + '_id' : '');
                header.foreignType = foreignType != null ? foreignType : undefined;
                header.v = relations.version;
            }
        }

        return itemType;
    };

    // convert values recursively
    const convertObject = (item, parents) => {

        const itemType = item?.type;
        const relation = relations[itemType];
        if (!relation)
            return;

        if (hasParentOfType(parents, itemType))
            return;

        // filter out duplicate rows by ID
        const itemId = getItemId(item);
        if (itemId != null) {
            if (relation.ids[itemId]) {
                return;
            } else {
                relation.ids[itemId] = true;
            }
        }

        const row = [];
        for (const column in relation.headers) {
            const header = relation.headers[column];
            const value = unboxValue(item.attributes[column]);

            const values = Array.isArray(value) ? value : [value];
            let fields = [];
            for (const unpackedValue of values) {
                if (header.foreignType) {
                    // nested type, just store id and row in a separate relation
                    const itemId = getItemId(unpackedValue);
                    fields.push(itemId);
                    convertObject(unpackedValue, parents.concat([item]));
                } else {
                    fields.push(unpackedValue);
                }
            }

            // plain value null/undefined -> ''
            // arrays joined by ',' (comma)
            row.push(fields.map(v => v == null ? '' : v).join(','));
        }

        relation.rows.push(row);
    };

    // #1 Build schema (scan data returned from API)
    loadSchema();
    ++relations.version;
    for (const item of data) {
        scanObject(item, []);
    }
    // Persist updated schema version
    saveSchema();

    // #2 Set headers (first rows for each relation)
    for (const relType in relations) {
        if (relType === 'version')
            continue;

        const relation = relations[relType];
        relation.rows.push(Object.values(relation.headers).map(header => header.title));
    }


    // #3 Convert data (rows)
    for (const item of data) {
        convertObject(item, []);
    }

    return relations;
}


/** Creates a "custom" sheet with the specified name (if it doesn't exist) and returns the sheet properties. */
function getOrAddSheet_(spreadsheetId, sheetTitle) {
    const spreadsheet = Sheets.Spreadsheets.get(spreadsheetId);
    if (!spreadsheet)
        throw new Error(`Specified spreadsheet ${spreadsheetId} does not exist or is not accessible`);

    const existingSheet = spreadsheet.sheets.find(sheet => sheet.properties.title === sheetTitle);
    if (existingSheet) {
        return existingSheet.properties;
    }

    const batch = {
        requests: [{
            addSheet: {
                properties: {
                    title: sheetTitle
                }
            }
        }]
    };
    const response = Sheets.Spreadsheets.batchUpdate(batch, spreadsheetId);
    return response.replies[0].addSheet.properties;
}


/** Write relations structure to the specified spreadsheet, creating the necessary sheets. */
function writeRelationsToSheet_(spreadsheetId, relations, valueInputOption) {

    // Extended Value for CellData fields see:
    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#extendedvalue
    const toExtendedValue = (value) => {
        switch (typeof value) {
            case 'boolean':
                return {boolValue: value};
            case 'number': // fall-through
            case 'bigint':
                return {numberValue: value};
            case 'string':
                return {stringValue: value};
            default:  // null, undefined, object, symbol, function
                return {errorValue: {type: 'NULL_VALUE'}};
        }
    };


    // Prepare Sheets API requests in advance
    const batch = {requests: []};

    for (const [relType, relation] of Object.entries(relations)) {

        if (relType === 'version')
            continue;

        const sheetProperties = getOrAddSheet_(spreadsheetId, relType);
        const sheetId = sheetProperties.sheetId;
        const rowCount = sheetProperties.gridProperties.rowCount;
        const columnCount = sheetProperties.gridProperties.columnCount;
        const targetRowCount = relation.rows.length;
        const targetColumnCount = relation.rows.length ? relation.rows[0].length : 0

        // #1 Ensure correct dimensions (enough rows/columns to insert data)
        if (targetRowCount > rowCount) {
            batch.requests.push({
                appendDimension: {
                    sheetId: sheetId,
                    dimension: 'ROWS',
                    length: targetRowCount - rowCount
                }
            });
        }
        if (targetColumnCount > columnCount) {
            batch.requests.push({
                appendDimension: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    length: targetColumnCount - columnCount
                }
            });
        }

        // #2 Clear the whole worksheet, preserving formats
        batch.requests.push(
            {
                updateCells: {
                    range: {
                        sheetId: sheetId
                    },
                    fields: 'userEnteredValue'
                }
            });

        // #3 Insert the data rows
        batch.requests.push(
            {
                updateCells: {
                    range: {
                        sheetId: sheetId,
                        startRowIndex: 0,
                        endRowIndex: relation.rows.length,
                        startColumnIndex: 0,
                        endColumnIndex: relation.rows.length ? relation.rows[0].length : 0
                    },
                    fields: 'userEnteredValue',
                    rows: relation.rows.map(row => ({
                        values: row.map(v => ({
                            userEnteredValue: toExtendedValue(v)
                        }))
                    }))
                }
            });
    }

    Sheets.Spreadsheets.batchUpdate(batch, spreadsheetId);
}
