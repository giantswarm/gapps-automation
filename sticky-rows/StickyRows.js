/**
 * Function to keep manually editable data around dynamic data aligned.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */


const TRIGGER_FUNCTION_NAME = 'onChangeStickyRows';


/**
 * Configure the stick rows feature.
 *
 * - store configuration for the current sheet
 * - enable/disable trigger
 *
 * @param {string} dynamic_range  The dynamic range (first column must be ID).
 * @param {string} sticky_range  The sticky range whose rows will follow the IDs from dynamic_range.
 * @customfunction
 */
function STICKY_ROWS(dynamic_range, sticky_range) {

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
        throw new Error("Must run in the context of a Spreadsheet");
    }

    const sheet = spreadsheet.getActiveRange().getSheet();

    // setup config for this sheet in the spreadsheet attributes
    const key = "stickyRows." + sheet.getSheetId();
    const config = JSON.parse(PropertiesService.getDocumentProperties().getProperty(key) || '{}');
    if (config.dynamic_range !== dynamic_range || config.sticky_range !== sticky_range) {
        config.dynamic_range = dynamic_range;
        config.sticky_range = sticky_range;
        PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(config));
    }

    // ensure onChange() trigger
    let haveTrigger = false;
    for (const trigger of ScriptApp.getProjectTriggers()) {
        if (trigger.getEventType() === ScriptApp.EventType.ON_CHANGE && trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
            haveTrigger = true;
        }
    }

    if (!haveTrigger) {
        throw new Error(`Please install function ${TRIGGER_FUNCTION_NAME}() as trigger in this Spreadsheet's script context (event source: From spreadsheet, event type: onChange)`)
    }
}


/** Uninstall triggers. */
function uninstall() {
    TriggerUtil.uninstall(TRIGGER_HANDLER_FUNCTION);
}


/** Install spreadsheet onChange trigger. */
function install() {

    TriggerUtil.uninstall(TRIGGER_HANDLER_FUNCTION);

    Logger.log("Installing onChange trigger for spreadsheet %s", spreadsheet.getId());

    const sheet = SpreadsheetApp.getActive();
    ScriptApp.newTrigger(TRIGGER_HANDLER_FUNCTION)
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId())
        .onChange()
        .create();
    Logger.log("Installed onChange trigger for spreadsheet %s", sheet.getSpreadsheet());
}


/** Perform the actual sticky rows logic.
 *
 * Supposed to be called in a SpreadSheet.onChange() trigger.
 */
function onChangeStickyRows(event) {

    const spreadsheet = event.source;
    if (!spreadsheet) {
        throw new Error("Must run in the context of a Spreadsheet");
    }

    const sheet = spreadsheet.getActiveRange().getSheet();
    const key = "stickyRows." + sheet.getSheetId();
    const config = JSON.parse(PropertiesService.getDocumentProperties().getProperty(key) || '{}');
    if (!config.dynamic_range || !config.sticky_range) {
        throw new Error("dynamic_range and sticky_range must be set");
    }

    const dynamic = sheet.getRange(config.dynamic_range);
    const sticky = sheet.getRange(config.sticky_range);

    const rowCount = sticky.getNumRows();
    if (dynamic.getNumRows() !== rowCount) {
        throw new Error("dynamic/sticky ranges must have the same number of rows");
    }

    if (dynamic.getColumn() < sticky.Right && dynamic.getColumn() + dynamic.getNumColumns() > sticky.getColumn() &&
        dynamic.getRow() > sticky.getRow() + rowCount && dynamic.getRow() + rowCount < sticky.getRow()) {
        throw new Error("dynamic/sticky data must not overlap");
    }

    const dynamicIds = sheet.getSheetValues(dynamic.getRow(), dynamic.getColumn(), rowCount, 1);
    const stickyRows = readRange_(sticky);

    // index sticky row indices by meta id
    let metaIdToStickyRowIndex = null;
    let modified = false;
    for (let i = 0; i < rowCount; ++i) {
        const id = dynamicIds[i][0];
        const metaId = stickyRows[i][0];
        if (id != null && id !== '' && id !== metaId) {
            if (!metaIdToStickyRowIndex) {
                // lazily index sticky rows metaId column
                metaIdToStickyRowIndex = buildIndex_(stickyRows);
            }
            const j = metaIdToStickyRowIndex.get(id);
            if (j !== undefined) {
                const row = stickyRows[j];
                stickyRows[j] = stickyRows[i];
                stickyRows[i] = row;
                if (metaId != null && metaId !== '') {
                    metaIdToStickyRowIndex.set(metaId, j);
                }
                metaIdToStickyRowIndex.delete(id);
                modified = true;
            } else {
                stickyRows[i][0] = id;
                modified = true;
            }
        }
    }

    if (modified) {
        sticky.setValues(stickyRows);
    }
}


/** Build a hash index (ES6 Map key to position) from the first column of the provided rows array.
 */
function buildIndex_(rows) {
    const rowCount = rows.length;
    const index = new Map();
    for (let i = 0; i < rowCount; ++i) {
        index.set(rows[i][0], i);
    }
    index.delete(null);
    index.delete('');

    return index;
}


/** Read one SpreadsheetApp Range object's cell values and formulas into an array usable with Range.setValues().
 *
 * @param {object} range The source Range to read from.
 * @return {array<array<string>>} Two dimensional array (rows with values).
 */
function readRange_(range) {
    const values = range.getValues();
    const rowCount = values.length;
    const relativeFormulas = range.getFormulasR1C1();
    const combined = [];
    combined.length = rowCount;
    for (let i = 0; i < rowCount; ++i) {
        const valueRow = values[i];
        const rowLength = valueRow.length;
        const formulaRow = relativeFormulas[i];
        const row = [];
        row.length = rowLength;
        for (let j = 0; j < rowLength; ++j) {
            if (formulaRow[j] !== '') {
                row[j] = formulaRow[j];
            } else {
                row[j] = valueRow[j];
            }
        }
        combined[i] = row;
    }

    return combined;
}
