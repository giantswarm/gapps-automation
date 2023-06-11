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
        console.log('Trigger:', trigger.getEventType().name(), trigger.getHandlerFunction());
        if (trigger.getEventType() === ScriptApp.EventType.ON_CHANGE && trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
            haveTrigger = true;
        }
    }

    if (!haveTrigger) {
        throw new Error(`Please install function ${TRIGGER_FUNCTION_NAME}() as trigger in this Spreadsheet's script context (event source: From spreadsheet, event type: onChange)`)
    }
}


/** Perform the actual sticky rows logic.
 *
 * Supposed to be called in a SpreadSheet.onChange() trigger.
 */
function onChangeStickyRows(event) {

    if (event.changeType !== 'EDIT') {
        // we only react to real user edits
        console.log(`Won't handle event type: ${event.changeType}`);
        return;
    }

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

    const dynamicIds = sheet.getSheetValues(dynamic.getRow(), dynamic.getColumn(), rowCount, 1).flat();
    const metaIds = sheet.getSheetValues(sticky.getRow(), sticky.getColumn(), rowCount, 1).flat();

    for (let i = 0; i < rowCount; ++i) {
        const id = dynamicIds[i];
        const metaId = metaIds[i];
        if (id != null && id !== '' && id !== metaId) {

            // TODO replace this with map index + lookup?
            let j;
            for (j = 0; j < rowCount && metaIds[j] !== id; ++j) {
            }

            if (j < rowCount) {
                const a = sheet.getRange(sticky.getRow() + i, sticky.getColumn(), 1, sticky.getNumColumns());
                const b = sheet.getRange(sticky.getRow() + j, sticky.getColumn(), 1, sticky.getNumColumns());
                const dataA = readRange_(a);
                a.setValues(readRange_(b));
                b.setValues(dataA);
                metaIds[i] = metaIds[j];
                metaIds[j] = metaId;
            } else if (metaId == null || metaId === '') {
                console.log(`[${i}]: init`);
                sheet.getRange(sticky.getRow() + i, sticky.getColumn(), 1, 1).setValue(id);
                metaIds[i] = id;
            }
        }
    }
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
