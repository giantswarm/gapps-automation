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
 * @param {string} dynamic_range The dynamic range (first column must be ID).
 * @param {string} sticky_range  The sticky range whose rows will follow the IDs from dynamic_range.
 * @param {number} toggle        A random number used to control re-calculation of this function. Choose any integer.
 * @customfunction
 */
function STICKY_ROWS(dynamic_range, sticky_range, toggle) {

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
        throw new Error("Must run in the context of a Spreadsheet");
    }

    const activeRange = spreadsheet.getActiveRange();
    const sheet = activeRange.getSheet();

    const argsValid = (dynamic_range === null || typeof dynamic_range === 'string')
        && (sticky_range === null || typeof sticky_range === 'string')
        && toggle !== undefined;

    // setup config for this sheet in the spreadsheet attributes
    const key = "stickyRows." + sheet.getSheetId();
    const config = mergeConfig_(key, {
        dynamic_range: dynamic_range,
        sticky_range: sticky_range,
        controlCell: argsValid ? activeRange.getA1Notation() : ''
    });

    if (!argsValid) {
        // We must always update config.controlCell first, to ensure onChangeStickyRows() doesn't overwrite
        // what was previously a control cell with a formula but may have been modified.
        throw new Error("Please specify all three arguments: dynamic_range (string), sticky_range (string) and some random integer as third argument");
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

    if (config.error) {
        throw new Error(config.error);
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

    const spreadsheet = event.source || SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
        throw new Error("Must run in the context of a Spreadsheet");
    }

    let sheet;
    try {
        sheet = spreadsheet.getActiveRange().getSheet();
    } catch (e) {
        sheet = spreadsheet.getActiveSheet();
    }

    const key = "stickyRows." + sheet.getSheetId();
    const config = mergeConfig_(key, {});
    if (!config.dynamic_range || !config.sticky_range || !config.controlCell) {
        // STICKY_ROWS() also reports this error, we can't forward since we know no control cell here
        throw new Error("Please call STICK_ROWS() with a valid dynamic_range and sticky_range in a formula on the sheet");
    }

    try {
        realignStickyRows_(sheet, config);
        if (config.error) {
            forceRecalculation_(sheet, mergeConfig_(key, {error: ''}));
        }
    } catch (e) {
        // store error message for displaying by STICKY_ROWS()
        const nextError = '' + e;
        if (nextError !== config.error) {
            forceRecalculation_(sheet, mergeConfig_(key, {error: nextError}));
        }
        throw e;
    }
}


/** Read, merge and update config properties.
 *
 * @return The updated config for the given key (sheet).
 */
function mergeConfig_(key, configProperties) {
    const config = JSON.parse(PropertiesService.getDocumentProperties().getProperty(key) || '{}');
    let changed = false;
    for (const property in configProperties) {
        if (config[property] !== configProperties[property]) {
            config[property] = configProperties[property];
        }
        changed = true;
    }

    if (changed) {
        PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(config));
    }

    return config;
}


/** Logic to re-align rows in the sticky area (defined in config) within sheet. */
function realignStickyRows_(sheet, config) {
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

    // index sticky row indices by id
    let idToStickyRowIndex = null;
    let modified = false;
    for (let i = 0; i < rowCount; ++i) {
        const id = dynamicIds[i][0];
        const stickyId = stickyRows[i][0];
        if (id != null && id !== '' && id !== stickyId) {
            if (!idToStickyRowIndex) {
                // lazily index
                idToStickyRowIndex = buildIndex_(stickyRows);
            }
            const j = idToStickyRowIndex.get(id);
            if (j !== undefined) {
                const row = stickyRows[j];
                stickyRows[j] = stickyRows[i];
                stickyRows[i] = row;
                if (stickyId != null && stickyId !== '') {
                    idToStickyRowIndex.set(stickyId, j);
                }
                idToStickyRowIndex.delete(id);
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


/** Force re-calculation of the control cell (by reading/writing the formula).
 */
function forceRecalculation_(sheet, config) {
    try {
        // force re-calculation of the formula to show the error
        if (config.controlCell) {
            const range = sheet.getRange(config.controlCell);
            const formula = range.getFormulaR1C1();
            if (formula !== '') {
                // Notes:
                // - Google Sheets analyzes formulas and ignores white-space changes in API setFormula() calls
                // - Formulas can have different argument separators (;,) depending on locale
                // - Boolean values are represented by different words based on locale (TRUE, WAHR, ...)
                // - There seems to be logic breaking execution cycles (cell change -> trigger -> cell change -> trigger)
                const functionName = 'STICKY_ROWS(';
                const iArgsBegin = formula.indexOf('STICKY_ROWS(') + functionName.length;
                const iArgsEnd = formula.indexOf(')', iArgsBegin);
                const iArg3Begin = Math.max(formula.lastIndexOf(',', iArgsEnd), formula.lastIndexOf(';', iArgsEnd));
                if (iArg3Begin >= 0 && iArgsEnd > iArg3Begin + 1) {
                    const arg3 = +(formula.slice(iArg3Begin + 1, iArgsEnd).trim()) || 0;
                    const nextArg3 = arg3 < 65535 ? arg3 + 1 : 0;
                    range.setFormula('' + formula.slice(0, iArg3Begin + 1) + nextArg3 + formula.slice(iArgsEnd));
                    return true;
                }
            }
        }
    } catch (e) {
        console.log("failed to force re-calculation: ", e);
    }
    return false;
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
