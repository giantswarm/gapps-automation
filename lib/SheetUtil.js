/** Google SheetUtil related utility functions.
 *
 * Note: This file/namespace cannot be called "Sheets" because that conflicts with the Advanced Sheets service name.
 */

/** The seconds in a day (24 * 60 * 60). */
const SECONDS_PER_DAY = 86400;

/** Milliseconds per day (24 * 60 * 60 * 1000). */
const MILLISECONDS_PER_DAY = 86400000;


class SheetUtil {

    /** Excel DateTime (AKA SERIAL_NUMBER date) to JS Date.
     *
     * @param excelSerialDateTime A SheetUtil/Excel serial date like:
     *   44876.641666666605  ( 2022-11-11T15:24:00.000Z )
     * @param timeZoneOffsetMillies The source time zone offset from which to convert the serial date time (usually the one of the sheet).
     *
     * @return A Date object representing that point in time.
     */
    static serialDateTimeToDate(excelSerialDateTime, timeZoneOffsetMillies) {
        const timeZoneOffsetDays = timeZoneOffsetMillies / MILLISECONDS_PER_DAY;
        const excelSerialDateTimeUtc = excelSerialDateTime + (-1 * timeZoneOffsetDays);
        const totalDays = Math.floor(excelSerialDateTimeUtc);
        const totalSeconds = excelSerialDateTimeUtc * SECONDS_PER_DAY; // convert to fractional seconds, to avoid precision issues
        const millies = Math.round((totalSeconds - (totalDays * SECONDS_PER_DAY)) * 1000);
        return new Date(Date.UTC(0, 0, totalDays - 1, 0, 0, 0, millies));
    }


    /** JS Date to Excel DateTime (AKA SERIAL_NUMBER date).
     *
     * @param date A JavaScript date object, possibly representing a datetime like 2022-11-11T15:24:00.000Z..
     * @param timeZoneOffsetMillies The time zone offset of the target serial date time (usually the one of the sheet).
     *
     * @return A SheetUtil/Excel serial date like 44876.641666666605.
     */
    static dateToSerialDateTime(date, timeZoneOffsetMillies) {
        const timeZoneOffsetDays = timeZoneOffsetMillies / MILLISECONDS_PER_DAY;
        return ((date.getTime() / MILLISECONDS_PER_DAY) + 25569) + timeZoneOffsetDays; // 1970-01-01 - 1900-01-01 = 25569
    }


    /** Extended Value for CellData fields.
     *
     * See: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#extendedvalue
     *
     * @param {any} value Some JS primitive or object.
     * @param timeZoneOffsetMillies The time zone offset to convert serial date times (usually the one of the sheet).
     *
     * @return The ExtendedValue for writing to a spreadsheet.
     */
    static toExtendedValue(value, timeZoneOffsetMillies) {
        switch (typeof value) {
            case 'boolean':
                return {boolValue: value};
            case 'number': // fall-through
            case 'bigint':
                return {numberValue: value};
            case 'string':
                return {stringValue: value};
            case 'object':
                if (value instanceof Date) {
                    return {numberValue: SheetUtil.dateToSerialDateTime(value, timeZoneOffsetMillies)};
                }
            // fall-through
            default:  // null, undefined, object, symbol, function
                return {numberValue: null};
        }
    }


    /** Default CellFormat based on javascript type.
     *
     * See: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#cellformat
     *
     * @param value Some JS primitive or object.
     * @return {{numberFormat: {type: string}}|undefined} The default CellFormat for writes (if required).
     */
    static toDefaultCellFormat(value) {
        switch (typeof value) {
            case 'object':
                if (value instanceof Date) {
                    return {numberFormat: {type: 'DATE_TIME'}};
                }
            // fall-through
            default: // use Sheet defaults
                return undefined;
        }
    }


    /** Get the timezone offset in milliseconds.
     *
     * @param timeZone The time zone in text format, ie. "Europe/Paris"
     * @return {number} Time zone offset in milliseconds.
     */
    static getTimeZoneOffset(timeZone) {
        const strOffset = Utilities.formatDate(new Date(), timeZone, "Z");
        const offsetSeconds = ((+(strOffset.substring(0, 3))) * 3600) + ((+strOffset.substring(3)) * 60);
        return offsetSeconds * 1000;
    }


    /** Get or create a sheet inside the spreadsheet with the specified name.
     * @param {object} spreadsheet A spreadsheet object returned from a SpreadsheetApp call.
     * @param {string} name A sheet name for the existing or newly created sheet in spreadsheet.
     * @return {object} A new SpreadsheetApp Sheet object.
     */
    static ensureSheet(spreadsheet, name) {
        let sheet = spreadsheet.getSheetByName(name)
        if (!sheet) {
            sheet = spreadsheet.insertSheet();
            sheet.setName(name);
        }

        return sheet;
    }


    /** Transform a column header into a column name compatible with common DB engines.
     *
     * @param {string} columnHeader The column header to transform.
     * @returns {string} Returns a sanitized column name.
     */
    static sanitizeColumnName(columnHeader) {
        return (('' + columnHeader) || '').trim().toLowerCase().replaceAll(/[^a-z0-9]+/g, '_');
    }


    /** Map an array of rows to an array of objects.
     *
     * NOTE: The resulting record objects have a member $columnIndex with a map of properties to their column position.
     *
     * @param {array<array<any>>} rows Array of rows to map (including header row, column names will be lower-cased, converted to alphanumeric, unique identifiers).
     * @param {array<string>} columns Columns to include, if empty array or undefined all columns are mapped.
     * @returns {array<object>} Returns an array of objects containing the mapped row fields.
     */
    static mapRows(rows, columns = undefined) {
        if (!Array.isArray(rows) || rows.length < 2) {
            return [];
        }

        const fields = [];
        const $columnIndex = {};  // we track this for our users
        for (let i = 0; i < rows[0].length; ++i) {
            const header = rows[0][i];
            let suffix = '';
            let id;
            do {
                id = SheetUtil.sanitizeColumnName(header) + suffix;
                suffix = '' + (+suffix + 1);
            }
            while (fields.includes(id));

            if (!Array.isArray(columns) || !columns.length || columns.includes(id) || columns.includes(header)) {
                fields.push(id);
                $columnIndex[id] = i;
            } else {
                fields.push(null);
            }
        }

        return rows.slice(1, rows.length).map(row => {
            const record = {$columnIndex: $columnIndex};
            for (let i = 0; i < fields.length; ++i) {
                const property = fields[i];
                if (property != null) {
                    record[property] = row[i];
                }
            }
            return record;
        });
    }


    /** Get all rows of a specified Sheet inside the given spreadsheet.
     *
     * @param {array<string>} columns An array of columns to map or empty array to map all columns to objects,
     *                                no mapping performed if this parameter is undefined.
     * @return {array<array<any>>|array<object>} An array of rows or an array of mapped objects, depending on the arguments.
     */
    static getSheetData(spreadsheet, sheetName, columns = undefined) {
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            throw new Error('Sheet with name ' + sheetName + ' not found in spreadsheet');
        }

        const rows = sheet.getSheetValues(1, 1, -1, -1);
        return SheetUtil.mapRows(rows, columns);
    }


    /** Unique index rows by key column.
     *
     * @param rows An array of row objects as returned by SheetUtil.mapRows().
     * @param keyCol The column name to index by.
     * @return {object} A map of rows indexed by the specified key column.
     */
    static indexUnique(rows, keyCol) {
        const keyColSane = SheetUtil.sanitizeColumnName(keyCol);
        const map = {};
        if (Array.isArray(rows) && rows.length && rows[0].hasOwnProperty(keyColSane)) {
            for (const row of rows) {
                const key = (String(row[keyColSane]) || '').trim();
                if (key !== '') {
                    if (map.hasOwnProperty(key)) {
                        throw new Error('Duplicate key "' + key + '" found when indexing rows by column "' + keyCol + '"');
                    }
                    map[key] = row;
                }
            }
        }
        return map;
    }


    /** Round to specified number of decimal places using bankers rounding/round half to even.
     *
     * @param {number} value The number to round.
     * @param {number} decimalPlaces The number of decimal places to round to (default: 2).
     * @return {number} The rounded number.
     */
    static roundBankers(value, decimalPlaces = 2) {
        const f = Math.pow(10, decimalPlaces);
        const n = value * f, o = Math.round(n);
        return (Math.abs(n) % 1 === .5 ? o % 2 === 0 ? o : o - 1 : o) / f
    }

    /** Round up (positive numbers) or down (negative numbers) to specified number of decimal places.
     *
     * @param {number} value The number to round.
     * @param {number} decimalPlaces The number of decimal places to round to (default: 2).
     * @return {number} The rounded number.
     */
    static round(value, decimalPlaces = 2) {
        const f = Math.pow(10, decimalPlaces);
        const n = Number(value);
        return Math.round(n * f) / f;
    }

    /** Round up (positive numbers) or down (negative numbers) to specified number of decimal places.
     *
     * @param {number} value The number to round.
     * @param {number} decimalPlaces The number of decimal places to round to (default: 2).
     * @return {number} The rounded number.
     */
    static roundUp(value, decimalPlaces = 2) {
        const f = Math.pow(10, decimalPlaces);
        const n = Number(value);
        return n >= 0 ? Math.ceil(n * f) / f : Math.floor(n * f) / f;
    }
}


/**
 * Returns the current sheets timezone.
 *
 * @return The current sheets timezone in IANA time zone database name format (ie Europe/Berlin).
 * @customfunction
 */
function TIMEZONE() {
    return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
}


/**
 * Returns the current sheets timezone offset in milliseconds.
 *
 * @return The current sheets timezone offset in milliseconds.
 * @customfunction
 */
function TIMEZONE_OFFSET() {
    const tz = TIMEZONE();
    return tz != null ? SheetUtil.getTimeZoneOffset(tz) : null;
}


/**
 * Convert ISO8601 timestamp strings (ie. 2022-11-22T14:47:01+0100) to a Sheets serial datetime.
 *
 * @param {string|Array<Array<string>>} input Input ISO8601 date string to parse.
 *
 * @return {number} The native sheets "serial datetime" as double (format the field as Number->Date Time manually).
 * @customfunction
 */
function PARSE_ISO8601(input) {

    const tzOffsetMillies = TIMEZONE_OFFSET();
    const parseIso8601 = ts => ts ? SheetUtil.dateToSerialDateTime(new Date(ts), tzOffsetMillies) : null;

    return Array.isArray(input) ? input.map(row => row.map(field => parseIso8601(field))) : parseIso8601(input);
}
