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


    /** Get the the timezone offset in milliseconds.
     *
     * @param timeZone The time zone in text format, ie. "Europe/Paris"
     * @return {number} Time zone offset in milliseconds.
     */
    static getTimeZoneOffset(timeZone) {
        const strOffset = Utilities.formatDate(new Date(), timeZone, "Z");
        const offsetSeconds = ((+(strOffset.substring(0, 3))) * 3600) + ((+strOffset.substring(3)) * 60);
        return offsetSeconds * 1000;
    }


    /** Get a spreadsheet by ID. */
    static getSpreadsheet(id) {
        const spreadsheet = Sheets.Spreadsheets.get(id);
        if (!spreadsheet)
            throw new Error(`Specified spreadsheet ${id} does not exist or is not accessible`);

        return spreadsheet;
    }

    /** Creates a "custom" sheet with the specified name (if it doesn't exist) and returns the sheet properties. */
    static ensureSheet(spreadsheet, sheetTitle) {

        const existingSheet = spreadsheet.sheets.find(sheet => sheet.properties.title === sheetTitle);
        if (existingSheet) {
            return existingSheet;
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
        const response = Sheets.Spreadsheets.batchUpdate(batch, spreadsheet.spreadsheetId);
        return response.replies[0].addSheet;
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
