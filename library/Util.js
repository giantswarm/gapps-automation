/** Basic utility functions (cheap version of lodash) ;) */

/** The seconds in a day. */
const SECONDS_PER_DAY = 86400;

/** Milliseconds per day. */
const MILLISECONDS_PER_DAY = 86400000;

/** Check if something is an object.
 *
 * @param a The variable to check.
 * @return True if the argument for parameter a is an object and not null, false otherwise.
 */
function isObject(a) {
    return (!!a) && (a.constructor === Object);
}

/** Excel DateTime (AKA SERIAL_NUMBER date) to JS Date.
 *
 * @param excelSerialDateTime A Sheets/Excel serial date like:
 *   44876.641666666605  ( 2022-11-11T15:24:00.000Z )
 *
 * @return A Date object representing that point in time.
 */
function serialDateTimeToDate(excelSerialDateTime) {
    const totalDays = Math.floor(excelSerialDateTime);
    const totalSeconds = excelSerialDateTime * SECONDS_PER_DAY; // convert to fractional seconds, to avoid precision issues
    const millies = Math.round((totalSeconds - (totalDays * SECONDS_PER_DAY)) * 1000);
    return new Date(Date.UTC(0, 0, totalDays - 1, 0, 0, 0, millies));
}


/** JS Date to Excel DateTime (AKA SERIAL_NUMBER date).
 *
 * @param date A JavaScript date object, possibly representing a datetime like 2022-11-11T15:24:00.000Z..
 *
 * @return A Sheets/Excel serial date like 44876.641666666605.
 */
function dateToSerialDateTime(date) {
    return (date.getTime() / MILLISECONDS_PER_DAY) + 25569; // 1970-01-01 - 1900-01-01 = 25569
}
