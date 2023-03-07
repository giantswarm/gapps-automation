/** Basic utility functions (cheap version of lodash) ;) */
class Util {

    /** Check if something is an object.
     *
     * @param {any} a The variable to check.
     * @return {boolean} True if the argument for parameter a is an object and not null, false otherwise.
     */
    static isObject(a) {
        return (!!a) && (a.constructor === Object);
    }


    /** Calculate UTF-8 encoded size of a string.
     *
     * @param {string} str The string to calculate the size for.
     * @return {number} Size of the string in bytes.
     */
    static calculateUtf8Size(str) {
        // returns the byte length of an utf8 string
        let s = str.length;
        for (let i = str.length - 1; i >= 0; i--) {
            let code = str.charCodeAt(i);
            if (code > 0x7f && code <= 0x7ff) s++;
            else if (code > 0x7ff && code <= 0xffff) s += 2;
            if (code >= 0xDC00 && code <= 0xDFFF) i--; //trail surrogate
        }
        return s;
    }

    /** Get the timezone offset in milliseconds of a full ISO8601 timestamp. */
    static getTimeZoneOffset(tsStr) {
        const ts = (tsStr || '').trim();
        if (!ts) {
            return undefined;
        }

        if (ts.endsWith('Z')) {
            return 0;
        }

        const tIndex = ts.indexOf('T');
        if (tIndex > 0) {
            let iSign = ts.indexOf('+', tIndex);
            if (iSign < 0) {
                iSign = ts.indexOf('-', tIndex);
            }
            if (iSign >= 0) {
                const hour = +ts.substring(iSign + 1, iSign + 3);
                const minute = +ts.substring(ts[iSign + 3] === ':' ? iSign + 4 : iSign + 3);
                return (ts[iSign] === '+' ? 1 : -1) * (hour * 60 * 60 * 1000) + (minute * 60 * 1000);
            }
        }

        return undefined;
    }


    /** Get the offset from UTC in milliseconds in the specified timezone (like 'Europe/Berlin') at the given date. */
    static getNamedTimeZoneOffset(timeZone, date) {

        // sv-SE is used, because it uses ISO8601 format (there is no explicit ISO8601 locale)
        const utcMillies = new Date((new Intl.DateTimeFormat('sv-SE', {
            year: 'numeric',
            month: 'numeric',
            timeZone: 'UTC',
            day: 'numeric',
            hour: 'numeric',
            minute: 'numeric',
            hour12: false
        })).format(date) + 'Z').valueOf();

        const tzLocalMillies = new Date((new Intl.DateTimeFormat('sv-SE', {
            year: 'numeric',
            month: 'numeric',
            timeZone: timeZone,
            day: 'numeric',
            hour: 'numeric',
            minute: 'numeric',
            hour12: false
        })).format(date) + 'Z').valueOf();

        // the difference is the time-zone offset
        return tzLocalMillies - utcMillies;
    }


    /** In-place add milliseconds to an existing date object.
     * @return {Date} The same date instance specified as argument for parameter date.
     */
    static addDateMillies(date, millies) {
        date.setTime(date.getTime() + millies);
        return date;
    }


    /** Generate a compliant UUIDv4 (with unproven collision resistance attributes).
     *
     * Derived from: https://stackoverflow.com/a/8809472
     */
    static generateUUIDv4() {
        let d = Date.now();
        // TODO Use a higher precision time source here, as soon as AppsScript provides one!
        // let d2 = (performance && performance.now && (performance.now() * 1000)) || 0;
        let d2 = 0;
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
            let r = Math.random() * 16;
            if (d > 0) {
                r = (d + r) % 16 | 0;
                d = Math.floor(d / 16);
            } else {
                r = (d2 + r) % 16 | 0;
                d2 = Math.floor(d2 / 16);
            }
            return (c === 'x' ? r : (r & 0x7 | 0x8)).toString(16);
        });
    };


    /** Shuffle an array, with somewhat proper random distribution.
     *
     * Implementation of https://en.wikipedia.org/wiki/Fisher%E2%80%93Yates_shuffle#The_modern_algorithm
     *
     * Derived from: https://stackoverflow.com/a/12646864
     */
    static shuffleArray(array) {
        for (let i = array.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            const temp = array[i];
            array[i] = array[j];
            array[j] = temp;
        }
    }
}
