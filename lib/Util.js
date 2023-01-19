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
    static getTimestampOffset(tsStr) {
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

    /** In-place add milliseconds to an existing date object.
     * @return {Date} The same date instance specified as argument for parameter date.
     */
    static addDateMillies(date, millies) {
        date.setTime(date.getTime() + millies);
        return date;
    }
}
