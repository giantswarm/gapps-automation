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
}
