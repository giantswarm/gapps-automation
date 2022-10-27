/** Basic utility functions (cheap version of lodash) ;) */
class Util {
    /** Check if something is an object. */
    static isObject(a) {
        return (!!a) && (a.constructor === Object);
    }
}
