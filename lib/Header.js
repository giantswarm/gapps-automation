if (!!global.UrlFetchApp) {
    // we are in GAS land, throw error to indicate invalid usage
    throw new Error('This module does not support the Google Apps Script runtime, please use NodeJS v18 or install node-fetch or target a browser that supports fetch()');
}

/** We export in trailer. */
const nativeModule = module;
module = {exports: {}};


/** Simple replacement for GAS UrlFetchApp. Requires a fetch() implementation to be present. */
class UrlFetchApp {
    static async fetch(url, options) {
        console.log("fetch()", url, options);
        try {
            const r = await fetch(url, options);
            const text = await r.text();
            r.getResponseCode = () => r.status;
            r.getContentText = () => text;
            r.getHeaders = () => r.headers;
            return r;
        } catch (e) {
            console.log("fetch error: ", e);
            e.response = {getResponseCode: () => 0, getResponseText: () => null};
            throw e;
        }
    }
}


/** Simple in-memory replacement for PropertiesService.getScriptProperties().
 * Initialized from environment.
 */
class ScriptProperties {

    constructor() {
        this.store = process?.env ? {...process.env} : {};
    }

    setProperties(properties, deleteAllOthers) {
        this.store = deleteAllOthers ? properties : {...this.store, ...properties};
    }

    setProperty(key, value) {
        this.store[key] = value;
    }

    getProperty(key) {
        return this.store[key];
    }

    getProperties() {
        return {...this.store};
    }
}


const _PropertiesService_scriptProperties = new ScriptProperties();


// Simple mock like implementations for various GAS APIs
class PropertiesService {
    static getScriptProperties() {
        return _PropertiesService_scriptProperties;
    }
}


class Lock {

    constructor() {
        this.isLocked = false;
    }

    tryLock(timeoutMillies) {
        if (!this.isLocked) {
            this.isLocked = true;
            return true;
        }
        return false;
    }

    waitLock(timeoutMillies) {
        if (this.isLocked) {
            throw new Error("Resource is locked. Waiting for a lock is not supported in this Script Lock implementation.");
        }
        this.isLocked = true;
    }

    hasLock() {
        return this.isLocked;
    }

    releaseLock() {
        if (!this.isLocked) {
            throw new Error("Tried to release a lock which wasn't acquired before.")
        }

        this.isLocked = false;
    }
}

const _LockService_scriptLock = new Lock();


class LockService {
    static getScriptLock() {
        return _LockService_scriptLock;
    }
}


class Cache {

    constructor() {
        this.store = {};
    }

    get(key) {
        return this.store[key];
    }
}

const _CacheService_scriptCache = new Cache();

class CacheService {
    static getScriptCache() {
        return _CacheService_scriptCache;
    }
}
