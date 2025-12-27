/**
 * Base class for UrlFetchApp based services with JSON responses.
 */
class UrlFetchJsonClient {

    /**
     * Construct client instance.
     * @param service Optional, configured OAuth2 service to use for all requests, Must have a getAccessToken() method.
     */
    constructor(service) {
        this.service = service;
    }


    /** Get the authenticated service for this instance.
     *
     * @return {*} The authenticated service, for example an OAuth2 or PersonioAuthV1 instance.
     */
    getService() {
        return this.service;
    }


    /** Helper to create an impersonating OAuth2 "service".
     *
     * Main users (the one executing the script) must have IAM roles:
     *  roles/iam.serviceAccountUser
     *  roles/iam.serviceAccountTokenCreator
     *
     * Additionally Domain-wide-Delegation must be enabled for the service account and configured with the right scopes.
     */
    static async createImpersonatingService(uniqueServiceName, serviceAccountCredentials, primaryEmail, scope) {

        const service = OAuth2.createService(uniqueServiceName)
            .setTokenUrl('https://accounts.google.com/o/oauth2/token')
            .setPrivateKey(serviceAccountCredentials.private_key)
            .setIssuer(serviceAccountCredentials.client_email)
            .setSubject(primaryEmail)
            .setPropertyStore(PropertiesService.getScriptProperties())
            .setParam('access_type', 'offline')
            .setParam('prompt', 'none')
            .setScope(scope);

        if (!await service.hasAccess()) {
            throw new Error("Failed to impersonate " + primaryEmail);
        }
        return service;
    }


    /** Post JSON, receiving JSON object or array. */
    async postJson(url, data, params) {
        return await this.fetchJson(url, {
            ...params,
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(data)
        })
    }


    /** Put JSON, receiving JSON object or array. */
    async putJson(url, data, params) {
        return await this.fetchJson(url, {
            ...params,
            method: 'put',
            contentType: 'application/json',
            payload: JSON.stringify(data)
        })
    }


    /** Get a JSON object or array. */
    async getJson(url, params) {
        return await this.fetchJson(url, {
            ...params,
            method: 'get'
        })
    }


    async fetchJson(url, params) {
        const body = (await this.fetch(url, params)).getContentText();
        if (body === '' || body == null) {
            return null;
        }
        return JSON.parse(body);
    }


    async fetch(url, params) {
        const response = await UrlFetchApp.fetch(url, {
            ...params,
            muteHttpExceptions: true,
            headers: this.service ? ({
                ...((params || {}).headers || {}),
                Authorization: 'Bearer ' + await this.service.getAccessToken()
            }) : ((params || {}).headers || {})
        });

        const code = response.getResponseCode();
        if (code < 200 || code >= 300) {
            const e = new Error('Request for ' + url + ' failed with code ' + code + ': ' + response.getContentText());
            e.response = response;
            throw e;
        }

        return response;
    }


    /** Assemble a URL query component from an object with (map).
     *
     * Supports array members as multiple parameters.
     *
     * Properties whose value is undefined are ignored.
     *
     * A '?' prefix is added, if the query parameter list is not empty.
     *
     * @param params Object whose properties are converted to a query string.
     * @return {string|string}
     */
    static buildQuery(params) {
        const encodeQueryParam = (key, value) => encodeURIComponent(key) + (value !== undefined ? '=' + encodeURIComponent(value) : '');

        const query = Object.entries(params)
            .filter(([key, value]) => value !== undefined)
            .map(([key, value]) => {
                if (Array.isArray(value)) {
                    // flatten arrays
                    return value.map(v => encodeQueryParam(key, v)).join('&');
                } else {
                    return encodeQueryParam(key, value);
                }
            })
            .join('&');

        return query.length > 0 ? '?' + query : '';
    };


    /** Explode URL query components into an object.
     *
     * Expects ? prefix and splits ignores fragment part.
     *
     * @param url URL whose query part to convert.
     * @return {object}
     */
    static parseQuery(url) {
        const encodeQueryParam = (key, value) => encodeURIComponent(key) + '=' + encodeURIComponent(value);

        const urlParts = url.split('?');
        if (urlParts.length < 2) {
            return {};
        }

        const query = urlParts[1].split('#')[0];
        return query.split('&').reduce((acc, part) => {
            const parts = part.split('=');
            const key = decodeURIComponent(parts[0]);
            const value = parts[1] === undefined ? undefined : parts[1] === 'null' ? null : decodeURIComponent(parts[1]);
            if (acc.hasOwnProperty(key)) {
                const existingValue = acc[key];
                if (!Array.isArray(existingValue)) {
                    acc[key] = [existingValue, value];
                } else {
                    existingValue.push(value);
                }
            } else {
                acc[key] = value;
            }
            return acc;
        }, {});
    };
}
