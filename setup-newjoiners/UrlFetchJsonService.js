/**
 * Base class for UrlFetchApp based services with JSON responses.
 */
class UrlFetchJsonService {
    #service;

    /**
     * Construct client instance.
     * @param service Optional OAuth2 service to use for all requests.
     */
    constructor(service) {
        this.#service = service;
    }

    /** Helper to create an impersonating OAuth2 "service". */
    static createImpersonatingService(serviceAccountCredentials, primaryEmail, scope) {

        const service = OAuth2.createService('AddCalendars')
            .setTokenUrl('https://accounts.google.com/o/oauth2/token')
            .setPrivateKey(serviceAccountCredentials.private_key)
            .setIssuer(serviceAccountCredentials.client_email)
            .setSubject(primaryEmail)
            .setPropertyStore(PropertiesService.getScriptProperties())
            .setParam('access_type', 'offline')
            .setScope(scope);

        service.reset();

        return service;
    }

    postJson(url, data, params) {
        return this.fetchJson(url, {
            ...params,
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(data)
        })
    }

    getJson(url, params) {
        return this.fetchJson(url, {
            ...params,
            method: 'get'
        })
    }

    fetchJson(url, params) {
        const response = UrlFetchApp.fetch(url, {
            ...params,
            headers: this.#service ? {
                ...((options || {}).headers || {}),
                Authorization: 'Bearer ' + this.#service.getAccessToken()
            } : ((options || {}).headers || {})
        });

        const code = response.getResponseCode();
        if (code < 200 || code >= 300) {
            throw new Error('Request for ' + url + ' failed with code ' + code);
        }

        const body = response.getContentText();
        if (!body) {
            throw new Error('Empty response for ' + url + ', code ' + code);
        }

        return JSON.parse(body)
    }
}
