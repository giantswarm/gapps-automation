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
        if (!service.hasAccess()) {
            throw new Error("Failed to impersonate " + primaryEmail);
        }
        return service;
    }


    /** Post JSON, receiving JSON object or array. */
    postJson(url, data, params) {
        return this._fetchJson(url, {
            ...params,
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(data)
        })
    }

    /** Get a JSON object or array. */
    getJson(url, params) {
        return this._fetchJson(url, {
            ...params,
            method: 'get'
        })
    }


    _fetchJson(url, params) {
        const body = this.fetch(url, params).getContentText();
        if (body === '' || body == null) {
            throw null;
        }
        return JSON.parse(body);
    }


    fetch(url, params) {
        const response = UrlFetchApp.fetch(url, {
            ...params,
            headers: this.service ? ({
                ...((params || {}).headers || {}),
                Authorization: 'Bearer ' + this.service.getAccessToken()
            }) : ((params || {}).headers || {})
        });

        const code = response.getResponseCode();
        if (code < 200 || code >= 300) {
            throw new Error('Request for ' + url + ' failed with code ' + code);
        }

        return response;
    }
}
