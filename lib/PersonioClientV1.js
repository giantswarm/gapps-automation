const PERSONIO_API_BASE_URL = 'https://api.personio.de/v1';


/** Simple wrapper around UrlFetchApp that performs authenticated requests against Personio API v1.
 *
 * Export to global namespace, otherwise ES6 declarative constructs are not available outside the library.
 */
class PersonioClientV1 {

    constructor(clientId, clientSecret) {
        this.clientId = clientId;
        this.clientSecret = clientSecret;
    }


    /** Requests an access token for this instance's clientId and clientSecret. */
    _fetchToken() {

        const response = UrlFetchApp.fetch(PERSONIO_API_BASE_URL + '/auth', {
            method: 'post',
            payload: {
                client_id: this.clientId,
                client_secret: this.clientSecret
            },
            headers: {
                Accept: 'application/json'
            }
        });

        const code = response.getResponseCode();
        if (code < 200 || code >= 300) {
            throw new Error("Personio auth failed with code " + code);
        }

        const body = response.getContentText();
        if (!body) {
            throw new Error("Empty auth response from Personio, code " + code);
        }

        const document = JSON.parse(body);
        if (!document || !document.data || !document.data.token) {
            throw new Error("No token received from Personio, code " + code);
        }

        return document.data.token;
    }


    /** Fetch JSON from the Personio API.
     *
     * @param url Absolute or relative URL below Personio API endpoint.
     * @param options Additional fetch parameters (same as UrlFetchApp.fetch()'s second parameter.
     * @return JSON document "data" member from Personio.
     */
    fetch(url, options) {
        const token = this._fetchToken();

        // we ensure only known Personio API endpoints can be contacted
        const pathAndQuery = (url || '').replace(PERSONIO_API_BASE_URL, '');

        const response = UrlFetchApp.fetch(PERSONIO_API_BASE_URL + pathAndQuery, {
            ...options,
            headers: {
                ...((options || {}).headers || {}),
                Authorization: 'Bearer ' + token
            }
        });

        const code = response.getResponseCode();
        if (code < 200 || code >= 300) {
            throw new Error('Personio request for ' + pathAndQuery + ' failed with code ' + code);
        }

        const body = response.getContentText();
        if (!body) {
            throw new Error('Empty response for ' + pathAndQuery + ' from Personio, code ' + code);
        }

        const document = JSON.parse(body);
        if (!document || !document.success) {
            const message = (document && document.error && document.error.message) ? document.error.message : '';
            throw new Error('Error response for ' + pathAndQuery + ' from Personio' + (message ? ': ' : '') + message);
        }

        if (!document.data) {
            throw new Error('Response for ' + pathAndQuery + ' from Personio doesn\'t contain data');
        }

        return document.data;
    }
}
