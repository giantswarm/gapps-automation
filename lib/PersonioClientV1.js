/** The Personio API v1 prefix. */
const PERSONIO_API_BASE_URL = 'https://api.personio.de/v1';


/** Simple wrapper around UrlFetchApp that performs authenticated requests against Personio API v1.
 *
 * Export to global namespace, otherwise ES6 declarative constructs are not available outside the library.
 */
class PersonioClientV1 extends UrlFetchJsonClient {

    constructor(service) {
        super(service);
    }


    /** Get a client authenticated with the specified API credentials. */
    static withApiCredentials(clientId, clientSecret) {
        return new PersonioClientV1(new PersonioAuthV1(clientId, clientSecret));
    }


    /** Override UrlFetchJsonClient.fetch() to grab the next access token from the response. */
    fetch(url, params) {
        const response = super.fetch(url, params);

        // intercept and remember access token
        const service = this.getService();
        if (service?.setAccessToken) {
            let nextToken = null;
            const auth = response.getHeaders()['authorization'];
            if (auth) {
                nextToken = auth.replace('Bearer ', '');
            }

            service.setAccessToken(nextToken);
        }

        return response;
    }


    /** Fetch JSON from the Personio API.
     *
     * @param url Absolute or relative URL below Personio API endpoint.
     * @param options Additional fetch parameters (same as UrlFetchApp.fetch()'s second parameter).
     * @return JSON document "data" member from Personio.
     */
    getPersonioJson(url, options) {

        // we ensure only known Personio API endpoints can be contacted
        const pathAndQuery = (url || '').replace(PERSONIO_API_BASE_URL, '');

        const document = this.getJson(PERSONIO_API_BASE_URL + pathAndQuery, options);
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
