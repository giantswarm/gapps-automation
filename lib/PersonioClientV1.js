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

        const absoluteUrl = PERSONIO_API_BASE_URL + (url || '').replace(PERSONIO_API_BASE_URL, '');
        let response = undefined;
        let error = undefined;
        try
        {
            response = super.fetch(absoluteUrl, params);
        }
        catch (e)
        {
            error = e;
            response = e.response;
        }

        // intercept and remember access token (even in case of errors to prevent double-faults)
        const service = this.getService();
        if (service?.setAccessToken) {
            let nextToken = null;
            const auth = response.getHeaders()['authorization'];
            if (auth) {
                nextToken = auth.replace('Bearer ', '');
            }

            service.setAccessToken(nextToken);
        }

        if (error)
        {
            // re-throw caught exception
            throw error;
        }

        return response;
    }


    /** Fetch JSON (transparently handling Personio auth). */
    fetchJson(url, params) {
        const body = this.fetch(url, params).getContentText();
        if (body === '' || body == null) {
            return null;
        }
        return JSON.parse(body);
    }


    /** Fetch JSON from the Personio API.
     *
     * Will automatically query all remaining pages if offset and limit are not specified.
     *
     * @param url Absolute or relative URL below Personio API endpoint.
     * @param options Additional fetch parameters (same as UrlFetchApp.fetch()'s second parameter).
     * @return JSON document "data" member from Personio.
     */
    getPersonioJson(url, options) {
        let data = [];
        let offset = null;
        do {
            // we ensure only known Personio API endpoints can be contacted
            let pathAndQuery = url;
            if (offset != null) {
                pathAndQuery += pathAndQuery.includes('?') ? '&offset=' + offset : '?offset=' + offset;
            }

            const document = this.getJson(pathAndQuery, options);
            if (!document || !document.success) {
                const message = (document && document.error && document.error.message) ? document.error.message : '';
                throw new Error('Error response for ' + pathAndQuery + ' from Personio' + (message ? ': ' : '') + message);
            }

            if (!document.data) {
                throw new Error('Response for ' + pathAndQuery + ' from Personio doesn\'t contain data');
            }

            if (!Array.isArray(document.data)
                || url.includes('limit=')
                || url.includes('offset=')
                || !document.limit) {
                data = document.data;
                break;
            }

            data = data.concat(document.data);

            // keep requesting remaining pages
            offset = document.data.length < document.limit ? offset = null : data.length;
        }
        while (offset != null);

        return data;
    }
}
