/** The Personio API v1 prefix. */
const PERSONIO_API_BASE_URL = 'https://api.personio.de/v1';

const PERSONIO_MAX_PAGE_SIZE = 100;


/** Simple wrapper around UrlFetchApp that performs authenticated requests against Personio API v1.
 *
 * Export to global namespace, otherwise ES6 declarative constructs are not available outside the library.
 */
class PersonioClientV1 extends UrlFetchJsonClient {

    constructor(service, baseUrl = undefined) {
        super(service);
        this.baseUrl = baseUrl || PERSONIO_API_BASE_URL;
    }


    /** Get a client authenticated with the specified API credentials. */
    static withApiCredentials(clientId, clientSecret) {
        return new PersonioClientV1(new PersonioAuthV1(clientId, clientSecret));
    }


    /** Override UrlFetchJsonClient.fetch() to grab the next access token from the response. */
    async fetch(url, params) {

        const absoluteUrl = this.baseUrl + (url || '').replace(this.baseUrl, '');
        let response = undefined;
        let error = undefined;
        try
        {
            response = await super.fetch(absoluteUrl, params);
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
            if (response) {
                const auth = response.getHeaders()['authorization'];
                if (auth) {
                    nextToken = auth.replace('Bearer ', '');
                }
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
    async fetchJson(url, params) {
        const body = (await this.fetch(url, params)).getContentText();
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
    async getPersonioJson(url, options) {
        const params = PersonioClientV1.parseQuery(url);
        let data = [];
        let offset = params.offset !== undefined ? +params.offset : null;
        let limit = params.limit !== undefined ? +params.limit : null;
        do {
            if (offset != null || limit != null) {
                offset = Math.floor(Math.max(+offset, 0));
                limit = Math.floor(Math.max(+limit, PERSONIO_MAX_PAGE_SIZE));
                params.offset = '' + offset;
                params.limit = '' + limit;
            }

            const finalUrl = url.split('?')[0] + PersonioClientV1.buildQuery(params);
            // we ensure only known Personio API endpoints can be contacted
            const document = await this.getJson(finalUrl, options);
            if (!document || !document.success) {
                const message = (document && document.error && document.error.message) ? document.error.message : '';
                throw new Error('Error response for ' + finalUrl + ' from Personio' + (message ? ': ' : '') + message);
            }

            if (!document.data) {
                throw new Error('Response for ' + finalUrl + ' from Personio doesn\'t contain data');
            }

            if (!Array.isArray(document.data)
                || !limit
                || !document.limit) {
                data = document.data;
                break;
            }

            data = data.concat(document.data);

            // keep requesting remaining pages
            limit = document.limit;
            offset = document.data.length < limit ? offset = null : data.length;
            if (offset && url.includes('company/time-offs')) {
                // special case: time-offs endpoint's offset parameters is a page index (not element index)
                offset = Math.floor(offset / document.limit);
            }
        }
        while (offset != null);

        return data;
    }
}
