/** Authenticated Service for Personio.
 *
 * Supports getAccessToken() just like OAuth2.
 */
class PersonioAuthV1 extends UrlFetchJsonClient {

    constructor(clientId, clientSecret, baseUrl = undefined) {
        super();

        this.clientId = clientId;
        this.clientSecret = clientSecret;
        this.baseUrl = baseUrl || PERSONIO_API_BASE_URL;
    }


    /** Get a valid access token
     *
     * Requests an access token for this instance's clientId and clientSecret.
     */
    async getAccessToken() {

        if (!this.accessToken) {
            const document = await this.postJson(this.baseUrl + '/auth', {
                client_id: this.clientId,
                client_secret: this.clientSecret
            });

            if (!document || !document.data || !document.data.token) {
                throw new Error("No token received from Personio, code " + code);
            }

            this.accessToken = document.data.token;
        }

        return this.accessToken;
    }

    setAccessToken(nextAccessToken) {
        this.accessToken = nextAccessToken;
    }
}
