/** Google People API client for reading the domain directory.
 *
 * Requires service account credentials with domain-wide delegation.
 * Scope 'directory.readonly' must be authorized in Workspace Admin Console.
 */
class DirectoryClient extends UrlFetchJsonClient {

    constructor(service) {
        super(service);
    }


    /** Helper to create an impersonating instance. */
    static async withImpersonatingService(serviceAccountCredentials, primaryEmail) {
        const service = await UrlFetchJsonClient.createImpersonatingService('DirectoryClient-' + primaryEmail,
            serviceAccountCredentials,
            primaryEmail,
            'https://www.googleapis.com/auth/directory.readonly');

        return new DirectoryClient(service);
    }


    /** List all domain directory people with their names and email addresses.
     *
     * @returns {Promise<Array>} Array of person objects with names and emailAddresses fields
     */
    async listDirectoryPeople() {
        const people = [];
        const queryParams = {
            readMask: 'names,emailAddresses',
            sources: 'DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE',
            pageSize: 1000,
            pageToken: undefined
        };
        do {
            const query = DirectoryClient.buildQuery(queryParams);
            const response = await this.getJson(`https://people.googleapis.com/v1/people:listDirectoryPeople${query}`);
            if (response?.people) {
                people.push(...response.people);
            }
            queryParams.pageToken = response?.nextPageToken;
        } while (queryParams.pageToken);

        return people;
    }
}
