/** Google Meet REST API v2 client for reading conference records and participant data.
 *
 * Requires service account credentials with domain-wide delegation.
 * Scope 'meetings.space.readonly' must be authorized in Workspace Admin Console.
 */
class MeetClient extends UrlFetchJsonClient {

    constructor(service) {
        super(service);
    }


    /** Helper to create an impersonating instance. */
    static async withImpersonatingService(serviceAccountCredentials, primaryEmail) {
        const service = await UrlFetchJsonClient.createImpersonatingService('MeetClient-' + primaryEmail,
            serviceAccountCredentials,
            primaryEmail,
            'https://www.googleapis.com/auth/meetings.space.readonly');

        return new MeetClient(service);
    }


    /** List conference records matching the given filter.
     *
     * @param filter EBNF filter string, e.g. 'space.meeting_code="abc-mnop-xyz"'
     * @returns {Promise<Array>} Array of conference record objects
     */
    async listConferenceRecords(filter) {
        const records = [];
        const queryParams = { filter, pageToken: undefined };
        do {
            const query = MeetClient.buildQuery(queryParams);
            const response = await this.getJson(`https://meet.googleapis.com/v2/conferenceRecords${query}`);
            if (response?.conferenceRecords) {
                records.push(...response.conferenceRecords);
            }
            queryParams.pageToken = response?.nextPageToken;
        } while (queryParams.pageToken);

        return records;
    }


    /** List participants of a conference record.
     *
     * @param conferenceRecordName The conference record resource name, e.g. 'conferenceRecords/abc123'
     * @returns {Promise<Array>} Array of participant objects
     */
    async listParticipants(conferenceRecordName) {
        const participants = [];
        const queryParams = { pageSize: 250, pageToken: undefined };
        do {
            const query = MeetClient.buildQuery(queryParams);
            const response = await this.getJson(`https://meet.googleapis.com/v2/${conferenceRecordName}/participants${query}`);
            if (response?.participants) {
                participants.push(...response.participants);
            }
            queryParams.pageToken = response?.nextPageToken;
        } while (queryParams.pageToken);

        return participants;
    }
}
