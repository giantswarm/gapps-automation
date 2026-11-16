/** Google Meet REST API v2 client for reading conference records and participant data.
 *
 * Requires service account credentials with domain-wide delegation.
 * Scope 'meetings.space.readonly' (or the one passed to withImpersonatingService()) must be authorized in
 * Workspace Admin Console.
 */
class MeetClient extends UrlFetchJsonClient {

    constructor(service) {
        super(service);
    }


    /** Helper to create an impersonating instance.
     *
     * @param serviceAccountCredentials The service account credentials (JSON object).
     * @param primaryEmail The user to impersonate.
     * @param scope Optional scope, default 'https://www.googleapis.com/auth/meetings.space.readonly'.
     *              Updating spaces created by other apps (like Calendar) requires 'https://www.googleapis.com/auth/meetings.space.settings'.
     */
    static async withImpersonatingService(serviceAccountCredentials, primaryEmail, scope) {
        const effectiveScope = scope || 'https://www.googleapis.com/auth/meetings.space.readonly';
        // the OAuth2 service (and its cached token) is per scope
        const serviceName = scope ? 'MeetClient-' + effectiveScope.split('/').pop() + '-' + primaryEmail : 'MeetClient-' + primaryEmail;
        const service = await UrlFetchJsonClient.createImpersonatingService(serviceName,
            serviceAccountCredentials,
            primaryEmail,
            effectiveScope);

        return new MeetClient(service);
    }


    /** Get a meeting space, including its config (access type, automatic artifacts).
     *
     * @param space The space resource name, 'spaces/{spaceId}' or 'spaces/{meetingCode}', e.g. 'spaces/abc-mnop-xyz'
     * @returns {Promise<Object>} The space object
     */
    async getSpace(space) {
        return await this.getJson(`https://meet.googleapis.com/v2/${space}`);
    }


    /** Update a meeting space (only the fields named in updateMask).
     *
     * @param spaceName The space resource name with the server generated ID ('spaces/{spaceId}', as returned by getSpace()),
     *                  meeting codes are not accepted here
     * @param space The (partial) space object holding the new values
     * @param updateMask Comma separated field paths to update, e.g. 'config.artifactConfig.recordingConfig.autoRecordingGeneration'
     * @returns {Promise<Object>} The updated space object
     */
    async patchSpace(spaceName, space, updateMask) {
        const query = MeetClient.buildQuery({updateMask: updateMask});
        return await this.patchJson(`https://meet.googleapis.com/v2/${spaceName}${query}`, space);
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
