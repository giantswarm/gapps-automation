/** Returns OAuth2 authenticated calendar scoped impersonation service for the specified user (by primary email).
 *
 * Requires service account credentials as exported from Google Cloud Admin Console.
 */
class CalendarClient extends UrlFetchJsonClient {

    constructor(service) {
        super(service);
    }


    /** Helper to create an impersonating instance. */
    static async withImpersonatingService(serviceAccountCredentials, primaryEmail) {
        const service = await UrlFetchJsonClient.createImpersonatingService('CalendarClient-' + primaryEmail,
            serviceAccountCredentials,
            primaryEmail,
            'https://www.googleapis.com/auth/calendar');

        return new CalendarClient(service);
    }


    async insert(calendarId, event) {
        return await this.postJson(`https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events`, event);
    }


    async update(calendarId, eventId, event) {
        return await this.putJson(`https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/${eventId}`, event);
    }


    async list(calendarId, params) {
        const events = [];
        const queryParams = {
            ...params,
            pageToken: undefined
        };
        do {
            const query = CalendarClient.buildQuery(queryParams);
            const list = await this.getJson(`https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events${query}`);

            queryParams.pageToken = list.nextPageToken;
            events.push(...list.items);
        }
        while (queryParams.pageToken);

        return events;
    }
}
