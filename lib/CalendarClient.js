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


    /** Patch an event (only the fields present in event are changed).
     *
     * @param calendarId The calendar ID, e.g. 'primary'
     * @param eventId The event ID
     * @param event The partial event resource
     * @param params Optional query parameters, e.g. {sendUpdates: 'none'}
     * @param etag Optional ETag of the event as read before, the patch fails (HTTP 412) if the event changed since
     */
    async patch(calendarId, eventId, event, params, etag) {
        const query = params ? CalendarClient.buildQuery(params) : '';
        return await this.patchJson(`https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/${eventId}${query}`, event,
            etag ? {headers: {'If-Match': etag}} : undefined);
    }


    async get(calendarId, eventId, params) {
        const query = params ? CalendarClient.buildQuery(params) : '';
        return await this.getJson(`https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/${eventId}${query}`);
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
