/** Returns OAuth2 authenticated calendar scoped impersonation service for the specified user (by primary email).
 *
 * Requires service account credentials as exported from Google Cloud Admin Console.
 */
class CalendarListClient extends UrlFetchJsonClient {

    constructor(service) {
        super(service);
    }


    /** Helper to create an impersonating instance. */
    static withImpersonatingService(serviceAccountCredentials, primaryEmail) {
        const service = UrlFetchJsonClient.createImpersonatingService('CalendarListClient-' + primaryEmail,
            serviceAccountCredentials,
            primaryEmail,
            'https://www.googleapis.com/auth/calendar');

        return new CalendarListClient(service);
    }


    insert(calendarListItem) {
        return this.postJson('https://www.googleapis.com/calendar/v3/users/me/calendarList', calendarListItem);
    }


    list() {
        const calendars = [];
        const params = {
            pageToken: undefined
        };
        do {
            const query = CalendarListClient.buildQuery(params);
            const list = this.getJson(`https://www.googleapis.com/calendar/v3/users/me/calendarList${query}`);

            params.pageToken = list.nextPageToken;
            calendars.push(...list.items);
        }
        while (params.pageToken);

        return calendars;
    }
}
