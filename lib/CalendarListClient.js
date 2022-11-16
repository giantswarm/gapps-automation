/** Returns OAuth2 authenticated calendar scoped impersonation service for the specified user (by primary email).
 *
 * Requires service account credentials (with domain-wide delegation enabled) to be set, for example via:
 *
 *   clasp run 'setProperties' --params '[{"AddCalendars.serviceAccountCredentials": "{...ESCAPED_JSON...}"}, false]'
 */
class CalendarListClient extends UrlFetchJsonClient {

    constructor(service) {
        super(service);
    }


    /** Helper to create an impersonating instance. */
    static withImpersonatingService(serviceAccountCredentials, primaryEmail) {
        const service = UrlFetchJsonClient.createImpersonatingService(serviceAccountCredentials,
            primaryEmail,
            'https://www.googleapis.com/auth/calendar');

        return new CalendarListClient(service);
    }


    insert(calendarListItem) {
        return this.postJson('https://www.googleapis.com/calendar/v3/users/me/calendarList', calendarListItem);
    }


    list() {
        const calendars = [];
        let list = null;
        do {
            list = this.getJson('https://www.googleapis.com/calendar/v3/users/me/calendarList', {
                pageToken: list?.nextPageToken
            });
            calendars.push(...list.items);
        }
        while (list?.nextPageToken);

        return calendars;
    }
}
