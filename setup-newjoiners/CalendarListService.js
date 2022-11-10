/** Returns OAuth2 authenticated calendar scoped impersonation service for the specified user (by primary email).
 *
 * Requires service account credentials (with domain-wide delegation enabled) to be set, for example via:
 *
 *   clasp run 'setProperties' --params '[{"AddCalendars.serviceAccountCredentials": "{...ESCAPED_JSON...}"}, false]'
 */
class CalendarListService extends UrlFetchJsonService {
    /** Create an instance impersonating the user specified by primaryEmail via the given service account . */
    constructor(serviceAccountCredentials, primaryEmail) {
        super(UrlFetchJsonService.createImpersonatingService(serviceAccountCredentials,
            primaryEmail,
            'https://www.googleapis.com/auth/calendar'));
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
