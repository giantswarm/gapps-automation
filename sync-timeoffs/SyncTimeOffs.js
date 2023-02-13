/**
 * Synchronize Time Offs / Out-of-Office events between Personio and the domains personal Google Calendars.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */


/** The prefix for properties specific to this script in the project. */
const PROPERTY_PREFIX = 'SyncTimeOffs.';

/** Personio clientId and clientSecret, separated by '|'. */
const PERSONIO_TOKEN_KEY = PROPERTY_PREFIX + 'personioToken';

/** Service account credentials (in JSON format, as downloaded from Google Management Console. */
const SERVICE_ACCOUNT_CREDENTIALS_KEY = PROPERTY_PREFIX + 'serviceAccountCredentials';

/** Filter for allowed domains (to avoid working and failing on users present on foreign domains). */
const ALLOWED_DOMAINS_KEY = PROPERTY_PREFIX + 'allowedDomains';

/** White-list to restrict synchronization to a few tester email accounts.
 *
 * Must be one email or a comma separated list.
 *
 * Default: null or empty
 */
const EMAIL_WHITELIST_KEY = PROPERTY_PREFIX + 'emailWhiteList';

/** Lookahead days for event/time-off synchronization.
 *
 * Default: 6 * 30 days, should scale up to ~18 months
 */
const LOOKAHEAD_DAYS_KEY = PROPERTY_PREFIX + 'lookaheadDays';

/** Lookback days for event/time-off synchronization.
 *
 * Default: 30 days, should scale up to ~6 months, avoid changing events too far back
 */
const LOOKBACK_DAYS_KEY = PROPERTY_PREFIX + 'lookbackDays';

/** Maximum number of attempts to retry doing something with the Personio API, per event.
 *
 * This is intended as a forward oriented measure to avoid wasting resources
 * in case there are "hanging" events (for example after configuration or API changes).
 *
 * In case of long-lasting Personio problems (outage or similar) a few events may stop being synced.
 *
 * There are ways to handle this situation:
 *    - increase MAX_SYNC_FAIL_COUNT by 1 (so one more retry is allowed)
 *    - reset fail count for the affected events (using scripting here or some tool)
 *    - delete the Gcal events and create new ones (failCount: 0)
 *    - ignore the situation (it's just a few events)
 */
const MAX_SYNC_FAIL_COUNT_KEY = PROPERTY_PREFIX + 'maxSyncFailCount';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'syncTimeOffs';

/** The minimum duration of a half-day out-of-office event. */
const MINIMUM_OUT_OF_OFFICE_DURATION_HALF_DAY_MILLIES = 3 * 60 * 60 * 1000;  // half-day >= 3h

/** The minimum duration of a whole day or longe out-of-office event. */
const MINIMUM_OUT_OF_OFFICE_DURATION_WHOLE_DAY_MILLIES = 6 * 60 * 60 * 1000; // whole-day >= 6h


/** Main entry point.
 *
 * Sync TimeOffs between Personio and personal Google Calendars in the organization.
 *
 * This function is intended for continuous operation for scalability reasons. This means that it won't synchronize
 * events that are very far in the past.
 *
 * This requires a configured Personio access token and the user impersonation API enabled for the cloud project:
 *  https://cloud.google.com/iam/docs/impersonating-service-accounts
 *
 * Requires the following script properties for operation:
 *
 *   SyncTimeOffs.personioToken              CLIENT_ID|CLIENT_SECRET
 *   SyncTimeOffs.serviceAccountCredentials  {...SERVICE_ACCOUNT_CREDENTIALS...}
 *   SyncTimeOffs.allowedDomains             giantswarm.io,giantswarm.com
 *
 * Requires service account credentials (with domain-wide delegation enabled) to be set, for example via:
 *
 *   $ clasp run 'setProperties' --params '[{"SyncTimeOffs.serviceAccountCredentials": "{...ESCAPED_JSON...}"}, false]'
 *
 * One may use the following command line to compress the service account creds into one line:
 *
 *   $ cat credentials.json | tr -d '\n '
 *
 * The service account must be configured correctly and have at least permission for these scopes:
 *   https://www.googleapis.com/auth/calendar
 */
function syncTimeOffs() {

    const scriptLock = LockService.getScriptLock();
    if (!scriptLock.tryLock(5000)) {
        throw new Error('Failed to acquire lock. Only one instance of this script can run at any given time!');
    }

    const allowedDomains = (getScriptProperties_().getProperty(ALLOWED_DOMAINS_KEY) || '')
        .split(',')
        .map(d => d.trim());

    const emailWhiteList = getEmailWhiteList_();
    const isEmailAllowed = email => (!emailWhiteList.length || emailWhiteList.includes(email))
        && allowedDomains.includes(email.substring(email.lastIndexOf('@') + 1));

    Logger.log('Configured to handle accounts %s on domains %s', emailWhiteList.length ? emailWhiteList : '', allowedDomains);

    // all timing related activities are relative to this EPOCH
    const epoch = new Date();

    // how far back to sync events/time-offs
    const lookbackMillies = -Math.round(getLookbackDays_() * 24 * 60 * 60 * 1000);
    // how far into the future to sync events/time-offs
    const lookaheadMillies = Math.round(getLookaheadDays_() * 24 * 60 * 60 * 1000);
    // how many Personio action retries per event?
    const maxFailCount = getMaxSyncFailCount_();
    // after how many milliseconds should this script stop by itself (to avoid forced termination/unclean state)?
    const maxRuntimeMillies = Math.round(320 * 1000); // 5:20 minutes (hard AppsScript kill comes at 6:00 minutes)

    const fetchTimeMin = Util.addDateMillies(new Date(epoch), lookbackMillies);
    fetchTimeMin.setUTCHours(0, 0, 0, 0); // round down to start of day
    const fetchTimeMax = Util.addDateMillies(new Date(epoch), lookaheadMillies);
    fetchTimeMax.setUTCHours(24, 0, 0, 0); // round up to end of day

    const personioCreds = getPersonioCreds_();
    const personio = PersonioClientV1.withApiCredentials(personioCreds.clientId, personioCreds.clientSecret);

    // load timeOffTypeDb
    const timeOffTypeDb = new TimeOffTypeDb(personio.getPersonioJson('/company/time-off-types'));

    // load and prepare list of employees to process
    const employees = personio.getPersonioJson('/company/employees').filter(employee =>
        employee.attributes.status.value !== 'inactive' && isEmailAllowed(employee.attributes.email.value)
    );
    Util.shuffleArray(employees);

    Logger.log('Syncing events between %s and %s for %s accounts', fetchTimeMin.toISOString(), fetchTimeMax.toISOString(), '' + employees.length);

    let firstError = null;
    let processedCount = 0;
    for (const employee of employees) {

        const email = employee.attributes.email.value;

        // we keep operating if handling calendar of a single user fails
        try {
            const calendar = CalendarClient.withImpersonatingService(getServiceAccountCredentials_(), email);
            if (!syncTimeOffs_(personio, calendar, employee, epoch, timeOffTypeDb, fetchTimeMin, fetchTimeMax, maxFailCount, maxRuntimeMillies)) {
                break;
            }
        } catch (e) {
            Logger.log('Failed to sync time-offs/out-of-offices of user %s: %s', email, e);
            firstError = firstError || e;
        }
        ++processedCount;
    }

    Logger.log('Completed synchronization for %s of %s accounts', '' + processedCount, '' + employees.length);

    // for completeness, also automatically released at exit
    scriptLock.releaseLock();

    if (firstError) {
        throw firstError;
    }
}


/** Uninstall triggers. */
function uninstall() {
    TriggerUtil.uninstall(TRIGGER_HANDLER_FUNCTION);
}


/** Install periodic execution trigger. */
function install(delayMinutes) {
    TriggerUtil.install(TRIGGER_HANDLER_FUNCTION, delayMinutes);
}


/** Allow setting properties. */
function setProperties(properties, deleteAllOthers) {
    TriggerUtil.setProperties(properties, deleteAllOthers);
}


/** Get script properties. */
function getScriptProperties_() {
    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties) {
        throw new Error('ScriptProperties not accessible');
    }

    return scriptProperties;
}


/** Get the service account credentials. */
function getServiceAccountCredentials_() {
    const creds = getScriptProperties_().getProperty(SERVICE_ACCOUNT_CREDENTIALS_KEY);
    if (!creds) {
        throw new Error("No service account credentials at script property " + SERVICE_ACCOUNT_CREDENTIALS_KEY);
    }

    return JSON.parse(creds);
}


/** Get the email account white-list (optional, leave empty to sync all suitable accounts). */
function getEmailWhiteList_() {
    return (getScriptProperties_().getProperty(EMAIL_WHITELIST_KEY) || '').trim()
        .split(',').map(email => email.trim()).filter(email => !!email);
}


/** Get the number of lookahead days (positive integer) or the default (6 * 30 days). */
function getLookaheadDays_() {
    const creds = (getScriptProperties_().getProperty(LOOKAHEAD_DAYS_KEY) || '').trim();
    const lookaheadDays = Math.abs(Math.round(+creds));
    if (!lookaheadDays || Number.isNaN(lookaheadDays)) {
        // use default: 6 months
        return 6 * 30;
    }

    return lookaheadDays;
}


/** Get the number of lookback days (positive integer) or the default (30 days). */
function getLookbackDays_() {
    const creds = (getScriptProperties_().getProperty(LOOKBACK_DAYS_KEY) || '').trim();
    const lookbackDays = Math.abs(Math.round(+creds));
    if (!lookbackDays || Number.isNaN(lookbackDays)) {
        // use default: 30 days
        return 30;
    }

    return lookbackDays;
}


/** Get the number of tries to act using the Personio API, per event. The default is 10. */
function getMaxSyncFailCount_() {
    const value = (getScriptProperties_().getProperty(MAX_SYNC_FAIL_COUNT_KEY) || '').trim();
    const maxFailCount = Math.abs(Math.round(+value));
    if (!maxFailCount || Number.isNaN(maxFailCount)) {
        // use default: try 10 times
        return 10;
    }

    return maxFailCount;
}


/** Get the Personio token. */
function getPersonioCreds_() {
    const credentialFields = (getScriptProperties_().getProperty(PERSONIO_TOKEN_KEY) || '|')
        .split('|')
        .map(field => field.trim());

    return {clientId: credentialFields[0], clientSecret: credentialFields[1]};
}


/** Small wrapper over a list of TimeOffType objects (for caching patterns). */
class TimeOffTypeDb {
    constructor(timeOffTypes) {

        this.timeOffTypes = timeOffTypes;

        // prepare patterns to guess time-off type from text
        // we dynamically create one RegExp per time-off, which we cache here (V8 can't)
        const pattern = [];
        for (let timeOffType of timeOffTypes) {
            const keyword = TimeOffTypeDb.extractKeyword(timeOffType.attributes?.name);
            pattern.push(new RegExp(`/(^|[\\s:-])(${keyword})([\\s:-]|$)/`, "sim"));
        }
        this.pattern = pattern;
    }

    /** Returns the keyword for the specified TimeOffType name (field timeOffType.attributes.name). */
    static extractKeyword(typeName) {
        return (typeName || '').attributes.name.split(' ')[0].trim() || undefined;
    }

    /** Guess timeOffType from a text (e.g. event summary) by keyword. */
    findByKeywordMatch(text) {
        const searchText = text || '';

        for (let i = 0; i < this.timeOffTypes; ++i) {
            if (this.pattern[i].test(searchText)) {
                return this.timeOffTypes[i];
            }
        }

        return undefined;
    }

    /** Find a timeOffType by its ID.
     *
     * @param {number} id The ID for looking up the TimeOffType.
     */
    findById(id) {
        return this.timeOffTypes.find(t => t.attributes.id === id);
    }
}


/** Subscribe a single account to all the specified calendars.
 *
 * @returns true if the specified employees account was fully processed, false if processing was aborted early.
 */
function syncTimeOffs_(personio, calendar, employee, epoch, timeOffTypeDb, fetchTimeMin, fetchTimeMax, maxFailCount, maxRuntimeMillies) {

    // test against dead-line first
    const deadlineTs = +epoch + maxRuntimeMillies;
    if (Date.now() >= deadlineTs) {
        return false;
    }

    const primaryEmail = employee.attributes.email.value;

    // We ignore working on events/time-offs that were updated too recently
    const updateDeadZoneMillies = 120 * 1000; // 120 seconds, to avoid races and workaround lack of transactions
    const updateMax = Util.addDateMillies(new Date(epoch), -updateDeadZoneMillies);

    // load timeOffs indexed by ID
    const timeOffs = queryPersonioTimeOffs_(personio, fetchTimeMin, fetchTimeMax, employee.attributes.id.value);

    const allEvents = queryCalendarEvents_(calendar, 'primary', fetchTimeMin, fetchTimeMax);
    for (const event of allEvents) {

        if (Date.now() >= deadlineTs) {
            return false;
        }

        const failCount = event.extendedProperties?.private?.syncFailCount || 0;
        let nextFailCount = failCount;
        const isEventCancelled = event.status === 'cancelled';
        const eventUpdatedAt = new Date(event.updated);
        const timeOffId = event.extendedProperties?.private?.timeOffId;
        if (timeOffId) {

            // we handle this time-off
            const timeOff = timeOffs[timeOffId];
            if (timeOff) {

                delete timeOffs[timeOffId];

                if (timeOff.updatedAt > updateMax || eventUpdatedAt > updateMax || failCount > maxFailCount) {
                    // dead zone
                    continue;
                }

                if (isEventCancelled) {
                    nextFailCount += !syncActionDeleteTimeOff_(personio, primaryEmail, timeOff);
                } else {
                    // need to convert to be able to compare start/end timestamps (Personio is whole-day/half-day only)
                    const updatedTimeOff = convertOutOfOfficeToTimeOff_(timeOffTypeDb, employee, event, timeOff);
                    if (updatedTimeOff && (!updatedTimeOff.startAt.equals(timeOff.startAt) || !updatedTimeOff.endAt.equals(timeOff.endAt))) {
                        // start/end timestamps differ, now check which (Personio/Google Calendar) has more recent changes
                        if (+timeOff.updatedAt >= +eventUpdatedAt) {
                            syncActionUpdateEvent_(calendar, primaryEmail, event, timeOff);
                        } else {
                            nextFailCount += !syncActionUpdateTimeOff_(personio, calendar, primaryEmail, event, timeOff, updatedTimeOff);
                        }
                    }
                }
            } else if (!isEventCancelled) {
                // check for dead zone
                // we allow event cancellation even in case maxFailCount was reached
                if (eventUpdatedAt <= updateMax) {
                    syncActionDeleteEvent_(calendar, primaryEmail, event);
                }
            }
        } else if (!isEventCancelled) {
            // check for dead zone, ignore events created by Cronofy
            if (eventUpdatedAt <= updateMax && failCount <= maxFailCount && !event.iCalUID.includes('cronofy.com')) {
                const newTimeOff = convertOutOfOfficeToTimeOff_(timeOffTypeDb, employee, event, undefined);
                if (newTimeOff) {
                    nextFailCount += !syncActionInsertTimeOff_(personio, calendar, primaryEmail, event, newTimeOff);
                }
            }
        }

        // register failure for Personio client circuit breaker
        if (failCount !== nextFailCount) {
            syncActionUpdateEventFailCount_(calendar, primaryEmail, event, nextFailCount);
        }
    }

    // Handle each remaining time-off, not handled above
    for (const timeOff of Object.values(timeOffs)) {

        if (Date.now() >= deadlineTs) {
            return false;
        }

        // check for dead zone
        if (timeOff.updatedAt <= updateMax) {
            syncActionInsertEvent_(calendar, primaryEmail, timeOffTypeDb, timeOff);
        }
    }

    return true;
}


/** Delete Personio TimeOffs for cancelled Google Calendar events */
function syncActionDeleteTimeOff_(personio, primaryEmail, timeOff) {
    try {
        // event deleted in google calendar, delete in Personio
        deletePersonioTimeOff_(personio, timeOff);
        Logger.log('Deleted TimeOff "%s" at %s for user %s', timeOff.typeName, String(timeOff.startAt), primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to delete TimeOff "%s" at %s for user %s: %s', timeOff.comment, String(timeOff.startAt), primaryEmail, e);
        return false;
    }
}


/** Update Personio TimeOff -> Google Calendar event */
function syncActionUpdateEvent_(calendar, primaryEmail, event, timeOff) {
    try {
        // Update event timestamps
        event.start.dateTime = timeOff.startAt.toISOString(timeOff.timeZoneOffset);
        event.start.date = null;
        event.start.timeZone = null;
        event.end.dateTime = timeOff.endAt.switchHour24ToHour0().toISOString(timeOff.timeZoneOffset);
        event.end.date = null;
        event.end.timeZone = null;
        calendar.update('primary', event.id, event);
        Logger.log('Updated event "%s" at %s for user %s', event.summary, event.start.dateTime || event.start.date, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to update event "%s" at %s for user %s: %s', event.summary, event.start.dateTime || event.start.date, primaryEmail, e);
        return false;
    }
}


/** Update Google Calendar event -> Personio TimeOff */
function syncActionUpdateTimeOff_(personio, calendar, primaryEmail, event, timeOff, updatedTimeOff) {
    try {
        // Create new Google Calendar Out-of-Office
        // updating by ID is not possible (according to docs AND trial and error)
        // since overlapping time-offs are not allowed (HTTP 400) no "more-safe" update operation is possible
        deletePersonioTimeOff_(personio, timeOff);
        const createdTimeOff = createPersonioTimeOff_(personio, updatedTimeOff);
        setEventPrivateProperty_(event, 'timeOffId', createdTimeOff.id);
        updateEventPersonioDeepLink_(event, createdTimeOff);
        calendar.update('primary', event.id, event);
        Logger.log('Updated TimeOff "%s" at %s for user %s', createdTimeOff.typeName, createdTimeOff.startAt, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to update TimeOff "%s" at %s for user %s: %s', timeOff.comment, String(timeOff.startAt), primaryEmail, e);
        return false;
    }
}


/** Personio TimeOff -> New Google Calendar event */
function syncActionInsertEvent_(calendar, primaryEmail, timeOffTypeDb, timeOff) {
    try {
        const newEvent = createEventFromTimeOff(timeOffTypeDb, timeOff);
        calendar.insert('primary', newEvent);
        Logger.log('Inserted Out-of-Office "%s" at %s for user %s', timeOff.typeName, String(timeOff.startAt), primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to insert Out-of-Office "%s" at %s for user %s: %s', timeOff.typeName, String(timeOff.startAt), primaryEmail, e);
        return false;
    }
}


/** Delete from Google Calendar */
function syncActionDeleteEvent_(calendar, primaryEmail, event) {
    try {
        event.status = 'cancelled';
        calendar.update('primary', event.id, event);
        Logger.log('Cancelled out-of-office "%s" at %s for user %s', event.summary, event.start.dateTime || event.start.date, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to cancel Out-Of-Office "%s" at %s for user %s: %s', event.summary, event.start.dateTime || event.start.date, primaryEmail, e);
        return false;
    }
}


/** Google Calendar -> New Personio TimeOff */
function syncActionInsertTimeOff_(personio, calendar, primaryEmail, event, newTimeOff) {
    try {
        const createdTimeOff = createPersonioTimeOff_(personio, newTimeOff);
        setEventPrivateProperty_(event, 'timeOffId', createdTimeOff.id);
        updateEventPersonioDeepLink_(event, createdTimeOff);
        calendar.update('primary', event.id, event);
        Logger.log('Inserted TimeOff "%s" at %s for user %s', createdTimeOff.typeName, String(createdTimeOff.startAt), primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to insert new TimeOff "%s" at %s for user %s: %s', event.summary, event.start.dateTime || event.start.date, primaryEmail, e);
        return false;
    }
}


/** Update the event's syncFailCount property. */
function syncActionUpdateEventFailCount_(calendar, primaryEmail, event, failCount) {
    try {
        setEventPrivateProperty_(event, 'syncFailCount', failCount);
        calendar.update('primary', event.id, event);
        return true;
    } catch (e) {
        Logger.log('Failed to set syncFailCount to %s for event at %s for user %s: %s', failCount, event.start.dateTime || event.start.date, primaryEmail, e);
        return false;
    }
}


/** Fetch Google Calendar events.
 *
 * @param {CalendarClient} calendar Initialized and authenticated Google Calendar custom client.
 * @param {string} calendarId Id of the calendar to query, or 'primary'.
 * @param {Date} timeMin Minimum time to fetch events for.
 * @param {Date} timeMax Maximum time to fetch events for.
 *
 * @return {Array<Object>} Array of Google Calendar event resources.
 */
function queryCalendarEvents_(calendar, calendarId, timeMin, timeMax) {
    const eventListParams = {
        singleEvents: true,
        showDeleted: true,
        // we use the local timezone to allow the simple Date constructor to correctly parse dates like "2016-05-16"
        timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        timeMin: timeMin.toISOString(),
        timeMax: timeMax.toISOString()
    };

    return calendar.list(calendarId, eventListParams);
}


/** Convert time-offs in Personio API format to intermediate format. */
function normalizePersonioTimeOffPeriod_(timeOffPeriod) {

    const attributes = timeOffPeriod.attributes || {};

    // parse start/end dates assuming whole-days
    const startAt = PeopleTime.fromISO8601(attributes.start_date, 0);
    const endAt = PeopleTime.fromISO8601(attributes.end_date, 24);

    // Handle "half day logic"
    // NOTE: Full addHours() calculations shouldn't be required,
    //       as we operate in the same time-zone (same day) and only care about wall-clock times.
    const halfDayStart = !!attributes.half_day_start;
    const halfDayEnd = !!attributes.half_day_end;
    if (startAt.isAtSameDay(endAt)) {
        // single-day events can have a have day in the morning or in the afternoon
        if (halfDayStart && !halfDayEnd) {
            endAt.hour = 12;
        } else if (!halfDayStart && halfDayEnd) {
            startAt.hour = 12;
        }
    } else {
        // for multi-day events, half-days can occur on each end
        if (halfDayStart) {
            startAt.hour = 12;
        }
        if (halfDayEnd) {
            endAt.hour = 12;
        }
    }

    // Web UI created Personio created_at/updated_at timestamps are shifted +1h.
    // see: https://community.personio.com/attendances-absences-87/absences-api-updated-at-and-created-at-timestamp-values-invalid-1743
    //   - checking for created_by isn't enough: later updates by the UI won't change the "created_by" field but still set a shifted updated_at :/
    //   - while this issue persists, one has to make sure this runs at least every 30 minutes and catch "updated_at" values that lie in the future
    const updatedAt = new Date(attributes.updated_at);
    if (attributes.created_by !== 'API' || +updatedAt >= Date.now()) {
        Util.addDateMillies(updatedAt, -1 * 60 * 60 * 1000);
    }

    return {
        id: attributes.id,
        startAt: startAt,
        endAt: endAt,
        typeId: attributes.time_off_type?.attributes.id,
        typeName: attributes.time_off_type?.attributes.name,
        comment: attributes.comment,
        status: attributes.status,
        timeZoneOffset: Util.getTimeZoneOffset(attributes.start_date),
        updatedAt: updatedAt,
        employeeId: attributes.employee?.attributes.id?.value,
        email: (attributes.employee?.attributes.email?.value || '').trim()
    };
}

/** Convert time-offs in Personio API format to intermediate format and index by ID.
 *
 * @param {PersonioClientV1} personio Initialized and authenticated PersonioClientV1 instance.
 * @param {Date} timeMin Minimum TimeOffPeriod fetch time.
 * @param {Date} timeMax Maximum TimeOffPeriod fetch time.
 * @param {number} employeeId Employee ID to query TimeOffPeriods for.
 *
 * @return {Object} Normalized TimeOff structures indexed by TimeOffPeriod ID.
 */
function queryPersonioTimeOffs_(personio, timeMin, timeMax, employeeId) {
    const params = {
        start_date: timeMin.toISOString().split("T")[0],
        end_date: timeMax.toISOString().split("T")[0],
        ['employees[]']: employeeId
    };
    const timeOffPeriods = personio.getPersonioJson('/company/time-offs' + UrlFetchJsonClient.buildQuery(params));
    const timeOffs = {};
    for (const timeOffPeriod of timeOffPeriods) {
        const timeOff = normalizePersonioTimeOffPeriod_(timeOffPeriod);
        timeOffs[timeOff.id] = timeOff;
    }

    return timeOffs;
}


/** Construct a matching TimeOff structure for a Google Calendar event. */
function convertOutOfOfficeToTimeOff_(timeOffTypeDb, employee, event, existingTimeOff) {

    let timeOffType = timeOffTypeDb.findByKeywordMatch(event.summary || '');
    if (!timeOffType) {
        if (!existingTimeOff) {
            return undefined;
        }

        const previousType = timeOffTypeDb.findById(existingTimeOff.typeId);
        if (!previousType) {
            return undefined;
        }

        timeOffType = previousType;
    }

    const halfDaysAllowed = !!timeOffType.attributes?.half_day_requests_enabled;
    if (!existingTimeOff) {
        // if we consider creating a new time-off (previously untracked),
        // we ignore events which do not cover a certain minimum of hours
        const minimumDurationMillies = halfDaysAllowed
            ? MINIMUM_OUT_OF_OFFICE_DURATION_HALF_DAY_MILLIES
            : MINIMUM_OUT_OF_OFFICE_DURATION_WHOLE_DAY_MILLIES;
        if (+(new Date(event.end.dateTime || event.end.date)) - +(new Date(event.start.dateTime || event.start.date)) < minimumDurationMillies) {
            return undefined;
        }
    }

    const startAt = PeopleTime.fromISO8601(event.start.dateTime || event.start.date).normalizeHalfDay(false, halfDaysAllowed);
    const endAt = PeopleTime.fromISO8601(event.end.dateTime || event.end.date).normalizeHalfDay(true, halfDaysAllowed);
    const timeZoneOffset = (event.start.dateTime || event.end.dateTime)
        ? Util.getTimeZoneOffset(event.start.dateTime || event.end.dateTime)
        : (((new Date()).getTimezoneOffset() * -1) * 60 * 1000);

    return {
        startAt: startAt,
        endAt: endAt,
        typeId: timeOffType.attributes.id,
        typeName: timeOffType.attributes.name,
        timeZoneOffset: timeZoneOffset,
        comment: event.summary.replace(' [synced]', ''),
        updatedAt: new Date(event.updated),
        employeeId: employee.attributes.id.value,
        email: employee.attributes.email.value,
        status: existingTimeOff ? existingTimeOff.status : 'pending'
    };
}


/** Delete the specified TimeOff from Personio. */
function deletePersonioTimeOff_(personio, timeOff) {
    return personio.fetchJson(`/company/time-offs/${timeOff.id.toFixed(0)}`, {
        method: 'delete'
    });
}


/** Insert a new Personio TimeOff. */
function createPersonioTimeOff_(personio, timeOff) {

    const isMultiDay = !timeOff.startAt.isAtSameDay(timeOff.endAt);
    const halfDayStart = isMultiDay ? timeOff.startAt.isHalfDay() : timeOff.endAt.isHalfDay();
    const halfDayEnd = isMultiDay ? timeOff.endAt.isHalfDay() : timeOff.startAt.isHalfDay();

    const result = personio.fetchJson('/company/time-offs', {
        method: 'post',
        payload: {
            employee_id: timeOff.employeeId.toFixed(0),
            time_off_type_id: timeOff.typeId.toFixed(0),
            // Reminder: there may be adjustments needed to handle timezones better to avoid switching days, here
            start_date: timeOff.startAt.toISODate(),
            end_date: timeOff.endAt.toISODate(),
            half_day_start: halfDayStart ? "1" : "0",
            half_day_end: halfDayEnd ? "1" : "0",
            comment: timeOff.comment,
            skip_approval: timeOff.status === 'approved' ? "1" : "0"
        }
    });

    if (!result?.data) {
        throw new Error(`Failed to create TimeOffPeriod: ${JSON.stringify(result)}`);
    }

    return normalizePersonioTimeOffPeriod_(result.data);
}


/** Create a new Gcal event to mirror the specified TimeOff. */
function createEventFromTimeOff(timeOffTypeDb, timeOff) {
    const newEvent = {
        kind: 'calendar#event',
        iCalUID: `${Util.generateUUIDv4()}-p-${timeOff.id}-sync-timeoffs@giantswarm.io`,
        start: {
            dateTime: timeOff.startAt.toISOString(timeOff.timeZoneOffset)
        },
        end: {
            dateTime: timeOff.endAt.switchHour24ToHour0().toISOString(timeOff.timeZoneOffset)
        },
        eventType: 'outOfOffice',  // left here for completeness, still not fully supported by Google Calendar
        extendedProperties: {
            private: {
                timeOffId: timeOff.id
            }
        },
        summary: `${timeOff.comment} [synced]`
    };

    // if we can't guess the corresponding time-off-type, prefix the event summary with its name
    const guessedType = timeOffTypeDb.findByKeywordMatch(newEvent.summary);
    if (!guessedType || guessedType.attributes.id !== timeOff.typeId) {
        const keyword = TimeOffTypeDb.extractKeyword(timeOff.typeName);
        newEvent.summary = `${keyword}: ${newEvent.summary}`;
    }

    // add a link to the correct Personio absence calendar page
    updateEventPersonioDeepLink_(newEvent, timeOff);

    return newEvent;
}


/** Generate and add a link to the Personio Absence Calendar page for the specified TimeOffPeriod to this event. */
function updateEventPersonioDeepLink_(event, timeOff) {
    const employeeIdSafe = (+timeOff.employeeId).toFixed(0);
    const timeOffTypeId = (+timeOff.typeId).toFixed(0);
    const year = (+timeOff.startAt.year).toFixed(0);
    const month = (+timeOff.startAt.month).toFixed(0);
    const deepLink = `https://giant-swarm.personio.de/time-off/employee/${employeeIdSafe}/monthly?absenceTypeId=${timeOffTypeId}&month=${month}&year=${year}`;
    const deepLinkHtml = `<a href="${deepLink}">Show in Personio</a>`;
    if (!event.description) {
        event.description = deepLinkHtml;
    } else if (event.description) {
        const deepLinkHtmlPattern = /<a.*href="https:\/\/giant-swarm.personio.de\/time-off\/employee\/.*".?>.+<\/a>/g;
        if (deepLinkHtmlPattern.test(event.description)) {
            event.description.replace(deepLinkHtmlPattern, deepLinkHtml);
        } else {
            event.description = `${event.description}<br/>${deepLinkHtml}`;
        }
    }
}


/** Set a private property on an event. */
function setEventPrivateProperty_(event, key, value) {
    const props = event.extendedProperties ? event.extendedProperties : event.extendedProperties = {};
    const privateProps = props.private ? props.private : props.private = {};
    privateProps[key] = value;
}
