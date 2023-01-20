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

    const allowedDomains = (getScriptProperties_().getProperty(ALLOWED_DOMAINS_KEY) || '')
        .split(',')
        .map(d => d.trim());
    const isEmailDomainAllowed = email => allowedDomains.includes(email.substring(email.lastIndexOf('@') + 1));
    Logger.log('Configured to handle users on domains: %s', allowedDomains);

    // all timing related activities are relative to this EPOCH
    const epoch = new Date();

    // how far back to sync events/time-offs
    const lookbackMillies = -Math.round(getLookbackDays_() * 24 * 60 * 60 * 1000);
    // how far into the future to sync events/time-offs
    const lookaheadMillies = Math.round(getLookaheadDays_() * 24 * 60 * 60 * 1000);
    // how many Personio action retries per event?
    const maxFailCount = getMaxSyncFailCount_();

    const fetchTimeMin = Util.addDateMillies(new Date(epoch), lookbackMillies);
    fetchTimeMin.setUTCHours(0, 0, 0, 0); // round down to start of day
    const fetchTimeMax = Util.addDateMillies(new Date(epoch), lookaheadMillies);
    fetchTimeMax.setUTCHours(24, 0, 0, 0); // round up to end of day
    Logger.log('Syncing events between %s and %s', fetchTimeMin.toISOString(), fetchTimeMax.toISOString());

    const personioCreds = getPersonioCreds_();
    const personio = PersonioClientV1.withApiCredentials(personioCreds.clientId, personioCreds.clientSecret);

    // load timeOffTypes
    const timeOffTypes = personio.getPersonioJson('/company/time-off-types');

    let firstError = null;

    for (const employee of personio.getPersonioJson('/company/employees')) {

        const email = employee.attributes.email.value;
        if (email !== 'jonas@giantswarm.io') {
            continue;
        }

        if (isEmailDomainAllowed(email)) {
            // we keep operating if handling calendar of a single user fails
            try {
                const calendar = CalendarClient.withImpersonatingService(getServiceAccountCredentials_(), email);
                syncTimeOffs_(personio, calendar, employee, epoch, timeOffTypes, fetchTimeMin, fetchTimeMax, maxFailCount)
            } catch (e) {
                Logger.log('Failed to sync time-offs/out-of-offices of user %s: %s', email, e);
                firstError = firstError || e;
            }
        } else {
            Logger.log('Not synchronizing employee with email %s: domain not white-listed', email);
        }
    }

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


/** Subscribe a single account to all the specified calendars. */
function syncTimeOffs_(personio, calendar, employee, epoch, timeOffTypes, fetchTimeMin, fetchTimeMax, maxFailCount) {

    const primaryEmail = employee.attributes.email.value;

    // We ignore working on events/time-offs that were updated too recently
    const updateDeadZoneMillies = 120 * 1000; // 120 seconds, to avoid races and workaround lack of transactions
    const updateMax = Util.addDateMillies(new Date(epoch), -updateDeadZoneMillies);

    // load timeOffs indexed by ID
    const timeOffs = queryPersonioTimeOffs_(personio, fetchTimeMin, fetchTimeMax, employee.attributes.id.value);

    const allEvents = queryCalendarEvents_(calendar, 'primary', fetchTimeMin, fetchTimeMax);
    for (const event of allEvents) {
        let failed = false;
        const failCount = event.extendedProperties?.private?.syncFailCount || 0;
        const eventUpdatedAt = new Date(event.updated);
        const timeOffId = event.extendedProperties?.private?.timeOffId;
        if (timeOffId) {

            // we handle this time-off
            const timeOff = timeOffs[timeOffId];
            if (timeOff) {

                delete timeOffs[timeOffId];

                if (timeOff.updatedAt > updateMax || eventUpdatedAt > updateMax || failCount > maxSyncFail) {
                    // dead zone
                    continue;
                }

                if (event.status === 'cancelled') {
                    failed = !syncActionDeleteTimeOff_(personio, primaryEmail, timeOff) || failed;
                } else {
                    // need to convert to be able to compare start/end timestamps (Personio is whole-day/half-day only)
                    const updatedTimeOff = convertOutOfOfficeToTimeOff_(timeOffTypes, employee, event, timeOff);
                    if (updatedTimeOff && (!updatedTimeOff.startAt.equals(timeOff.startAt) || !updatedTimeOff.endAt.equals(timeOff.endAt))) {
                        // start/end timestamps differ by more than 2s, now compare which (Personio/Gcal) is more up-to-date
                        if (+timeOff.updatedAt < +(eventUpdatedAt)) {
                            console.log("events differed OoO is newer: ", event, updatedTimeOff, timeOff);
                            syncActionUpdateEvent_(calendar, primaryEmail, event, timeOff);
                        } else {
                            console.log("events differed TimeOff is newer: ", event, updatedTimeOff, timeOff);
                            failed = !syncActionUpdateTimeOff_(personio, calendar, primaryEmail, event, timeOff, updatedTimeOff) || failed;
                        }
                    }
                }
            } else if (event.status !== 'cancelled') {
                // check for dead zone
                // we allow event cancellation even in case maxSyncFail was reached
                if (eventUpdatedAt <= updateMax) {
                    syncActionDeleteEvent_(calendar, primaryEmail, event);
                }
            }
        } else if (event.status !== 'cancelled') {

            if (event.summary.trim() === 'Out of office') {
                // ignore events created by Cronofy for now (summary always exactly "Out of office")
                continue;
            }

            // check for dead zone
            if (eventUpdatedAt <= updateMax || failCount > maxSyncFail) {
                const newTimeOff = convertOutOfOfficeToTimeOff_(timeOffTypes, employee, event, undefined);
                if (newTimeOff) {
                    failed = !syncActionInsertTimeOff_(personio, calendar, primaryEmail, event, newTimeOff) || failed;
                }
            }
        }

        // register failure for Personio client circuit breaker
        if (failed) {
            syncActionUpdateEventFailCount_(calendar, primaryEmail, event, failCount + 1);
        }
    }

    // Handle each remaining time-off, not handled above
    for (const timeOff of Object.values(timeOffs)) {
        // check for dead zone
        if (timeOff.updatedAt <= updateMax) {
            syncActionInsertEvent_(calendar, primaryEmail, timeOffTypes, timeOff);
        }
    }
}


/** Delete Personio TimeOffs for cancelled Google Calendar events */
function syncActionDeleteTimeOff_(personio, primaryEmail, timeOff) {
    try {
        // event deleted in google calendar, delete in Personio
        deletePersonioTimeOff_(personio, timeOff);
        Logger.log('Deleted TimeOff "%s" at %s for user %s', timeOff.typeName, timeOff.startAt, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to delete TimeOff "%s" at %s for user %s: %s', timeOff.comment, timeOff.startAt, primaryEmail, e);
        return false;
    }
}


/** Update Personio TimeOff -> Google Calendar event */
function syncActionUpdateEvent_(calendar, primaryEmail, event, timeOff) {
    try {
        // Update event timestamps
        event.start.dateTime = timeOff.startAt.toISOString(timeOff.timeZoneOffset);
        event.end.dateTime = timeOff.endAt.toISOString(timeOff.timeZoneOffset);
        calendar.update('primary', event.id, event);
        Logger.log('Updated event "%s" at %s for user %s', event.summary, event.start.dateTime, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to update event "%s" at %s for user %s: %s', event.summary, event.start.dateTime, primaryEmail, e);
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
        const createdTimeOff = createPersonioTimeOff_(personio, updatedTimeOff)?.data;
        setEventPrivateProperty_(event, 'timeOffId', createdTimeOff.attributes.id);
        calendar.update('primary', event.id, event);
        Logger.log('Updated TimeOff "%s" at %s for user %s', updatedTimeOff.typeName, updatedTimeOff.startAt, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to update TimeOff "%s" at %s for user %s: %s', timeOff.comment, timeOff.startAt, primaryEmail, e);
        return false;
    }
}


/** Personio TimeOff -> New Google Calendar event */
function syncActionInsertEvent_(calendar, primaryEmail, timeOffTypes, timeOff) {
    try {
        const newEvent = createEventFromTimeOff(timeOffTypes, timeOff);
        calendar.insert('primary', newEvent);
        Logger.log('Inserted Out-of-Office "%s" at %s for user %s', timeOff.typeName, timeOff.startAt, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to insert Out-of-Office "%s" at %s for user %s: %s', timeOff.typeName, timeOff.startAt, primaryEmail, e);
        return false;
    }
}


/** Delete from Google Calendar */
function syncActionDeleteEvent_(calendar, primaryEmail, event) {
    try {
        event.status = 'cancelled';
        calendar.update('primary', event.id, event);
        Logger.log('Cancelled out-of-office "%s" at %s for user %s', event.summary, event.start.dateTime, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to cancel Out-Of-Office "%s" at %s for user %s: %s', event.summary, event.start.dateTime, primaryEmail, e);
        return false;
    }
}


/** Google Calendar -> New Personio TimeOff */
function syncActionInsertTimeOff_(personio, calendar, primaryEmail, event, newTimeOff) {
    try {
        const createdTimeOff = createPersonioTimeOff_(personio, newTimeOff)?.data;
        setEventPrivateProperty_(event, 'timeOffId', createdTimeOff.attributes.id);
        calendar.update('primary', event.id, event);
        Logger.log('Inserted TimeOff "%s" at %s for user %s', newTimeOff.typeName, newTimeOff.startAt, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to insert new TimeOff "%s" at %s for user %s: %s', event.summary, event.start.dateTime, primaryEmail, e);
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
        Logger.log('Failed to set syncFailCount to %s for event at %s for user %s: %s', failCount, event.start.dateTime, primaryEmail, e);
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
        timeMin: timeMin.toISOString(),
        timeMax: timeMax.toISOString()
    };
    return calendar.list(calendarId, eventListParams);
}

/** Guess timeOffType from a text (description of out-of-office events usually). */
function guessTimeOffType_(timeOffTypes, event) {

    const text = ((event?.source?.title || '') + ' '
        + (event?.location || '') + ' '
        + (event?.description || '') + ' '
        + (event?.summary || '')
    ).toLowerCase();

    let matchedTimeOfType = undefined;
    for (const timeOffType of timeOffTypes) {
        const typeNameFirstWord = timeOffType.attributes.name.split(' ')[0].toLowerCase();
        if (text.toLowerCase().includes(typeNameFirstWord)) {
            matchedTimeOfType = timeOffType;
        }
    }

    return matchedTimeOfType;
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
    const updatedAtOffset = (attributes.created_by !== 'API') ? -1 * 60 * 60 * 1000 : 0; // -1h if creator isn't API
    const updatedAt = Util.addDateMillies(new Date(attributes.updated_at), updatedAtOffset)

    return {
        id: attributes.id,
        startAt: startAt,
        endAt: endAt,
        typeId: attributes.time_off_type?.attributes.id,
        typeName: attributes.time_off_type?.attributes.name,
        comment: attributes.comment,
        status: attributes.status,
        timeZoneOffset: Util.getTimestampOffset(attributes.start_date),
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
function convertOutOfOfficeToTimeOff_(timeOffTypes, employee, event, existingTimeOff) {

    let timeOffType = guessTimeOffType_(timeOffTypes, event);
    if (!timeOffType) {
        if (!existingTimeOff) {
            return undefined;
        }

        const previousType = timeOffTypes.find(t => timeOffType.attributes.id === existingTimeOff.typeId);
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
        if (+(new Date(event.end.dateTime)) - +(new Date(event.start.dateTime)) < minimumDurationMillies) {
            return undefined;
        }
    }

    const startAt = PeopleTime.fromISO8601(event.start.dateTime).normalizeHalfDay(false, halfDaysAllowed);
    const endAt = PeopleTime.fromISO8601(event.end.dateTime).normalizeHalfDay(true, halfDaysAllowed);
    const timeZoneOffset = Util.getTimestampOffset(event.start.dateTime);

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

    return personio.fetchJson('/company/time-offs', {
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
}


/** Create a new Gcal event to mirror the specified TimeOff. */
function createEventFromTimeOff(timeOffTypes, timeOff) {
    const newEvent = {
        kind: 'calendar#event',
        start: {
            dateTime: timeOff.startAt.toISOString(timeOff.timeZoneOffset)
        },
        end: {
            dateTime: timeOff.endAt.switchHour24ToHour0().toISOString(timeOff.timeZoneOffset)
        },
        eventType: 'outOfOffice',
        extendedProperties: {
            private: {
                timeOffId: timeOff.id
            }
        },
        summary: '' + timeOff.comment + ' [synced]'
    };

    // if we can't guess the corresponding time-off-type, prefix the event summary with its name
    const guessedType = guessTimeOffType_(timeOffTypes, newEvent);
    if (!guessedType || guessedType.attributes.id !== timeOff.typeId) {
        newEvent.summary = timeOff.typeName.split('(')[0].trim() + ': ' + newEvent.summary
    }

    return newEvent;
}

/** Set a private property on an event. */
function setEventPrivateProperty_(event, key, value) {
    const props = event.extendedProperties ? event.extendedProperties : event.extendedProperties = {};
    const privateProps = props.private ? props.private : props.private = {};
    privateProps[key] = value;
}