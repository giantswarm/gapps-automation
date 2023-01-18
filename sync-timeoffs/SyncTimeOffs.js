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

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'syncTimeOffs';

/** The minimum duration of a half-day out-of-office event. */
const MINIMUM_OUT_OF_OFFICE_DURATION_HALF_DAY_MILLIES = 3 * 60 * 60 * 1000;  // half-day >= 3h

/** The minimum duration of a whole day or longe out-of-office event. */
const MINIMUM_OUT_OF_OFFICE_DURATION_WHOLE_DAY_MILLIES = 6 * 60 * 60 * 1000; // whole-day >= 6h




/** People time is supposed to make handling rough local times a bit easier
 * without having to pull in Joda Time or similar libs.
 */
class PeopleTime {
    constructor(year, month, day, hour) {
        this.year = year;
        this.month = month;
        this.day = day;
        this.hour = hour;
    }

    isFirstHalfDay() {
        return this.hour < 12;
    }

    /** Get a PeopleTime instance from a ISO8601 timestamp like (ie. "2016-05-13" or "2018-09-24T20:15:13.123+01:00" .*/
    fromISO8601(ts) {
        // we care about local date-time
        if (ts) {
            const dateAndTime = ts.trim().split('T');
            const ymd = dateAndTime[0].split('-').map(field => +field);
            const hour = dateAndTime[1] ? +(dateAndTime[1].substring(0, 2)) : 0;

            return new PeopleTime(ymd[0], ymd[1], ymd[2], hour);
        }

        return undefined;
    }
}


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

        if (!isEmailDomainAllowed(email)) {
            Logger.log('Not synchronizing employee with email %s: mail domain not white-listed', email);
            continue;
        }

        // we keep operating if handling calendar of a single user fails
        try {
            syncUserTimeOffs_(personio, employee, epoch, timeOffTypes, fetchTimeMin, fetchTimeMax)
        } catch (e) {
            Logger.log('Failed to sync time-offs/out-of-offices of user %s: %s', email, e);
            firstError = firstError || e;
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

/** Get the Personio token. */
function getPersonioCreds_() {
    const credentialFields = (getScriptProperties_().getProperty(PERSONIO_TOKEN_KEY) || '|')
        .split('|')
        .map(field => field.trim());

    return {clientId: credentialFields[0], clientSecret: credentialFields[1]};
}


/** Subscribe a single account to all the specified calendars. */
function syncUserTimeOffs_(personio, employee, epoch, timeOffTypes, fetchTimeMin, fetchTimeMax) {

    const primaryEmail = employee.attributes.email.value;

    // We ignore working on events/time-offs that were updated too recently
    const updateDeadZoneMillies = 120 * 1000; // 120 seconds, to avoid races and workaround lack of transactions
    const updateMax = Util.addDateMillies(new Date(epoch), -updateDeadZoneMillies);

    // load timeOffs
    const params = {
        start_date: fetchTimeMin.toISOString().split("T")[0],
        end_date: fetchTimeMax.toISOString().split("T")[0]
    };
    params['employees[]'] = employee.attributes.id.value;  // WTF Personio, ugly query params
    const timeOffs = indexPersonioTimeOffs_(personio.getPersonioJson('/company/time-offs' + UrlFetchJsonClient.buildQuery(params)));

    const calendar = CalendarClient.withImpersonatingService(getServiceAccountCredentials_(), primaryEmail);
    const eventListParams = {
        singleEvents: true,
        showDeleted: true,
        timeMin: fetchTimeMin.toISOString(),
        timeMax: fetchTimeMax.toISOString()
    };
    const allEvents = calendar.list('primary', eventListParams);
    for (const event of allEvents) {
        const eventUpdatedAt = new Date(event.updated);
        const timeOffId = event.extendedProperties?.private?.timeOffId;
        if (timeOffId) {

            // we handle this time-off
            const timeOff = timeOffs[timeOffId];
            if (timeOff) {

                delete timeOffs[timeOffId];

                if (timeOff.updatedAt > updateMax || eventUpdatedAt > updateMax) {
                    // dead zone
                    continue;
                }

                if (event.status === 'cancelled') {
                    syncActionDeleteTimeOff_(personio, primaryEmail, timeOff);
                } else {
                    // need to convert to be able to compare start/end timestamps (Personio is whole-day/half-day only)
                    const updatedTimeOff = convertOutOfOfficeToTimeOff_(timeOffTypes, employee, event, timeOff);
                    if (updatedTimeOff
                        && (Math.abs(+updatedTimeOff.startAt - +timeOff.startAt) > 2000 || Math.abs(+updatedTimeOff.endAt - +timeOff.endAt) > 2000)) {
                        // start/end timestamps differ by more than 2s, now compare which (Personio/Gcal) is more up-to-date
                        if (+timeOff.updatedAt > +(eventUpdatedAt)) {
                            console.log("events differed OoO is newer: ", event, updatedTimeOff, timeOff);
                            syncActionUpdateEvent_(calendar, primaryEmail, event, timeOff);
                        } else {
                            console.log("events differed TimeOff is newer: ", event, updatedTimeOff, timeOff);
                            syncActionUpdateTimeOff_(personio, calendar, primaryEmail, event, timeOff, updatedTimeOff);
                        }
                    }
                }
            } else if (event.status !== 'cancelled') {
                // check for dead zone
                if (eventUpdatedAt <= updateMax) {
                    syncActionDeleteEvent_(personio, calendar, primaryEmail, event);
                }
            }
        } else if (event.status !== 'cancelled') {

            if (event.summary.trim() === 'Out of office') {
                // ignore events created by Cronofy for now (summary always exactly "Out of office")
                continue;
            }

            // check for dead zone
            if (eventUpdatedAt <= updateMax) {
                const newTimeOff = convertOutOfOfficeToTimeOff_(timeOffTypes, employee, event, undefined);
                if (newTimeOff) {
                    syncActionInsertTimeOff_(personio, calendar, primaryEmail, event, newTimeOff);
                }
            }
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
    } catch (e) {
        Logger.log('Failed to delete TimeOff "%s" at %s for user %s: %s', timeOff.comment, timeOff.startAt, primaryEmail, e);
    }
}

/** Update Personio TimeOff -> Google Calendar event */
function syncActionUpdateEvent_(calendar, primaryEmail, event, timeOff) {
    try {
        // Update event timestamps
        event.start.dateTime = timeOff.startAt.toISOString();
        event.end.dateTime = timeOff.endAt.toISOString();
        calendar.update('primary', event.id, event);
        Logger.log('Updated event "%s" at %s for user %s', event.summary, event.start.dateTime, primaryEmail);
    } catch (e) {
        Logger.log('Failed to update event "%s" at %s for user %s: %s', event.summary, event.start.dateTime, primaryEmail, e);
    }
}

/** Update Google Calendar event -> Personio TimeOff */
function syncActionUpdateTimeOff_(personio, calendar, primaryEmail, event, timeOff, updatedTimeOff) {
    try {
        // Create new Google Calendar Out-of-Office
        // updating by ID is not possible (according to docs AND trial and error)
        // since overlapping time-offs are not allowed (HTTP 400) no "more-safe" update operation is possible
        deletePersonioTimeOff_(personio, timeOff);
        insertPersonioTimeOffAndUpdateEvent_(calendar, personio, updatedTimeOff, event);
        Logger.log('Updated TimeOff "%s" at %s for user %s', updatedTimeOff.typeName, updatedTimeOff.startAt, primaryEmail);
    } catch (e) {
        Logger.log('Failed to update TimeOff "%s" at %s for user %s: %s', timeOff.comment, timeOff.startAt, primaryEmail, e);
    }
}

/** Personio TimeOff -> New Google Calendar event */
function syncActionInsertEvent_(calendar, primaryEmail, timeOffTypes, timeOff) {
    try {
        const newEvent = createEventFromTimeOff(timeOffTypes, timeOff);
        calendar.insert('primary', newEvent);
        Logger.log('Created Out-of-Office "%s" at %s for user %s', newEvent.summary, timeOff.startAt, primaryEmail);
    } catch (e) {
        Logger.log('Failed to create Out-of-Office "%s" at %s for user %s: %s', timeOff.comment, timeOff.startAt, primaryEmail, e);
    }
}

/** Delete from Google Calendar */
function syncActionDeleteEvent_(personio, calendar, primaryEmail, event) {
    try {
        event.status = 'cancelled';
        calendar.update('primary', event.id, event);
        Logger.log('Cancelled out-of-office "%s" at %s for user %s', event.summary, event.startAt, primaryEmail);
    } catch (e) {
        Logger.log('Failed to cancel Out-Of-Office "%s" at %s for user %s: %s', event.summary, event.start.dateTime, primaryEmail, e);
    }
}

/** Google Calendar -> New Personio TimeOff */
function syncActionInsertTimeOff_(personio, calendar, primaryEmail, event, newTimeOff) {
    try {
        insertPersonioTimeOffAndUpdateEvent_(calendar, personio, newTimeOff, event);
        Logger.log('Inserted TimeOff "%s" at %s for user %s', newTimeOff.typeName, newTimeOff.startAt, primaryEmail);
    } catch (e) {
        Logger.log('Failed to create new TimeOff "%s" at %s for user %s: %s', event.summary, event.start.dateTime, primaryEmail, e);
    }
}

/** Convert time-offs in Personio API format to intermediate format and index by ID. */
function indexPersonioTimeOffs_(timeOffs) {
    const userTimeOffs = {};
    for (const item of timeOffs) {
        const attributes = item.attributes || {};

        const startAt = new Date(attributes.start_date);
        const endAt = new Date(attributes.end_date);
        const timeZoneOffset = Util.getTimestampOffset(attributes.start_date);
        const daysCount = 0.0 + attributes.days_count;
        const halfDayStart = !!attributes.half_day_start;
        const halfDayEnd = !!attributes.half_day_end;

        // TODO Include work schedule info for correct translation, shouldn't matter for us now
        const halfDayMillies = 12 * 60 * 60 * 1000;
        const fullDayMillies = (24 * 60 * 60 * 1000) - 1000; // minus one second to stay at same day
        if (daysCount > 1.0) {
            Util.addDateMillies(startAt, halfDayStart ? halfDayMillies : 0);
            Util.addDateMillies(endAt, halfDayEnd ? halfDayMillies : fullDayMillies);
        } else {
            Util.addDateMillies(startAt, halfDayEnd && !halfDayStart ? halfDayMillies : 0);
            Util.addDateMillies(endAt, halfDayStart && !halfDayEnd ? halfDayMillies : fullDayMillies);
        }

        // Web UI created Personio created_at/updated_at timestamps are shifted +1h.
        // see: https://community.personio.com/attendances-absences-87/absences-api-updated-at-and-created-at-timestamp-values-invalid-1743
        const updatedAt = (attributes.created_by !== 'API')
            ? Util.addDateMillies(new Date(attributes.updated_at), -1 * 60 * 60 * 1000) // -1h
            : new Date(attributes.updated_at);

        const timeOff = {
            id: attributes.id,
            startAt: startAt,
            endAt: endAt,
            typeId: attributes.time_off_type?.attributes.id,
            typeName: attributes.time_off_type?.attributes.name,
            timezoneOffset: timeZoneOffset,
            comment: attributes.comment,
            status: attributes.status,
            updatedAt: updatedAt,
            employeeId: attributes.employee?.attributes.id?.value,
            email: (attributes.employee?.attributes.email?.value || '').trim()
        };

        userTimeOffs[timeOff.id] = timeOff;
    }

    return userTimeOffs;
}

/** Guess timeOffType from a text (description of out-of-office events usually). */
function guessTimeOffType_(timeOffTypes, event) {

    const start = new Date(event.start.dateTime);
    const end = new Date(event.end.dateTime);
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

    const start = new Date(event.start.dateTime);
    const end = new Date(event.end.dateTime);
    if (!existingTimeOff) {
        // if we create a new time-off, we don't accept events which do not cover a half-day/whole-day respectively
        const minimumDurationMillies = timeOffType.attributes?.half_day_requests_enabled
            ? MINIMUM_OUT_OF_OFFICE_DURATION_HALF_DAY_MILLIES
            : MINIMUM_OUT_OF_OFFICE_DURATION_WHOLE_DAY_MILLIES;
        if (end - start < minimumDurationMillies) {
            return undefined;
        }
    }

    const startDateParts = event.start.dateTime.split('T');
    const endDateParts = event.end.dateTime.split('T');
    const startDate = startDateParts[0];
    const endDate = endDateParts[0];
    const startHourRaw = +startDateParts[1].split(':')[0];
    const endHourRaw = +endDateParts[1].split(':')[0];
    const startHour = (startHourRaw > 12 && startHourRaw <= 22) ? 12 : 0;
    const endHour = (endHourRaw >= 2 && endHourRaw <= 12) ? 12 : 23;

    // different day?
    // This is all based on Berlin/Paris TZ (+01:00), we may have to do some minor adjustments in future!
    // (Personio needs to become more "conformant"/"precise" in their API work.)

    // default: whole-day
    let half_day_start = false;
    let half_day_end = false;

    if (timeOffType.attributes?.half_day_requests_enabled) {
        if (startDate !== endDate) {
            if (startHour >= 12) {
                half_day_start = true;
            }
            if (endHour <= 12) {
                half_day_end = true;
            }
        } else {
            if (startHour >= 12) {
                half_day_end = true;
            } else if (endHour <= 12) {
                half_day_start = true;
            }
        }
    }

    const startAt = new Date(start);
    startAt.setHours(startHour, 0, 0, 0);
    const endAt = new Date(end);
    if (endHour === 23) {
        endAt.setHours(23, 59, 59, 0);
    } else {
        endAt.setHours(endHour, 0, 0, 0);
    }

    // We try to adjust for Personio's weird/incorrect ISO8601 timestamp handling here.
    // TODO Can we handle timestamps correctly somehow?
    const timeZoneOffset = Util.getTimestampOffset(event.start.dateTime);
    //Util.addDateMillies(startAt, timeZoneOffset);
    // event.end maybe at 0 o'clock on the _next day_
    if (endAt.getHours() === 24 && endAt.getMinutes() === 0 && endAt.getSeconds() === 0) {
        Util.addDateMillies(endAt, -1000); // minus one second to roll back into previous day
    }

    return {
        startAt: startAt,
        endAt: endAt,
        typeId: timeOffType.attributes.id,
        typeName: timeOffType.attributes.name,
        timezoneOffset: timeZoneOffset,
        comment: event.summary.replace(' [synced]', ''),
        updatedAt: new Date(event.updated),
        employeeId: employee.attributes.id.value,
        email: employee.attributes.email.value,
        half_day_start: half_day_start,
        half_day_end: half_day_end
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
    return personio.fetchJson('/company/time-offs', {
        method: 'post',
        payload: {
            employee_id: timeOff.employeeId.toFixed(0),
            time_off_type_id: timeOff.typeId.toFixed(0),
            // Reminder: there may be adjustments needed to handle timezones better to avoid switching days, here
            start_date: timeOff.startAt.toISOString().split('T')[0],
            end_date: timeOff.endAt.toISOString().split('T')[0],
            half_day_start: timeOff.half_day_start ? "1" : "0",
            half_day_end: timeOff.half_day_end ? "1" : "0",
            comment: timeOff.comment,
            skip_approval: timeOff.status === 'approved'
        }
    });
}


/** Create a new Gcal event to mirror the specified TimeOff. */
function createEventFromTimeOff(timeOffTypes, timeOff) {
    const newEvent = {
        kind: 'calendar#event',
        start: {
            dateTime: timeOff.startAt.toISOString()
        },
        end: {
            dateTime: timeOff.endAt.toISOString()
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
    if (guessTimeOffType_(timeOffTypes, newEvent) === undefined) {
        newEvent.summary = timeOff.typeName + ': ' + newEvent.summary
    }

    return newEvent;
}


/** Insert Personio time-off and update "timeOffId" property of existing Gcal event. */
function insertPersonioTimeOffAndUpdateEvent_(calendar, personio, newTimeOff, event) {
    const createdTimeOff = createPersonioTimeOff_(personio, newTimeOff)?.data;
    if (!event.extendedProperties) {
        event.extendedProperties = {};
    }
    if (!event.extendedProperties.private) {
        event.extendedProperties.private = {};
    }
    event.extendedProperties.private.timeOffId = createdTimeOff.attributes.id;
    calendar.update('primary', event.id, event);
}
