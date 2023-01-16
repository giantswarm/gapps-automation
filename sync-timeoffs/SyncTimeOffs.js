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
            syncUserTimeOffs_(personio, employee, epoch, timeOffTypes)
        } catch (e) {
            Logger.log('Failed to sync time-offs/out-of-offices of user %s: %s', email, e.message);
            firstError = firstError || e;
        }
    }

    if (firstError) {
        throw firstError;
    }
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

    const minimumDurationMillies = matchedTimeOfType?.attributes?.half_day_requests_enabled
        ? MINIMUM_OUT_OF_OFFICE_DURATION_WHOLE_DAY_MILLIES
        : MINIMUM_OUT_OF_OFFICE_DURATION_HALF_DAY_MILLIES;
    if (end - start < minimumDurationMillies) {
        matchedTimeOfType = undefined;
    }

    return matchedTimeOfType;
}


/** Create a matching timeOff for an out-of-office event. */
function convertOutOfOfficeToTimeOff_(timeOffTypes, employee, event) {

    const timeOffType = guessTimeOffType_(timeOffTypes, event);
    if (!timeOffType) {
        throw undefined;
    }

    // We try to adjust for Personio's weird/incorrect ISO8601 timestamp handling here.
    // TODO Can we handle timestamps correctly somehow?
    const timeZoneOffset = Util.getTimestampOffset(event.start.dateTime);
    const start = Util.addDateMillies(new Date(event.start.dateTime), timeZoneOffset);
    const end = new Date(event.end.dateTime); // maybe at 0 o'clock on the next day
    if ((end.getHours() === 0 || end.getHours() === 24) && end.getMinutes() === 0 && end.getSeconds() === 0)
    {
        Util.addDateMillies(end,-1000); // minus one second to roll back into previous day
    }

    const startDateParts = event.start.dateTime.split('T');
    const endDateParts = event.end.dateTime.split('T');
    const startDate = startDateParts[0];
    const endDate = endDateParts[0];
    const startHour = +startDateParts[1].split(':')[0];
    const endHour = +endDateParts[1].split(':')[0];

    // guess begin/end half-days
    const timeOff = {
        startAt: start,
        endAt: end,
        typeId: timeOffType.attributes.id,
        typeName: timeOffType.attributes.name,
        timezoneOffset: timeZoneOffset,
        comment: event.summary.replace(' [synced]', ''),
        updatedAt: new Date(event.updated),
        employeeId: employee.attributes.id.value,
        email: employee.attributes.email.value
    };

    // different day?
    // This is all based on Berlin/Paris TZ (+01:00), we may have to do some minor adjustments in future!
    // (Personio needs to become more "conformant"/"precise" in their API work.)

    // default: whole-day
    timeOff.half_day_start = false;
    timeOff.half_day_end = false;

    if (timeOffType.attributes?.half_day_requests_enabled) {
        if (startDate !== endDate) {
            timeOff.half_day_start = (startHour > 12);
            timeOff.half_day_end = (endHour < 14);
        } else {
            if (startHour > 12) {
                timeOff.half_day_end = true;
            } else if (endHour < 14) {
                timeOff.half_day_start = true;
            }
        }
    }

    return timeOff;
}


function deletePersonioTimeOff_(personio, timeOff) {
    return personio.fetchJson(`/company/time-offs/${timeOff.id.toFixed(0)}`, {
        method: 'delete'
    });
}


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


function createPersonioTimeOffAndUpdateEvent_(calendar, personio, newTimeOff, event) {
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


/** Subscribe a single account to all the specified calendars. */
function syncUserTimeOffs_(personio, employee, epoch, timeOffTypes) {

    const primaryEmail = employee.attributes.email.value;

    // TODO Make look back and look ahead and dead zone configurable (within limits)?
    const lookBackMillies = 30 * 24 * 60 * 60 * 1000; // 30 days
    const startTsBacklog = Util.addDateMillies(new Date(epoch), -lookBackMillies);

    // We ignore working on events that were updated too recently
    const updateDeadZoneMillies = 120 * 1000; // 120 seconds
    const updateMax = Util.addDateMillies(new Date(epoch), -updateDeadZoneMillies);

    // we fetch Google Calendar events further into the past, to make sure we don't miss already synced timeOffs
    const maxLookaheadMillies = Math.round(1.5 * 365 * 24 * 60 * 60 * 1000); // 1.5 years, limits event fetch count
    const safeTimeMin = Util.addDateMillies(new Date(epoch), -(lookBackMillies * 2));
    const safeTimeMax = Util.addDateMillies(new Date(epoch), maxLookaheadMillies);

    // load timeOffs
    const params = {
        start_date: startTsBacklog.toISOString().split("T")[0]
    };
    params['employees[]'] = employee.attributes.id.value;  // WTF Personio, ugly query params
    const timeOffs = indexPersonioTimeOffs_(personio.getPersonioJson('/company/time-offs' + UrlFetchJsonClient.buildQuery(params)));

    const eventListParams = {
        singleEvents: true,
        showDeleted: true,
        timeMin: safeTimeMin.toISOString(),
        timeMax: safeTimeMax.toISOString()
    };

    const calendar = CalendarClient.withImpersonatingService(getServiceAccountCredentials_(), primaryEmail);
    const eventByTimeOffId = {};
    const eventWithoutTimeOffId = [];
    const eventWithDeletedTimeOff = [];
    const allEvents = calendar.list('primary', eventListParams);
    for (const event of allEvents) {

        // We can't currently use proper Out-of-Office events, due to:
        //   https://issuetracker.google.com/issues/112063903
        //   https://issuetracker.google.com/issues/122674818
        //if (event.eventType !== 'outOfOffice') {
        // continue;
        //}

        const timeOffId = event.extendedProperties?.private?.timeOffId;
        if (timeOffId) {
            if (timeOffs[timeOffId]) {
                eventByTimeOffId[timeOffId] = event;
            } else if (event.status !== 'cancelled') {
                // TimeOff deleted in personio
                eventWithDeletedTimeOff.push(event);
            }
        } else if (event.status !== 'cancelled') {
            // we always fetch less timeOffs than events
            const start = new Date(event.start.dateTime);
            if (start > startTsBacklog && guessTimeOffType_(timeOffTypes, event) !== undefined) {
                eventWithoutTimeOffId.push(event);
            }
        }
    }

    for (const timeOff of Object.values(timeOffs)) {

        if (timeOff.updatedAt > updateMax) {
            // dead zone
            continue;
        }

        try {
            const event = eventByTimeOffId[timeOff.id];
            if (event) {
                const start = new Date(event.start.dateTime);
                const end = new Date(event.end.dateTime);
                if (event.status === 'cancelled') {
                    // event deleted in google calendar, delete in Personio
                    deletePersonioTimeOff_(personio, timeOff);
                    Logger.log('Deleted TimeOff "%s" at %s for user %s', timeOff.typeName, start, primaryEmail);
                } else if (Math.abs(+start - +timeOff.startAt) > 2000 || Math.abs(+end - +timeOff.endAt) > 2000) {
                    // start/end timestamps differ by more than 2s
                    if (+timeOff.updatedAt > +(new Date(event.updated))) {
                        // TODO restore timezone/offset here?
                        event.start.dateTime = timeOff.startAt.toISOString();
                        event.end.dateTime = timeOff.endAt.toISOString();
                        calendar.update('primary', event.id, event);
                        Logger.log('Updated Out-of-Office event "%s" at %s for user %s', event.summary, start, primaryEmail);
                    } else {
                        // Create new Google Calendar Out-of-Office
                        const updatedTimeOff = convertOutOfOfficeToTimeOff_(timeOffTypes, employee, event);
                        if (!updatedTimeOff) {
                            // TODO Send emails? How to notify users?
                            // Scenarios:
                            //    - Overlapping events
                            //    - TimeOffType not mentioned in event text: assume some default (which)?
                            Logger.log("Cannot convert event to TimeOff: %s, %s, %s", employee.attributes?.email?.value, event.start?.dateTime, event.summary);
                            continue;
                        }

                        // updating by ID is not possible (according to docs AND trial and error)
                        // since overlapping time-offs are not allowed (HTTP 400) no "more-safe" update operation is possible
                        deletePersonioTimeOff_(personio, timeOff);
                        createPersonioTimeOffAndUpdateEvent_(calendar, personio, updatedTimeOff, event);

                        Logger.log('Updated TimeOff "%s" at %s for user %s', timeOff.typeName, start, primaryEmail);
                    }
                }
            } else {
                // Create new Google Calendar Out-of-Office
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

                calendar.insert('primary', newEvent);

                Logger.log('Created out-of-office "%s" at %s for user %s', newEvent.summary, timeOff.startAt, primaryEmail);
            }
        } catch (e) {
            Logger.log('Failed to process Personio TimeOff %s for user %s: %s', timeOff.id.toFixed(0), primaryEmail, e);
        }
    }

    // Delete from Google Calendar
    for (const event of eventWithDeletedTimeOff) {

        if (new Date(event.updated) > updateMax) {
            // dead zone
            continue;
        }

        try {
            event.status = 'cancelled';
            calendar.update('primary', event.id, event);
        } catch (e) {
            Logger.log('Failed to cancel Out-Of-Office "%s" at %s for user %s: %s', event.summary, event.start.dateTime, primaryEmail, e);
        }
    }

    // Google Calendar -> New Personio TimeOff
    for (const event of eventWithoutTimeOffId) {

        if (new Date(event.updated) > updateMax) {
            // dead zone
            continue;
        }

        try {
            const newTimeOff = convertOutOfOfficeToTimeOff_(timeOffTypes, employee, event);
            if (!newTimeOff) {
                // TODO Send emails? How to notify users?
                // Scenarios:
                //    - Overlapping events
                //    - TimeOffType not mentioned in event text: assume some default (which)?
                Logger.log("Cannot convert event to TimeOff: %s, %s, %s", employee.attributes?.email?.value, event.start?.dateTime, event.summary);
                continue;
            }

            createPersonioTimeOffAndUpdateEvent_(calendar, personio, newTimeOff, event);
        } catch (e) {
            Logger.log('Failed to create new TimeOff "%s" at %s for user %s: %s', event.summary, event.start.dateTime, primaryEmail, e);
        }
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


/** Get the Personio token. */
function getPersonioCreds_() {
    const credentialFields = (getScriptProperties_().getProperty(PERSONIO_TOKEN_KEY) || '|')
        .split('|')
        .map(field => field.trim());

    return {clientId: credentialFields[0], clientSecret: credentialFields[1]};
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

        // Certain (all?) Personio timestamps are invalid. They must be shifted by "-01:00" (1 hour in the past) or the other way for uploads.
        // see: https://community.personio.com/attendances-absences-87/absences-api-updated-at-and-created-at-timestamp-values-invalid-1743?postid=5870#post5870
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
