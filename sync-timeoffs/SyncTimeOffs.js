/**
 * Synchronize Time Offs / Out-of-Office events between Personio and the domains personal Google Calendars.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */


/** The prefix for properties specific to this script in the project. */
const PROPERTY_PREFIX = 'SyncTimeOffs.';

/** Personio clientId and clientSecret, separated by '|'. */
const PERSONIO_TOKEN_KEY = PROPERTY_PREFIX + 'personioToken';

/** Service account credentials (in JSON format, as downloaded from Google Management Console). */
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

/** Black-List of keywords for skipping approval on insert (those may then require approval).
 *
 * Must be one word or a comma separated list.
 *
 * Default: null or empty
 */
const SKIP_APPROVAL_BLACKLIST_KEY = PROPERTY_PREFIX + 'skipApprovalBlackList';

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

/** A setting configuring if larger bulk requests are to be preferred.
 *
 * This can help with Personio's long per-request processing time and UrlFetchApp quota limits.
 *
 * The behavior should be configurable because:
 *  - Personio's backend is expected to change in the near future.
 *  - If the TimeOff dataset is becoming too large (or the range is configured to be greater),
 *    smaller requests may be required.
 *
 * The value can be true (default) or false.
 */
const PREFER_BULK_REQUESTS = PROPERTY_PREFIX + 'preferBulkRequests';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'syncTimeOffs';

/** The dead-zone after a sync failure where the saved original event.updated timestamp is being preferred. */
const SYNC_FAIL_UPDATED_DEAD_ZONE = 30 * 1000; // 30s

/** Do not touch failed events too often, otherwise we may exceed our quotas. */
const MAX_SYNC_FAIL_DELAY = 60 * 60 * 1000; // 1h


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
async function syncTimeOffs() {

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
    // 4:50 minutes (hard AppsScript kill comes at 6:00 minutes)
    // stay under 5 min. to ensure termination before the next instances starts if operating at 5 min. job delay
    const maxRuntimeMillies = Math.round(290 * 1000);

    const fetchTimeMin = Util.addDateMillies(new Date(epoch), lookbackMillies);
    fetchTimeMin.setUTCHours(0, 0, 0, 0); // round down to start of day
    const fetchTimeMax = Util.addDateMillies(new Date(epoch), lookaheadMillies);
    fetchTimeMax.setUTCHours(24, 0, 0, 0); // round up to end of day

    const personioCreds = getPersonioCreds_();
    const personio = PersonioClientV1.withApiCredentials(personioCreds.clientId, personioCreds.clientSecret);

    // load timeOffTypeConfig
    const timeOffTypeConfig = new TimeOffTypeConfig(await personio.getPersonioJson('/company/time-off-types'), getSkipApprovalBlackList_());

    // load and prepare list of employees to process
    const employees = await personio.getPersonioJson('/company/employees').filter(employee =>
        employee.attributes.status.value !== 'inactive' && isEmailAllowed(employee.attributes.email.value)
    );
    Util.shuffleArray(employees);

    // if bulk requests are preferred, prefetch all Personio time-offs
    const allTimeOffs = isPreferBulkRequestsEnabled_()
        ? await queryPersonioTimeOffs_(personio, fetchTimeMin, fetchTimeMax, undefined)
        : undefined;

    Logger.log('Syncing events between %s and %s for %s accounts', fetchTimeMin.toISOString(), fetchTimeMax.toISOString(), '' + employees.length);

    let firstError = null;
    let processedCount = 0;
    for (const employee of employees) {

        const email = employee.attributes.email.value;

        // we keep operating if handling calendar of a single user fails
        try {
            const calendar = await CalendarClient.withImpersonatingService(getServiceAccountCredentials_(), email);
            const isCompleted = await syncTimeOffs_(personio, calendar, employee, epoch, timeOffTypeConfig, fetchTimeMin, fetchTimeMax, maxFailCount, maxRuntimeMillies, allTimeOffs);
            if (!isCompleted) {
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


/** Utility function to unsynchronize certain events and delete the associated absence from Personio.
 *
 * @note This is a destructive operation, USE WITH UTMOST CARE!
 *
 * @param title The event title (only events whose title includes this string are de-synced).
 */
async function unsyncTimeOffs_(title) {

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

    // after how many milliseconds should this script stop by itself (to avoid forced termination/unclean state)?
    // 4:50 minutes (hard AppsScript kill comes at 6:00 minutes)
    // stay under 5 min. to ensure termination before the next instances starts if operating at 5 min. job delay
    const maxRuntimeMillies = Math.round(290 * 1000);

    const fetchTimeMin = Util.addDateMillies(new Date(epoch), lookbackMillies);
    fetchTimeMin.setUTCHours(0, 0, 0, 0); // round down to start of day
    const fetchTimeMax = Util.addDateMillies(new Date(epoch), lookaheadMillies);
    fetchTimeMax.setUTCHours(24, 0, 0, 0); // round up to end of day

    const personioCreds = getPersonioCreds_();
    const personio = PersonioClientV1.withApiCredentials(personioCreds.clientId, personioCreds.clientSecret);

    // load and prepare list of employees to process
    const employees = await personio.getPersonioJson('/company/employees').filter(employee =>
        employee.attributes.status.value !== 'inactive' && isEmailAllowed(employee.attributes.email.value)
    );
    Util.shuffleArray(employees);

    Logger.log('Unsyncing events with title containing "%s" between %s and %s for %s accounts', title, fetchTimeMin.toISOString(), fetchTimeMax.toISOString(), '' + employees.length);

    let firstError = null;
    let processedCount = 0;
    for (const employee of employees) {

        const email = employee.attributes.email.value;

        // we keep operating if handling calendar of a single user fails
        try {
            const calendar = await CalendarClient.withImpersonatingService(getServiceAccountCredentials_(), email);

            // test against dead-line first
            const deadlineTs = +epoch + maxRuntimeMillies;
            let now = Date.now();
            if (now >= deadlineTs) {
                return false;
            }

            const allEvents = await queryCalendarEvents_(calendar, 'primary', fetchTimeMin, fetchTimeMax);
            Util.shuffleArray(allEvents);

            for (const event of allEvents) {
                const timeOffId = +event.extendedProperties?.private?.timeOffId;
                if (timeOffId && (event.summary || '').includes(title)) {
                    let now = Date.now();
                    if (now >= deadlineTs) {
                        break;
                    }

                    try {
                        await deletePersonioTimeOff_(personio, {id: timeOffId});
                    } catch (e) {
                        Logger.log('Failed to remove time-off for de-synced event of user %s: %s', email, e);
                    }

                    setEventPrivateProperty_(event, 'timeOffId', undefined);

                    await calendar.update('primary', event.id, event);
                    Logger.log('De-synced event "%s" at %s for user %s', event.summary, event.start.dateTime || event.start.date, email);
                }
            }
        } catch (e) {
            Logger.log('Failed to unsync matching time-offs/out-of-offices of user %s: %s', email, e);
            firstError = firstError || e;
        }
        ++processedCount;
    }

    Logger.log('Completed de-synchronization for %s of %s accounts', '' + processedCount, '' + employees.length);

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


/** White-list (active/onboarding) employees by Team.
 *
 * Will only do something if the white-list is not empty (already enabled).
 *
 * @param {string|Array<string>} teams Team name or array of team names that should be white-listed
 *
 * @return {string} Returns the updated white-list.
 */
function whiteListTeam(teams) {
    if (!teams) {
        throw new Error('No team name(s) specified');
    }

    const teamNames = Array.isArray(teams) ? teams : [teams];
    const whiteList = getEmailWhiteList_();
    if (!Array.isArray(whiteList) || !whiteList.length) {
        throw new Error('White-list is empty (disabled), cannot add members by team');
    }

    // load and prepare list of employees to process
    const personioCreds = getPersonioCreds_();
    const personio = PersonioClientV1.withApiCredentials(personioCreds.clientId, personioCreds.clientSecret);
    const employees = personio.getPersonioJson('/company/employees')
        .filter(employee => employee.attributes.status.value !== 'inactive');

    for (const employee of employees) {
        const email = (employee.attributes?.email?.value || '').trim();
        const team = employee.attributes?.team?.value?.attributes?.name;
        if (email && team && teamNames.includes(team) && !whiteList.includes(email)) {
            whiteList.push(email);
        }
    }

    const whiteListStr = whiteList.join(',');
    TriggerUtil.setProperties({[EMAIL_WHITELIST_KEY]: whiteListStr}, false);

    return whiteListStr;
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


/** Get the TimeOffType keyword skip approval black-list (optional, leave empty to skip approval for all types). */
function getSkipApprovalBlackList_() {
    return (getScriptProperties_().getProperty(SKIP_APPROVAL_BLACKLIST_KEY) || '').trim()
        .split(',').map(keyword => keyword.trim().toLowerCase()).filter(keyword => !!keyword);
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


/** Get the "preferBulkRequests" flag. */
function isPreferBulkRequestsEnabled_() {
    const value = (getScriptProperties_().getProperty(PREFER_BULK_REQUESTS) || '').trim().toLowerCase();
    return value !== 'false' && value !== '0' && value !== 'off' && value !== 'no';
}


/** Get the Personio token. */
function getPersonioCreds_() {
    const credentialFields = (getScriptProperties_().getProperty(PERSONIO_TOKEN_KEY) || '|')
        .split('|')
        .map(field => field.trim());

    return {clientId: credentialFields[0], clientSecret: credentialFields[1]};
}


/** Small wrapper over a list of TimeOffType objects (for caching patterns). */
class TimeOffTypeConfig {
    constructor(timeOffTypes, skipApprovalBlackList) {

        this.timeOffTypes = timeOffTypes;

        // prepare patterns to guess time-off type from text
        // we dynamically create one RegExp per time-off, which we cache here (V8 can't)
        const pattern = [];
        for (let timeOffType of timeOffTypes) {
            const keyword = TimeOffTypeConfig.extractKeyword(timeOffType.attributes.name);
            pattern.push(new RegExp(`(^|[\\s:\\-[(])(${keyword})([\\s:\\-\\])]|$)`, "sim"));
        }
        this.pattern = pattern;
        this.skipApprovalBlackList = skipApprovalBlackList || [];
    }

    /** Returns the keyword for the specified TimeOffType name (field timeOffType.attributes.name). */
    static extractKeyword(typeName) {
        return (typeName || '').split(' ')[0].trim() || undefined;
    }

    /** Guess timeOffType from a text (e.g. event summary) by keyword. */
    findByKeywordMatch(text) {
        const searchText = text || '';

        for (let i = 0; i < this.timeOffTypes.length; ++i) {
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

    /** If approval may be skipped for a certain TimeOffType.
     *
     * We default to skipping approvals.
     */
    isSkippingApprovalAllowed(id) {
        const timeOffType = this.findById(id);
        if (timeOffType) {
            const keyword = TimeOffTypeConfig.extractKeyword(timeOffType.attributes.name).toLowerCase();
            return !this.skipApprovalBlackList.includes(keyword);
        }

        return true;
    }
}


/** Subscribe a single account to all the specified calendars.
 *
 * @returns true if the specified employees account was fully processed, false if processing was aborted early.
 */
async function syncTimeOffs_(personio, calendar, employee, epoch, timeOffTypeConfig, fetchTimeMin, fetchTimeMax, maxFailCount, maxRuntimeMillies, allTimeOffs) {

    // test against dead-line first
    const deadlineTs = +epoch + maxRuntimeMillies;
    let now = Date.now();
    if (now >= deadlineTs) {
        return false;
    }

    const primaryEmail = employee.attributes.email.value;

    // We ignore working on events/time-offs that were updated too recently
    const updateDeadZoneMillies = 120 * 1000; // 120 seconds, to avoid races and workaround lack of transactions
    const updateMax = Util.addDateMillies(new Date(epoch), -updateDeadZoneMillies);

    // load or filter timeOffs indexed by ID
    const employeeId = employee.attributes.id.value;
    const timeOffs = Util.isObject(allTimeOffs) ? allTimeOffs : await queryPersonioTimeOffs_(personio, fetchTimeMin, fetchTimeMax, employeeId);

    const failedSyncs = getFailedSyncs_(primaryEmail);

    const allEvents = await queryCalendarEvents_(calendar, 'primary', fetchTimeMin, fetchTimeMax);
    Util.shuffleArray(allEvents);

    let failCount = 0;
    const processedTimeOffIds = {};
    for (const event of allEvents) {

        if (now >= deadlineTs) {
            if (failCount > 0) {
                putFailedSyncs_(primaryEmail, failedSyncs);
            }
            return false;
        }

        if (failCount >= maxFailCount) {
            break;
        }

        const syncFailUpdatedAt = failedSyncs['e' + event.id];
        const eventUpdatedAt = new Date(event.updated);
        const skipDueToFail = syncFailUpdatedAt != null && syncFailUpdatedAt === +eventUpdatedAt;
        const isEventCancelled = event.status === 'cancelled';
        const timeOffId = event.extendedProperties?.private?.timeOffId;
        let isOk = true;
        if (timeOffId) {

            // we handle this time-off
            const timeOff = timeOffs[timeOffId];
            if (timeOff && timeOff.employeeId === employeeId) {

                // mark as handled
                processedTimeOffIds[timeOffId] = true;

                if (timeOff.updatedAt > updateMax || eventUpdatedAt > updateMax || skipDueToFail) {
                    // dead zone
                    continue;
                }

                if (isEventCancelled) {
                    isOk = await syncActionDeleteTimeOff_(personio, primaryEmail, timeOff);
                    now = Date.now();
                } else {
                    // need to convert to be able to compare start/end timestamps (Personio is whole-day/half-day only)
                    const updatedTimeOff = convertOutOfOfficeToTimeOff_(timeOffTypeConfig, employee, event, timeOff);
                    if (updatedTimeOff
                        && (!updatedTimeOff.startAt.equals(timeOff.startAt) || !updatedTimeOff.endAt.equals(timeOff.endAt) || updatedTimeOff.typeId !== timeOff.typeId)) {
                        // start/end timestamps differ, now check which (Personio/Google Calendar) has more recent changes
                        if (timeOff.updatedAt >= eventUpdatedAt) {
                            isOk = await syncActionUpdateEvent_(calendar, primaryEmail, event, timeOff);
                        } else {
                            isOk = await syncActionUpdateTimeOff_(personio, calendar, primaryEmail, event, timeOff, updatedTimeOff);
                        }
                        now = Date.now();
                    }
                }
            } else if (!isEventCancelled) {
                // check for dead zone
                // we allow event cancellation even in case maxFailCount was reached
                if (eventUpdatedAt <= updateMax) {
                    isOk = await syncActionDeleteEvent_(calendar, primaryEmail, event);
                    now = Date.now();
                }
            }
        } else if (!isEventCancelled) {
            // check for dead zone, ignore events created by Cronofy
            if (eventUpdatedAt <= updateMax && !skipDueToFail && !event.iCalUID.includes('cronofy.com')) {
                const newTimeOff = convertOutOfOfficeToTimeOff_(timeOffTypeConfig, employee, event, undefined);
                if (newTimeOff) {
                    isOk = await syncActionInsertTimeOff_(personio, calendar, primaryEmail, event, newTimeOff);
                    now = Date.now();
                }
            }
        }

        if (!isOk) {
            failedSyncs['e' + event.id] = +eventUpdatedAt;
            ++failCount;
        }
    }

    // Handle each remaining time-off, not handled above
    for (const timeOff of Object.values(timeOffs)) {

        if (failCount >= maxFailCount) {
            break;
        }

        const syncFailUpdatedAt = failedSyncs['t' + timeOff.id];
        const skipDueToFail = syncFailUpdatedAt != null && syncFailUpdatedAt === +timeOff.updatedAt;

        // check for dead zone
        if (!skipDueToFail && timeOff.employeeId === employeeId && timeOff.updatedAt <= updateMax && !processedTimeOffIds[timeOff.id]) {

            if (Date.now() >= deadlineTs) {
                if (failCount > 0) {
                    putFailedSyncs_(primaryEmail, failedSyncs);
                }
                return false;
            }

            if (!await syncActionInsertEvent_(calendar, primaryEmail, timeOffTypeConfig, timeOff)) {
                failedSyncs['t' + timeOff.id] = +timeOff.updatedAt;
                ++failCount;
            }
        }
    }

    if (failCount > 0) {
        putFailedSyncs_(primaryEmail, failedSyncs);
    }
    return true;
}


/** Get the cached map of event id to updatedAt mapping for failed synchronizations by email. */
function getFailedSyncs_(primaryEmail) {
    const cache = CacheService.getScriptCache();
    const item = cache.get("syncFailed_" + primaryEmail);
    return item != null ? JSON.parse(item) : {};
}


/** Cache the map of event id to updatedAt millies for failed synchronizations by email. */
function putFailedSyncs_(primaryEmail, failedSyncs) {
    const cache = CacheService.getScriptCache();
    const expiration = (MAX_SYNC_FAIL_DELAY / 2) + (MAX_SYNC_FAIL_DELAY * Math.random());
    cache.put("syncFailed_" + primaryEmail, JSON.stringify(failedSyncs), Math.round(expiration / 1000));
}


/** Delete Personio TimeOffs for cancelled Google Calendar events */
async function syncActionDeleteTimeOff_(personio, primaryEmail, timeOff) {
    try {
        // event deleted in google calendar, delete in Personio
        await deletePersonioTimeOff_(personio, timeOff);
        Logger.log('Deleted TimeOff "%s" at %s for user %s', timeOff.typeName, String(timeOff.startAt), primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to delete TimeOff "%s" at %s for user %s: %s', timeOff.comment, String(timeOff.startAt), primaryEmail, e);
        return false;
    }
}


/** Update Personio TimeOff -> Google Calendar event */
async function syncActionUpdateEvent_(calendar, primaryEmail, event, timeOff) {
    try {
        // Update event timestamps
        event.start.dateTime = timeOff.startAt.toISOString();
        event.start.date = null;
        event.start.timeZone = null;
        event.end.dateTime = timeOff.endAt.switchHour24ToHour0().toISOString();
        event.end.date = null;
        event.end.timeZone = null;

        await calendar.update('primary', event.id, event);
        Logger.log('Updated event "%s" at %s for user %s', event.summary, event.start.dateTime || event.start.date, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to update event "%s" at %s for user %s: %s', event.summary, event.start.dateTime || event.start.date, primaryEmail, e);
        return false;
    }
}


/** Update Google Calendar event -> Personio TimeOff */
async function syncActionUpdateTimeOff_(personio, calendar, primaryEmail, event, timeOff, updatedTimeOff) {
    try {
        // Create new Google Calendar Out-of-Office
        // updating by ID is not possible (according to docs AND trial and error)
        // since overlapping time-offs are not allowed (HTTP 400) no "more-safe" update operation is possible
        await deletePersonioTimeOff_(personio, timeOff);
        const createdTimeOff = await createPersonioTimeOff_(personio, updatedTimeOff);
        setEventPrivateProperty_(event, 'timeOffId', createdTimeOff.id);
        updateEventPersonioDeepLink_(event, createdTimeOff);
        if (!/ ?⇵$/.test(event.summary)) {
            event.summary = event.summary.replace(' [synced]', '') + ' ⇵';
        }
        await calendar.update('primary', event.id, event);
        Logger.log('Updated TimeOff "%s" at %s for user %s', createdTimeOff.typeName, String(createdTimeOff.startAt), primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to update TimeOff "%s" at %s for user %s: %s', timeOff.comment, String(timeOff.startAt), primaryEmail, e);
        return false;
    }
}


/** Personio TimeOff -> New Google Calendar event */
async function syncActionInsertEvent_(calendar, primaryEmail, timeOffTypeConfig, timeOff) {
    try {
        const newEvent = createEventFromTimeOff_(timeOffTypeConfig, timeOff);
        await calendar.insert('primary', newEvent);
        Logger.log('Inserted Out-of-Office "%s" at %s for user %s', timeOff.typeName, String(timeOff.startAt), primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to insert Out-of-Office "%s" at %s for user %s: %s', timeOff.typeName, String(timeOff.startAt), primaryEmail, e);
        return false;
    }
}


/** Delete from Google Calendar */
async function syncActionDeleteEvent_(calendar, primaryEmail, event) {
    try {
        event.status = 'cancelled';
        await calendar.update('primary', event.id, event);
        Logger.log('Cancelled out-of-office "%s" at %s for user %s', event.summary, event.start.dateTime || event.start.date, primaryEmail);
        return true;
    } catch (e) {
        Logger.log('Failed to cancel Out-Of-Office "%s" at %s for user %s: %s', event.summary, event.start.dateTime || event.start.date, primaryEmail, e);
        return false;
    }
}


/** Google Calendar -> New Personio TimeOff */
async function syncActionInsertTimeOff_(personio, calendar, primaryEmail, event, newTimeOff) {
    try {
        const createdTimeOff = await createPersonioTimeOff_(personio, newTimeOff);
        setEventPrivateProperty_(event, 'timeOffId', createdTimeOff.id);
        updateEventPersonioDeepLink_(event, createdTimeOff);
        if (!/ ?⇵$/.test(event.summary)) {
            event.summary = event.summary.replace(' [synced]', '') + ' ⇵';
        }
        await calendar.update('primary', event.id, event);
        Logger.log('Inserted TimeOff "%s" at %s for user %s: %s', createdTimeOff.typeName, String(createdTimeOff.startAt), primaryEmail, createdTimeOff.comment);
        return true;
    } catch (e) {
        Logger.log('Failed to insert new TimeOff "%s" at %s for user %s: %s', event.summary, event.start.dateTime || event.start.date, primaryEmail, e);
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
async function queryCalendarEvents_(calendar, calendarId, timeMin, timeMax) {
    const eventListParams = {
        singleEvents: true,
        showDeleted: true,
        timeMin: timeMin.toISOString(),
        timeMax: timeMax.toISOString()
    };

    return await calendar.list(calendarId, eventListParams);
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
async function queryPersonioTimeOffs_(personio, timeMin, timeMax, employeeId) {
    const params = {
        start_date: timeMin.toISOString().split("T")[0],
        end_date: timeMax.toISOString().split("T")[0],
        ['employees[]']: employeeId
    };

    if (employeeId != null) {
        params['employees[]'] = employeeId;
    }

    const timeOffPeriods = await personio.getPersonioJson('/company/time-offs' + UrlFetchJsonClient.buildQuery(params));
    Util.shuffleArray(timeOffPeriods);

    const timeOffs = {};
    for (const timeOffPeriod of timeOffPeriods) {
        const timeOff = normalizePersonioTimeOffPeriod_(timeOffPeriod);
        timeOffs[timeOff.id] = timeOff;
    }

    return timeOffs;
}


/** Construct a matching TimeOff structure for a Google Calendar event. */
function convertOutOfOfficeToTimeOff_(timeOffTypeConfig, employee, event, existingTimeOff) {

    // skip events created by other users
    if (event?.creator?.email && event?.creator?.email !== employee.attributes.email.value) {
        return undefined;
    }

    let timeOffType = timeOffTypeConfig.findByKeywordMatch(event.summary || '');
    if (!timeOffType && existingTimeOff) {
        const previousType = timeOffTypeConfig.findById(existingTimeOff.typeId);
        if (previousType) {
            timeOffType = previousType;
        }
    }

    if (!timeOffType && event.eventType === 'outOfOffice') {
        timeOffType = timeOffTypeConfig.findByKeywordMatch('out');
    }

    if (!timeOffType) {
        return undefined;
    }

    const halfDaysAllowed = !!timeOffType.attributes?.half_day_requests_enabled;
    const localTzOffsetStart = event.start.date ? Util.getNamedTimeZoneOffset(event.start.timeZone, new Date(event.start.date)) : undefined;
    const localTzOffsetEnd = event.end.date ? Util.getNamedTimeZoneOffset(event.end.timeZone, new Date(event.end.date)) : undefined;
    const startAt = PeopleTime.fromISO8601(event.start.dateTime || event.start.date, undefined, localTzOffsetStart)
        .normalizeHalfDay(false, halfDaysAllowed);
    const endAt = PeopleTime.fromISO8601(event.end.dateTime || event.end.date, undefined, localTzOffsetEnd)
        .normalizeHalfDay(true, halfDaysAllowed);

    const skipApproval = timeOffTypeConfig.isSkippingApprovalAllowed(timeOffType.attributes.id);

    return {
        startAt: startAt,
        endAt: endAt,
        typeId: timeOffType.attributes.id,
        typeName: timeOffType.attributes.name,
        comment: event.summary.replace(' [synced]', '').replace(/ ?⇵$/, ''),
        updatedAt: new Date(event.updated),
        employeeId: employee.attributes.id.value,
        email: employee.attributes.email.value,
        status: existingTimeOff ? existingTimeOff.status : (skipApproval ? 'approved' : 'pending')
    };
}


/** Delete the specified TimeOff from Personio. */
async function deletePersonioTimeOff_(personio, timeOff) {
    return await personio.fetchJson(`/company/time-offs/${timeOff.id.toFixed(0)}`, {
        method: 'delete'
    });
}


/** Generate payload for a personio time-off request. */
function generatePersonioTimeOffPayload_(timeOff) {
    const isMultiDay = !timeOff.startAt.isAtSameDay(timeOff.endAt);
    const halfDayStart = isMultiDay ? timeOff.startAt.isHalfDay() : timeOff.endAt.isHalfDay() && timeOff.startAt.isFirstHalfDay();
    const halfDayEnd = isMultiDay ? timeOff.endAt.isHalfDay() : timeOff.startAt.isHalfDay() && !timeOff.endAt.isFirstHalfDay();

    const payload = {
        employee_id: timeOff.employeeId.toFixed(0),
        time_off_type_id: timeOff.typeId.toFixed(0),
        // Reminder: there may be adjustments needed to handle timezones better to avoid switching days, here
        start_date: timeOff.startAt.toISODate(),
        end_date: timeOff.endAt.toISODate(),
        half_day_start: halfDayStart ? "1" : "0",
        half_day_end: halfDayEnd ? "1" : "0",
        comment: timeOff.comment
    };

    if (timeOff.status === 'approved') {
        payload.skip_approval = "1";
    }

    return payload;
}


/** Insert a new Personio TimeOff. */
async function createPersonioTimeOff_(personio, timeOff) {

    const payload = generatePersonioTimeOffPayload_(timeOff);
    const result = await personio.fetchJson('/company/time-offs', {
        method: 'post',
        payload: payload
    });

    if (!result?.data) {
        throw new Error(`Failed to create TimeOffPeriod: ${JSON.stringify(result)}`);
    }

    return normalizePersonioTimeOffPeriod_(result.data);
}


/** Create a new Gcal event to mirror the specified TimeOff. */
function createEventFromTimeOff_(timeOffTypeConfig, timeOff) {
    const newEvent = {
        kind: 'calendar#event',
        iCalUID: `${Util.generateUUIDv4()}-p-${timeOff.id}-sync-timeoffs@giantswarm.io`,
        start: {
            dateTime: timeOff.startAt.toISOString()
        },
        end: {
            dateTime: timeOff.endAt.switchHour24ToHour0().toISOString()
        },
        eventType: 'outOfOffice',  // left here for completeness, still not fully supported by Google Calendar
        extendedProperties: {
            private: {
                timeOffId: timeOff.id
            }
        },
        summary: `${timeOff.comment} ⇵`
    };

    // if we can't guess the corresponding time-off-type, prefix the event summary with its name
    const guessedType = timeOffTypeConfig.findByKeywordMatch(newEvent.summary);
    if (!guessedType || guessedType.attributes.id !== timeOff.typeId) {
        const keyword = TimeOffTypeConfig.extractKeyword(timeOff.typeName);
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
