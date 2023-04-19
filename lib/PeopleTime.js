/** PeopleTime is supposed to make handling date-times from Personio a bit easier
 * without having to pull in Joda Time or similar libs and building complex constructs.
 */
class PeopleTime {
    constructor(year, month, day, hour, tzOffset) {
        this.year = year;
        this.month = month;
        this.day = day;
        this.hour = hour;
        this.tzOffset = tzOffset;
    }

    toString() {
        return this.toISODate() + 'T' + ('' + this.hour).padStart(2, '0') + ':00:00';
    }

    ['util.inspect.custom'](depth) {
        return this.toString()
    }

    get [Symbol.toStringTag]() {
        return this.toString();
    }

    toISODate() {
        return ('' + this.year).padStart(4, '0')
            + '-' + ('' + this.month).padStart(2, '0')
            + '-' + ('' + this.day).padStart(2, '0');
    }

    /** Convert to ISO8601 timestamp at the given time-zone offset. */
    toISOString() {
        const hours = Math.abs(Math.round(this.tzOffset / (60 * 60 * 1000)));
        const minutes = Math.abs(Math.round((this.tzOffset % (60 * 60 * 1000)) / (60 * 1000)));
        if (hours || minutes) {
            return this.toString() + (this.tzOffset < 0 ? '-' : '+')
                    + ('' + hours).padStart(2, '0')
                    + ':' + ('' + minutes).padStart(2, '0');
        }
        return this.toString() + 'Z';
    }

    isHalfDay() {
        return this.hour !== 0 && this.hour !== 24;
    }

    /** Returns a new PeopleTime instance shifted by the specified amount of hours (obeying calendar rules).
     *
     * NOTE: Avoid using this private helper function across DST switch boundaries,
     *       as it will not adjust the timezone offset (for now).
     */
    addHours_(hours) {
        const date = new Date(Date.UTC(this.year, this.month - 1, this.day, this.hour, 0, 0, 0)
            + (hours * 60 * 60 * 1000));

        return new PeopleTime(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate(), date.getUTCHours(), this.tzOffset);
    }

    /** Convert from 00:00 to 24:00 on the previous day.
     *  This is only done to satisfy Personio, which stores the end date on the last day of the event.
     */
    switchHour0ToHour24() {
        if (this.hour === 0) {
            const switched = this.addHours_(-1);
            switched.hour = 24;
            return switched;
        }
        return this;
    }

    /** Convert from 24:00 to 00:00 on the previous day.
     *  This is usually done to switch from Personio event date format to the regular way event ranges are recorded.
     */
    switchHour24ToHour0() {
        if (this.hour === 24) {
            const switched = this.addHours_(+1);
            switched.hour = 0;
            return switched;
        }
        return this;
    }

    /** Normalize hour to half-days according to this instances role as a start or end timestamp in a range.
     *
     * @param {boolean} isEndOfEvent Does this PeopleTime describe the end of an event (true) or the start (false)?
     * @param {boolean} halfDaysAllowed Are half-days allowed in this context?
     *
     * @return {PeopleTime} This or a new PeopleTime instance guaranteed to be normalized according to the specified inputs.
     */
    normalizeHalfDay(isEndOfEvent, halfDaysAllowed) {
        // ensure new normalized instance
        let normal = isEndOfEvent ? this.switchHour0ToHour24() : this;
        normal = normal === this ? new PeopleTime(this.year, this.month, this.day, this.hour, this.tzOffset) : normal;

        if (isEndOfEvent) {
            if (normal.hour > 12 || !halfDaysAllowed) {
                normal.hour = 24;
            } else {
                normal.hour = 12;
            }
        } else {
            if (normal.hour < 12 || !halfDaysAllowed) {
                normal.hour = 0;
            } else {
                normal.hour = 12;
            }
        }

        return normal;
    }

    /** Does the other PeopleTime instance describe a point in time that falls on the same calendar day as this one? .*/
    isAtSameDay(other) {
        return this.year === other.year && this.month === other.month && this.day === other.day;
    }

    /** Does this PeopleTime instance refer to the same point in time as the other? */
    equals(other) {
        return other && this.year === other.year && this.month === other.month && this.day === other.day && this.hour === other.hour;
    }

    /** Is the hour in the first or second half of the day? .*/
    isFirstHalfDay() {
        return this.hour < 12;
    }

    /** Get a PeopleTime instance from a ISO8601 timestamp like (ie. "2016-05-13" or "2018-09-24T20:15:13.123+01:00".
     * @param {string} ts The ISO8601 timestamp to convert to PeopleTime (lossy).
     * @param {number} hour Hour override. Will override the hour component if present, may be undefined.
     * @param {number} tzOffset Timezone offset override in milliseconds. Will override the parsed offset if present, may be undefined.
     *
     * @return {PeopleTime} A people time instance with a local datetime matching the timestamps and hour override's.
     */
    static fromISO8601(ts, hour = undefined, tzOffset = undefined) {
        // we care about local date-time
        if (!ts) {
            throw new Error('Cannot convert null/empty timestamp to PeopleTime');
        }

        const dateAndTime = ts.trim().split('T');
        const h = hour != null ? hour : (dateAndTime[1] ? Math.round(+(dateAndTime[1].substring(0, 2))) : 0);
        const ymd = dateAndTime[0].split('-');
        const year = Math.round(+ymd[0]);
        const month = Math.round(+ymd[1]);
        const day = Math.round(+ymd[2]);
        const offset = tzOffset != null ? tzOffset : Util.getTimeZoneOffset(ts);

        // test for NaN and invalid values
        if (!(year + month + day + h) || month <= 0 || month > 12 || day <= 0 || day > 31 || h < 0 || h > 24 ||
            offset == null || offset < -12 * 60 * 60 * 1000 || offset > 14 * 60 * 60 * 1000) {
            throw new Error('Invalid ISO8601 timestamp, hour override or tzOffset specified: ts=' + ts + ', hour=' + hour + ', tzOffset=' + tzOffset);
        }

        return new PeopleTime(year, month, day, h, offset);
    }
}
