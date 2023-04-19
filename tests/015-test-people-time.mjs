import assert from 'node:assert/strict';

const {PeopleTime} = (await import('../lib-output/lib.js')).default;

const t1 = PeopleTime.fromISO8601('2016-05-16', 0, 29 * 1000);
const t2 = PeopleTime.fromISO8601('2016-05-16T20:00:00+05:00');

assert.equal(t1.switchHour0ToHour24().toString(), '2016-05-15T24:00:00', 'switching to hour 24 does not work');
assert.equal(t1.switchHour0ToHour24().switchHour24ToHour0().toString(), '2016-05-16T00:00:00', 'switching between hour 24 and hour 0 does not work');
assert.equal(t1.toISOString(), '2016-05-16T00:00:00Z', 'constructed string does not match correct time');
assert(PeopleTime.fromISO8601(t1.toISOString()).equals(t1), 'PeopleTime.equals() does not work or compares unnecessary fields');
assert(t2.isAtSameDay(t1), 't1 and t2 are on the same day but not recognized as such');
assert(t2.isHalfDay(), 't2 is not recognized as half-day');
assert(t1.isFirstHalfDay(), 't1 is at 00:00 which means lies in the first half of the day');
assert(!t2.isFirstHalfDay(), 't2 is at 20:00 which means the second half of the day already started');
assert(t2.normalizeHalfDay(true, true), 't2 not normalized to 2016-05-16T12:00:00 at end of event with half-days');
assert(t2.normalizeHalfDay(false, false), 't2 not normalized to 2016-05-16T00:00:00 at start of event without half-days');
