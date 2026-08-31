import assert from 'node:assert/strict';
import fs from 'fs';
const {PeopleTime} = (await import('../lib-output/lib.js')).default;

/** Call some function in the code under test which doesn't export anything, without polluting our scope. */
function createTestFunctionWrapper(fileName, functionName) {
    return Function(`"use strict"; ${fs.readFileSync(fileName)}; return ${functionName}(...arguments);`);
}

// test generatePersonioTimeOffPayload_
const generatePersonioTimeOffPayload_ = createTestFunctionWrapper('sync-timeoffs/SyncTimeOffs.js', 'generatePersonioTimeOffPayload_');
for (const [timeOffInput, expectedPayload] of [
    [{startAt: PeopleTime.fromISO8601('2016-05-16T00:00:00+05:00'), endAt: PeopleTime.fromISO8601('2016-05-16T24:00:00+05:00')}, {half_day_start: "0", half_day_end: "0"}],
    [{startAt: PeopleTime.fromISO8601('2016-05-16T00:00:00+05:00'), endAt: PeopleTime.fromISO8601('2016-05-17T24:00:00+05:00')}, {half_day_start: "0", half_day_end: "0"}],
    [{startAt: PeopleTime.fromISO8601('2016-05-16T00:00:00+05:00'), endAt: PeopleTime.fromISO8601('2016-05-17T17:00:00+05:00')}, {half_day_start: "0", half_day_end: "0"}],
    [{startAt: PeopleTime.fromISO8601('2016-05-16T01:00:00+05:00'), endAt: PeopleTime.fromISO8601('2016-05-16T11:00:00+05:00')}, {half_day_start: "1", half_day_end: "0"}],
    [{startAt: PeopleTime.fromISO8601('2016-05-16T12:00:00+05:00'), endAt: PeopleTime.fromISO8601('2016-05-16T19:00:00+05:00')}, {half_day_start: "0", half_day_end: "1"}],
    [{startAt: PeopleTime.fromISO8601('2016-05-16T12:00:00+05:00'), endAt: PeopleTime.fromISO8601('2016-05-16T12:15:00+05:00')}, {half_day_start: "0", half_day_end: "1"}],
    [{startAt: PeopleTime.fromISO8601('2016-05-16T11:00:00+05:00'), endAt: PeopleTime.fromISO8601('2016-05-16T11:45:00+05:00')}, {half_day_start: "1", half_day_end: "0"}]
]) {
    const timeOff = {employeeId: 123, typeId: 456, ...timeOffInput};
    const payload = generatePersonioTimeOffPayload_(timeOff);

    for (const property of Object.keys(expectedPayload)) {
        if (payload[property] !== expectedPayload[property]) {
            assert.equal(payload[property], expectedPayload[property], `returned incorrect payload ${JSON.stringify(payload)} properties for timeOffInput ${JSON.stringify(timeOffInput)}`);
        }
    }
}
