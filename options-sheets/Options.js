/**
 * Functions for use in Spreadsheets dealing with options/shares.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */

const DAYS_PER_MONTH = 30.42;

const MILLISECONDS_PER_DAY = 60 * 60 * 24 * 1000;

const VESTING_CLIFF_MONTHS = 24;


/**
 * Calculates the shares vested at a certain point in time.
 *
 * @param input {Array<Array<any>>} The shares and westing info, the following fields are required:
 *         Amount of shares,
 *         Vesting Start Date,
 *         25% vested,
 *         50% vested,
 *         100% Vested,
 *         Excluded Months,
 *         Reference Date  (pass string 'NOW' to use the current date)
 * @param richOutput Output reasons for vested shares == 0
 * @return The amount of vested shares at the specified point in time (one column per input row).
 * @customfunction
 */
function VESTING(input, richOutput) {
    if (!Array.isArray(input) || (input.length > 0 && input[0].length < 7)) {
        throw new Error('Invalid input range, expecting rows with at least 7 columns');
    }

    // transform each row into new row with one column (# of shares vested)
    return input.map(([shares, start, end1, end2, end3, excludedMonths, nowArg]) => {

        //const dumpArgs = () => `shares=${shares}, start=${start}, end1=${end1}, end2=${end2}, end3=${end3}, excludedMonths=${excludedMonths}, now=${nowArg}`;
        try {
            return calculateVesting(+shares,
                new Date(start),
                new Date(end1),
                new Date(end2),
                new Date(end3),
                +excludedMonths,
                (nowArg === 'NOW' || nowArg === 'now') ? new Date() : new Date(nowArg),
                !!richOutput);
        } catch (e) {
            return richOutput ? '' + e.message : null;
        }
    });
}


/** Calculate shares vested at referenceDate. */
const calculateVesting = function (shares, start, end1, end2, end3, excludedMonths, now, richOutput) {

    const isMonotonic = values => values.every((value, index, array) => (index) ? value >= array[index - 1] : true);

    const milliesToDays = seconds => seconds / MILLISECONDS_PER_DAY;

    const milliesToMonths = seconds => Math.ceil(milliesToDays(seconds)) / DAYS_PER_MONTH;

    const monthsToMillies = months => months * DAYS_PER_MONTH * MILLISECONDS_PER_DAY;

    // plausibility checks
    if (![start, end1, end2, end3, now].every(d => d instanceof Date && !isNaN(d))) {
        throw new Error('an input date is NULL or invalid');
    } else if (!isMonotonic([start, end1, end2, end3])) {
        throw new Error('vesting period dates not set or not monotonic');
    } else if (typeof shares !== 'number' || shares < 0) {
        throw new Error('parameter shares must be a number >= 0');
    } else if (typeof excludedMonths !== 'number' || excludedMonths < 0) {
        throw new Error('parameter excludedMonths must be a number >= 0');
    }

    if (now < start) {
        return richOutput ? 'vesting not started yet' : 0;
    } else if (now < start + monthsToMillies(VESTING_CLIFF_MONTHS)) {
        const cliffMonthsRemaining = milliesToMonths((start + monthsToMillies(VESTING_CLIFF_MONTHS)) - now);
        return richOutput ? 'vesting cliff not reached, yet: remaining months: ' + Math.ceil(cliffMonthsRemaining) : 0;
    }

    // TODO Probably need some rounding (to days?) here?
    const vesting25_months_total = Math.max(0, milliesToMonths(end1 - start));
    const vesting50_months_total = Math.max(0, milliesToMonths(end2 - end1));
    const vesting100_months_total = Math.max(0, milliesToMonths(end3 - end2));

    let vestedMonths = Math.max(0, milliesToMonths(now - start) - excludedMonths);
    let vestedShares = 0.0;

    // phase 1 (0 -> 25%)
    const phase1Months = Math.min(vestedMonths, vesting25_months_total);
    vestedShares += (phase1Months * (shares * 0.25)) / vesting25_months_total;
    vestedMonths -= phase1Months;

    // phase 2 (25 -> 50%)
    const phase2Months = Math.min(vestedMonths, vesting50_months_total);
    vestedShares += (phase2Months * (shares * 0.25)) / vesting50_months_total;
    vestedMonths -= phase2Months;

    // phase 3 (50 -> 100%)
    const phase3Months = Math.min(vestedMonths, vesting100_months_total);
    vestedShares += (phase3Months * (shares * 0.50)) / vesting100_months_total;
    vestedMonths -= phase3Months;

    return +(new Number(vestedShares).toFixed(2));
};