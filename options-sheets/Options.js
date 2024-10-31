/**
 * Functions for use in Spreadsheets dealing with options/shares.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */

const VESTING_CLIFF_MONTHS = 24;


/**
 * Calculates the shares vested at a certain point in time.
 *
 * @param input {Array<Array<any>>} The shares and vesting info, the following fields are required:
 *         Amount of shares,
 *         Vesting Start Date,
 *         25% vested,
 *         50% vested,
 *         100% Vested,
 *         Reference Date  (pass string 'NOW' to use the current date)
 * @param richOutput Output reasons for vested shares == 0
 * @return The amount of vested shares at the specified point in time and the total finished months considered (two columns per input row).
 * @customfunction
 */
function VESTING(input, richOutput) {
    if (!Array.isArray(input) || (input.length > 0 && input[0].length < 6)) {
        throw new Error('Invalid input range, expecting rows with at least 6 columns');
    }

    // transform each row into new row with one column (# of shares vested)
    return input.map(([shares, start, end1, end2, end3, nowArg]) => {

        //const dumpArgs = () => `shares=${shares}, start=${start}, end1=${end1}, end2=${end2}, end3=${end3}, now=${nowArg}`;
        try {
            return calculateVesting(+shares,
                new Date(start),
                new Date(end1),
                new Date(end2),
                new Date(end3),
                (nowArg === 'NOW' || nowArg === 'now') ? new Date() : new Date(nowArg),
                !!richOutput);
        } catch (e) {
            return richOutput ? '' + e.message : null;
        }
    });
}


/**
 * Calculates the shares vested at a certain point in time.
 *
 * This variant assumes that the shares are vested in consecutive batches and vesting periods and omits "excluded months".
 *
 * Vesting period duration is assumed to be 12 months each.
 *
 * @param input {Array<Array<any>>} The shares and vesting info, the following fields are required:
 *         Vesting Start Date,
 *         Reference Date  (pass string 'NOW' to use the current date),
 *         Amount of shares
 * @param richOutput Output reasons for vested shares == 0
 * @return The amount of vested shares at the specified point in time and the total finished months considered (two columns per input row).
 * @customfunction
 */
function VESTING_SIMPLE(input, richOutput) {
    if (!Array.isArray(input) || (input.length > 0 && input[0].length < 3)) {
        throw new Error('Invalid input range, expecting rows with at least 3 columns (start date, reference date, shares)');
    }

    // transform each row into new row with two columns (# of shares vested, full months considered)
    return input.map(([start, nowArg, shares]) => {

        //const dumpArgs = () => `start=${start}, nowArg=${nowArg}, shares=${shares}`;
        try {
            return calculateVesting(+shares,
                new Date(start),
                Util.addMonths(start, VESTING_CLIFF_MONTHS),
                Util.addMonths(start, VESTING_CLIFF_MONTHS + 12),
                Util.addMonths(start, VESTING_CLIFF_MONTHS + 12 + 12),
                (nowArg === 'NOW' || nowArg === 'now') ? new Date() : new Date(nowArg),
                !!richOutput);
        } catch (e) {
            return richOutput ? '' + e.message : null;
        }
    });
}


/** Calculate shares vested at referenceDate. */
const calculateVesting = function (shares, start, end1, end2, end3, now, richOutput) {

    const isMonotonic = values => values.every((value, index, array) => (index) ? value >= array[index - 1] : true);

    // plausibility checks
    if (![start, end1, end2, end3, now].every(d => d instanceof Date && !isNaN(d))) {
        throw new Error('an input date is NULL or invalid');
    } else if (!isMonotonic([start, end1, end2, end3])) {
        throw new Error('vesting period dates not set or not monotonic');
    } else if (typeof shares !== 'number' || shares < 0) {
        throw new Error('parameter shares must be a number >= 0');
    }

    const endOfCliff = Util.addMonths(start, VESTING_CLIFF_MONTHS);
    if (now < start) {
        return richOutput ? [0, 'vesting not started yet'] : [0, 0];
    } else if (now < endOfCliff) {
        const cliffMonthsRemaining = Util.monthDiff(endOfCliff, now);
        return richOutput ? [0, 'vesting cliff not reached, yet: remaining months: ' + Math.ceil(cliffMonthsRemaining)] : [0, 0];
    }

    const vesting25_months_total = Math.max(0, Util.monthDiff(start, end1));
    const vesting50_months_total = Math.max(0, Util.monthDiff(end1, end2));
    const vesting100_months_total = Math.max(0, Util.monthDiff(end2, end3));

    let totalVestedMonths = Math.max(0, Util.monthDiff(start, now));
    let vestedMonths = totalVestedMonths;
    let vestedShares = 0.0;

    // phase 1 (0 -> 25%)
    const phase1Months = Math.min(vestedMonths, vesting25_months_total);
    const phase1Shares = shares * 0.25;
    vestedShares += vesting25_months_total > 0 ? ((phase1Months * phase1Shares) / vesting25_months_total) : phase1Shares;
    vestedMonths -= phase1Months;

    // phase 2 (25 -> 50%)
    const phase2Months = Math.min(vestedMonths, vesting50_months_total);
    const phase2Shares = shares * 0.25;
    vestedShares += vesting50_months_total > 0 ? ((phase2Months * phase2Shares) / vesting50_months_total) : phase2Shares;
    vestedMonths -= phase2Months;

    // phase 3 (50 -> 100%)
    const phase3Months = Math.min(vestedMonths, vesting100_months_total);
    const phase3Shares = shares * 0.50;
    vestedShares += vesting100_months_total > 0 ? ((phase3Months * phase3Shares) / vesting100_months_total) : phase3Shares;

    return [+(new Number(vestedShares).toFixed(2)), +(new Number(totalVestedMonths).toFixed(2))];
};
