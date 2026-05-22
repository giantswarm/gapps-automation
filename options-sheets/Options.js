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


/**
 * Computes the share allocation for a single batch.
 *
 * Translation of the spreadsheet formula:
 *   IF(employeeId = "")
 *     -> NaN                                         // manual override or no person, no recalc
 *   ELSE
 *     -> totalShares(quarter(startDate))
 *        * (4 * optionsSelect * salaryAt(employeeId, startDate) * locationFactor)
 *        / valuation(quarter(startDate))
 *
 * where quarter is derived from startDate as "YYYY Qn". The salary at startDate is used to pick the salary
 * at that point in time.
 *
 * Ranges are passed in (rather than fetched via SpreadsheetApp.openById) because custom functions
 * cannot use scoped services. Stage cross-spreadsheet data with =IMPORTRANGE(...) in a sheet of the
 * host spreadsheet and pass that range in.
 *
 * salaryHistoryRange:  Salary_History_Import-style wide table including the header row. Col 1 is
 *                      "Personio ID" (the match key). Salary columns are identified by header cells
 *                      that are actual Date values (e.g. "2020 - Jan", "2021 - Jul"); non-date headers
 *                      (ID, First/Last Name, Team, Position, Status, OTP Q4 ..., Notes) are ignored.
 *                      The salary used is the value in the column whose header date is the latest
 *                      one <= referenceDate that is non-empty for that person.
 * valuationRange:      col 0 = quarter key (sorted ascending), col 5 = valuation, col 6 = total shares
 *                      (approximate-match lookup, like VLOOKUP(..., TRUE)). If startDate falls past
 *                      the last quarter in the range, the last populated row is used as fallback.
 *
 * @param employeeId {string|number} Personio ID (matches Salary_History_Import col B).
 * @param startDate {Date|string|number} Vesting period start date; selects the valuation quarter.
 * @param optionsSelect {number} Options select multiplier.
 * @param locationFactor {number} Location factor multiplier.
 * @param salaryHistoryRange {Array<Array<any>>} Salary_History_Import range incl. header row (e.g. =Salary_History_Import!A:Z).
 * @param valuationRange {Array<Array<any>>} Valuation table
 * @param strictSalary {boolean} Require strict historical salary (true) or take the one before or after (false)
 * @return {number} The computed shares for this batch.
 * @customfunction
 */
function CALCULATE_SHARES(employeeId, startDate, optionsSelect, locationFactor, salaryHistoryRange, valuationRange, strictSalary = false) {
    const id = (employeeId == null) ? '' : String(employeeId).trim();
    if (!id) {
        return NaN;
    }

    const date = startDate instanceof Date ? startDate : new Date(startDate);
    if (isNaN(date)) {
        throw new Error('CALCULATE_SHARES: startDate (arg 2) is invalid');
    }

    if (!Array.isArray(salaryHistoryRange) || !Array.isArray(salaryHistoryRange[0])) {
        throw new Error('CALCULATE_SHARES: salaryHistoryRange (arg 5) must be a 2D range, got '
            + describeRange_(salaryHistoryRange) + '. Pass e.g. Salary_History!A:Z (with the header row).');
    }
    if (!Array.isArray(valuationRange) || !Array.isArray(valuationRange[0])) {
        throw new Error('CALCULATE_SHARES: valuationRange (arg 6) must be a 2D range, got '
            + describeRange_(valuationRange) + '. Pass e.g. Valuation!A7:G37.');
    }
    const quarterKey = date.getFullYear() + ' Q' + (Math.floor(date.getMonth() / 3) + 1);

    const salary = lookupSalaryAtDate_(salaryHistoryRange, id, startDate, strictSalary);
    if (salary == null) {
        throw new Error('CALCULATE_SHARES: no salary for Personio ID "' + id + '" at or before ' + startDate.toISOString().slice(0, 10));
    }

    const totalSharesRaw = vlookupApprox_(valuationRange, quarterKey, 6);
    const valuationRaw = vlookupApprox_(valuationRange, quarterKey, 5);
    const totalShares = toNumber_(totalSharesRaw);
    const valuation = toNumber_(valuationRaw);
    if (totalShares == null) {
        throw new Error('CALCULATE_SHARES: no Total Shares for "' + quarterKey + '" in valuationRange ('
            + describeRange_(valuationRange) + '). Make sure the range spans through column G (Total Shares); e.g. A7:G37. Raw col 6 lookup: ' + JSON.stringify(totalSharesRaw));
    }
    if (!valuation) {
        throw new Error('CALCULATE_SHARES: no Valuation for "' + quarterKey + '" in valuationRange ('
            + describeRange_(valuationRange) + '). Raw col 5 lookup: ' + JSON.stringify(valuationRaw));
    }

    return totalShares * (4 * Number(optionsSelect) * salary * Number(locationFactor)) / valuation;
}


/** Coerce a cell value to a number, tolerating currency symbols (€/$/£/¥) and US-style thousand
 *  separators (commas). Returns null for null/empty/'-'/Date/unparseable. Numbers pass through. */
function toNumber_(value) {
    if (typeof value === 'number') return isNaN(value) ? null : value;
    if (value == null || value instanceof Date) return null;
    const s = String(value).trim().replace(/[€$£¥\s]/g, '').replace(/,/g, '');
    if (s === '' || s === '-') return null;
    const n = Number(s);
    return isNaN(n) ? null : n;
}


/** Returns the salary for `employeeId` (matched against col 1 = "Personio ID") from a
 *  Salary_History_Import-style range. The salary column is the one whose Date header is the latest
 *  <= `date` with a numeric value in the matching row. Non-Date header cells are skipped.
 *
 *  If `strictSalary` is false and no salary on/before `date` is found, falls back to the earliest
 *  salary at a date > `date` (assumes salary unchanged outside the known range). Returns null only
 *  if no salary value exists for the person at any date. */
function lookupSalaryAtDate_(range, employeeId, date, strictSalary) {
    if (!Array.isArray(range) || range.length < 2) return null;

    const header = range[0];
    const dateCols = [];
    for (let i = 0; i < header.length; ++i) {
        const h = header[i];
        if (h instanceof Date && !isNaN(h)) {
            dateCols.push({col: i, date: h});
        }
    }
    dateCols.sort((a, b) => a.date - b.date);

    const target = String(employeeId).trim();
    let personRow = null;
    for (let r = 1; r < range.length; ++r) {
        const row = range[r];
        if (!row) continue;
        if (String(row[1] == null ? '' : row[1]).trim() === target) {
            personRow = row;
            break;
        }
    }
    if (!personRow) return null;

    let salary = null;
    let firstFutureSalary = null;
    for (const c of dateCols) {
        const n = toNumber_(personRow[c.col]);
        if (n == null) continue;
        if (c.date <= date) {
            salary = n;
        } else if (firstFutureSalary == null) {
            firstFutureSalary = n;
        }
    }
    if (salary != null) return salary;
    if (!strictSalary && firstFutureSalary != null) return firstFutureSalary;
    return null;
}


/** VLOOKUP approximate match on column 0 (range assumed sorted ascending, like Sheets VLOOKUP ... TRUE).
 *  Returns the value at column `col` of the row with the largest key <= `key`. If no key matches
 *  (target is past the last key, or before the first key, or the matching row has no value at `col`),
 *  falls back to the last row in the range whose `col` holds a non-null/empty value. */
function vlookupApprox_(range, key, col) {
    if (!Array.isArray(range)) return null;
    const target = String(key).trim();
    let bestMatch = null;
    let lastWithValue = null;
    for (const row of range) {
        if (!row) continue;
        if (toNumber_(row[col]) == null) continue;
        lastWithValue = row;
        if (row[0] == null) continue;
        const k = String(row[0]).trim();
        if (k === '') continue;
        if (k <= target) {
            bestMatch = row;
        }
    }
    const pick = bestMatch || lastWithValue;
    return pick ? pick[col] : null;
}


/** Returns a short description of a 2D-array range for diagnostic error messages. */
function describeRange_(range) {
    if (!Array.isArray(range)) return 'not an array';
    const rows = range.length;
    const cols = (rows > 0 && Array.isArray(range[0])) ? range[0].length : 0;
    return rows + ' rows × ' + cols + ' cols';
}
