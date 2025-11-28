/** Salary calculation functions */


/** Required employee row headers for salary calculation */
const EMPLOYEE_REQUIRED_FIELDS = [
    "ID",
    "Firstname",
    "Lastname",
    "Role",
    "Experience",
    "Resp",
    "Working Years",
    "LocFac",
    "Location Factor Top Up",
    "Wiggle Room",
    "Var Comp",
    "Options Select",
    "Ops Duty Select",
    "Perks",
    "Kids",
    "Entry Date",
    "Contract",
    "FTE"
];


/** Helper to format responses */
function _respond(statusCode, payload) {
    let responseBody;
    if (statusCode === 200) {
        responseBody = {
            status: statusCode,
            data: payload
        };
    } else {
        responseBody = {
            status: statusCode,
            error: statusCode === 400 ? "Bad Request" : "Internal Server Error",
            message: payload.message || "An error occurred"
        };
    }

    return ContentService.createTextOutput(JSON.stringify(responseBody));
}


/** Web app GET endpoint - exports available input enumerations and lookup tables.
 *
 * Note: All data is returned regardless of query parameters.
 *
 * Returned data format (content-type text/plain with JSON body):
 * {
 *  status: <HTTP status code>,
 *  error: <error type string, e.g. "Bad Request", "Internal Server Error">,
 *  message: <error message string>,
 *  data: {
 *    contractTypes: <array of contract types from query parameter>,
 *    data: {
 *      positions: <object with constant values from Data sheet>,
 *      contracts: <array of all contract-specific data from Data sheet>,
 *      experienceLevels: <array of experience level factors>,
 *      responsibilityAmounts: <array of responsibility level amounts>,
 *      ageLevels: <array of age/working years factors>
 *    }
 *  }
 * }
 *
 * @param e {Object} The event parameter with query parameters.
 * @return {ContentService.TextOutput} The JSON-encoded enumeration data with status info.
 */
function doGet(e) {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const Data = SheetUtil.getSheetData(spreadsheet, "Data", []);
        const Target = SheetUtil.getSheetData(spreadsheet, "TargetSalaries", []);
        const ExpLvl = SheetUtil.getSheetData(spreadsheet, "ExperienceLevel", []);
        const RespAmt = SheetUtil.getSheetData(spreadsheet, "ResponsibilityAmount", []);
        const AgeTbl = SheetUtil.getSheetData(spreadsheet, "Age", []);
        const targetByRole = SheetUtil.indexUnique(Target, "Position");
        const expByStep = SheetUtil.indexUnique(ExpLvl, "Step");
        const respByStep = SheetUtil.indexUnique(RespAmt, "STEP");
        const ageByStep = SheetUtil.indexUnique(AgeTbl, "STEP");
        const dataByContract = SheetUtil.indexUnique(Data, "Contract");
        const result = {
            data: {
                positions: Object.keys(targetByRole).sort(),
                contracts: Object.keys(dataByContract).sort(),
                experienceLevels: Object.keys(expByStep).sort(),
                responsibilityAmounts: Object.keys(respByStep).sort(),
                ageLevels: Object.keys(ageByStep).sort()
            }
        };
        return _respond(200, result);
    } catch (err) {
        return _respond(500, { message: err.message || "An unexpected error occurred" });
    }
}


/** Web app endpoint.
 *
 * Input data format (MUST be content-type text/plain with JSON body):
 * {
 *   employees: <array of employee rows, first row is header>
 * }
 *
 * Returned data format (content-type text/plain with JSON body):
 * {
 *  status: <HTTP status code>,
 *  error: <error type string, e.g. "Bad Request", "Internal Server Error">,
 *  message: <error message string>,
 *  data: <output rows array on success>
 * }
 *
 * @param e {Object} The event parameter, post data must contain "params" field with JSON-encoded input employee header and row(s).
 * @return {ContentService.TextOutput} The JSON-encoded output rows with status info.
 */
function doPost(e) {
    try {
        // Validate request has POST data
        if (!e?.postData?.contents) {
            return _respond(400, "Missing POST data");
        }

        // Parse JSON input
        let params;
        try {
            params = JSON.parse(e.postData.contents);
        } catch (parseError) {
            return _respond(400, "Invalid JSON in request body: " + parseError.message);
        }

        // Validate employees parameter
        const employees = params.employees;
        if (!employees) {
            return _respond(400, "Missing 'employees' parameter in request");
        }

        if (!Array.isArray(employees)) {
            return _respond(400, "'employees' parameter must be an array");
        }
        if (employees.length < 1) {
            return _respond(400, "'employees' array must contain at least a header row"
                + " with the following required fields: " + EMPLOYEE_REQUIRED_FIELDS.join(", "));
        }
        for (const requiredField of EMPLOYEE_REQUIRED_FIELDS) {
            if (!employees[0].includes(requiredField)) {
                return _respond(400, "Missing required field '" + requiredField + "' in 'employees' header row");
            }
        }

        const result = computeSalary(employees);
        return _respond(200, result);
    } catch (err) {
        return _respond(500, err.message || "An unexpected error occurred");
    }
}


/**
 * Calculates the salary information for the given employee records.
 *
 * @param employees {Array<Array<any>>} The employee record (see Employees sheet).
 * @return The current individual style output values for the given employees.
 * @customfunction
 */
function computeSalary(employees) {

    // sanitize number
    const NUM = (v) => {
        if (v == null) return 0;
        if (typeof v === "number") return v;
        let s = String(v).trim();
        if (!s) return 0;
        // strip currency + whitespace, then handle "€117.300" / "1,4"
        s = s.replace(/[€$]/g, "").replace(/\s/g, "");
        s = s.replace(/\./g, "").replace(",", ".");
        if (/%$/.test(s)) return parseFloat(s.replace("%", "")) / 100;
        return parseFloat(s);
    };

    // sanitize percentage (number/factor)
    const PCT = (v) => {
        if (v == null) return 0;
        if (typeof v === "number") return v;
        let s = String(v).trim().replace(",", ".").replace(/\s/g, "");
        if (!s) return 0;
        return s.endsWith("%") ? parseFloat(s.slice(0, -1)) / 100 : parseFloat(s);
    };

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const Employees = SheetUtil.mapRows(employees);
    const Data = SheetUtil.getSheetData(spreadsheet, "Data");
    const Target = SheetUtil.getSheetData(spreadsheet, "TargetSalaries");
    const ExpLvl = SheetUtil.getSheetData(spreadsheet, "ExperienceLevel");
    const RespAmt = SheetUtil.getSheetData(spreadsheet, "ResponsibilityAmount");
    const AgeTbl = SheetUtil.getSheetData(spreadsheet, "Age");
    const targetByRole = SheetUtil.indexUnique(Target, "Position");
    const expByStep = SheetUtil.indexUnique(ExpLvl, "Step");
    const respByStep = SheetUtil.indexUnique(RespAmt, "STEP");
    const ageByStep = SheetUtil.indexUnique(AgeTbl, "STEP");
    const dataByContract = SheetUtil.indexUnique(Data, "Contract");

    const outHeader = [
        "ID",
        "Firstname", "Lastname",
        "Cash Per month", "Var Comp Real", "Var Indiv Real", "Options Real", "Ops Duty Real",
        "Perks Real", "Perks per Month ", "Kids Real", "Kids per Month",
        "Time at GS (in years)", "Cash all in", "HR all in", "Target Salary",
        "Company Costs (annual incl. Options)", "monthly company costs excl. options",
        "LocationFactor GS Soft", "Contract Adjust", "Final Location Factor",  // recruiting needs these displayed
        "Input Error"
    ];

    if (!Data.length) {
        throw new Error("No records found in static data sheet \"Data\"")
    }

    const todayStr = Data[0][SheetUtil.sanitizeColumnName("Todays Date")] || null;
    const today = todayStr ? new Date(todayStr) : new Date();

    const rowsOut = [outHeader];

    const isEmptyRow = (row) => {
        for (const col in row) {
            if (col !== "id" && col !== "$columnIndex" && String(row[col] ?? "").trim()) {
                return false;
            }
        }
        return true;
    };

    const getField = (row, columnName) => {
        const sanitizedColumn = SheetUtil.sanitizeColumnName(columnName);
        if (!row.hasOwnProperty(sanitizedColumn)) {
            throw new Error("Input values invalid or data for column \"" + columnName + "\" missing");
        }
        return row[sanitizedColumn];
    };

    for (const emp of Employees) {
        if (isEmptyRow(emp)) {
            continue;
        }

        let rowOut = null;
        let id = '';
        let firstname = '';
        let lastname = '';
        try {
            // Constants
            const OPS_DUTY_ANNUAL_ONCALL_BASE = NUM(getField(Data[0], "OPS DUTY ANNUAL ONCALL BASE"));
            const PERKS_DEFAULT = NUM(getField(Data[0], "PERKS DEFAULT"));
            const KIDS_ALLOWANCE_PER_MONTH = NUM(getField(Data[0], "KIDS ALLOWANCE PER MONTH"));

            // Fields from Employees
            id = getField(emp, "ID");
            firstname = getField(emp, "Firstname");
            lastname = getField(emp, "Lastname");
            const role = String(getField(emp, "Role") ?? "").trim();
            const expKey = String(getField(emp, "Experience") ?? "").trim();
            const respKey = String(getField(emp, "Resp") ?? "").trim();
            const ageKey = String(getField(emp, "Working Years") ?? "").trim();
            const locFac = PCT(getField(emp, "LocFac")) || 1;
            const topUp = PCT(getField(emp, "Location Factor Top Up"));
            const wiggleRoom = PCT(getField(emp, "Wiggle Room"));
            const varCompField = getField(emp, "Var Comp");
            const optionsSelect = getField(emp, "Options Select");
            const opsDutySelect = String(getField(emp, "Ops Duty Select") || "").toLowerCase();
            const perks = NUM(getField(emp, "Perks"));
            const kids = NUM(getField(emp, "Kids"));
            const entryDateField = getField(emp, "Entry Date");
            const contract = String(getField(emp, "Contract") || "").trim().toLowerCase();
            const fte = NUM(getField(emp, "FTE")) || 1;

            // Lookups
            const expFactor = NUM(getField((expByStep[expKey] || {}), "Factor")) || 1;
            const respAmount = NUM(getField((respByStep[respKey] || {}), "Amount")) || 0;
            const ageFactor = NUM(getField((ageByStep[ageKey] || {}), "Factor")) || 1;
            const target = targetByRole[role] || {};
            const baseTarget = NUM(getField(target, "Target Salary"));
            const varCompPct = (varCompField ? PCT(varCompField) : PCT(getField(target, "Variable \nCompany"))) || 0;
            const varIndivPct = PCT(getField(target, "Variable Individual")) || 0;
            const optionsPct = PCT(optionsSelect || getField(target, "Options \nSelect")) || 0;
            const contractAdj = NUM(getField((dataByContract[contract] || {}), "Contract Adjust")) || 0;
            const costMultiplier = NUM(getField((dataByContract[contract] || {}), "Cost Multiplier")) || 1;

            // Location factors
            const locFacRaw = locFac || 1;
            const locSoft = 1 + (locFacRaw - 1) / 2;
            const locFactorTotal = SheetUtil.roundUp(locSoft + contractAdj + topUp);
            const locFactorCash = SheetUtil.roundUp(locSoft + topUp + wiggleRoom + contractAdj);
            const locPerksKidsFactor = locSoft + contractAdj;

            // TG base (no location), then split into cash/var/options
            const tgTotalBase = (((baseTarget * expFactor) + respAmount) * ageFactor) * fte;

            const varCompBase = tgTotalBase * varCompPct;
            const varIndivBase = tgTotalBase * varIndivPct;
            const optionsBase = tgTotalBase * optionsPct;
            const cashBase = tgTotalBase - (varCompBase + varIndivBase + optionsBase);

            // Apply location like Current Individual
            const cashAnnual = cashBase * locFactorCash;
            const varCompAmt = varCompBase * locFactorTotal;
            const varIndivAmt = varIndivBase * locFactorTotal;
            const optionsAmt = optionsBase * locSoft;

            // Ops Duty
            const opsDutyBase = opsDutySelect === "oncall" ? OPS_DUTY_ANNUAL_ONCALL_BASE : 0;
            const opsDuty = opsDutyBase * locFactorTotal;

            // Perks
            const perksBase = perks || PERKS_DEFAULT;
            const perksYear = perksBase * locPerksKidsFactor;
            const perksMonth = perksYear / 12;

            // Kids
            const kidsBase = kids * KIDS_ALLOWANCE_PER_MONTH * 12;
            const kidsYear = kidsBase * locPerksKidsFactor;
            const kidsMonth = kidsYear / 12;

            // Time at GS
            const entryDate = entryDateField ? new Date(entryDateField) : null;
            const yearsAt = entryDate
                ? ((today - entryDate) / (24 * 60 * 60 * 1000)) / 365
                : 0;

            // Aggregates
            const cashAllIn = cashAnnual + opsDuty + perksYear + kidsYear;
            const hrAllIn = cashAllIn + optionsAmt;

            const companyCostsInclOptions = hrAllIn * costMultiplier;

            const monthlyExclOptions = (companyCostsInclOptions - optionsAmt) / 12;

            rowOut = [
                id,
                firstname, lastname,
                (cashAnnual / 12),
                varCompAmt, varIndivAmt, optionsAmt, opsDuty,
                perksYear, perksMonth, kidsYear, kidsMonth,
                yearsAt,
                cashAllIn, hrAllIn, baseTarget,
                companyCostsInclOptions, monthlyExclOptions,
                locSoft, contractAdj, locFactorTotal,
                ''
            ].map(v => (typeof v === "number" ? SheetUtil.roundBankers(v, 2) : v));
        } catch (e) {
            rowOut = [
                id,
                firstname, lastname,
                null,
                null, null, null, null,
                null, null, null, null,
                null,
                null, null, null,
                null, null,
                null, null, null,
                e.message
            ]
        }

        rowsOut.push(rowOut);
    }

    return rowsOut;
}
