/**
 * Functions for use in Spreadsheets dealing with time-off.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */


/**
 * List upcoming time-offs in tabular form, optionally filtering.
 *
 * The function accesses data exported from Personio which must reside as sheets inside the current spreadsheet.
 *
 * Use =IMPORTRANGE(PERSONIO_SPREADSHEET_ID, TimeOffPeriod!A:ZZZ) and similar to import the needed sheet's data.
 *
 * @param {array<string>|string} typeFilter  Words that the TimeOffType name must match to be displayed.
 * @param {boolean} includeAbsent  Also list currently entires for currently absent employees.
 * @return Rows listing the upcoming leave.
 * @customfunction
 */
function UPCOMING_LEAVE(typeFilter, includeAbsent) {

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
        throw new Error("Must run in the context of a Spreadsheet");
    }

    // sanitize args
    if (typeFilter && !Array.isArray(typeFilter)) {
        typeFilter = [typeFilter];
    }
    includeAbsent = !!includeAbsent;

    const now = new Date();

    const matchingTypes = SheetUtil.getSheetData(spreadsheet, 'TimeOffType', [])
        .filter(timeOffType => !Array.isArray(typeFilter)
            || typeFilter.some(t => timeOffType.name.toLowerCase().includes(t) || timeOffType.category.toLowerCase().includes(t)))
        .reduce((map, timeOffType) => ({...map, [timeOffType.id]: timeOffType.name}), {});

    const timeOffColumns = ['id', 'start_date', 'comment', 'end_date', 'time_off_type_timeofftype_id', 'employee_employee_id'];
    const timeOffPeriods = SheetUtil.getSheetData(spreadsheet, 'TimeOffPeriod', timeOffColumns)
        .filter(timeOff => matchingTypes[timeOff.time_off_type_timeofftype_id] != null)
        .map(timeOff => ({...timeOff, start_date: new Date(timeOff.start_date), end_date: new Date(timeOff.end_date)}))
        .filter(timeOff => timeOff.end_date >= now && (includeAbsent || timeOff.start_date >= now))
        .sort((t1, t2) => {
            return t1 - t2;
        });

    const employees = SheetUtil.getSheetData(spreadsheet, 'Employee', ['id', 'first_name', 'last_name', 'email'])
        .reduce((map, employee) => {
            map[employee.id] = employee;
            return map;
        }, {});

    // id name	start_date	end_date	First name	Last name	Email
    const result = [['Id', 'Name', 'Start Date', 'End Date', 'First Name', 'Last Name', 'Email', 'Comment']];
    for (const timeOff of timeOffPeriods) {
        const employee = employees[timeOff.employee_employee_id];
        result.push([
            timeOff.id,
            matchingTypes[timeOff.time_off_type_timeofftype_id],
            timeOff.start_date,
            timeOff.end_date,
            employee?.first_name,
            employee?.last_name,
            employee?.email,
            timeOff.comment
        ])
    }

    return result;
}
