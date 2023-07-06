/** Functions for listing and modifying Google Groups. */

/** Export Google groups and group membership to the active spreadsheet.
 *
 * Requirements:
 *   - script deployed in the context of a spreadsheet (script parent)
 *   - executing user has privileges to access AdminDirectory service
 *   - AdminDirectory service is enabled and accessible
 *
 * Run this in a trigger to keep updating the values in the spreadsheet.
 */
function listGoogleGroups() {

    const options = {
        customer: "my_customer",
        maxResults: 100,
        orderBy: "email"
    };

    const groupRows = [];
    const memberRows = [];
    do {
        const groupsResponse = AdminDirectory.Groups.list(options);
        const groups = groupsResponse.groups || [];
        Logger.log(`Listing members for ${groups.length} visible groups`);
        for (const group of groups) {
            groupRows.push([group.name, group.email, group.directMembersCount, group.description]);
            const membersOptions = {
                maxResults: 500
            };
            do {
                const membersResponse = AdminDirectory.Members.list(group.id, membersOptions);
                const members = membersResponse.members || [];
                for (const member of members) {
                    var row = [group.email, member.email, member.role, member.type, member.status];
                    memberRows.push(row);
                }
                membersOptions.pageToken = membersResponse.nextPageToken;
            } while (membersOptions.pageToken);
        }
        options.pageToken = groupsResponse.nextPageToken;
    } while (options.nextPageToken);

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    const groupsSheet = SheetUtil.ensureSheet(spreadsheet, "Groups");
    groupRows.unshift(["Group Name", "Group Email", "Direct Member Count", "Description"]);
    groupsSheet.getRange(1, 1, groupsSheet.getMaxRows(), groupRows[0].length).clearContent();
    groupsSheet.getRange(1, 1, groupRows.length, groupRows[0].length).setValues(groupRows);

    const membersSheet = SheetUtil.ensureSheet(spreadsheet, "Group_Members");
    memberRows.unshift(["Group Email", "Member Email", "Member Role", "Member Type", "Member Account Status"]);
    membersSheet.getRange(1, 1, membersSheet.getMaxRows(), memberRows[0].length).clearContent();
    membersSheet.getRange(1, 1, memberRows.length, memberRows[0].length).setValues(memberRows);
}
