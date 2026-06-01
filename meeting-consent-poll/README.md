# Meeting Recording Consent Poll

Posts an interactive GDPR consent poll card into Google Meet in-meeting chat for recording consent collection.

## Created Files

### `appsscript.json`
The Apps Script manifest with:
- `"chat": {}` to register as a Chat App (enables `onCardClick`)
- OAuth scopes for Chat bot, spaces, memberships, and standard script scopes
- `executionApi` access for `clasp run`

### `ConsentPoll.js` (~550 lines)
The main application:

**Setup functions** — `install()`, `uninstall()`, `setProperties()` following the `PersonioToGroup.js` pattern

**Chat App event handlers:**
- `onAddToSpace()` — welcome message
- `onCardClick()` — routes to consent/decline/refresh handlers

**Core trigger: `checkUpcomingMeetings()`** — scans employee calendars via Personio + CalendarClient impersonation, finds qualifying events (Meet link, ≥N attendees), sends consent polls

**Chat space discovery: `findMeetingChatSpace_()`** — lists organizer's GROUP_CHAT spaces, filters by createTime window, matches by ≥80% member overlap with calendar attendees

**Chat App installation: `installAppInSpace_()`** — adds bot member, handles "already exists" gracefully

**Card building: `buildConsentCard_()`** — header with videocam icon, meeting info, consent/decline buttons (green/red), response status section with progress bar and name lists, refresh button

**Response handling** — `LockService.getScriptLock()` for concurrency, `UPDATE_MESSAGE` response to update card in-place

**State management** — JSON in ScriptProperties, 48h cleanup cycle

**All helpers** match existing patterns: `getScriptProperties_()`, `getServiceAccountCredentials_()`, `getPersonioCreds_()`, `getEmailWhiteList_()`, etc.

## Deployment Steps

1. Create the Apps Script project:
   ```bash
   cd meeting-consent-poll && clasp create --type standalone --title "Meeting Consent Poll" --rootDir .
   ```

2. Build and push:
   ```bash
   make meeting-consent-poll/
   ```

3. Configure Chat App in GCP Console:
   - Go to Chat API → Configuration
   - Set the Apps Script deployment ID
   - Enable interactive features (slash commands not needed, just card clicks)

4. Add Domain-Wide Delegation scopes to the service account:
   - `https://www.googleapis.com/auth/calendar.readonly`
   - `https://www.googleapis.com/auth/chat.spaces.readonly`
   - `https://www.googleapis.com/auth/chat.memberships.app`
   - `https://www.googleapis.com/auth/chat.memberships.readonly`
   - `https://www.googleapis.com/auth/chat.bot`
    
5. Set properties:
   ```bash
   clasp run 'setProperties' --params '[{
     "ConsentPoll.personioToken": "CLIENT_ID|CLIENT_SECRET",
     "ConsentPoll.serviceAccountCredentials": "{...JSON...}",
     "ConsentPoll.allowedDomains": "example.com",
     "ConsentPoll.emailWhiteList": "tester@example.com"
   }, false]'
   ```

6. Test manually:
   ```bash
   clasp run 'checkUpcomingMeetings'
   ```
   Create a test meeting with a Meet link starting in ~10 minutes, verify the poll appears in meeting chat, click consent buttons, verify the card updates in-place.

7. Install the trigger:
   ```bash
   clasp run 'install' --params '[5]'
   ```
