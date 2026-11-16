# Meeting Recording Consent Poll

Asks every attendee of Google Meet meetings that may be recorded (automatic recording, transcription or
Gemini notes enabled, or opted in with `#consent-poll` in the event description) for consent **before
the meeting**, by email with a personal link, and documents every answer. Runs entirely in Google Apps
Script, no external service.

Legal background: [giantswarm/giantswarm#35534](https://github.com/giantswarm/giantswarm/issues/35534)
(GDPR, § 201 StGB): consent must be explicit, informed, given before recording, specific to the purpose,
voluntary, documented and revocable.

## Why no bot in the meeting

Apps Script cannot join a Meet call, and posting into the in-meeting chat via the Chat API is not
reliable: meeting chats are system generated group DMs that Chat apps can't be added to and that have
no API link to the calendar event (see the history of
[#368](https://github.com/giantswarm/gapps-automation/pull/368)). An earlier version drove a
browser-based bot; that was dropped in favour of asking by email ahead of time.

## How it works

1. **Scan** (`checkUpcomingMeetings()`, time trigger): loads active employees from Personio (filtered by
   `allowedDomains` and `emailWhiteList`). For each one it reads the primary calendar for events starting
   within `lookaheadHours` (ongoing ones included) that the employee organizes. Each event must be
   timed, not cancelled, have a Google Meet link and at least `minAttendees` attendees.
2. **Qualify**: with `joinMode` `recorded` (the default), a meeting qualifies when automatic recording,
   transcription or Gemini notes are `ON` for its Meet space, read via Meet REST API `spaces.get` as the
   organizer. It also qualifies when the description contains `#consent-poll`. `joinMode` `all` polls
   every meeting.
3. **Poll state**: stored in ScriptProperties as `state.<eventId>`. It holds the meeting data and, for
   each attendee, a random personal token (122-bit secure UUID), the name, the invitation time and the
   answer. Rooms, declined attendees and the sender account are skipped. Updates happen under a script
   lock. States larger than about 8.5 KB are split over `chunk.<eventId>.<n>` properties.
4. **Calendar notice** (`calendarNotice`, on by default): a short, generic notice (`Recording consent
   notice: ...`) is appended once to the event description, as the organizer. The patch uses
   `sendUpdates=none`, so attendees get no update email, and `If-Match`, so concurrent edits are never
   overwritten. The notice is never duplicated: the marker text is detected, and the state remembers
   that the notice was added.
5. **Invitations**: every attendee, internal or external, gets one email from `senderEmail` (Gmail API,
   impersonated). `Reply-To` is the organizer. It contains the meeting title, the start time (in the
   event's time zone plus UTC), the organizer, what may be recorded and why (`recordingPurpose`), that
   saying no has no disadvantages, that consent can be withdrawn, the policy link and the personal
   link, marked as personal and not to be forwarded. Each person is invited exactly once per event:
   `invitedAt` is stored right after each mail. Attendees added later are invited on the next run. A
   meeting created at short notice is handled on the next run.
6. **Answering**: the link opens the web app (`doGet()`), which identifies the respondent by the token.
   An answer is recorded only by pressing a button (`submitConsent()` via `google.script.run`), never
   by opening the link, so link scanners and previews can't answer. The page shows only the
   respondent's own answer and aggregate counts, never other attendees. Answers can be changed or
   withdrawn until the meeting ends. After that, or if the meeting was cancelled, the page says the poll
   is closed and to contact the organizer.
7. **Documentation**: every answer is appended to the **consent log** spreadsheet *before* it counts. If
   the log can't be written, the answer is rejected and the respondent sees an error. Columns:
   responded at, email, name, consent, meeting title/start/link, organizer, event ID, previous answer,
   how the respondent was identified (the personal link mailed on date X), and the signed-in Google
   account if Apps Script reveals it (usually only for internal users). Values are guarded against
   formula injection.
8. **Objection**: when someone answers no (including withdrawing consent):
   - Automatic recording, transcription and Gemini notes that are `ON` are switched `OFF` via Meet REST
     API `spaces.patch` as the organizer (`disableRecordingOnObjection`, on by default). This is recorded
     in the state, and the poll keeps running even though the meeting no longer looks "recorded". If
     someone switches them on again while an objection stands, the next run switches them off again and
     tells the organizer. The script never switches anything on.
   - The organizer is emailed right away **without the objector's name**. Such emails often pop up
     during the meeting, possibly on a shared screen, and the organizer only needs to know *that* they
     must not record. The email says whether switching off worked. If it failed, it says so and asks
     the organizer to check the settings; nothing is faked. If sending fails, the next trigger run
     retries.
9. **Organizer summary**: in the first run within `summaryMinutes` before the start, the organizer gets
   one summary email. It lists the counts and who consented, who did not and who has not answered,
   because the organizer is responsible for compliance and needs to act, for example by asking
   individuals. It is marked confidential. Each later change before the meeting ends (late answers,
   objections, new attendees) triggers one update email in the next run.
10. **Moved or cancelled meetings**: polls are keyed by event ID. A moved meeting keeps its poll; the
    times are updated, links stay valid and invitations are not resent. A meeting that is cancelled,
    deleted, no longer qualifies (for example recording switched off by the organizer, Meet link
    removed) or changes organizer before it ends gets no more invitations and no summary. Invitees are
    not notified; the poll page shows the meeting as cancelled or closed.
11. **Cleanup**: poll state is deleted 24 hours after the meeting ended. Malformed or outdated states
    and orphaned parts are deleted too. The consent log is kept: define its retention in the recording
    policy.

## Files

- `appsscript.json`: manifest with web app settings (`access: ANYONE_ANONYMOUS`, `executeAs: USER_DEPLOYING`)
- `ConsentPoll.js`: the script (written synchronously, see the note in the file header)
- `lib.js`: generated by `make`, never edit

## Script properties

Required:

| Property | Description |
|---|---|
| `ConsentPoll.personioToken` | Personio `CLIENT_ID\|CLIENT_SECRET` |
| `ConsentPoll.serviceAccountCredentials` | Service account key JSON (domain-wide delegation, see below) |
| `ConsentPoll.allowedDomains` | Comma separated organizer domains, e.g. `giantswarm.io` |
| `ConsentPoll.senderEmail` | Account the emails are sent from, e.g. `automation@giantswarm.io` (must be a mailbox) |
| `ConsentPoll.consentLogSpreadsheetId` | ID of the consent log spreadsheet |
| `ConsentPoll.webAppUrl` | The web app's `/exec` URL (defaults to `ScriptApp.getService().getUrl()`, set it explicitly) |

Optional:

| Property | Default | Description |
|---|---|---|
| `ConsentPoll.emailWhiteList` | (all) | Comma separated organizers to handle (for testing) |
| `ConsentPoll.joinMode` | `recorded` | `recorded` or `all` |
| `ConsentPoll.lookaheadHours` | `24` | Invitations are sent for meetings starting within this window |
| `ConsentPoll.summaryMinutes` | `15` | Organizer summary this many minutes before the start |
| `ConsentPoll.minAttendees` | `2` | Minimum number of attendees |
| `ConsentPoll.calendarNotice` | `true` | Add the generic notice to the event description |
| `ConsentPoll.disableRecordingOnObjection` | `true` | Switch off automatic recording/transcription/notes on objection |
| `ConsentPoll.recordingPurpose` | (none) | Completes "may be recorded ...", e.g. `for AI generated meeting minutes` |
| `ConsentPoll.policyUrl` | (none) | Link to the recording policy |

## Domain-wide delegation scopes

Add these to the service account in Admin Console > Security > API controls > Domain-wide delegation:

- `https://www.googleapis.com/auth/calendar`: read organizer calendars and add the notice (shared
  `CalendarClient`)
- `https://www.googleapis.com/auth/directory.readonly`: attendee names (People API)
- `https://www.googleapis.com/auth/meetings.space.settings`: read and switch off the automatic
  artifacts of Meet spaces, impersonating the organizer
- `https://www.googleapis.com/auth/gmail.send`: send as `senderEmail`
- With `disableRecordingOnObjection=false`: `https://www.googleapis.com/auth/meetings.space.readonly`
  instead of `meetings.space.settings`

About `meetings.space.settings`: `spaces.patch` accepts `meetings.space.created` or
`meetings.space.settings`. Only `meetings.space.settings` covers spaces created by other apps such as
Google Calendar (Meet API
[release notes](https://developers.google.com/workspace/meet/release-notes): new scope 2025-02-07,
preconfiguring auto-recording/transcripts/notes GA since 2025-04-29). Only the meeting organizer can
change these settings, co-hosts can't
([configure meeting spaces](https://developers.google.com/workspace/meet/api/guides/meeting-spaces-configuration)).
`spaces.patch` needs the server-generated space name (`spaces/{id}`), not the meeting code, so the
space is read by meeting code first.

Enable these APIs in the service account's GCP project: Google Calendar API, People API, Google Meet
REST API, Gmail API.

## Deployment

1. Create or choose the sender account (e.g. `automation@giantswarm.io`, a licensed mailbox) and make
   sure external recipients accept mail from it (SPF/DKIM are already covered for the domain).

2. Add the domain-wide delegation scopes above and enable the APIs.

3. Create the consent log spreadsheet in an HR shared drive with restricted access, give the account
   deploying the script edit access and note its ID (from the URL). The sheet `Consent Log` is created on
   first use.

4. Build and push:
   ```bash
   make meeting-consent-poll/
   ```

5. Deploy the web app (execute as you, access: anyone, *including anonymous users*, as set in
   `appsscript.json`):
   ```bash
   cd meeting-consent-poll && clasp deploy --description "consent web app"
   ```
   Note the `/exec` URL in the form `https://script.google.com/macros/s/<deployment id>/exec`. Don't use
   the `/a/macros/<domain>/` form, which requires a company sign-in. The Workspace admin setting for
   Apps Script web app sharing must allow "Anyone". When changing the code later, update the same
   deployment (`clasp deploy -i <deployment id>`) so existing links keep working.

6. Set the properties:
   ```bash
   clasp run 'setProperties' --params '[{
     "ConsentPoll.personioToken": "CLIENT_ID|CLIENT_SECRET",
     "ConsentPoll.serviceAccountCredentials": "{...JSON...}",
     "ConsentPoll.allowedDomains": "giantswarm.io",
     "ConsentPoll.emailWhiteList": "tester@giantswarm.io",
     "ConsentPoll.senderEmail": "automation@giantswarm.io",
     "ConsentPoll.webAppUrl": "https://script.google.com/macros/s/.../exec",
     "ConsentPoll.consentLogSpreadsheetId": "SPREADSHEET ID",
     "ConsentPoll.recordingPurpose": "for AI generated meeting minutes",
     "ConsentPoll.policyUrl": "https://handbook.giantswarm.io/..."
   }, false]'
   ```

7. Test (see the test plan below), then install the trigger (every 15 minutes recommended; the interval
   must not exceed `summaryMinutes`, so that the summary arrives before the meeting):
   ```bash
   clasp run 'install' --params '[15]'
   ```
   Remove the `emailWhiteList` when the test is done.

## Test plan

With `emailWhiteList` set to the tester as organizer:

1. Create a meeting starting in 2 hours with a Meet link, `#consent-poll` in the description, a second
   internal attendee and a private external address as guest. Run `clasp run checkUpcomingMeetings`.
   Expect: one invitation per attendee from the sender (check subject umlauts, `Reply-To`, time zone),
   the notice in the event description, and no update email to attendees. Running again sends nothing
   new.
2. Add another guest and run again: only the new guest is invited.
3. Open the external guest's link in a private window (not signed in). The page should show only that
   guest's own answer and the counts. Answer yes, then check the consent log row.
4. Change the internal attendee's answer to no. Expect the organizer email (without the name) right
   away and a consent log row with previous answer `consent`.
5. Turn on automatic recording for the meeting (Calendar > event > Meet video conferencing settings) and
   run the trigger: since an objection stands, recording is switched off again (check Calendar) and the
   organizer email says so.
6. Set `ConsentPoll.summaryMinutes` to `150` and run: exactly one summary with correct lists. Change an
   answer and run: one update.
7. Move the meeting to tomorrow and run: no new mails, and the poll page shows the new time. Then delete
   the meeting and run: the poll page says cancelled, and no summary is sent.
8. After the end time, the page says the poll is closed.

## Known limitations

- Only people on the invitation are asked. Guests who join through a forwarded invitation, the Meet link
  or a dial-in were not asked; the summary reminds the organizer. A Google Group invited as a single
  attendee gets one link shared by all its members.
- Identity is possession of the personal link mailed to an address. A forwarded link lets someone else
  answer for that person. The emails ask not to forward the link, and the consent log records how
  identification worked.
- Recording or transcription started manually during the meeting isn't prevented. Switching off
  automatic recording only changes the space configuration for when someone with recording rights
  joins; a recording that is already running isn't stopped. The organizer remains responsible. Meetings
  that are only recorded manually are covered with `#consent-poll` or `joinMode` `all`.
- Switching off fails if the organizer doesn't own the Meet space (for example a link pasted from
  another meeting). The organizer is then told to check the settings themselves.
- Only meetings organized by active employees on `allowedDomains` are handled. Recurring meetings are
  handled per instance, and the notice is added to each instance.
- Invitations already sent are not updated when a meeting moves (the poll page shows the current time),
  and invitees are not told about cancellations.
- Limits: ScriptProperties hold 500 KB in total. A poll uses about 100 bytes per attendee, and a meeting
  is limited to about 100 KB (several hundred attendees). Gmail limits the sender to 2,000 recipients a
  day, and runs are limited to 6 minutes. Keep `lookaheadHours` modest with `joinMode` `all`.
