# meet-bot

Joins Google Meet meetings **visibly** as a Workspace bot account (e.g. `automation@giantswarm.io`) and
talks into the meeting chat when remote controlled by the
[meeting consent poll](../meeting-consent-poll/README.md) Apps Script.

Meet has no API to join a call as a participant and Chat apps cannot post into meeting
conversations, so the bot drives the Meet web UI in a real browser (Playwright, Chromium on Xvfb),
like note taker bots do. It never grants camera or microphone permissions: it can neither be seen
nor heard, and it records nothing.

## What it does

1. The Apps Script `PUT`s a session per calendar event (re-sent on every trigger run, idempotent).
2. At `joinAt` (1 minute before the meeting) the bot opens the Meet link signed in as the bot
   account, continues without camera and microphone and joins (or asks to join and waits up to
   10 minutes to be admitted).
3. It posts the `greeting` (the consent question with a link to the poll web app) into the chat.
4. Messages sent with `PATCH {say}` are posted right away (e.g. "a participant does not consent").
5. At `leaveAt` (default 10 minutes, at the latest at the meeting end) it posts the latest
   `farewell` (the tally) and leaves. It leaves without a farewell when removed from the call or on `DELETE`.

## API

All `/v1` requests are signed with a shared secret, see [`src/auth.mjs`](src/auth.mjs).

| Request | Body | Effect |
| --- | --- | --- |
| `GET /healthz` | | liveness (unauthenticated) |
| `GET /v1/sessions` | | list sessions |
| `PUT /v1/sessions/{eventId}` | `{meetingUri, title, joinAt, leaveAt, greeting, farewell}` | create, reschedule, or retry a failed session |
| `PATCH /v1/sessions/{eventId}` | `{say, farewell}` | post a message, replace the farewell |
| `DELETE /v1/sessions/{eventId}` | | cancel or leave now |

Only `https://meet.google.com/xxx-xxxx-xxx` links are accepted, the bot browses nowhere else.
Sessions are kept in memory: after a restart the next trigger run re-creates them (meetings that
started up to 10 minutes ago are joined again).

## Bot account

- A regular Workspace user (license with Meet), named so participants recognize it, e.g. display
  name "Recording Consent Bot" with a matching profile picture.
- Joins meetings of the same domain directly when the meeting access type is the default
  ("trusted"). For restricted meetings the bot knocks and has to be admitted.
- Its credentials (password, 2FA) live in 1Password.

### Signing in (storage state)

The bot uses a Playwright [storage state](https://playwright.dev/docs/auth) (the session cookies)
of the signed in bot account. Create it on a machine with a display:

```bash
cd meet-bot
npm ci
npx playwright-core install chromium
node login.mjs storage-state.json      # sign in as the bot account in the opened window
kubectl -n <namespace> create secret generic meet-bot \
  --from-file=storage-state.json \
  --from-literal=secret="$(op read 'op://<vault>/meet-bot/secret')"
rm storage-state.json
```

The file contains live session cookies: never commit it (it is git-ignored) and delete it after
creating the secret. With `MEETBOT_STATE_DIR` on a persistent volume the bot keeps the rotated
cookies after every join, so the session lasts as long as Google does not revoke it. When it is
revoked, joins fail with `bot account is signed out` (visible in `GET /v1/sessions` and in the Apps
Script log): repeat the steps above.

## Configuration

| Variable | Default | |
| --- | --- | --- |
| `MEETBOT_SECRET` | | shared HMAC secret, 32+ characters, same as `ConsentPoll.botSecret` (required) |
| `MEETBOT_STORAGE_STATE` | `/secrets/storage-state.json` | signed in storage state |
| `MEETBOT_STATE_DIR` | `/state` (image) | writable directory for the refreshed storage state |
| `MEETBOT_MAX_SESSIONS` | `4` | concurrently joined meetings (~300-500 MB memory each) |
| `MEETBOT_HEADLESS` | `false` | Meet may refuse headless browsers, the image runs headed on Xvfb |
| `MEETBOT_BROWSER_CHANNEL` | bundled Chromium | e.g. `chrome` |
| `PORT` | `8080` | |

## Build, test, run

```bash
cd meet-bot
npm ci && npm test
docker build -t meet-bot .
docker run --rm -p 8080:8080 -e MEETBOT_SECRET=... \
  -v $PWD/storage-state.json:/secrets/storage-state.json:ro meet-bot
```

The `playwright-core` version in `package.json` and the Playwright image tag in the `Dockerfile`
must match.

## Deployment

Run a single replica (sessions are in memory) with:

- the secret mounted at `/secrets/storage-state.json` and `MEETBOT_SECRET` from the same secret,
- a small persistent volume at `/state`,
- a TLS ingress reachable from Google Apps Script (`UrlFetchApp`, i.e. from the internet; the
  request signatures are the authentication, restrict the ingress to the `/v1/` and `/healthz` paths),
- `terminationGracePeriodSeconds: 30` so the bot leaves its meetings on shutdown,
- memory requests of about 1 GiB + 500 MiB per concurrent session.

Then set `ConsentPoll.botUrl` to the ingress URL, see the
[consent poll deployment steps](../meeting-consent-poll/README.md#deployment).

## Known limitations

- The Meet UI is automated via accessible names of the English UI (forced with `hl=en`, see
  [`src/meet.mjs`](src/meet.mjs)). Meet UI changes can break joining or chatting.
- The bot cannot post when the host disabled chat for participants.
- Answers of external guests in the chat are not collected automatically, the organizer has to read them.
