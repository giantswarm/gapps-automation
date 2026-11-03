/**
 * meet-bot: joins Google Meet meetings as the bot account when told so by the consent poll Apps Script.
 *
 * Environment:
 *   MEETBOT_SECRET           shared HMAC secret (same as script property ConsentPoll.botSecret), required
 *   MEETBOT_STORAGE_STATE    signed in Playwright storage state of the bot account (default: /secrets/storage-state.json)
 *   MEETBOT_STATE_DIR        writable directory for the refreshed storage state (optional)
 *   MEETBOT_MAX_SESSIONS     concurrently joined meetings (default: 4)
 *   MEETBOT_HEADLESS         "true" to run without a display (Meet may refuse headless browsers, default: false)
 *   MEETBOT_BROWSER_CHANNEL  Playwright browser channel, e.g. "chrome" (default: bundled Chromium)
 *   PORT                     listen port (default: 8080)
 */
import {createApiServer} from './api.mjs';
import {RequestVerifier} from './auth.mjs';
import {MeetDriver} from './meet.mjs';
import {SessionManager} from './sessions.mjs';

const log = {
    info: message => console.log(new Date().toISOString(), 'INFO', message),
    warn: message => console.warn(new Date().toISOString(), 'WARN', message),
    error: message => console.error(new Date().toISOString(), 'ERROR', message)
};

const verifier = new RequestVerifier(process.env.MEETBOT_SECRET);
const driver = new MeetDriver({
    storageStatePath: process.env.MEETBOT_STORAGE_STATE || '/secrets/storage-state.json',
    stateDir: process.env.MEETBOT_STATE_DIR || null,
    headless: process.env.MEETBOT_HEADLESS === 'true',
    channel: process.env.MEETBOT_BROWSER_CHANNEL || undefined,
    log: log
});
const sessions = new SessionManager({
    driver: driver,
    maxActive: +(process.env.MEETBOT_MAX_SESSIONS || 4),
    log: log
});

const port = +(process.env.PORT || 8080);
const server = createApiServer({sessions, verifier, log}).listen(port, () => log.info(`listening on :${port}`));

async function shutdown(signal) {
    log.info(`${signal}: leaving all meetings`);
    server.close();
    await sessions.shutdown();
    await driver.stop();
    process.exit(0);
}

process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));
