/**
 * Google Meet driver: joins meetings visibly in a real browser, signed in as the bot account.
 *
 * Meet has no API to join a call as a participant, so the web UI is remote controlled with
 * Playwright. Selectors use accessible names of the English UI (forced via hl=en), which are far
 * more stable than Meet's generated class names. If Meet changes its UI, this is the file to fix.
 */
import {access, mkdir, rename, writeFile} from 'node:fs/promises';
import {dirname, join} from 'node:path';
import {chromium} from 'playwright-core';

/** Meet's in-call chat rejects longer messages. */
export const MAX_CHAT_MESSAGE_LENGTH = 500;

const LEAVE_CALL = /^Leave call$/i;
const JOIN = /^(Join now|Ask to join|Join anyway)$/i;
const CONTINUE_WITHOUT_MEDIA = /^Continue without (microphone and camera|camera|microphone)$/i;
const TURN_OFF_MEDIA = /^Turn off (microphone|camera)\b/i;
const OPEN_CHAT = /^Chat with everyone$/i;
const CHAT_INPUT = /^Send a message/i;
const REFUSED = /can.t join this (video )?call|denied your request|removed from the meeting/i;
const GUEST_NAME = /^Your name$/i;
const SIGNED_OUT_HOST = 'accounts.google.com';


/** Split text into chunks Meet accepts, preferring line boundaries. */
export function splitChatMessage(text, maxLength = MAX_CHAT_MESSAGE_LENGTH) {
    const chunks = [];
    let current = '';
    for (const line of text.split('\n')) {
        for (let rest = line; ;) {
            const candidate = current ? current + '\n' + rest : rest;
            if (candidate.length <= maxLength) {
                current = candidate;
                break;
            }
            if (current) {
                chunks.push(current);
                current = '';
                continue;
            }
            // a single line longer than the limit: cut at the last space (never inside a link)
            const cut = rest.lastIndexOf(' ', maxLength) > 0 ? rest.lastIndexOf(' ', maxLength) : maxLength;
            chunks.push(rest.substring(0, cut));
            rest = rest.substring(cut).trimStart();
            if (!rest) {
                break;
            }
        }
    }
    if (current) {
        chunks.push(current);
    }
    return chunks;
}


export class MeetDriver {

    /**
     * @param storageStatePath Playwright storage state of the signed in bot account (read only, see login.mjs)
     * @param stateDir Optional writable directory, refreshed storage state is kept there and preferred
     * @param headless Meet treats headless browsers with suspicion, the container runs headed under Xvfb
     * @param admitTimeoutMs How long to wait in the lobby when the bot has to ask to join
     */
    constructor({storageStatePath, stateDir = null, headless = false, channel = undefined, admitTimeoutMs = 10 * 60 * 1000, log = console}) {
        this.storageStatePath = storageStatePath;
        this.refreshedStatePath = stateDir ? join(stateDir, 'storage-state.json') : null;
        this.headless = headless;
        this.channel = channel;
        this.admitTimeoutMs = admitTimeoutMs;
        this.log = log;
        this.browser = null;
    }


    async start() {
        this.browser = await chromium.launch({
            headless: this.headless,
            channel: this.channel,
            args: ['--disable-blink-features=AutomationControlled', '--autoplay-policy=no-user-gesture-required']
        });
        const browser = this.browser;
        browser.on('disconnected', () => {
            if (this.browser === browser) {
                this.log.error('browser disconnected unexpectedly');
                this.browser = null;
            }
        });
    }


    async stop() {
        const browser = this.browser;
        this.browser = null;
        if (browser) {
            await browser.close();
        }
    }


    async join(session) {
        if (!this.browser) {
            await this.start();
        }

        // No camera/microphone permission: the bot can neither be seen nor heard
        const context = await this.browser.newContext({
            storageState: await this.currentStorageStatePath_(),
            locale: 'en-US',
            permissions: [],
            viewport: {width: 1280, height: 800}
        });

        try {
            const page = await context.newPage();
            const url = new URL(session.meetingUri);
            url.searchParams.set('hl', 'en');
            await page.goto(url.toString(), {waitUntil: 'domcontentloaded'});

            const joinButton = page.getByRole('button', {name: JOIN}).first();
            const refused = page.getByRole('heading', {name: REFUSED}).first();
            const guestName = page.getByRole('textbox', {name: GUEST_NAME}).first();
            await joinButton.or(refused).or(guestName).first().waitFor({timeout: 60000})
                .catch(() => undefined);   // the checks below report what went wrong

            if (new URL(page.url()).hostname === SIGNED_OUT_HOST || await guestName.isVisible()) {
                throw new Error('bot account is signed out, renew the storage state (see meet-bot/README.md)');
            }
            await throwIfRefused_(refused);

            await clickIfVisible_(page, CONTINUE_WITHOUT_MEDIA, 3000);
            await turnOffMedia_(page);
            await joinButton.click({timeout: 10000});

            // Either joined directly (same domain, trusted access) or waiting in the lobby until admitted
            const leaveButton = page.getByRole('button', {name: LEAVE_CALL});
            await leaveButton.or(refused).first().waitFor({timeout: this.admitTimeoutMs});
            await throwIfRefused_(refused);
            await turnOffMedia_(page);

            await this.saveStorageState_(context);
            return new MeetCall(context, page, this.log);
        } catch (e) {
            await context.close().catch(() => undefined);
            throw e;
        }
    }


    async currentStorageStatePath_() {
        if (this.refreshedStatePath) {
            try {
                await access(this.refreshedStatePath);
                return this.refreshedStatePath;
            } catch (e) {
                // not refreshed yet, fall through
            }
        }
        return this.storageStatePath;
    }


    /** Keep the rotated Google session cookies, so the bot stays signed in longer than the initial state. */
    async saveStorageState_(context) {
        if (!this.refreshedStatePath) {
            return;
        }
        try {
            await mkdir(dirname(this.refreshedStatePath), {recursive: true});
            const temporaryPath = this.refreshedStatePath + '.tmp';
            await writeFile(temporaryPath, JSON.stringify(await context.storageState()), {mode: 0o600});
            await rename(temporaryPath, this.refreshedStatePath);
        } catch (e) {
            this.log.warn(`failed to save refreshed storage state: ${e.message}`);
        }
    }
}


class MeetCall {

    constructor(context, page, log) {
        this.context = context;
        this.page = page;
        this.log = log;
    }


    async isActive() {
        return !this.page.isClosed() && await this.page.getByRole('button', {name: LEAVE_CALL}).count() > 0;
    }


    async say(text) {
        const input = this.page.getByRole('textbox', {name: CHAT_INPUT});
        if (!await input.isVisible()) {
            await this.page.getByRole('button', {name: OPEN_CHAT}).click({timeout: 10000});
            await input.waitFor({timeout: 10000});
        }
        if (await input.isDisabled()) {
            throw new Error('chat is disabled in this meeting');
        }
        for (const chunk of splitChatMessage(text)) {
            await input.fill(chunk);
            await input.press('Enter');
        }
    }


    async leave() {
        try {
            if (await this.isActive()) {
                await this.page.getByRole('button', {name: LEAVE_CALL}).click({timeout: 5000});
            }
        } finally {
            await this.context.close();
        }
    }
}


async function clickIfVisible_(page, name, timeout) {
    const button = page.getByRole('button', {name: name}).first();
    try {
        await button.waitFor({state: 'visible', timeout: timeout});
        await button.click();
        return true;
    } catch (e) {
        return false;
    }
}


async function throwIfRefused_(refused) {
    if (await refused.isVisible()) {
        throw new Error('Meet refused to let the bot join: ' + await refused.innerText());
    }
}


/** Defensive: without permissions Meet has no devices, but never risk an open microphone or camera. */
async function turnOffMedia_(page) {
    const buttons = page.getByRole('button', {name: TURN_OFF_MEDIA});
    for (let i = await buttons.count(); i > 0; --i) {
        await buttons.first().click({timeout: 5000}).catch(() => undefined);
    }
}
