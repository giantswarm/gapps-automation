/**
 * Sign in the bot account interactively and save its Playwright storage state.
 *
 * Usage: node login.mjs [storage-state.json]
 *
 * A browser window opens on the Google sign-in page. Sign in as the bot account (password and 2FA
 * from 1Password), the script saves the session once the account page is reached. The file contains
 * live session cookies: never commit it, store it as a Kubernetes secret (see README.md) and delete it.
 */
import {chromium} from 'playwright-core';

const path = process.argv[2] || 'storage-state.json';

const browser = await chromium.launch({
    headless: false,
    channel: process.env.MEETBOT_BROWSER_CHANNEL || undefined,
    args: ['--disable-blink-features=AutomationControlled']
});
const context = await browser.newContext({locale: 'en-US'});
const page = await context.newPage();

await page.goto('https://accounts.google.com/ServiceLogin?hl=en&continue=https://myaccount.google.com/');
console.log('Sign in as the bot account in the browser window (waiting up to 10 minutes) ...');
await page.waitForURL(url => url.hostname === 'myaccount.google.com', {timeout: 10 * 60 * 1000});

// Visit Meet once, so its cookies are part of the state as well
await page.goto('https://meet.google.com/?hl=en', {waitUntil: 'domcontentloaded'});
await page.waitForTimeout(5000);

await context.storageState({path});
console.log(`Saved storage state to ${path}`);
await browser.close();
