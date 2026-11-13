/**
 * Request authentication for the meet-bot control API.
 *
 * Every request carries:
 *   X-MeetBot-Timestamp: <unix seconds>
 *   X-MeetBot-Signature: sha256=<hex HMAC-SHA256(secret, timestamp + "\n" + METHOD + "\n" + path + "\n" + body)>
 *
 * The Apps Script side (meeting-consent-poll/ConsentPoll.js, signBotRequest_()) computes the same.
 */
import {createHmac, timingSafeEqual} from 'node:crypto';

/** Maximum clock skew (and replay window) accepted for signed requests. */
export const MAX_SKEW_MS = 5 * 60 * 1000;


/** Compute the hex signature of a request. */
export function sign(secret, timestamp, method, path, body) {
    return createHmac('sha256', secret)
        .update([timestamp, method.toUpperCase(), path, body || ''].join('\n'), 'utf8')
        .digest('hex');
}


/** Verifies signed requests and rejects replays of already seen signatures within the skew window. */
export class RequestVerifier {

    constructor(secret, now = () => Date.now()) {
        if (!secret || secret.length < 32) {
            throw new Error('Shared secret must have at least 32 characters');
        }
        this.secret = secret;
        this.now = now;
        this.seen = new Map();   // signature -> expiry (ms)
    }


    /** Returns null if the request is authentic, else a reason string. */
    verify(headers, method, path, body) {
        const timestamp = headers['x-meetbot-timestamp'];
        const signature = (headers['x-meetbot-signature'] || '').replace(/^sha256=/, '');
        if (!timestamp || !signature) {
            return 'missing signature headers';
        }

        const now = this.now();
        const seconds = Number(timestamp);
        if (!Number.isFinite(seconds) || Math.abs(now - seconds * 1000) > MAX_SKEW_MS) {
            return 'timestamp outside of the accepted window';
        }

        const expected = Buffer.from(sign(this.secret, timestamp, method, path, body), 'hex');
        const actual = Buffer.from(signature, 'hex');
        if (expected.length !== actual.length || !timingSafeEqual(expected, actual)) {
            return 'invalid signature';
        }

        for (const [seenSignature, expiry] of this.seen) {
            if (expiry < now) {
                this.seen.delete(seenSignature);
            }
        }
        if (this.seen.has(signature)) {
            return 'replayed request';
        }
        this.seen.set(signature, now + 2 * MAX_SKEW_MS);

        return null;
    }
}
