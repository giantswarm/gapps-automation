import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'fs';

const {Util} = (await import('../lib-output/lib.js')).default;

/** Call some function in the code under test which doesn't export anything, without polluting our scope. */
function createTestFunctionWrapper(fileName, functionName) {
    return Function(`"use strict"; ${fs.readFileSync(fileName)}; return ${functionName}(...arguments);`);
}

/** Minimal replacement for the GAS Utilities service (only what the code under test uses). */
const toBuffer = data => typeof data === 'string'
    ? Buffer.from(data, 'utf8')
    : Buffer.from(Uint8Array.from(data.map(b => b < 0 ? b + 256 : b)));
const toSignedBytes = buffer => [...buffer].map(b => b > 127 ? b - 256 : b);

globalThis.Utilities = {
    Charset: {UTF_8: 'utf-8'},
    base64Encode: data => toBuffer(data).toString('base64'),
    base64EncodeWebSafe: data => toBuffer(data).toString('base64url'),
    newBlob: data => ({getBytes: () => toSignedBytes(toBuffer(data))}),
    getUuid: () => crypto.randomUUID()
};
globalThis.Util = Util;

const CONSENT_POLL = 'meeting-consent-poll/ConsentPoll.js';
const decodeWords = text => text.replace(/\r\n /g, '')
    .replace(/=\?UTF-8\?B\?([^?]*)\?=/g, (_, b64) => Buffer.from(b64, 'base64').toString('utf8'));


// test escapeHtml_
const escapeHtml_ = createTestFunctionWrapper(CONSENT_POLL, 'escapeHtml_');
assert.equal(escapeHtml_('<script>alert("x" & \'y\')</script>'),
    '&lt;script&gt;alert(&quot;x&quot; &amp; &#39;y&#39;)&lt;/script&gt;');
assert.equal(escapeHtml_(null), '');


// test isValidEmail_ (addresses end up in mail headers)
const isValidEmail_ = createTestFunctionWrapper(CONSENT_POLL, 'isValidEmail_');
for (const [email, valid] of [
    ['jane.doe@example.com', true],
    ['jane+tag@sub.example.co', true],
    ['jane@example.com\r\nBcc: evil@example.com', false],
    ['Jane <jane@example.com>', false],
    ['jane@localhost', false],
    ['', false],
    [null, false]
]) {
    assert.equal(isValidEmail_(email), valid, `isValidEmail_ failed for ${email}`);
}


// test encodeHeaderText_
const encodeHeaderText_ = createTestFunctionWrapper(CONSENT_POLL, 'encodeHeaderText_');
assert.equal(encodeHeaderText_('Plain ASCII subject'), 'Plain ASCII subject');
assert.equal(encodeHeaderText_('Injected\r\nBcc: evil@example.com'), 'Injected Bcc: evil@example.com');
const subject = 'Recording consent: Planung 🎥 ' + 'ä'.repeat(40);
const encoded = encodeHeaderText_(subject);
assert.equal(decodeWords(encoded), subject);
for (const line of ('Subject: ' + encoded).split('\r\n')) {
    assert.ok(line.length <= 78, `header line too long: ${line}`);
}


// test buildMimeMessage_
const buildMimeMessage_ = createTestFunctionWrapper(CONSENT_POLL, 'buildMimeMessage_');
const raw = buildMimeMessage_({
    from: 'automation@example.com',
    fromName: 'Meeting Recording Consent',
    to: 'guest@partner.example.org',
    replyTo: 'organizer@example.com',
    subject: 'Recording consent: Grüße',
    text: 'Hällo\nhttps://example.com/exec?e=1&t=abc',
    html: '<p>Hällo</p>'
});
const mime = Buffer.from(raw, 'base64url').toString('utf8');
const [head, ...body] = mime.split('\r\n\r\n');
assert.ok(head.includes('\r\nTo: guest@partner.example.org\r\n'));
assert.ok(head.includes('\r\nReply-To: organizer@example.com\r\n'));
assert.ok(head.startsWith('From: "Meeting Recording Consent" <automation@example.com>\r\n'));
assert.equal(decodeWords(head.match(/\r\nSubject: (.*(\r\n .*)*)/)[1]), 'Recording consent: Grüße');
const boundary = head.match(/boundary="([^"]+)"/)[1];
const parts = body.join('\r\n\r\n').split('--' + boundary);
assert.equal(parts.length, 4);   // preamble, text, html, closing
assert.equal(parts[3].trim(), '--');
const decodePart = part => Buffer.from(part.split('\r\n\r\n')[1].replace(/\r\n/g, ''), 'base64').toString('utf8');
assert.ok(parts[1].includes('text/plain; charset="UTF-8"'));
assert.equal(decodePart(parts[1]), 'Hällo\nhttps://example.com/exec?e=1&t=abc');
assert.ok(parts[2].includes('text/html; charset="UTF-8"'));
assert.equal(decodePart(parts[2]), '<p>Hällo</p>');
assert.throws(() => buildMimeMessage_({from: 'a@example.com', to: 'b@example.com\r\nBcc: c@example.com', subject: '', text: '', html: ''}));


// test splitUtf8_ (large poll states are split over several properties)
const splitUtf8_ = createTestFunctionWrapper(CONSENT_POLL, 'splitUtf8_');
const text = ('{"name":"Jörg 🎥"}').repeat(500);
const chunks = splitUtf8_(text, 1000);
assert.equal(chunks.join(''), text);
assert.ok(chunks.length > 1);
assert.ok(chunks.every(chunk => Buffer.byteLength(chunk, 'utf8') <= 1000));


// test findAttendeeByToken_ and countAnswers_
const findAttendeeByToken_ = createTestFunctionWrapper(CONSENT_POLL, 'findAttendeeByToken_');
const countAnswers_ = createTestFunctionWrapper(CONSENT_POLL, 'countAnswers_');
const state = {
    v: 2, id: 'event1',
    att: {
        'a@example.com': {t: 'token-a', c: true},
        'b@partner.org': {t: 'token-b', c: false},
        'c@example.com': {t: 'token-c'},
        'gone@example.com': {t: 'token-gone', c: false, x: 1}
    }
};
assert.equal(findAttendeeByToken_(state, 'token-b'), 'b@partner.org');
assert.equal(findAttendeeByToken_(state, 'token-x'), null);
assert.equal(findAttendeeByToken_(state, ''), null);
assert.equal(findAttendeeByToken_({...state, v: 1}, 'token-a'), null);
assert.deepEqual(countAnswers_(state), {invited: 3, consented: 1, declined: 1, pending: 1});
