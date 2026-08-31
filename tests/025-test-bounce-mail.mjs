import assert from 'node:assert/strict';
import fs from 'fs';

const {GmailClientV1} = (await import('../lib-output/lib.js')).default;

/** Call some function in the code under test which doesn't export anything, without polluting our scope. */
function createTestFunctionWrapper(fileName, functionName) {
    return Function(`"use strict"; ${fs.readFileSync(fileName)}; return ${functionName}(...arguments);`);
}

/** Minimal replacement for the GAS Utilities service (only what the code under test uses). */
const toBuffer = data => typeof data === 'string'
    ? Buffer.from(data, 'utf8')
    : Buffer.from(Uint8Array.from(data.map(b => b < 0 ? b + 256 : b)));
const toSignedBytes = buffer => [...buffer].map(b => b > 127 ? b - 256 : b);
const CHARSETS = {'utf-8': 'utf8', 'utf8': 'utf8', 'us-ascii': 'ascii', 'iso-8859-1': 'latin1', 'windows-1252': 'latin1'};
const toEncoding = charset => {
    const encoding = charset ? CHARSETS[charset.toLowerCase()] : 'utf8';
    if (!encoding) {
        throw new Error('unsupported charset: ' + charset);  // like Utilities does
    }
    return encoding;
};

globalThis.Utilities = {
    Charset: {UTF_8: 'utf-8'},
    base64Encode: data => toBuffer(data).toString('base64'),
    base64EncodeWebSafe: data => toBuffer(data).toString('base64url'),
    base64Decode: data => toSignedBytes(Buffer.from(data, 'base64')),
    base64DecodeWebSafe: data => toSignedBytes(Buffer.from(data, 'base64url')),
    newBlob: data => ({
        getBytes: () => toSignedBytes(toBuffer(data)),
        getDataAsString: charset => toBuffer(data).toString(toEncoding(charset))
    })
};

globalThis.GmailClientV1 = GmailClientV1;

const BOUNCE_MAIL = '../bounce-mail/BounceMail.js';


// test parseAddress_
const parseAddress_ = createTestFunctionWrapper(BOUNCE_MAIL, 'parseAddress_');
for (const [value, expected] of [
    ['Jane Doe <Jane.Doe@example.com>', {name: 'Jane Doe', email: 'jane.doe@example.com'}],
    ['"Doe, Jane" <jane@example.com>', {name: 'Doe, Jane', email: 'jane@example.com'}],
    ['jane@example.com', {name: '', email: 'jane@example.com'}],
    ['&quot;Jane&quot; &lt;jane@example.com&gt;', {name: 'Jane', email: 'jane@example.com'}],
    ['Jane <mailto:jane@example.com>', {name: 'Jane', email: 'jane@example.com'}],
    ['=?UTF-8?B?SsO2cmc=?= <joerg@example.com>', {name: 'Jörg', email: 'joerg@example.com'}],
    ['=?ISO-8859-1?Q?J=F6rg?= <joerg@example.com>', {name: 'Jörg', email: 'joerg@example.com'}],
    ['a@example.com, b@example.com', {name: '', email: 'a@example.com'}],
    ['no address here', {name: 'no address here', email: ''}]
]) {
    assert.deepEqual(parseAddress_(value), expected, `parseAddress_ failed for ${value}`);
}


// test sanitizeEmailAddress_ (this address ends up in the To header of an automatically sent mail)
const sanitizeEmailAddress_ = createTestFunctionWrapper(BOUNCE_MAIL, 'sanitizeEmailAddress_');
for (const [value, expected] of [
    ['Jane@Example.COM', 'jane@example.com'],
    ['<jane@example.com>', 'jane@example.com'],
    ['  jane+applications@sub.example.co.uk  ', 'jane+applications@sub.example.co.uk'],
    ['jane@example.com\r\nBcc: victim@example.com', ''],
    ['jane@example.com, victim@example.com', ''],
    ['jane@example.com victim@example.com', ''],
    ['"jane"@example.com', ''],
    ['jane..doe@example.com', ''],
    ['jane@example', ''],
    ['jane@-example.com', ''],
    ['@example.com', ''],
    ['', ''],
    ['j'.repeat(65) + '@example.com', '']
]) {
    assert.equal(sanitizeEmailAddress_(value), expected, `sanitizeEmailAddress_ failed for '${value}'`);
}


// test sanitizeText_
const sanitizeText_ = createTestFunctionWrapper(BOUNCE_MAIL, 'sanitizeText_');
assert.equal(sanitizeText_('Application\r\nX-Evil: yes', 100), 'Application X-Evil: yes');
assert.equal(sanitizeText_('  spaced \t out\n', 100), 'spaced out');
assert.equal(sanitizeText_('abcdefghij', 5), 'abcd…');
assert.equal(sanitizeText_(null, 10), '');


// test stripSubjectPrefixes_
const stripSubjectPrefixes_ = createTestFunctionWrapper(BOUNCE_MAIL, 'stripSubjectPrefixes_');
assert.equal(stripSubjectPrefixes_('Fwd: Re: Application'), 'Application');
assert.equal(stripSubjectPrefixes_('AW: WG: Bewerbung'), 'Bewerbung');
assert.equal(stripSubjectPrefixes_('Re[2]: Application'), 'Application');
assert.equal(stripSubjectPrefixes_('Golang: my application'), 'Golang: my application');


// test parseForwardedHeaderBlock_
const parseForwardedHeaderBlock_ = createTestFunctionWrapper(BOUNCE_MAIL, 'parseForwardedHeaderBlock_');

const gmailForward = `Please handle this one.

---------- Forwarded message ---------
From: Jane Doe <jane.doe@example.com>
Date: Wed, 5 Aug 2026 at 10:11
Subject: Application for Platform Engineer
To: <someone@giantswarm.io>

Dear Sir or Madam, please find my CV attached.
`;
assert.deepEqual(parseForwardedHeaderBlock_(gmailForward), {
    name: 'Jane Doe', email: 'jane.doe@example.com', subject: 'Application for Platform Engineer', messageId: null
});

const germanForward = `FYI

-------- Weitergeleitete Nachricht --------
Von: Erika Mustermann <erika@example.de>
Datum: Mi., 5. Aug. 2026
Betreff: Initiativbewerbung
An: someone@giantswarm.io
`;
assert.deepEqual(parseForwardedHeaderBlock_(germanForward), {
    name: 'Erika Mustermann', email: 'erika@example.de', subject: 'Initiativbewerbung', messageId: null
});

const outlookForward = `*From:* John Smith <john.smith@example.org>
*Sent:* Wednesday, August 5, 2026 10:11
*To:* Someone <someone@giantswarm.io>
*Subject:* CV John Smith

Hello,
`;
assert.deepEqual(parseForwardedHeaderBlock_(outlookForward), {
    name: 'John Smith', email: 'john.smith@example.org', subject: 'CV John Smith', messageId: null
});

const quotedForward = `> ---------- Forwarded message ---------
> From: Jane Doe <jane.doe@example.com>
> Subject: Application
`;
assert.equal(parseForwardedHeaderBlock_(quotedForward).email, 'jane.doe@example.com');

// prose mentioning an address must not be mistaken for a forwarded header block
assert.equal(parseForwardedHeaderBlock_('I got this from: someone@example.com, what now?'), null);
assert.equal(parseForwardedHeaderBlock_('Just a plain message without any forward.'), null);


// test extractForwardedOrigin_ for messages forwarded as attachment
const extractForwardedOrigin_ = createTestFunctionWrapper(BOUNCE_MAIL, 'extractForwardedOrigin_');
const attachedMessage = {
    payload: {
        mimeType: 'multipart/mixed',
        headers: [{name: 'From', value: 'Employee <employee@giantswarm.io>'}],
        parts: [
            {mimeType: 'text/plain', body: {data: Buffer.from('FYI', 'utf8').toString('base64url')}},
            {
                mimeType: 'message/rfc822',
                parts: [{
                    mimeType: 'multipart/alternative',
                    headers: [
                        {name: 'From', value: '=?UTF-8?B?SsO2cmc=?= <joerg@example.com>'},
                        {name: 'Subject', value: 'Fwd: =?UTF-8?B?QmV3ZXJidW5n?='},
                        {name: 'Message-ID', value: '<abc123@mail.example.com>'}
                    ]
                }]
            }
        ]
    }
};
assert.deepEqual(extractForwardedOrigin_(attachedMessage), {
    name: 'Jörg', email: 'joerg@example.com', subject: 'Fwd: Bewerbung', messageId: '<abc123@mail.example.com>'
});

// test extractForwardedOrigin_ for automatically forwarded mail
const autoForwarded = {
    payload: {
        mimeType: 'text/plain',
        headers: [
            {name: 'From', value: 'employee@giantswarm.io'},
            {name: 'Subject', value: 'Fwd: Speculative application'},
            {name: 'X-Forwarded-For', value: 'jane@example.com someone@giantswarm.io'}
        ],
        body: {data: Buffer.from('CV attached', 'utf8').toString('base64url')}
    }
};
assert.deepEqual(extractForwardedOrigin_(autoForwarded), {
    name: '', email: 'jane@example.com', subject: 'Speculative application', messageId: null
});

// test extractForwardedOrigin_ for an HTML only inline forward
const htmlForward = {
    payload: {
        mimeType: 'multipart/alternative',
        headers: [{name: 'From', value: 'employee@giantswarm.io'}],
        parts: [{
            mimeType: 'text/html',
            body: {
                data: Buffer.from('<div>FYI</div><div>---------- Forwarded message ---------<br>'
                    + 'From: Jane Doe &lt;jane@example.com&gt;<br>Subject: Application<br></div>', 'utf8').toString('base64url')
            }
        }]
    }
};
assert.equal(extractForwardedOrigin_(htmlForward).email, 'jane@example.com');

// a message that isn't a forward at all
assert.equal(extractForwardedOrigin_({
    payload: {
        mimeType: 'text/plain',
        headers: [{name: 'From', value: 'employee@giantswarm.io'}],
        body: {data: Buffer.from('Who is handling jobs@ these days?', 'utf8').toString('base64url')}
    }
}), null);


// test isBlockedAddress_ (mail loop protection)
const isBlockedAddress_ = createTestFunctionWrapper(BOUNCE_MAIL, 'isBlockedAddress_');
const config = {mailbox: 'automation@giantswarm.io', groupEmail: 'jobs@giantswarm.io', fromAddress: 'jobs@giantswarm.io', replyTo: ''};
for (const [email, expected] of [
    ['jane@example.com', false],
    ['', true],
    ['jobs@giantswarm.io', true],
    ['automation@giantswarm.io', true],
    ['no-reply@example.com', true],
    ['noreply+jobs@example.com', true],
    ['MAILER-DAEMON@example.com', true],
    ['postmaster@example.com', true]
]) {
    assert.equal(isBlockedAddress_(email, config), expected, `isBlockedAddress_ failed for '${email}'`);
}


// test isInternalAddress_
const isInternalAddress_ = createTestFunctionWrapper(BOUNCE_MAIL, 'isInternalAddress_');
const domains = {internalDomains: ['giantswarm.io', 'giantswarm.com']};
assert.equal(isInternalAddress_('someone@giantswarm.io', domains), true);
assert.equal(isInternalAddress_('someone@notgiantswarm.io', domains), false);
assert.equal(isInternalAddress_('', domains), false);


// test substitute_
const substitute_ = createTestFunctionWrapper(BOUNCE_MAIL, 'substitute_');
assert.equal(substitute_('Hi ${senderName}, re ${subject}', {senderName: 'Jane', subject: 'CV'}), 'Hi Jane, re CV');
// values must never be interpreted as part of the template
assert.equal(substitute_('${subject}', {subject: '${senderEmail}', senderEmail: 'x@y.z'}), '${senderEmail}');


// test encodeHeaderText_
const encodeHeaderText_ = createTestFunctionWrapper(BOUNCE_MAIL, 'encodeHeaderText_');
assert.equal(encodeHeaderText_('Re: Application'), 'Re: Application');
assert.equal(encodeHeaderText_('Bewerbung Jörg'), '=?UTF-8?B?QmV3ZXJidW5nIErDtnJn?=');
for (const line of encodeHeaderText_('Jörg '.repeat(20)).split('\r\n')) {
    assert.ok(line.trim().length <= 75, `encoded word exceeds 75 characters: ${line}`);
}


// test buildAutoReply_
const buildAutoReply_ = createTestFunctionWrapper(BOUNCE_MAIL, 'buildAutoReply_');
const raw = buildAutoReply_({
    fromAddress: 'jobs@giantswarm.io',
    fromName: 'Giant Swarm',
    replyTo: 'jobs@giantswarm.io',
    to: 'jane@example.com',
    toName: 'Jane Doe',
    subject: 'Re: Application',
    body: 'Hi there,\n\nunfortunately...\n',
    inReplyTo: '<abc123@mail.example.com>'
});
const mime = Buffer.from(raw, 'base64url').toString('utf8');
const [rawHeaders, rawBody] = mime.split('\r\n\r\n');

for (const expectedHeader of [
    'From: "Giant Swarm" <jobs@giantswarm.io>',
    'To: "Jane Doe" <jane@example.com>',
    'Subject: Re: Application',
    'Reply-To: jobs@giantswarm.io',
    'In-Reply-To: <abc123@mail.example.com>',
    'References: <abc123@mail.example.com>',
    'Auto-Submitted: auto-replied',
    'X-Auto-Response-Suppress: All',
    'Precedence: bulk',
    'Content-Type: text/plain; charset="UTF-8"',
    'Content-Transfer-Encoding: base64'
]) {
    assert.ok(rawHeaders.split('\r\n').includes(expectedHeader), `missing header '${expectedHeader}' in:\n${rawHeaders}`);
}

assert.equal(Buffer.from(rawBody.replace(/\r\n/g, ''), 'base64').toString('utf8'), 'Hi there,\n\nunfortunately...\n');

// long bodies must be wrapped into lines of at most 76 characters
const longBody = buildAutoReply_({
    fromAddress: 'jobs@giantswarm.io', to: 'jane@example.com', subject: 'Hi', body: 'Ä'.repeat(500)
});
for (const line of Buffer.from(longBody, 'base64url').toString('utf8').split('\r\n\r\n')[1].split('\r\n')) {
    assert.ok(line.length <= 76, `body line exceeds 76 characters: ${line.length}`);
}

// an address without display name must not be quoted
assert.ok(Buffer.from(longBody, 'base64url').toString('utf8').startsWith('From: jobs@giantswarm.io\r\n'));

console.log('bounce-mail tests passed');
