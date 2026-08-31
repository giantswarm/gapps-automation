import assert from 'node:assert/strict';
import fs from 'fs';

const {SlackWebClient, GmailClientV1} = (await import('../lib-output/lib.js')).default;

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
    base64Decode: data => toSignedBytes(Buffer.from(data, 'base64')),
    base64DecodeWebSafe: data => toSignedBytes(Buffer.from(data, 'base64url')),
    newBlob: data => ({
        getBytes: () => toSignedBytes(toBuffer(data)),
        getDataAsString: () => toBuffer(data).toString('utf8')
    })
};

globalThis.SlackWebClient = SlackWebClient;
globalThis.GmailClientV1 = GmailClientV1;

const DOMAIN_HARVESTER = 'domain-harvester/DomainHarvester.js';


// test extractEmailDomain_
const extractEmailDomain_ = createTestFunctionWrapper(DOMAIN_HARVESTER, 'extractEmailDomain_');
for (const [value, expected] of [
    ['Jane@Example.COM', 'example.com'],
    ['jane@sub.example.co.uk', 'sub.example.co.uk'],
    ['', ''],
    ['jane', ''],
    ['@example.com', ''],
    ['jane@', ''],
    ['jane@-example.com', ''],
    ['jane@example', ''],
    [undefined, '']
]) {
    assert.equal(extractEmailDomain_(value), expected, `extractEmailDomain_ failed for '${value}'`);
}


// test isExcludedDomain_
const isExcludedDomain_ = createTestFunctionWrapper(DOMAIN_HARVESTER, 'isExcludedDomain_');
const config = {internalDomains: ['giantswarm.io'], excludedDomains: ['mailinator.com']};
assert.equal(isExcludedDomain_('giantswarm.io', config), true, 'internal domain must be excluded');
assert.equal(isExcludedDomain_('gmail.com', config), true, 'public provider must be excluded');
assert.equal(isExcludedDomain_('mailinator.com', config), true, 'explicitly excluded domain must be excluded');
assert.equal(isExcludedDomain_('fleetio.com', config), false, 'genuine customer domain must not be excluded');


// test extractDomains_
const extractDomains_ = createTestFunctionWrapper(DOMAIN_HARVESTER, 'extractDomains_');
const users = [
    {profile: {email: 'employee@giantswarm.io'}},
    {profile: {email: 'gui@fleetio.com'}},
    {profile: {email: 'GUI2@Fleetio.com'}},  // different case, same domain
    {profile: {email: 'someone@adidas.com'}},
    {profile: {email: 'freemail@gmail.com'}},
    {profile: {email: 'excluded@mailinator.com'}},
    {is_bot: true, profile: {email: 'bot@fleetio.com'}},
    {deleted: true, profile: {email: 'gone@fleetio.com'}},
    {profile: {}},
    {profile: {email: ''}}
];
assert.deepEqual(extractDomains_(users, config), ['adidas.com', 'fleetio.com'],
    'extractDomains_ must dedupe, lower-case and filter internal/public/excluded/bot/deleted users');


// test buildReport_ and buildReportMail_
const buildReport_ = createTestFunctionWrapper(DOMAIN_HARVESTER, 'buildReport_');
const reportConfig = {addressListName: 'urgent-group-allowed-sender-domains', adminConsoleUrl: 'https://admin.example.com/list'};
const report = buildReport_(reportConfig, ['fleetio.com'], ['adidas.com', 'fleetio.com']);
assert.ok(report.subject.includes('1 new candidate domain(s)'), `unexpected subject: ${report.subject}`);
assert.ok(report.body.includes('fleetio.com'), 'report body must mention the new domain');
assert.ok(report.body.includes('adidas.com'), 'report body must mention the full current set');
assert.ok(report.body.includes(reportConfig.adminConsoleUrl), 'report body must link to the admin console');

const buildReportMail_ = createTestFunctionWrapper(DOMAIN_HARVESTER, 'buildReportMail_');
const mailConfig = {mailbox: 'automation@giantswarm.io', fromName: 'Domain Harvester', reportTo: ['it@giantswarm.io', 'security@giantswarm.io']};
const raw = buildReportMail_(mailConfig, report);
const mime = Buffer.from(raw, 'base64url').toString('utf8');
const [rawHeaders, rawBody] = mime.split('\r\n\r\n');

for (const expectedHeader of [
    'From: "Domain Harvester" <automation@giantswarm.io>',
    'To: it@giantswarm.io, security@giantswarm.io',
    'Subject: ' + report.subject,
    'Content-Type: text/plain; charset="UTF-8"',
    'Content-Transfer-Encoding: base64'
]) {
    assert.ok(rawHeaders.split('\r\n').includes(expectedHeader), `missing header '${expectedHeader}' in:\n${rawHeaders}`);
}
assert.equal(Buffer.from(rawBody.replace(/\r\n/g, ''), 'base64').toString('utf8'), report.body);

console.log('domain-harvester tests passed');
