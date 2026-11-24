/**
 * Harvest external email domains from Slack workspace profiles and report new ones by mail,
 * so they can be reviewed and added to a Gmail address list (for example the
 * "urgent-group-allowed-sender-domains" list backing a shared support escalation mailbox alias).
 *
 * There is no Google API to manage Gmail "address lists" (Admin console > Apps > Google Workspace
 * > Gmail > Advanced settings > Manage address lists), so this script cannot add the domains itself.
 * Instead it mails a ready-to-paste, de-duplicated list of newly seen domains for a human to add via
 * the console's "Bulk add addresses" dialog. Domains are only reported once; already reported domains
 * are not repeated even if they were never actually added to the list.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 * Context: https://github.com/giantswarm/giantswarm/issues/37417
 */


/** The prefix for properties specific to this script in the project. */
const PROPERTY_PREFIX = 'DomainHarvester.';

/** Service account credentials (in JSON format, as downloaded from Google Cloud Console). */
const SERVICE_ACCOUNT_CREDENTIALS_KEY = PROPERTY_PREFIX + 'serviceAccountCredentials';

/** Slack bot token (needs the users:read and users:read.email scopes). */
const SLACK_BOT_TOKEN_KEY = PROPERTY_PREFIX + 'slackBotToken';

/** Primary email of the Workspace user account the report is sent from (impersonated). */
const MAILBOX_KEY = PROPERTY_PREFIX + 'mailbox';

/** Comma separated list of report recipient addresses. */
const REPORT_TO_KEY = PROPERTY_PREFIX + 'reportTo';

/** Display name of the report sender. */
const FROM_NAME_KEY = PROPERTY_PREFIX + 'fromName';

/** Comma separated list of internal domains (never reported as external customer domains). */
const INTERNAL_DOMAINS_KEY = PROPERTY_PREFIX + 'internalDomains';

/** Comma separated list of additional domains to always exclude (on top of the built-in public providers). */
const EXCLUDED_DOMAINS_KEY = PROPERTY_PREFIX + 'excludedDomains';

/** Name of the Gmail address list the report should mention. */
const ADDRESS_LIST_NAME_KEY = PROPERTY_PREFIX + 'addressListName';

/** Admin console URL of the address list management page. */
const ADMIN_CONSOLE_URL_KEY = PROPERTY_PREFIX + 'adminConsoleUrl';

/** If truthy, log what would be reported without sending mail or updating state. */
const DRY_RUN_KEY = PROPERTY_PREFIX + 'dryRun';

/** Prefix of the properties remembering which domains were already reported. */
const SEEN_DOMAIN_PREFIX = PROPERTY_PREFIX + 'seen.';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'harvestDomains';


/** Default name of the Gmail address list the report should mention. */
const DEFAULT_ADDRESS_LIST_NAME = 'urgent-group-allowed-sender-domains';

/** Default admin console URL of the address list management page. */
const DEFAULT_ADMIN_CONSOLE_URL = 'https://admin.google.com/ac/apps/gmail/manageaddresslist';

/** Well known public/free email providers, never useful as an allowlisted customer domain. */
const PUBLIC_EMAIL_DOMAINS = [
    'gmail.com', 'googlemail.com', 'yahoo.com', 'yahoo.co.uk', 'yahoo.de', 'yahoo.fr', 'ymail.com',
    'hotmail.com', 'hotmail.co.uk', 'hotmail.de', 'hotmail.fr', 'outlook.com', 'outlook.de', 'live.com', 'msn.com',
    'icloud.com', 'me.com', 'mac.com', 'aol.com',
    'gmx.net', 'gmx.de', 'gmx.com', 'gmx.at', 'gmx.ch', 'web.de', 't-online.de', 'freenet.de',
    'mail.ru', 'yandex.com', 'yandex.ru',
    'protonmail.com', 'proton.me', 'pm.me', 'tutanota.com', 'tutanota.de', 'zoho.com', 'fastmail.com',
    'qq.com', '163.com', '126.com', 'naver.com'
];


/** Main entry point.
 *
 * Harvest external email domains out of Slack workspace user profiles and mail a report of the
 * domains not yet reported before, so they can be reviewed and added to a Gmail address list.
 *
 * Requires the following script properties for operation:
 *
 *   DomainHarvester.serviceAccountCredentials  {...SERVICE_ACCOUNT_CREDENTIALS...}
 *   DomainHarvester.slackBotToken              xoxb-...
 *   DomainHarvester.mailbox                    automation@example.com
 *   DomainHarvester.reportTo                   it@example.com,security@example.com
 *   DomainHarvester.internalDomains            example.com,example.io
 *
 * Optional script properties (defaults in parentheses):
 *
 *   DomainHarvester.fromName                   ("") Display name of the report sender.
 *   DomainHarvester.excludedDomains             ("") Extra domains to never report, on top of common
 *                                               public/free mail providers already built in.
 *   DomainHarvester.addressListName            ("urgent-group-allowed-sender-domains")
 *   DomainHarvester.adminConsoleUrl             ("https://admin.google.com/ac/apps/gmail/manageaddresslist")
 *   DomainHarvester.dryRun                      ("false") If true, only log what would be reported.
 *
 * Requires service account credentials (with domain-wide delegation enabled) to be set, for example via:
 *
 *   $ clasp run 'setProperties' --params '[{"DomainHarvester.serviceAccountCredentials": "{...ESCAPED_JSON...}"}, false]'
 *
 * The service account must be configured correctly and have at least permission for the scope:
 *   https://mail.google.com/
 *
 * The Slack bot token needs the users:read and users:read.email scopes and must be a member of the
 * workspace (no channel membership is required, since users.list is a workspace wide API call).
 *
 * Ensure that this methods ExecutionApi scope is limited to MYSELF in the manifest,
 * to prevent other unauthorized domain users from using its features.
 */
async function harvestDomains() {
    return await harvestDomains_(false);
}


/** Log what would be reported without sending mail or updating any state (for testing the configuration). */
async function previewHarvestDomains() {
    return await harvestDomains_(true);
}


/** Harvest domains and mail a report of the ones not yet reported before.
 *
 * @param {boolean} forceDryRun If true, don't send mail and don't modify any state, no matter what is configured.
 *
 * @return {Object} Statistics about this run.
 */
async function harvestDomains_(forceDryRun) {

    const config = getConfig_();
    const dryRun = forceDryRun || config.dryRun;

    const slack = new SlackWebClient(config.slackBotToken);
    const users = await slack.listAllUsers();

    const currentDomains = extractDomains_(users, config);
    const seenDomains = getSeenDomains_();
    const newDomains = currentDomains.filter(domain => !seenDomains.has(domain));

    const stats = {
        usersSeen: users.length,
        domainsFound: currentDomains.length,
        newDomains: newDomains.length,
        reportSent: false
    };

    Logger.log('Found %s distinct external domain(s) among %s Slack user(s), %s new%s',
        '' + stats.domainsFound, '' + stats.usersSeen, '' + stats.newDomains, dryRun ? ' (DRY RUN)' : '');

    if (!newDomains.length) {
        return stats;
    }

    const report = buildReport_(config, newDomains, currentDomains);

    if (dryRun) {
        Logger.log('DRY RUN: would send report to %s:\n%s', config.reportTo.join(', '), report.body);
        return stats;
    }

    const gmail = await GmailClientV1.withImpersonatingService(getServiceAccountCredentials_(), config.mailbox);
    await gmail.sendUserMessage(config.mailbox, buildReportMail_(config, report), null);

    markDomainsSeen_(newDomains);
    stats.reportSent = true;

    Logger.log('Reported %s new domain(s) to %s', '' + newDomains.length, config.reportTo.join(', '));

    return stats;
}


/** Extract the de-duplicated, sorted set of external email domains out of Slack user profiles.
 *
 * @param {Array<Object>} users Slack user objects, as returned by users.list.
 * @param {Object} config The script configuration.
 *
 * @return {Array<string>} Sorted, de-duplicated external domains.
 */
function extractDomains_(users, config) {
    const domains = new Set();

    for (const user of users || []) {
        if (user.deleted || user.is_bot) {
            continue;
        }

        const domain = extractEmailDomain_((user.profile || {}).email);
        if (!domain || isExcludedDomain_(domain, config)) {
            continue;
        }

        domains.add(domain);
    }

    return Array.from(domains).sort();
}


/** Extract and validate the domain part of an email address.
 *
 * @return {string} The lower-cased domain, or an empty string if the address is not plausible.
 */
function extractEmailDomain_(email) {
    const address = (email || '').trim().toLowerCase();
    const at = address.lastIndexOf('@');
    if (at <= 0) {
        return '';
    }

    const domain = address.substring(at + 1);
    if (!/^([a-z0-9]([a-z0-9-]*[a-z0-9])?\.)+[a-z]{2,63}$/.test(domain)) {
        return '';
    }

    return domain;
}


/** Check whether a domain must never be reported (internal, public/free provider or explicitly excluded). */
function isExcludedDomain_(domain, config) {
    return config.internalDomains.includes(domain)
        || PUBLIC_EMAIL_DOMAINS.includes(domain)
        || config.excludedDomains.includes(domain);
}


/** Build the (plain text) report content. */
function buildReport_(config, newDomains, currentDomains) {
    const subject = `[DomainHarvester] ${newDomains.length} new candidate domain(s) for ${config.addressListName}`;

    const body = `Discovered ${newDomains.length} new email domain(s) from Slack workspace profiles that are `
        + `not yet reported for the "${config.addressListName}" Gmail address list.\n\n`
        + `Review them and add the legitimate customer/partner domains at:\n${config.adminConsoleUrl}\n\n`
        + `Use "Bulk add addresses" there and paste (comma separated):\n\n`
        + `${newDomains.join(', ')}\n\n`
        + `For reference, the full current set of ${currentDomains.length} non-internal domain(s) seen on `
        + `Slack is:\n\n${currentDomains.join(', ')}\n\n`
        + `Domains are only reported once: this list won't repeat a domain in a future report, even if it `
        + `wasn't added to the address list.\n\n`
        + `This is an automated report from the DomainHarvester script, see `
        + `https://github.com/giantswarm/giantswarm/issues/37417\n`;

    return {subject: subject, body: body};
}


/** Assemble the report as a base64url encoded RFC 5322 message ready for the Gmail API. */
function buildReportMail_(config, report) {
    const from = config.fromName ? `"${config.fromName.replace(/["\\]/g, '')}" <${config.mailbox}>` : config.mailbox;

    const headers = [
        'From: ' + from,
        'To: ' + config.reportTo.join(', '),
        'Subject: ' + report.subject,
        'MIME-Version: 1.0',
        'Content-Type: text/plain; charset="UTF-8"',
        'Content-Transfer-Encoding: base64'
    ];

    const body = Utilities.base64Encode(report.body, Utilities.Charset.UTF_8).match(/.{1,76}/g) || [];
    const mime = headers.join('\r\n') + '\r\n\r\n' + body.join('\r\n') + '\r\n';

    return Utilities.base64EncodeWebSafe(Utilities.newBlob(mime).getBytes());
}


/** Get the set of domains already reported in a previous run. */
function getSeenDomains_() {
    const properties = getScriptProperties_().getProperties() || {};
    const seen = new Set();

    for (const key in properties) {
        if (key.startsWith(SEEN_DOMAIN_PREFIX)) {
            seen.add(key.substring(SEEN_DOMAIN_PREFIX.length));
        }
    }

    return seen;
}


/** Remember that the specified domains have now been reported. */
function markDomainsSeen_(domains) {
    const properties = {};
    for (const domain of domains) {
        properties[SEEN_DOMAIN_PREFIX + domain] = '' + Date.now();
    }

    getScriptProperties_().setProperties(properties, false);
}


/** Forget that the specified domains were reported, so they are reported again on the next run.
 *
 * Useful to re-surface a domain that was excluded from the address list on purpose but should be
 * reconsidered later, or to recover from a report that was accidentally missed.
 */
function forgetDomains(domains) {
    const scriptProperties = getScriptProperties_();
    for (const domain of domains || []) {
        scriptProperties.deleteProperty(SEEN_DOMAIN_PREFIX + ('' + domain).trim().toLowerCase());
    }
}


/** Forget previously reported domains. Generate a full list and report all domains again on the next run. */
function forgetAllDomains() {
    const scriptProperties = getScriptProperties_();
    for (const key of scriptProperties.getKeys() || []) {
        if (key.startsWith(SEEN_DOMAIN_PREFIX)) {
            scriptProperties.deleteProperty(key);
        }
    }
}


/** Read and validate the script configuration. */
function getConfig_() {

    const properties = getScriptProperties_();
    const get = (key, fallback) => {
        const value = (properties.getProperty(key) || '').trim();
        return value || fallback;
    };
    const getList = (key) => get(key, '').split(',').map(value => value.trim().toLowerCase()).filter(value => !!value);

    const config = {
        slackBotToken: get(SLACK_BOT_TOKEN_KEY, ''),
        mailbox: get(MAILBOX_KEY, '').toLowerCase(),
        reportTo: getList(REPORT_TO_KEY),
        fromName: get(FROM_NAME_KEY, ''),
        internalDomains: getList(INTERNAL_DOMAINS_KEY),
        excludedDomains: getList(EXCLUDED_DOMAINS_KEY),
        addressListName: get(ADDRESS_LIST_NAME_KEY, DEFAULT_ADDRESS_LIST_NAME),
        adminConsoleUrl: get(ADMIN_CONSOLE_URL_KEY, DEFAULT_ADMIN_CONSOLE_URL),
        dryRun: get(DRY_RUN_KEY, 'false').toLowerCase() === 'true'
    };

    if (!config.slackBotToken) {
        throw new Error('No Slack bot token configured at script property ' + SLACK_BOT_TOKEN_KEY);
    }

    if (!config.mailbox) {
        throw new Error('No sending mailbox configured at script property ' + MAILBOX_KEY);
    }

    if (!config.reportTo.length) {
        throw new Error('No report recipients configured at script property ' + REPORT_TO_KEY);
    }

    if (!config.internalDomains.length) {
        throw new Error('No internal domains configured at script property ' + INTERNAL_DOMAINS_KEY);
    }

    return config;
}


/** Uninstall triggers. */
function uninstall() {
    TriggerUtil.uninstall(TRIGGER_HANDLER_FUNCTION);
}


/** Install periodic execution trigger. */
function install(delayMinutes) {
    TriggerUtil.install(TRIGGER_HANDLER_FUNCTION, delayMinutes);
}


/** Allow setting properties. */
function setProperties(properties, deleteAllOthers) {
    TriggerUtil.setProperties(properties, deleteAllOthers);
}


/** Get script properties. */
function getScriptProperties_() {
    const scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties) {
        throw new Error('ScriptProperties not accessible');
    }

    return scriptProperties;
}


/** Get the service account credentials. */
function getServiceAccountCredentials_() {
    const creds = getScriptProperties_().getProperty(SERVICE_ACCOUNT_CREDENTIALS_KEY);
    if (!creds) {
        throw new Error("No service account credentials at script property " + SERVICE_ACCOUNT_CREDENTIALS_KEY);
    }

    return JSON.parse(creds);
}
