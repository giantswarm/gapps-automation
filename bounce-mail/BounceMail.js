/**
 * Auto-reply to job applications that employees forward to a shared Google Group.
 *
 * Mail sent to the group by external senders is already answered by the group's own
 * auto-responder. That responder does not fire for mail forwarded by members (employees),
 * which is exactly the case this script handles: it picks up newly forwarded messages,
 * extracts the *original* sender and subject out of the forwarded content and sends the
 * configured auto-reply to that original sender.
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */


/** The prefix for properties specific to this script in the project. */
const PROPERTY_PREFIX = 'BounceMail.';

/** Service account credentials (in JSON format, as downloaded from Google Cloud Console). */
const SERVICE_ACCOUNT_CREDENTIALS_KEY = PROPERTY_PREFIX + 'serviceAccountCredentials';

/** Primary email of the Workspace user account whose mailbox receives the group's messages. */
const MAILBOX_KEY = PROPERTY_PREFIX + 'mailbox';

/** Email address of the Google Group applications are forwarded to. */
const GROUP_EMAIL_KEY = PROPERTY_PREFIX + 'groupEmail';

/** Optional override for the Gmail search expression selecting candidate messages. */
const SEARCH_EXPRESSION_KEY = PROPERTY_PREFIX + 'searchExpression';

/** Comma separated list of internal domains (their users are considered forwarding employees). */
const INTERNAL_DOMAINS_KEY = PROPERTY_PREFIX + 'internalDomains';

/** Whether to only handle messages forwarded by an internal (employee) account. */
const REQUIRE_INTERNAL_FORWARDER_KEY = PROPERTY_PREFIX + 'requireInternalForwarder';

/** The From address of the auto-reply (must be a verified "send as" alias of the mailbox). */
const FROM_ADDRESS_KEY = PROPERTY_PREFIX + 'fromAddress';

/** The From display name of the auto-reply. */
const FROM_NAME_KEY = PROPERTY_PREFIX + 'fromName';

/** Optional Reply-To address of the auto-reply. */
const REPLY_TO_KEY = PROPERTY_PREFIX + 'replyTo';

/** Subject template of the auto-reply. */
const REPLY_SUBJECT_KEY = PROPERTY_PREFIX + 'replySubject';

/** Body template of the auto-reply. */
const REPLY_TEXT_KEY = PROPERTY_PREFIX + 'replyText';

/** Name of the Gmail label marking messages this script is done with. */
const HANDLED_LABEL_KEY = PROPERTY_PREFIX + 'handledLabel';

/** Upper bound of auto-replies sent per run (safety net against mail loops). */
const MAX_REPLIES_PER_RUN_KEY = PROPERTY_PREFIX + 'maxRepliesPerRun';

/** Hours during which the same original sender is not auto-replied to twice. */
const REPLY_COOLDOWN_HOURS_KEY = PROPERTY_PREFIX + 'replyCooldownHours';

/** If truthy, log what would be sent without actually sending anything. */
const DRY_RUN_KEY = PROPERTY_PREFIX + 'dryRun';

/** Prefix of the properties remembering when an address was last auto-replied to. */
const REPLIED_AT_PREFIX = PROPERTY_PREFIX + 'repliedAt.';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'bounceMail';


/** Default Gmail search expression selecting the messages to consider. */
const DEFAULT_SEARCH_EXPRESSION = 'newer_than:2d -in:chats -in:trash -in:spam';

/** Default name of the label marking handled messages. */
const DEFAULT_HANDLED_LABEL = 'BounceMail/handled';

/** Default subject template of the auto-reply. */
const DEFAULT_REPLY_SUBJECT = 'Re: ${subject}';

/** Subject used in place of ${subject} if the original subject could not be determined. */
const FALLBACK_SUBJECT = 'Your application';

/** Default maximum number of auto-replies sent per run. */
const DEFAULT_MAX_REPLIES_PER_RUN = 25;

/** Default number of hours during which the same original sender is auto-replied to only once. */
const DEFAULT_REPLY_COOLDOWN_HOURS = 24;

/** Default body of the auto-reply (agreed wording, see the issue referenced in README.md). */
const DEFAULT_REPLY_TEXT = `Hi there,

Thank you for your interest in Giant Swarm - we really appreciate you taking the time to reach out.

A quick note on how our hiring works: we review applications exclusively through our careers portal, where you'll find every role we currently have open: https://www.giantswarm.io/careers, this is to ensure a fast, fair and smooth process for you.

We'd love to see your application come through that channel. Our careers page is updated regularly, so feel free to check back as new opportunities open up. Unfortunately we're not able to consider applications sent by email.

Wishing you all the best with your search.

Warm regards,
Giant Swarm
`;

/** Local parts we never auto-reply to (mail loop and bounce storm protection). */
const BLOCKED_LOCAL_PARTS = [
    'noreply', 'no-reply', 'no_reply', 'donotreply', 'do-not-reply', 'do_not_reply',
    'mailer-daemon', 'mailerdaemon', 'postmaster', 'abuse', 'bounce', 'bounces', 'daemon'
];

/** Header names indicating that a message is itself automated (never auto-reply to those).
 *
 * Deliberately excludes 'list-id'/'list-unsubscribe': the Google Group adds those (and
 * Precedence: list) to every message that passes through it, including genuine employee
 * forwards, so they aren't a usable signal here.
 */
const AUTOMATED_MESSAGE_HEADERS = [
    'auto-submitted', 'x-autoreply', 'x-autorespond', 'x-auto-response-suppress'
];

/** Markers introducing a forwarded (or quoted original) message in a mail body. */
const FORWARD_MARKERS = [
    'forwarded message', 'original message', 'weitergeleitete nachricht', 'ursprüngliche nachricht',
    'message transféré', 'message d\'origine', 'mensaje reenviado', 'messaggio inoltrato',
    'doorgestuurd bericht', 'begin forwarded message'
];

/** Field labels introducing the sender in a forwarded message header block. */
const FROM_LABELS = ['from', 'von', 'de', 'da', 'van', 'expéditeur', 'mittente', 'remitente'];

/** Field labels introducing the subject in a forwarded message header block. */
const SUBJECT_LABELS = ['subject', 'betreff', 'objet', 'asunto', 'oggetto', 'assunto', 'onderwerp'];

/** Subject prefixes stripped off the original subject before use. */
const SUBJECT_PREFIXES = ['re', 'aw', 'fwd', 'fw', 'wg', 'antw', 'rif', 'tr', 'sv', 'vs'];


/** Main entry point.
 *
 * Auto-reply to the original senders of job applications forwarded to a shared Google Group.
 *
 * The Google Group itself cannot be read via the Gmail API. Instead this script reads the mailbox
 * of a regular Workspace user account that receives the group's messages, which is either
 * a member of the group or owns the group address as an alias (for example automation@example.com).
 *
 * Requires the following script properties for operation:
 *
 *   BounceMail.serviceAccountCredentials  {...SERVICE_ACCOUNT_CREDENTIALS...}
 *   BounceMail.mailbox                    automation@example.com
 *   BounceMail.groupEmail                 jobs@example.com
 *   BounceMail.internalDomains            example.com,example.io
 *   BounceMail.fromAddress                jobs@example.com
 *
 * Optional script properties (defaults in parentheses):
 *
 *   BounceMail.fromName                   ("") Display name of the auto-reply sender.
 *   BounceMail.replyTo                    ("") Reply-To address of the auto-reply.
 *   BounceMail.replySubject               ("Re: ${subject}") Subject template.
 *   BounceMail.replyText                  (built-in wording) Body template.
 *   BounceMail.searchExpression           ("newer_than:2d -in:chats -in:trash -in:spam")
 *                                         Overrides the Gmail search expression selecting candidates.
 *                                         The recipient and handled label conditions are always added.
 *   BounceMail.handledLabel               ("BounceMail/handled") Label marking messages already decided about.
 *   BounceMail.requireInternalForwarder   ("true") Only handle messages sent by internal accounts.
 *   BounceMail.maxRepliesPerRun           ("25") Safety limit on auto-replies per run.
 *   BounceMail.replyCooldownHours         ("24") Don't auto-reply to the same address twice within this window
 *                                         (0 disables the cooldown and its bookkeeping).
 *   BounceMail.dryRun                     ("false") If true, only log what would be sent.
 *
 * The subject and body templates support the placeholders ${senderName}, ${senderEmail} and ${subject},
 * which are substituted with the sanitized values extracted from the forwarded message.
 *
 * Requires service account credentials (with domain-wide delegation enabled) to be set, for example via:
 *
 *   $ clasp run 'setProperties' --params '[{"BounceMail.serviceAccountCredentials": "{...ESCAPED_JSON...}"}, false]'
 *
 * One may use the following command line to compress the service account creds into one line:
 *
 *   $ cat credentials.json | tr -d '\n '
 *
 * The service account must be configured correctly and have at least permission for these scopes:
 *   https://mail.google.com/
 *
 * Ensure that this methods ExecutionApi scope is limited to MYSELF in the manifest,
 * to prevent other unauthorized domain users from using its features.
 */
async function bounceMail() {
    return await bounceMail_(false);
}


/** Log what would be done without sending or modifying anything (for testing the configuration). */
async function previewBounceMail() {
    return await bounceMail_(true);
}


/** Auto-reply to all not yet handled forwarded applications.
 *
 * @param {boolean} forceDryRun If true, don't send mail and don't modify any message, no matter what is configured.
 *
 * @return {Object} Statistics about the messages seen in this run.
 */
async function bounceMail_(forceDryRun) {

    const config = getConfig_();
    const dryRun = forceDryRun || config.dryRun;

    const gmail = await GmailClientV1.withImpersonatingService(getServiceAccountCredentials_(), config.mailbox);

    const searchExpression = buildSearchExpression_(config);
    Logger.log('Searching mailbox %s for: %s%s', config.mailbox, searchExpression, dryRun ? ' (DRY RUN)' : '');

    // messages in spam/trash are never applications we must answer
    const messages = await gmail.listUserMessages(config.mailbox, searchExpression, null, false);

    const stats = {seen: messages.length, replied: 0, skipped: 0, failed: 0};
    if (!messages.length) {
        Logger.log('No unhandled messages found');
        return stats;
    }

    const handledLabelId = dryRun ? null : await gmail.ensureUserLabel(config.mailbox, config.handledLabel);

    let firstError = null;

    // we keep operating if handling a single message fails
    for (const messageRef of messages) {

        if (stats.replied >= config.maxRepliesPerRun) {
            Logger.log('Reached the limit of %s auto-replies per run, %s message(s) left for the next run',
                '' + config.maxRepliesPerRun, '' + (stats.seen - stats.replied - stats.skipped - stats.failed));
            break;
        }

        try {
            const result = await handleMessage_(gmail, config, messageRef.id, dryRun);

            if (result.replied) {
                ++stats.replied;
            } else {
                ++stats.skipped;
                Logger.log('Skipped message %s: %s', messageRef.id, result.reason);
            }

            if (handledLabelId && result.handled) {
                await gmail.modifyUserMessage(config.mailbox, messageRef.id, [handledLabelId], null);
            }
        } catch (e) {
            ++stats.failed;
            Logger.log('Failed to handle message %s: %s', messageRef.id, e.message);
            firstError = firstError || e;
        }
    }

    Logger.log('Handled %s message(s): %s replied, %s skipped, %s failed',
        '' + stats.seen, '' + stats.replied, '' + stats.skipped, '' + stats.failed);

    if (!dryRun) {
        cleanupRepliedAtState_(config.replyCooldownHours);
    }

    if (firstError) {
        throw firstError;
    }

    return stats;
}


/** Handle a single candidate message.
 *
 * @param {GmailClientV1} gmail The Gmail client impersonating the configured mailbox.
 * @param {Object} config The script configuration.
 * @param {string} messageId The ID of the message to handle.
 * @param {boolean} dryRun If true, don't send anything.
 *
 * @return {{replied: boolean, handled: boolean, reason: string}} Outcome, where handled marks a final decision.
 */
async function handleMessage_(gmail, config, messageId, dryRun) {

    const message = await gmail.getUserMessage(config.mailbox, messageId, 'full');
    const headers = (message.payload || {}).headers;

    // never react to automated mail, this is the most important mail loop protection
    for (const headerName of AUTOMATED_MESSAGE_HEADERS) {
        if (GmailClientV1.getHeader(headers, headerName) !== null) {
            return {replied: false, handled: true, reason: 'automated message (has header ' + headerName + ')'};
        }
    }

    const precedence = (GmailClientV1.getHeader(headers, 'Precedence') || '').trim().toLowerCase();
    if (['bulk', 'junk', 'auto_reply'].includes(precedence)) {
        return {replied: false, handled: true, reason: 'automated message (Precedence: ' + precedence + ')'};
    }

    const forwarder = parseAddress_(GmailClientV1.getHeader(headers, 'From') || '');
    if (isBlockedAddress_(forwarder.email, config)) {
        return {replied: false, handled: true, reason: 'sent by a blocked address: ' + forwarder.email};
    }

    if (config.requireInternalForwarder && !isInternalAddress_(forwarder.email, config)) {
        // external senders are answered by the group's own auto-responder
        return {replied: false, handled: true, reason: 'not forwarded by an internal account: ' + forwarder.email};
    }

    const origin = extractForwardedOrigin_(message);
    if (!origin || !origin.email) {
        return {replied: false, handled: true, reason: 'no forwarded original sender found'};
    }

    const senderEmail = sanitizeEmailAddress_(origin.email);
    if (!senderEmail) {
        return {replied: false, handled: true, reason: 'original sender is not a valid address: ' + sanitizeText_(origin.email, 100)};
    }

    if (isBlockedAddress_(senderEmail, config)) {
        return {replied: false, handled: true, reason: 'original sender is a blocked address: ' + senderEmail};
    }

    if (isInternalAddress_(senderEmail, config)) {
        return {replied: false, handled: true, reason: 'original sender is an internal address: ' + senderEmail};
    }

    const lastRepliedAt = getRepliedAt_(senderEmail);
    if (lastRepliedAt && Date.now() - lastRepliedAt < config.replyCooldownHours * 3600000) {
        return {
            replied: false, handled: true,
            reason: 'already auto-replied to ' + senderEmail + ' at ' + new Date(lastRepliedAt).toISOString()
        };
    }

    const senderName = sanitizeText_(origin.name, 120);
    const subject = sanitizeText_(stripSubjectPrefixes_(origin.subject || ''), 150) || FALLBACK_SUBJECT;
    const substitutions = {senderName: senderName, senderEmail: senderEmail, subject: subject};

    const raw = buildAutoReply_({
        fromAddress: config.fromAddress || config.mailbox,
        fromName: config.fromName,
        replyTo: config.replyTo,
        to: senderEmail,
        toName: senderName,
        subject: sanitizeText_(substitute_(config.replySubject, substitutions), 200),
        body: substitute_(config.replyText, substitutions),
        inReplyTo: origin.messageId
    });

    if (dryRun) {
        Logger.log('DRY RUN: would auto-reply to message %s:\n%s',
            messageId, Utilities.newBlob(Utilities.base64DecodeWebSafe(raw)).getDataAsString());
        return {replied: false, handled: false, reason: 'dry run'};
    }

    await gmail.sendUserMessage(config.mailbox, raw, null);
    setRepliedAt_(senderEmail, config.replyCooldownHours);

    Logger.log('Auto-replied to %s (application forwarded by %s, subject: %s)', senderEmail, forwarder.email, subject);

    return {replied: true, handled: true, reason: 'auto-replied'};
}


/** Build the Gmail search expression selecting candidate messages. */
function buildSearchExpression_(config) {
    const terms = [];

    if (config.groupEmail) {
        terms.push('to:' + config.groupEmail);
    }

    terms.push(config.searchExpression);
    terms.push('-label:"' + config.handledLabel + '"');

    return terms.join(' ');
}


/** Extract the original sender and subject out of a forwarded message.
 *
 * Handles messages forwarded as attachment (message/rfc822), inline forwards of the usual
 * Gmail/Outlook/Thunderbird flavors and mail forwarded automatically by a Gmail filter.
 *
 * @param {Object} message A Gmail message resource (retrieved with format 'full').
 *
 * @return {null|{name: string, email: string, subject: string, messageId: null|string}} The original sender or null.
 */
function extractForwardedOrigin_(message) {

    const payload = message.payload || {};

    // 1. forwarded as attachment: the attached message carries its own headers
    const attached = GmailClientV1.findPart(payload, part => part.mimeType === 'message/rfc822');
    if (attached) {
        // the Gmail API exposes the attached message as the single child part of the message/rfc822 part
        const attachedHeaders = (attached.parts || []).map(part => part.headers || []).find(headers => headers.length);
        const from = GmailClientV1.getHeader(attachedHeaders, 'From');
        if (from) {
            const address = parseAddress_(from);
            if (address.email) {
                return {
                    name: address.name,
                    email: address.email,
                    subject: decodeMimeWords_(GmailClientV1.getHeader(attachedHeaders, 'Subject') || ''),
                    messageId: sanitizeMessageId_(GmailClientV1.getHeader(attachedHeaders, 'Message-ID'))
                };
            }
        }
    }

    // 2. inline forward: parse the forwarded message's header block out of the body
    const body = GmailClientV1.getMessageBody(payload);
    const text = body.isHtml ? htmlToText_(body.text) : body.text;
    const fromBody = parseForwardedHeaderBlock_(text);
    if (fromBody) {
        return fromBody;
    }

    // 3. automatically forwarded by a Gmail filter: the original sender is in X-Forwarded-For
    const forwardedFor = parseAddress_(GmailClientV1.getHeader(payload.headers, 'X-Forwarded-For') || '');
    if (forwardedFor.email) {
        return {
            name: forwardedFor.name,
            email: forwardedFor.email,
            subject: stripSubjectPrefixes_(decodeMimeWords_(GmailClientV1.getHeader(payload.headers, 'Subject') || '')),
            messageId: null
        };
    }

    return null;
}


/** Parse the header block of an inline forwarded message ("From: ...", "Subject: ...").
 *
 * Blocks introduced by a forward marker are trusted directly. Without such a marker a
 * sender line is only accepted if a subject line follows closely, to avoid picking up
 * an address that merely appears somewhere in the prose.
 *
 * @param {string} text The plain text body of the forwarding message.
 *
 * @return {null|{name: string, email: string, subject: string, messageId: null}} The original sender or null.
 */
function parseForwardedHeaderBlock_(text) {

    const lines = (text || '').split(/\r?\n/).map(line => line.replace(/^[>\s]+/, '').replace(/\*/g, '').trim());

    let markerSeen = false;
    for (let i = 0; i < lines.length; ++i) {

        if (isForwardMarker_(lines[i])) {
            markerSeen = true;
            continue;
        }

        const from = matchLabeledLine_(lines[i], FROM_LABELS);
        if (from === null) {
            continue;
        }

        // the subject usually follows within the next few header lines
        let subject = null;
        for (let j = i + 1; j < Math.min(lines.length, i + 7); ++j) {
            subject = matchLabeledLine_(lines[j], SUBJECT_LABELS);
            if (subject !== null) {
                break;
            }
        }

        if (!markerSeen && subject === null) {
            continue;  // not confident enough that this is a forwarded header block
        }

        const address = parseAddress_(from);
        if (address.email) {
            return {
                name: address.name,
                email: address.email,
                subject: decodeMimeWords_(subject || ''),
                messageId: null
            };
        }
    }

    return null;
}


/** Check whether a line introduces a forwarded message. */
function isForwardMarker_(line) {
    const normalized = (line || '').replace(/[-_=\s]+/g, ' ').trim().toLowerCase();
    return FORWARD_MARKERS.some(marker => normalized === marker || normalized === marker + ':');
}


/** Get the value of a "Label: value" line, if its label is one of the specified ones.
 *
 * @return {null|string} The value or null if the line has no matching label.
 */
function matchLabeledLine_(line, labels) {
    const separator = (line || '').indexOf(':');
    if (separator <= 0) {
        return null;
    }

    const label = line.substring(0, separator).trim().toLowerCase();
    if (!labels.includes(label)) {
        return null;
    }

    return line.substring(separator + 1).trim();
}


/** Split an address header value into display name and email address.
 *
 * @param {string} value An address like 'Jane Doe <jane@example.com>', 'jane@example.com' or 'jane@example.com (Jane)'.
 *
 * @return {{name: string, email: string}} The parsed parts (email is empty if there is none).
 */
function parseAddress_(value) {

    // HTML derived text may contain entities and mailto: links
    const raw = decodeMimeWords_((value || '')
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/&quot;/gi, '"')
        .replace(/&amp;/gi, '&')
        .replace(/mailto:/gi, '')
        .trim());

    let name = '';
    let email = '';

    // for address lists only the first address is of interest
    const angleBracketed = raw.match(/<([^<>]*)>/);
    if (angleBracketed) {
        email = angleBracketed[1].trim();
        name = raw.substring(0, angleBracketed.index).trim();
    } else {
        const bare = raw.match(/[^\s<>(),;:"']+@[^\s<>(),;:"']+/);
        email = bare ? bare[0] : '';
        name = bare ? raw.replace(bare[0], '').trim() : raw;

        // without angle brackets the remainder may be the rest of an address list rather than a name
        if (/[@,;<>]/.test(name)) {
            name = '';
        }
    }

    return {
        name: name.replace(/^[("']+|[)"']+$/g, '').trim(),
        email: email.toLowerCase()
    };
}


/** Decode RFC 2047 encoded words ("=?UTF-8?B?...?=") in a header value. */
function decodeMimeWords_(value) {

    // adjacent encoded words are separated by whitespace that must not be part of the result
    return (value || '').replace(/\?=\s+=\?/g, '?==?')
        .replace(/=\?([\w-]+)\?([BbQq])\?([^?]*)\?=/g, (match, charset, encoding, encoded) => {
            try {
                const bytes = encoding.toUpperCase() === 'B'
                    ? Utilities.base64Decode(encoded)
                    : decodeQuotedPrintableWord_(encoded);

                return Utilities.newBlob(bytes).getDataAsString(charset);
            } catch (e) {
                return match;  // unknown charset or broken encoding, keep as is
            }
        });
}


/** Decode the "Q" encoding of an RFC 2047 encoded word into signed bytes. */
function decodeQuotedPrintableWord_(encoded) {
    const bytes = [];
    for (let i = 0; i < encoded.length; ++i) {
        const c = encoded[i];
        if (c === '_') {
            bytes.push(32);
        } else if (c === '=' && i + 2 < encoded.length) {
            bytes.push(parseInt(encoded.substring(i + 1, i + 3), 16) || 0);
            i += 2;
        } else {
            bytes.push(encoded.charCodeAt(i) & 0xff);
        }
    }

    return bytes.map(b => b > 127 ? b - 256 : b);
}


/** Crude conversion of an HTML body to plain text (just good enough to find header blocks). */
function htmlToText_(html) {
    return (html || '')
        .replace(/<(style|script|head)[\s\S]*?<\/\1>/gi, ' ')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/(p|div|tr|li|h[1-6]|blockquote)\s*>/gi, '\n')
        .replace(/<[^>]*>/g, '')
        .replace(/&nbsp;/gi, ' ')
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/&quot;/gi, '"')
        .replace(/&#39;/g, "'")
        .replace(/&amp;/gi, '&');
}


/** Strip leading reply/forward prefixes ("Re:", "Fwd:", "AW:", ...) off a subject. */
function stripSubjectPrefixes_(subject) {
    let result = (subject || '').trim();

    for (let stripped = true; stripped;) {
        stripped = false;
        const prefix = result.match(/^([A-Za-z]{2,4})(\[\d+])?\s*:\s*/);
        if (prefix && SUBJECT_PREFIXES.includes(prefix[1].toLowerCase())) {
            result = result.substring(prefix[0].length);
            stripped = true;
        }
    }

    return result.trim();
}


/** Sanitize a string extracted from an untrusted message for use in a mail header or body.
 *
 * Removes control characters (which would otherwise allow header injection), collapses
 * whitespace and limits the length.
 *
 * @param {string} value The untrusted value.
 * @param {number} maxLength Maximum number of characters to keep.
 *
 * @return {string} The sanitized value.
 */
function sanitizeText_(value, maxLength) {
    const sanitized = (value || '')
        .replace(/[\u0000-\u001F\u007F-\u009F\u2028\u2029]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

    return sanitized.length > maxLength ? sanitized.substring(0, maxLength - 1).trim() + '…' : sanitized;
}


/** Sanitize and validate an email address extracted from an untrusted message.
 *
 * @param {string} value The untrusted address (without display name).
 *
 * @return {string} The sanitized address or an empty string if it isn't a plausible address.
 */
function sanitizeEmailAddress_(value) {

    const address = (value || '').trim().replace(/^<|>$/g, '').trim().toLowerCase();

    // deliberately strict: this address ends up in the To header of a mail we send automatically
    if (!/^[a-z0-9!#$%&'*+/=?^_`{|}~-]+(\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@([a-z0-9]([a-z0-9-]*[a-z0-9])?\.)+[a-z]{2,63}$/.test(address)) {
        return '';
    }

    const localPart = address.substring(0, address.lastIndexOf('@'));
    if (address.length > 254 || localPart.length > 64) {
        return '';
    }

    return address;
}


/** Sanitize a Message-ID header value for use in In-Reply-To/References.
 *
 * @return {null|string} The message ID (including angle brackets) or null if it is unusable.
 */
function sanitizeMessageId_(value) {
    const messageId = (value || '').trim();

    return /^<[^<>\s]+@[^<>\s]+>$/.test(messageId) ? messageId : null;
}


/** Check whether an address belongs to one of the configured internal domains. */
function isInternalAddress_(email, config) {
    const domain = (email || '').substring((email || '').lastIndexOf('@') + 1);

    return !!domain && config.internalDomains.includes(domain);
}


/** Check whether we must never send an auto-reply to the specified address. */
function isBlockedAddress_(email, config) {
    const address = (email || '').toLowerCase();
    if (!address) {
        return true;
    }

    // replying to ourselves or to the group would create a mail loop
    if ([config.mailbox, config.groupEmail, config.fromAddress, config.replyTo].some(own => own && own.toLowerCase() === address)) {
        return true;
    }

    const localPart = address.substring(0, address.lastIndexOf('@'));

    return BLOCKED_LOCAL_PARTS.some(blocked => localPart === blocked || localPart.startsWith(blocked + '+'));
}


/** Substitute the ${...} placeholders in a template with the specified values.
 *
 * Substitution happens in a single pass, hence values are never interpreted as part of the template.
 * Placeholders without a matching value are left alone.
 */
function substitute_(template, values) {
    return (template || '').replace(/\$\{(\w+)}/g, (placeholder, key) =>
        Object.prototype.hasOwnProperty.call(values, key) ? values[key] : placeholder);
}


/** Assemble the auto-reply as a base64url encoded RFC 5322 message.
 *
 * All values interpolated into headers must have been sanitized before.
 *
 * @return {string} The message, base64url (web-safe) encoded as expected by the Gmail API.
 */
function buildAutoReply_(reply) {

    const headers = [
        'From: ' + formatAddressHeader_(reply.fromAddress, reply.fromName),
        'To: ' + formatAddressHeader_(reply.to, reply.toName),
        'Subject: ' + encodeHeaderText_(reply.subject)
    ];

    if (reply.replyTo) {
        headers.push('Reply-To: ' + formatAddressHeader_(reply.replyTo, ''));
    }

    if (reply.inReplyTo) {
        headers.push('In-Reply-To: ' + reply.inReplyTo);
        headers.push('References: ' + reply.inReplyTo);
    }

    // mark as automatic response, so well behaved mail systems don't answer it
    headers.push('Auto-Submitted: auto-replied');
    headers.push('X-Auto-Response-Suppress: All');
    headers.push('Precedence: bulk');
    headers.push('MIME-Version: 1.0');
    headers.push('Content-Type: text/plain; charset="UTF-8"');
    headers.push('Content-Transfer-Encoding: base64');

    const body = Utilities.base64Encode(reply.body, Utilities.Charset.UTF_8).match(/.{1,76}/g) || [];
    const mime = headers.join('\r\n') + '\r\n\r\n' + body.join('\r\n') + '\r\n';

    return Utilities.base64EncodeWebSafe(Utilities.newBlob(mime).getBytes());
}


/** Format an address header value, encoding and quoting the display name as needed. */
function formatAddressHeader_(email, name) {
    if (!name) {
        return email;
    }

    const displayName = /^[\x20-\x7E]*$/.test(name) ? '"' + name.replace(/["\\]/g, '') + '"' : encodeHeaderText_(name);

    return displayName + ' <' + email + '>';
}


/** Encode a header value as RFC 2047 encoded words, if it isn't plain ASCII.
 *
 * Encoded words are limited to 75 characters, hence longer values are split up.
 */
function encodeHeaderText_(value) {
    const text = value || '';
    if (/^[\x20-\x7E]*$/.test(text)) {
        return text;
    }

    // 15 characters of UTF-8 encode to 20 base64 characters, staying well below the 75 character limit
    const chunks = text.match(/[\s\S]{1,15}/g) || [];

    return chunks
        .map(chunk => '=?UTF-8?B?' + Utilities.base64Encode(chunk, Utilities.Charset.UTF_8) + '?=')
        .join('\r\n ');
}


/** Get the time an address was last auto-replied to.
 *
 * @return {number} Epoch milliseconds or 0 if the address wasn't auto-replied to recently.
 */
function getRepliedAt_(email) {
    return +(getScriptProperties_().getProperty(REPLIED_AT_PREFIX + email) || 0);
}


/** Remember that an address was just auto-replied to (skipped if the cooldown is disabled). */
function setRepliedAt_(email, cooldownHours) {
    if (cooldownHours > 0) {
        getScriptProperties_().setProperty(REPLIED_AT_PREFIX + email, '' + Date.now());
    }
}


/** Forget all auto-reply timestamps that are older than the cooldown period. */
function cleanupRepliedAtState_(cooldownHours) {

    const scriptProperties = getScriptProperties_();
    const properties = scriptProperties.getProperties() || {};
    const expiredBefore = Date.now() - (cooldownHours * 3600000);

    for (const key in properties) {
        if (key.startsWith(REPLIED_AT_PREFIX) && (cooldownHours <= 0 || +(properties[key] || 0) < expiredBefore)) {
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

    const config = {
        mailbox: get(MAILBOX_KEY, '').toLowerCase(),
        groupEmail: get(GROUP_EMAIL_KEY, '').toLowerCase(),
        searchExpression: get(SEARCH_EXPRESSION_KEY, DEFAULT_SEARCH_EXPRESSION),
        internalDomains: get(INTERNAL_DOMAINS_KEY, '').split(',').map(domain => domain.trim().toLowerCase()).filter(domain => !!domain),
        requireInternalForwarder: get(REQUIRE_INTERNAL_FORWARDER_KEY, 'true').toLowerCase() !== 'false',
        fromAddress: get(FROM_ADDRESS_KEY, '').toLowerCase(),
        fromName: sanitizeText_(get(FROM_NAME_KEY, ''), 120),
        replyTo: get(REPLY_TO_KEY, '').toLowerCase(),
        replySubject: get(REPLY_SUBJECT_KEY, DEFAULT_REPLY_SUBJECT),
        replyText: get(REPLY_TEXT_KEY, DEFAULT_REPLY_TEXT),
        handledLabel: get(HANDLED_LABEL_KEY, DEFAULT_HANDLED_LABEL),
        maxRepliesPerRun: Math.max(0, +get(MAX_REPLIES_PER_RUN_KEY, '' + DEFAULT_MAX_REPLIES_PER_RUN) || 0),
        replyCooldownHours: Math.max(0, +get(REPLY_COOLDOWN_HOURS_KEY, '' + DEFAULT_REPLY_COOLDOWN_HOURS) || 0),
        dryRun: get(DRY_RUN_KEY, 'false').toLowerCase() === 'true'
    };

    if (!config.mailbox) {
        throw new Error('No mailbox to read configured at script property ' + MAILBOX_KEY);
    }

    if (config.requireInternalForwarder && !config.internalDomains.length) {
        throw new Error('No internal domains configured at script property ' + INTERNAL_DOMAINS_KEY
            + ' (required unless ' + REQUIRE_INTERNAL_FORWARDER_KEY + ' is false)');
    }

    if (config.handledLabel.includes('"')) {
        throw new Error('Invalid label name at script property ' + HANDLED_LABEL_KEY + ': ' + config.handledLabel);
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
