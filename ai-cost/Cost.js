/**
 * Fetch AI usage and cost statistics from LLM provider APIs and write to a Google Sheet.
 *
 * Supported sources: anthropic, claude-code, openai, Claude Team invoices (PDF attachments in a Gmail mailbox)
 *
 * All costs are written in EUR (we are billed in EUR). Provider APIs report USD, which is
 * converted using the ECB reference rate of the usage day (see Currency.convertCurrency).
 *
 * Seat based (subscription) costs live in the manually maintained Static-Data-<year> sheet.
 * The script keeps the current month's rows present (carrying seat counts and EUR prices forward
 * from the previous month) and fills in the informational EUR->USD reference rate. Seat tiers of
 * the claude.ai Team plan are not exposed by any API. They, and the extra usage of the Team plan, are
 * taken from the invoices Anthropic sends by mail (see fetchClaudeInvoices_). Without invoices,
 * user_count must be corrected by hand from the claude.ai members export.
 *
 * Script Properties:
 *   AiCost.anthropicAdminKey          Anthropic admin API key (for anthropic + claude-code sources)
 *   AiCost.openaiAdminKey             OpenAI admin API key (for openai source)
 *   AiCost.invoiceMailbox             Mailbox receiving the Claude Team invoices (e.g. automation@example.com)
 *   AiCost.serviceAccountCredentials  Service account credentials JSON (domain-wide delegation, scope
 *                                     https://mail.google.com/), for reading invoiceMailbox
 *   AiCost.invoiceSearchExpression    Optional, overrides INVOICE_DEFAULT_SEARCH (the date range is always added)
 *
 * Managed via: https://github.com/giantswarm/gapps-automation
 */

// --- Section A: Constants & configuration ---

const PROPERTY_PREFIX = 'AiCost.';

/** The trigger handler function to call in time based triggers. */
const TRIGGER_HANDLER_FUNCTION = 'fetchAiCosts';

const ANTHROPIC_ADMIN_KEY_PROP = PROPERTY_PREFIX + 'anthropicAdminKey';
const OPENAI_ADMIN_KEY_PROP = PROPERTY_PREFIX + 'openaiAdminKey';

const COLUMNS = [
    'date', 'source', 'record_type', 'model', 'actor',
    'input_tokens', 'output_tokens', 'cache_read_tokens', 'cache_creation_tokens',
    'num_requests', 'cost_eur', 'cost_type', 'sessions', 'metadata',
];

/** Currency all costs are written in. */
const COST_CURRENCY = 'EUR';
/** Currency assumed for provider amounts that do not carry an explicit currency. */
const PROVIDER_DEFAULT_CURRENCY = 'USD';

/** Name prefix of the sheet holding manually maintained seat based costs. */
const STATIC_SHEET_PREFIX = 'Static-Data-';
/** Header of the static data sheet (created if the sheet is missing). */
const STATIC_COLUMNS = ['month', 'provider', 'user_count', 'type', 'cost_cost_per_seat', 'eur-to-usd'];
/** Provider whose seat rows are maintained by the script. */
const STATIC_PROVIDER = 'claude-code';
/** Prefix of cell notes written by the script (cells with other, manual notes are never overwritten). */
const STATIC_AUTO_NOTE_PREFIX = 'Auto:';
/** Static data type of the Team plan extra usage (monthly sum of the invoices, user_count 1). */
const STATIC_EXTRA_USAGE_TYPE = 'Extra Usage';
/** Static data type of prorated seat charges and credits (monthly sum of the invoices, user_count 1). */
const STATIC_SEAT_ADJUSTMENT_TYPE = 'Seat Adjustment';
/** Static data types summed up from invoices, which are never carried forward into a new month. */
const STATIC_INVOICE_TYPES = [STATIC_EXTRA_USAGE_TYPE, STATIC_SEAT_ADJUSTMENT_TYPE];

const INVOICE_MAILBOX_PROP = PROPERTY_PREFIX + 'invoiceMailbox';
const SERVICE_ACCOUNT_CREDENTIALS_PROP = PROPERTY_PREFIX + 'serviceAccountCredentials';
const INVOICE_SEARCH_PROP = PROPERTY_PREFIX + 'invoiceSearchExpression';
/** Gmail search expression selecting candidate invoice mails (sent by Stripe on behalf of Anthropic, each with
 *  an Invoice-<number>.pdf and a Receipt-<number>.pdf attachment). */
const INVOICE_DEFAULT_SEARCH = 'from:invoice+statements@mail.anthropic.com has:attachment filename:pdf';
/** Name prefix of the sheets listing every parsed invoice line item (one sheet per invoice year). */
const INVOICE_SHEET_PREFIX = 'Invoices-';
const INVOICE_COLUMNS = ['date', 'invoice', 'kind', 'tier', 'description', 'quantity', 'unit_price', 'amount',
    'currency', 'amount_eur', 'message_id'];

const ANTHROPIC_BASE = 'https://api.anthropic.com';
const OPENAI_BASE = 'https://api.openai.com';

/** Anthropic standard-tier pricing per token (USD). Used to estimate per-model costs from the usage
 *  report (record_type usage, cost_type estimated), as a cross-check of the billed cost report rows.
 *  Cache reads are 0.1x base input (0.025x on Fable 5.1 / Mythos 5.1, 0.05x on Opus 5.5), 5m cache
 *  writes are 1.25x, 1h cache writes are 2x. The 1M context window is billed at standard rates.
 *  Source: https://platform.claude.com/docs/en/about-claude/pricing (checked 2026-10-07)
 */
const ANTHROPIC_PRICING = {
    'claude-fable-5-1':  { input: 10.00 / 1e6, cache_read: 0.25 / 1e6, cache_5m: 12.50 / 1e6, cache_1h: 20.00 / 1e6, output: 50.00 / 1e6 },
    'claude-mythos-5-1': { input: 10.00 / 1e6, cache_read: 0.25 / 1e6, cache_5m: 12.50 / 1e6, cache_1h: 20.00 / 1e6, output: 50.00 / 1e6 },
    'claude-fable-5':    { input: 10.00 / 1e6, cache_read: 1.00 / 1e6, cache_5m: 12.50 / 1e6, cache_1h: 20.00 / 1e6, output: 50.00 / 1e6 },
    'claude-mythos-5':   { input: 10.00 / 1e6, cache_read: 1.00 / 1e6, cache_5m: 12.50 / 1e6, cache_1h: 20.00 / 1e6, output: 50.00 / 1e6 },
    'claude-opus-5-5':   { input:  4.00 / 1e6, cache_read: 0.20 / 1e6, cache_5m:  5.00 / 1e6, cache_1h:  8.00 / 1e6, output: 20.00 / 1e6 },
    'claude-opus-5':     { input:  5.00 / 1e6, cache_read: 0.50 / 1e6, cache_5m:  6.25 / 1e6, cache_1h: 10.00 / 1e6, output: 25.00 / 1e6 },
    'claude-opus-4-8':   { input:  5.00 / 1e6, cache_read: 0.50 / 1e6, cache_5m:  6.25 / 1e6, cache_1h: 10.00 / 1e6, output: 25.00 / 1e6 },
    'claude-opus-4-7':   { input:  5.00 / 1e6, cache_read: 0.50 / 1e6, cache_5m:  6.25 / 1e6, cache_1h: 10.00 / 1e6, output: 25.00 / 1e6 },
    'claude-opus-4-6':   { input:  5.00 / 1e6, cache_read: 0.50 / 1e6, cache_5m:  6.25 / 1e6, cache_1h: 10.00 / 1e6, output: 25.00 / 1e6 },
    'claude-opus-4-5':   { input:  5.00 / 1e6, cache_read: 0.50 / 1e6, cache_5m:  6.25 / 1e6, cache_1h: 10.00 / 1e6, output: 25.00 / 1e6 },
    'claude-opus-4-1':   { input: 15.00 / 1e6, cache_read: 1.50 / 1e6, cache_5m: 18.75 / 1e6, cache_1h: 30.00 / 1e6, output: 75.00 / 1e6 },
    'claude-opus-4':     { input: 15.00 / 1e6, cache_read: 1.50 / 1e6, cache_5m: 18.75 / 1e6, cache_1h: 30.00 / 1e6, output: 75.00 / 1e6 },
    'claude-sonnet-5-5': { input:  2.00 / 1e6, cache_read: 0.20 / 1e6, cache_5m:  2.50 / 1e6, cache_1h:  4.00 / 1e6, output: 10.00 / 1e6 },
    'claude-sonnet-5':   { input:  2.00 / 1e6, cache_read: 0.20 / 1e6, cache_5m:  2.50 / 1e6, cache_1h:  4.00 / 1e6, output: 10.00 / 1e6 },
    'claude-sonnet-4-6': { input:  3.00 / 1e6, cache_read: 0.30 / 1e6, cache_5m:  3.75 / 1e6, cache_1h:  6.00 / 1e6, output: 15.00 / 1e6 },
    'claude-sonnet-4-5': { input:  3.00 / 1e6, cache_read: 0.30 / 1e6, cache_5m:  3.75 / 1e6, cache_1h:  6.00 / 1e6, output: 15.00 / 1e6 },
    'claude-sonnet-4':   { input:  3.00 / 1e6, cache_read: 0.30 / 1e6, cache_5m:  3.75 / 1e6, cache_1h:  6.00 / 1e6, output: 15.00 / 1e6 },
    'claude-haiku-4-5':  { input:  1.00 / 1e6, cache_read: 0.10 / 1e6, cache_5m:  1.25 / 1e6, cache_1h:  2.00 / 1e6, output:  5.00 / 1e6 },
    'claude-haiku-3-5':  { input:  0.80 / 1e6, cache_read: 0.08 / 1e6, cache_5m:  1.00 / 1e6, cache_1h:  1.60 / 1e6, output:  4.00 / 1e6 },
};

/** OpenAI standard-tier pricing per token (USD). Used to estimate per-model costs
 *  from the usage endpoint, since the costs endpoint has no model breakdown.
 *  Source: https://developers.openai.com/api/docs/pricing
 */
const OPENAI_PRICING = {
    'gpt-5':        { input: 1.25 / 1e6, cached: 0.125 / 1e6, output: 10.00 / 1e6 },
    'gpt-5-mini':   { input: 0.25 / 1e6, cached: 0.025 / 1e6, output: 2.00  / 1e6 },
    'gpt-5.1':      { input: 1.25 / 1e6, cached: 0.125 / 1e6, output: 10.00 / 1e6 },
    'gpt-4.1':      { input: 2.00 / 1e6, cached: 0.50  / 1e6, output: 8.00  / 1e6 },
    'gpt-4.1-mini': { input: 0.40 / 1e6, cached: 0.10  / 1e6, output: 1.60  / 1e6 },
    'gpt-4.1-nano': { input: 0.10 / 1e6, cached: 0.025 / 1e6, output: 0.40  / 1e6 },
    'o3':           { input: 2.00 / 1e6, cached: 0.50  / 1e6, output: 8.00  / 1e6 },
    'o3-mini':      { input: 1.10 / 1e6, cached: 0.275 / 1e6, output: 4.40  / 1e6 },
    'o4-mini':      { input: 1.10 / 1e6, cached: 0.275 / 1e6, output: 4.40  / 1e6 },
};


// --- Section B: Entry points ---

/** Main entry point. Fetches cost data for the current UTC day and writes to the configured sheet.
 *
 * Existing rows for the fetched dates are replaced (de-duplicated), all other rows are preserved.
 */
function fetchAiCosts() {
    const endDate = defaultEndDate_();
    // Query the last 7 days to account for reporting delays in analytics/cost
    // APIs.  Rows are de-duplicated by (date, source).
    const d = new Date(toIso8601_(defaultStartDate_()));
    d.setUTCDate(d.getUTCDate() - 6);
    const startDate = d.toISOString().slice(0, 10);

    fetchAiCostsForRange_(startDate, endDate);
}

function backfillAiCosts(startDate, endDate) {
    if (!startDate || !endDate) {
        throw new Error('startDate and endDate are required (YYYY-MM-DD, endDate exclusive)');
    }
    fetchAiCostsForRange_(startDate, endDate);
}

/** Throwaway: re-fetch 2026 from 2026-01-01 up to today, one call per month to stay within API bucket limits.
 *  Completed months are remembered in a script property, so re-running after a timeout resumes. Remove once done. */
function backfill2026() {
    const props = getScriptProperties_();
    const doneProp = PROPERTY_PREFIX + 'backfill2026Done';
    const endDate = defaultEndDate_();
    const pad = function(m) { return (m < 10 ? '0' : '') + m; };
    for (let month = 1; month <= 12; month++) {
        const start = '2026-' + pad(month) + '-01';
        if (start >= endDate) break;
        const done = (props.getProperty(doneProp) || '').split(',').filter(String);
        if (done.indexOf(String(month)) >= 0) continue;
        const nextMonth = month === 12 ? '2027-01-01' : '2026-' + pad(month + 1) + '-01';
        backfillAiCosts(start, nextMonth < endDate ? nextMonth : endDate);
        props.setProperty(doneProp, done.concat(String(month)).join(','));
    }
    Logger.log('Backfill 2026 complete');
}

function fetchAiCostsForRange_(startDate, endDate) {

    const props = getScriptProperties_();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    const anthropicKey = props.getProperty(ANTHROPIC_ADMIN_KEY_PROP) || '';
    const openaiKey = props.getProperty(OPENAI_ADMIN_KEY_PROP) || '';

    Logger.log('Fetching AI costs for %s to %s', startDate, endDate);

    const rows = [];
    let firstError = null;

    // Anthropic API usage + costs
    if (anthropicKey) {
        try {
            Logger.log('Fetching anthropic...');
            const anthropicRows = fetchAnthropicUsage_(anthropicKey, startDate, endDate)
                .concat(fetchAnthropicCosts_(anthropicKey, startDate, endDate));
            rows.push(...anthropicRows);
            Logger.log('  anthropic: %s rows', anthropicRows.length);
        } catch (e) {
            Logger.log('Failed to fetch anthropic: %s', e.message);
            firstError = firstError || e;
        }
    } else {
        Logger.log('Skipping anthropic: %s not set', ANTHROPIC_ADMIN_KEY_PROP);
    }

    // Claude Code usage (uses Anthropic key)
    if (anthropicKey) {
        try {
            Logger.log('Fetching claude-code...');
            const ccRows = fetchClaudeCodeUsage_(anthropicKey, startDate, endDate);
            rows.push(...ccRows);
            Logger.log('  claude-code: %s rows', ccRows.length);
        } catch (e) {
            Logger.log('Failed to fetch claude-code: %s', e.message);
            firstError = firstError || e;
        }
    } else {
        Logger.log('Skipping claude-code: %s not set', ANTHROPIC_ADMIN_KEY_PROP);
    }

    // OpenAI usage + costs
    if (openaiKey) {
        try {
            Logger.log('Fetching openai...');
            const openaiRows = fetchOpenaiUsage_(openaiKey, startDate, endDate)
                .concat(fetchOpenaiCosts_(openaiKey, startDate, endDate));
            rows.push(...openaiRows);
            Logger.log('  openai: %s rows', openaiRows.length);
        } catch (e) {
            Logger.log('Failed to fetch openai: %s', e.message);
            firstError = firstError || e;
        }
    } else {
        Logger.log('Skipping openai: %s not set', OPENAI_ADMIN_KEY_PROP);
    }

    // Collect (date, source) keys present in fetched data for targeted dedup
    const fetchedKeys = new Set();
    for (const row of rows) {
        fetchedKeys.add(row.date + '|' + row.source);
    }

    if (rows.length > 0) {
        rows.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
        appendToSheet_(spreadsheet, rows, fetchedKeys);
    }

    // Seat based costs (static data): keep the current month's rows and exchange rate up to date
    try {
        updateStaticData_(spreadsheet);
    } catch (e) {
        Logger.log('Failed to update static data: %s', e.message);
        firstError = firstError || e;
    }

    // Claude Team invoices (seats + extra usage), after the static data rows of the month exist
    const invoiceMailbox = props.getProperty(INVOICE_MAILBOX_PROP) || '';
    if (invoiceMailbox) {
        try {
            Logger.log('Fetching Claude Team invoices from %s...', invoiceMailbox);
            fetchClaudeInvoices_(spreadsheet, invoiceMailbox, startDate, endDate);
        } catch (e) {
            Logger.log('Failed to fetch Claude Team invoices: %s', e.message);
            firstError = firstError || e;
        }
    } else {
        Logger.log('Skipping Claude Team invoices: %s not set', INVOICE_MAILBOX_PROP);
    }

    Logger.log('Done. %s total new rows.', rows.length);

    if (firstError) {
        throw firstError;
    }
}


/** Import the Claude Team invoices of a date range (e.g. from another mailbox the invoices were sent to before).
 *
 * @param {string} startDate First invoice day (YYYY-MM-DD).
 * @param {string} endDate Day after the last invoice day (YYYY-MM-DD, exclusive).
 * @param {undefined|string} mailbox Optional mailbox to search instead of the configured AiCost.invoiceMailbox.
 * @param {undefined|string} searchExpression Optional Gmail search expression replacing the configured one
 *                                            (e.g. 'filename:Invoice-ABCD1234' for invoices forwarded by someone).
 */
function backfillClaudeInvoices(startDate, endDate, mailbox, searchExpression) {
    if (!startDate || !endDate) {
        throw new Error('startDate and endDate are required (YYYY-MM-DD, endDate exclusive)');
    }
    const effectiveMailbox = mailbox || getScriptProperties_().getProperty(INVOICE_MAILBOX_PROP);
    if (!effectiveMailbox) {
        throw new Error('No mailbox given and ' + INVOICE_MAILBOX_PROP + ' not set');
    }
    fetchClaudeInvoices_(SpreadsheetApp.getActiveSpreadsheet(), effectiveMailbox, startDate, endDate, searchExpression);
}


/** Throwaway: mailbox the Claude Team invoices were sent to before they went to AiCost.invoiceMailbox.
 *  Set it only in the deployed copy (Apps Script editor or a local edit followed by `make ai-cost/`), never commit
 *  the actual address. */
const INVOICE_BACKFILL_MAILBOX = 'mailbox@example.com';

/** Throwaway: import the 2026 Claude Team invoices of INVOICE_BACKFILL_MAILBOX, one call per month.
 *
 *  Only mails matching INVOICE_DEFAULT_SEARCH (Anthropic's invoice sender) are read. Invoices already recorded
 *  (e.g. also received by AiCost.invoiceMailbox) are skipped by invoice number. Completed months are remembered in a
 *  script property, so re-running after a timeout resumes. The impersonation token of the mailbox is deleted
 *  afterwards. Remove once done. */
function backfillClaudeInvoices2026() {
    if (!/^[^@\s]+@[^@\s]+$/.test(INVOICE_BACKFILL_MAILBOX) || /@example\.com$/.test(INVOICE_BACKFILL_MAILBOX)) {
        throw new Error('Set INVOICE_BACKFILL_MAILBOX in the deployed script first');
    }
    const props = getScriptProperties_();
    const doneProp = PROPERTY_PREFIX + 'invoiceBackfill2026Done';
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const pad = function(m) { return (m < 10 ? '0' : '') + m; };
    try {
        for (let month = 1; month <= 10; month++) {
            const done = (props.getProperty(doneProp) || '').split(',').filter(String);
            if (done.indexOf(String(month)) >= 0) continue;
            // AiCost.invoiceMailbox has the invoices from 2026-10-14 on, the overlap is skipped by invoice number
            const end = month === 10 ? '2026-10-16' : '2026-' + pad(month + 1) + '-01';
            fetchClaudeInvoices_(spreadsheet, INVOICE_BACKFILL_MAILBOX, '2026-' + pad(month) + '-01', end, INVOICE_DEFAULT_SEARCH);
            props.setProperty(doneProp, done.concat(String(month)).join(','));
        }
    } finally {
        // don't keep an access token to the mailbox in the script properties
        OAuth2.createService('GmailClientV1-' + INVOICE_BACKFILL_MAILBOX)
            .setPropertyStore(PropertiesService.getScriptProperties())
            .reset();
    }
    Logger.log('Invoice backfill 2026 complete');
}


/** Ensure the current month's seat rows exist in Static-Data-<year> and refresh their exchange rate. */
function updateStaticData() {
    updateStaticData_(SpreadsheetApp.getActiveSpreadsheet());
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


// --- Section C: Utility functions ---

function getScriptProperties_() {
    return PropertiesService.getScriptProperties();
}

function defaultStartDate_() {
    return new Date().toISOString().slice(0, 10);
}

function defaultEndDate_() {
    const d = new Date();
    d.setUTCDate(d.getUTCDate() + 1);
    return d.toISOString().slice(0, 10);
}

function toIso8601_(dateStr) {
    return dateStr + 'T00:00:00Z';
}

function toUnixSeconds_(dateStr) {
    return Math.floor(new Date(toIso8601_(dateStr)).getTime() / 1000);
}

/** Build a row object with defaults for all columns. Only non-default fields need to be passed. */
function makeRow_(fields) {
    return {
        date: '', source: '', record_type: '', model: '', actor: '',
        input_tokens: 0, output_tokens: 0, cache_read_tokens: 0, cache_creation_tokens: 0,
        num_requests: 0, cost_eur: 0, cost_type: '', sessions: 0, metadata: '',
        ...fields,
    };
}

function dateRange_(start, end) {
    const dates = [];
    const cur = new Date(toIso8601_(start));
    const stop = new Date(toIso8601_(end));
    while (cur < stop) {
        dates.push(cur.toISOString().slice(0, 10));
        cur.setUTCDate(cur.getUTCDate() + 1);
    }
    return dates;
}


/** Convert a provider reported amount to COST_CURRENCY at the reference rate of the given day.
 *
 * @param {number} amount The amount as reported by the provider.
 * @param {string} currency ISO 4217 code of the amount (falls back to PROVIDER_DEFAULT_CURRENCY).
 * @param {string} date Usage day (YYYY-MM-DD) whose exchange rate to apply.
 * @return {number} The amount in COST_CURRENCY (not rounded).
 */
function toCostCurrency_(amount, currency, date) {
    const value = +amount || 0;
    if (value === 0) return 0;
    return Currency.convertCurrency(currency || PROVIDER_DEFAULT_CURRENCY, COST_CURRENCY, value, date);
}


/** Estimate cost in COST_CURRENCY from token counts using the ANTHROPIC_PRICING table (USD list
 *  prices, converted at the reference rate of the usage day).
 *  Returns 0 if the model is not in the table.
 *  Model IDs like "claude-sonnet-4-5-20250929" (date suffix) or "claude-opus-4-8[1m]"
 *  (long-context suffix) are matched by stripping the suffix.
 */
function estimateAnthropicCost_(date, model, inputTokens, cacheReadTokens, cache5mTokens, cache1hTokens, outputTokens) {
    const base = (model || '').replace(/\[1m\]$/, '').replace(/-\d{8}$/, '');
    const p = ANTHROPIC_PRICING[base];
    if (!p) return 0;
    const usd = (inputTokens * p.input) + (cacheReadTokens * p.cache_read)
        + (cache5mTokens * p.cache_5m) + (cache1hTokens * p.cache_1h)
        + (outputTokens * p.output);
    return toCostCurrency_(usd, 'USD', date);
}


/** Estimate cost in COST_CURRENCY from token counts using the OPENAI_PRICING table (USD list
 *  prices, converted at the reference rate of the usage day).
 *  Returns 0 if the model is not in the table.
 *  Model IDs like "gpt-5.1-2025-11-13" are matched by stripping the date suffix.
 */
function estimateOpenaiCost_(date, model, inputTokens, cachedTokens, outputTokens) {
    const base = (model || '').replace(/-\d{4}-\d{2}-\d{2}$/, '');
    const p = OPENAI_PRICING[base];
    if (!p) return 0;
    const usd = (inputTokens * p.input) + (cachedTokens * p.cached) + (outputTokens * p.output);
    return toCostCurrency_(usd, 'USD', date);
}


// --- Section D: Generic paginated fetch ---

function fetchAllPages_(url, headers, parsePageFn, buildNextUrlFn) {
    const allItems = [];
    let currentUrl = url;

    while (currentUrl) {
        Logger.log('  GET %s', currentUrl.replace(/key=[^&]+/, 'key=***'));

        let response;
        for (let attempt = 0; attempt < 3; attempt++) {
            response = UrlFetchApp.fetch(currentUrl, {
                method: 'get',
                headers: headers,
                muteHttpExceptions: true,
            });

            if (response.getResponseCode() === 429) {
                const retryAfter = parseInt(response.getHeaders()['retry-after'] || '5', 10);
                Logger.log('  Rate limited, retrying after %ss...', retryAfter);
                Utilities.sleep(retryAfter * 1000);
                continue;
            }
            break;
        }

        const code = response.getResponseCode();
        if (code < 200 || code >= 300) {
            throw new Error('HTTP ' + code + ' from ' + currentUrl + ': '
                + response.getContentText().substring(0, 500));
        }

        const json = JSON.parse(response.getContentText());
        const items = parsePageFn(json);
        allItems.push(...items);

        currentUrl = buildNextUrlFn(json);
    }

    return allItems;
}


// --- Section E: Anthropic API source ---

function anthropicHeaders_(apiKey) {
    return {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
    };
}

function fetchAnthropicUsage_(apiKey, startDate, endDate) {
    const url = ANTHROPIC_BASE + '/v1/organizations/usage_report/messages'
        + '?starting_at=' + toIso8601_(startDate)
        + '&ending_at=' + toIso8601_(endDate)
        + '&bucket_width=1d&group_by[]=model';

    return fetchAllPages_(url, anthropicHeaders_(apiKey),
        function(json) {
            const rows = [];
            for (const bucket of json.data || []) {
                const date = (bucket.starting_at || '').slice(0, 10);
                for (const result of bucket.results || []) {
                    const inputTok = result.uncached_input_tokens || 0;
                    const outputTok = result.output_tokens || 0;
                    const cacheReadTok = result.cache_read_input_tokens || 0;
                    const cache5mTok = result.cache_creation?.ephemeral_5m_input_tokens || 0;
                    const cache1hTok = result.cache_creation?.ephemeral_1h_input_tokens || 0;
                    rows.push(makeRow_({
                        date: date, source: 'anthropic', record_type: 'usage',
                        model: result.model || '',
                        input_tokens: inputTok,
                        output_tokens: outputTok,
                        cache_read_tokens: cacheReadTok,
                        cache_creation_tokens: cache5mTok + cache1hTok,
                        cost_eur: estimateAnthropicCost_(date, result.model, inputTok, cacheReadTok, cache5mTok, cache1hTok, outputTok),
                        cost_type: 'estimated',
                    }));
                }
            }
            return rows;
        },
        function(json) { return json.has_more ? url + '&page=' + encodeURIComponent(json.next_page) : null; }
    );
}

/** Fetch actual (billed) Anthropic costs, one row per (day, model, cost_type).
 *
 * Grouping by description makes the API return one item per model and token type (input, output, cache ...);
 * these are summed into a single row, keeping the per token type breakdown (EUR) in the metadata column.
 */
function fetchAnthropicCosts_(apiKey, startDate, endDate) {
    const url = ANTHROPIC_BASE + '/v1/organizations/cost_report'
        + '?starting_at=' + toIso8601_(startDate)
        + '&ending_at=' + toIso8601_(endDate)
        + '&bucket_width=1d&limit=31&group_by[]=description';

    return fetchAllPages_(url, anthropicHeaders_(apiKey),
        function(json) {
            const rows = [];
            for (const bucket of json.data || []) {
                const date = (bucket.starting_at || '').slice(0, 10);
                const byKey = new Map();
                for (const result of bucket.results || []) {
                    const model = result.model || '';
                    const costType = result.cost_type || '';
                    const key = model + '|' + costType;
                    let row = byKey.get(key);
                    if (!row) {
                        row = makeRow_({date: date, source: 'anthropic', record_type: 'cost', model: model, cost_type: costType});
                        row.breakdown = {};
                        byKey.set(key, row);
                    }
                    const eur = toCostCurrency_(parseFloat(result.amount || '0') / 100, result.currency, date);
                    const detail = result.token_type || result.description || 'other';
                    row.cost_eur += eur;
                    row.breakdown[detail] = (row.breakdown[detail] || 0) + eur;
                }
                for (const row of byKey.values()) {
                    for (const detail of Object.keys(row.breakdown)) {
                        row.breakdown[detail] = SheetUtil.round(row.breakdown[detail], 4);
                    }
                    row.metadata = JSON.stringify(row.breakdown);
                    delete row.breakdown;
                    rows.push(row);
                }
            }
            return rows;
        },
        function(json) { return json.has_more ? url + '&page=' + encodeURIComponent(json.next_page) : null; }
    );
}


// --- Section F: Claude Code source ---

function fetchClaudeCodeUsage_(apiKey, startDate, endDate) {
    const allRows = [];
    const days = dateRange_(startDate, endDate);

    for (const day of days) {
        Logger.log('  Claude Code: fetching day %s', day);
        const ccUrl = ANTHROPIC_BASE + '/v1/organizations/usage_report/claude_code?starting_at=' + day + '&limit=1000';
        const dayRows = fetchAllPages_(
            ccUrl,
            anthropicHeaders_(apiKey),
            function(json) {
                const rows = [];
                for (const record of json.data || []) {
                    const actor = record.actor?.type === 'user_actor'
                        ? record.actor?.email_address || ''
                        : record.actor?.api_key_name || '';
                    const core = record.core_metrics || {};
                    const meta = {
                        lines_added: core.lines_of_code?.added ?? null,
                        lines_removed: core.lines_of_code?.removed ?? null,
                        commits: core.commits_by_claude_code ?? null,
                        pull_requests: core.pull_requests_by_claude_code ?? null,
                        terminal_type: record.terminal_type ?? null,
                        customer_type: record.customer_type ?? null,
                        tool_actions: record.tool_actions ?? null,
                    };

                    for (const mb of record.model_breakdown || []) {
                        const tokens = mb.tokens || {};
                        const costCents = mb.estimated_cost?.amount || 0;
                        const costCurrency = mb.estimated_cost?.currency;
                        rows.push(makeRow_({
                            date: day, source: 'claude-code', record_type: 'usage',
                            model: mb.model || '', actor: actor,
                            input_tokens: tokens.input || 0,
                            output_tokens: tokens.output || 0,
                            cache_read_tokens: tokens.cache_read || 0,
                            cache_creation_tokens: tokens.cache_creation || 0,
                            cost_eur: toCostCurrency_(costCents / 100, costCurrency, day), cost_type: 'tokens',
                            sessions: core.num_sessions || 0,
                            metadata: JSON.stringify(meta),
                        }));
                    }

                    // If no model_breakdown, still emit a row with aggregate data
                    if (!record.model_breakdown || record.model_breakdown.length === 0) {
                        rows.push(makeRow_({
                            date: day, source: 'claude-code', record_type: 'usage',
                            actor: actor,
                            sessions: core.num_sessions || 0,
                            metadata: JSON.stringify(meta),
                        }));
                    }
                }
                return rows;
            },
            function(json) { return json.has_more ? ccUrl + '&page=' + encodeURIComponent(json.next_page) : null; }
        );
        allRows.push(...dayRows);
    }

    return allRows;
}


// --- Section G: OpenAI source ---

function fetchOpenaiUsage_(apiKey, startDate, endDate) {
    const url = OPENAI_BASE + '/v1/organization/usage/completions'
        + '?start_time=' + toUnixSeconds_(startDate)
        + '&end_time=' + toUnixSeconds_(endDate)
        + '&bucket_width=1d&group_by[]=model';

    return fetchAllPages_(url,
        {'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json'},
        function(json) {
            const rows = [];
            for (const bucket of json.data || []) {
                for (const result of bucket.results || []) {
                    const inputTok = result.input_tokens || 0;
                    const cachedTok = result.input_cached_tokens || 0;
                    const outputTok = result.output_tokens || 0;
                    const date = new Date((bucket.start_time || 0) * 1000).toISOString().slice(0, 10);
                    rows.push(makeRow_({
                        date: date,
                        source: 'openai', record_type: 'usage',
                        model: result.model || '',
                        input_tokens: inputTok,
                        output_tokens: outputTok,
                        cache_read_tokens: cachedTok,
                        num_requests: result.num_model_requests || 0,
                        cost_eur: estimateOpenaiCost_(date, result.model, inputTok, cachedTok, outputTok),
                        cost_type: 'estimated',
                    }));
                }
            }
            return rows;
        },
        function(json) { return json.has_more && json.next_page ? url + '&page=' + encodeURIComponent(json.next_page) : null; }
    );
}

function fetchOpenaiCosts_(apiKey, startDate, endDate) {
    const url = OPENAI_BASE + '/v1/organization/costs'
        + '?start_time=' + toUnixSeconds_(startDate)
        + '&end_time=' + toUnixSeconds_(endDate)
        + '&bucket_width=1d';

    return fetchAllPages_(url,
        {'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json'},
        function(json) {
            const rows = [];
            for (const bucket of json.data || []) {
                const date = new Date((bucket.start_time || 0) * 1000).toISOString().slice(0, 10);
                for (const result of bucket.results || []) {
                    rows.push(makeRow_({
                        date: date, source: 'openai', record_type: 'cost',
                        cost_eur: toCostCurrency_(result.amount?.value || 0, result.amount?.currency, date),
                        cost_type: result.line_item || '',
                    }));
                }
            }
            return rows;
        },
        function(json) { return json.has_more && json.next_page ? url + '&page=' + encodeURIComponent(json.next_page) : null; }
    );
}


// --- Section H: Sheet operations ---

/** Write fetched rows to the target sheet.
 *
 * Removes any existing rows whose (date, source) key is in fetchedKeys, then appends all new rows.
 */
function appendToSheet_(spreadsheet, newRows, fetchedKeys) {
    const sheet = SheetUtil.ensureSheet(spreadsheet, 'Data-' + new Date().getUTCFullYear());

    let lastRow = sheet.getLastRow();

    // Ensure header row exists
    if (lastRow === 0) {
        sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
        lastRow = 1;
    }

    // Remove existing rows matching fetched (date, source) keys (contiguous ranges, bottom-to-top)
    if (lastRow > 1 && fetchedKeys.size > 0) {
        const tz = spreadsheet.getSpreadsheetTimeZone();
        const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        let i = data.length - 1;
        while (i >= 0) {
            const val = data[i][0];
            const dateStr = val instanceof Date
                ? Utilities.formatDate(val, tz, 'yyyy-MM-dd') : String(val);
            const key = dateStr + '|' + String(data[i][1]);
            if (fetchedKeys.has(key)) {
                const rangeEnd = i;
                while (i > 0) {
                    const prev = data[i - 1][0];
                    const prevStr = prev instanceof Date
                        ? Utilities.formatDate(prev, tz, 'yyyy-MM-dd') : String(prev);
                    const prevKey = prevStr + '|' + String(data[i - 1][1]);
                    if (!fetchedKeys.has(prevKey)) break;
                    i--;
                }
                sheet.deleteRows(i + 2, rangeEnd - i + 1);
            }
            i--;
        }
    }

    // Append new rows
    if (newRows.length > 0) {
        const rowArrays = newRows.map(function(obj) {
            return COLUMNS.map(function(col) { return obj[col] !== undefined ? obj[col] : ''; });
        });
        sheet.getRange(sheet.getLastRow() + 1, 1, rowArrays.length, COLUMNS.length).setValues(rowArrays);
    }

    Logger.log('Sheet updated: %s data rows', sheet.getLastRow() - 1);
}


// --- Section I: Static data (seat based costs) ---

/** Keep the seat cost rows of the current (UTC) month in Static-Data-<year> up to date.
 *
 * - If no seat rows for (current month, STATIC_PROVIDER) exist, they are created by copying the seat rows of the most
 *   recent earlier month (falling back to the previous year's sheet in January). Seat counts and EUR prices
 *   are carried forward unchanged; user_count gets a note asking for a manual update from the members export,
 *   because no API exposes the claude.ai Team seat tiers.
 * - The eur-to-usd column of the current month's rows is set to the latest EUR->USD reference rate (4 decimals).
 *   It is informational only (seat prices are EUR) and refreshed on every run until the month is over.
 *   Cells carrying a note that does not start with STATIC_AUTO_NOTE_PREFIX are treated as manual and left alone.
 *
 * Past months are never modified.
 *
 * @param {Spreadsheet} spreadsheet The spreadsheet holding the Static-Data-<year> sheets.
 */
function updateStaticData_(spreadsheet) {
    const now = new Date();
    const year = now.getUTCFullYear();
    const month = now.getUTCMonth() + 1;
    const today = now.toISOString().slice(0, 10);

    const sheet = SheetUtil.ensureSheet(spreadsheet, STATIC_SHEET_PREFIX + year);
    if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, STATIC_COLUMNS.length).setValues([STATIC_COLUMNS]);
    }
    const table = readStaticTable_(sheet);

    const rate = SheetUtil.round(Currency.convertCurrency('EUR', 'USD', 1), 4);
    const rateNote = STATIC_AUTO_NOTE_PREFIX + ' ECB reference rate EUR\u2192USD of ' + today
        + ', refreshed daily while the month is current. Informational only, seat prices are EUR.'
        + ' Replace the note to pin a manual value (e.g. the rate implied by the invoice).';

    const current = table.rows.filter(function(r) { return isCarriedForward_(r) && r.month === month; });
    if (current.length > 0) {
        let updated = 0;
        for (const r of current) {
            const cell = sheet.getRange(r.rowNumber, table.col.rate + 1);
            const note = cell.getNote() || '';
            if (note && note.indexOf(STATIC_AUTO_NOTE_PREFIX) !== 0) {
                continue; // manual value, leave alone
            }
            if (r.rate !== rate || note !== rateNote) {
                cell.setValue(rate).setNote(rateNote);
                updated++;
            }
        }
        Logger.log('Static data: %s/%s rows for %s-%s refreshed (EUR->USD %s)', updated, current.length, year, month, rate);
        return;
    }

    // No rows for the current month yet: carry the most recent earlier month forward
    let template = latestMonthRows_(table.rows, month);
    let templateSource = STATIC_SHEET_PREFIX + year;
    if (template.length === 0) {
        const previousSheet = spreadsheet.getSheetByName(STATIC_SHEET_PREFIX + (year - 1));
        if (previousSheet && previousSheet.getLastRow() > 1) {
            template = latestMonthRows_(readStaticTable_(previousSheet).rows, 13);
            templateSource = STATIC_SHEET_PREFIX + (year - 1);
        }
    }
    if (template.length === 0) {
        Logger.log('Static data: no earlier %s rows to carry forward into %s-%s, add them manually', STATIC_PROVIDER, year, month);
        return;
    }

    const countNote = STATIC_AUTO_NOTE_PREFIX + ' carried forward from ' + templateSource + ' month ' + template[0].month
        + ' on ' + today + '. Update from the claude.ai members export (Seat Tier column) and remove this note.';
    const width = table.header.length;
    const values = template.map(function(t) {
        const row = new Array(width).fill('');
        row[table.col.month] = month;
        row[table.col.provider] = t.provider;
        row[table.col.count] = t.count;
        row[table.col.type] = t.type;
        row[table.col.cost] = t.cost;
        row[table.col.rate] = rate;
        return row;
    });

    const firstRow = sheet.getLastRow() + 1;
    const range = sheet.getRange(firstRow, 1, values.length, width);
    range.setValues(values);
    sheet.getRange(firstRow, table.col.month + 1, values.length, 1).setNumberFormat('0');
    sheet.getRange(firstRow, table.col.count + 1, values.length, 1).setNumberFormat('0').setNote(countNote);
    sheet.getRange(firstRow, table.col.cost + 1, values.length, 1).setNumberFormat('#,##0.00"\u20ac"');
    sheet.getRange(firstRow, table.col.rate + 1, values.length, 1).setNumberFormat('0.0000').setNote(rateNote);

    Logger.log('Static data: added %s %s rows for %s-%s (carried forward from %s month %s, EUR->USD %s)',
        values.length, STATIC_PROVIDER, year, month, templateSource, template[0].month, rate);
}


/** Read a Static-Data sheet into {header, col, rows}. Columns are located by header name. */
function readStaticTable_(sheet) {
    const lastColumn = Math.max(sheet.getLastColumn(), STATIC_COLUMNS.length);
    const header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(SheetUtil.sanitizeColumnName);
    const indexOf = function(name) {
        const i = header.indexOf(SheetUtil.sanitizeColumnName(name));
        if (i < 0) throw new Error('Sheet ' + sheet.getName() + ' is missing column "' + name + '"');
        return i;
    };
    const col = {
        month: indexOf(STATIC_COLUMNS[0]),
        provider: indexOf(STATIC_COLUMNS[1]),
        count: indexOf(STATIC_COLUMNS[2]),
        type: indexOf(STATIC_COLUMNS[3]),
        cost: indexOf(STATIC_COLUMNS[4]),
        rate: indexOf(STATIC_COLUMNS[5]),
    };

    const lastRow = sheet.getLastRow();
    const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues() : [];
    const rows = [];
    data.forEach(function(r, i) {
        const m = +r[col.month];
        if (!(m >= 1 && m <= 12)) return;
        rows.push({
            rowNumber: i + 2,
            month: m,
            provider: String(r[col.provider] || '').trim(),
            count: r[col.count],
            type: r[col.type],
            cost: r[col.cost],
            rate: r[col.rate],
        });
    });

    return {header: header, col: col, rows: rows};
}


/** Whether a static data row is a seat row of STATIC_PROVIDER, carried forward into each new month. */
function isCarriedForward_(row) {
    const type = String(row.type || '').trim().toLowerCase();
    return row.provider === STATIC_PROVIDER
        && !STATIC_INVOICE_TYPES.some(function(t) { return t.toLowerCase() === type; });
}


/** Seat rows of STATIC_PROVIDER for the latest month strictly before the given month (empty if none). */
function latestMonthRows_(rows, beforeMonth) {
    const candidates = rows.filter(function(r) { return isCarriedForward_(r) && r.month < beforeMonth; });
    if (candidates.length === 0) return [];
    const latest = Math.max.apply(null, candidates.map(function(r) { return r.month; }));
    return candidates.filter(function(r) { return r.month === latest; });
}


// --- Section J: Claude Team invoices ---

/** Parse the Claude Team invoices received in a mailbox and update the Static-Data-<year> sheets.
 *
 * Anthropic bills the Team plan seats and every auto recharge of extra usage by invoice (several per day,
 * depending on usage). Each invoice is mailed with a PDF attachment, which is parsed (see parseClaudeInvoice_)
 * and its line items are appended to the Invoices-<year> sheet. That sheet is the record of the invoices seen,
 * keyed by invoice number (receipts and repeated mails of an invoice are skipped) and message ID (mails
 * already handled are not fetched again).
 *
 * The Static-Data-<year> rows of each month with new invoices are then recalculated from all invoice line items
 * of that month (see updateStaticDataFromInvoices_).
 *
 * @param {Spreadsheet} spreadsheet The spreadsheet holding the Invoices-<year> and Static-Data-<year> sheets.
 * @param {string} mailbox The mailbox to search (impersonated using the service account credentials).
 * @param {string} startDate First invoice day to consider (YYYY-MM-DD).
 * @param {string} endDate Day after the last invoice day to consider (YYYY-MM-DD, exclusive).
 * @param {undefined|string} searchExpression Optional Gmail search expression (defaults to the configured one).
 */
function fetchClaudeInvoices_(spreadsheet, mailbox, startDate, endDate, searchExpression) {
    const props = getScriptProperties_();
    const credentials = props.getProperty(SERVICE_ACCOUNT_CREDENTIALS_PROP);
    if (!credentials) {
        throw new Error('No service account credentials at script property ' + SERVICE_ACCOUNT_CREDENTIALS_PROP);
    }
    const gmail = GmailClientV1.withImpersonatingService(JSON.parse(credentials), mailbox);

    // mails may be received a day before/after the date of issue (time zones), the invoice date decides
    const query = (searchExpression || props.getProperty(INVOICE_SEARCH_PROP) || INVOICE_DEFAULT_SEARCH)
        + ' after:' + (toUnixSeconds_(startDate) - 86400) + ' before:' + (toUnixSeconds_(endDate) + 86400);
    Logger.log('  Searching %s for: %s', mailbox, query);
    const messages = gmail.listUserMessages(mailbox, query, null, false);

    const startYear = +startDate.slice(0, 4);
    const endYear = +endDate.slice(0, 4);
    const ledger = {};
    const knownMessages = new Set();
    const knownInvoices = new Set();
    for (let year = startYear - 1; year <= endYear; year++) {
        ledger[year] = readInvoiceLedger_(spreadsheet, year);
        for (const item of ledger[year]) {
            knownMessages.add(item.message_id);
            knownInvoices.add(item.invoice);
        }
    }

    const newItems = [];
    let firstError = null;
    for (const messageRef of messages) {
        if (knownMessages.has(messageRef.id)) {
            continue;
        }
        try {
            for (const invoice of readClaudeInvoices_(gmail, mailbox, messageRef.id)) {
                if (invoice.date < startDate || invoice.date >= endDate || knownInvoices.has(invoice.number)) {
                    continue;
                }
                knownInvoices.add(invoice.number);
                for (const item of invoice.items) {
                    const ratio = invoice.currency === COST_CURRENCY ? 1 : toCostCurrency_(1, invoice.currency, invoice.date);
                    newItems.push({
                        date: invoice.date,
                        invoice: invoice.number,
                        kind: item.kind,
                        tier: item.tier,
                        description: item.description,
                        quantity: item.quantity,
                        unit_price: item.unitPrice,
                        amount: item.amount,
                        currency: invoice.currency,
                        amount_eur: SheetUtil.round(item.amount * ratio, 4),
                        message_id: messageRef.id,
                    });
                }
                Logger.log('  Invoice %s of %s: %s %s (%s line items)', invoice.number, invoice.date,
                    invoice.total, invoice.currency, invoice.items.length);
            }
        } catch (e) {
            Logger.log('  Failed to read invoice mail %s: %s', messageRef.id, e.message);
            firstError = firstError || e;
        }
    }

    Logger.log('  Claude Team invoices: %s mails, %s new line items', messages.length, newItems.length);

    // append to the ledger and recalculate the static data of the months concerned
    const months = new Set();
    newItems.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    for (const item of newItems) {
        const year = +item.date.slice(0, 4);
        (ledger[year] = ledger[year] || []).push(item);
        months.add(item.date.slice(0, 7));
    }
    for (const year of new Set(newItems.map(function(item) { return +item.date.slice(0, 4); }))) {
        appendInvoiceLedger_(spreadsheet, year, newItems.filter(function(item) { return +item.date.slice(0, 4) === year; }));
    }
    for (const yearMonth of Array.from(months).sort()) {
        const year = +yearMonth.slice(0, 4);
        updateStaticDataFromInvoices_(spreadsheet, year, +yearMonth.slice(5, 7), ledger[year]);
    }

    if (firstError) {
        throw firstError;
    }
}


/** Read and parse all invoice PDFs attached to a mail (receipts and other PDFs are ignored). */
function readClaudeInvoices_(gmail, mailbox, messageId) {
    const message = gmail.getUserMessage(mailbox, messageId, 'full');
    const invoices = [];
    const visit = function(part) {
        if (!part) return;
        const body = part.body || {};
        const isPdf = /\.pdf$/i.test(part.filename || '') || part.mimeType === 'application/pdf';
        if (isPdf && (body.attachmentId || body.data)) {
            const bytes = body.attachmentId
                ? gmail.getUserMessageAttachmentBytes(mailbox, messageId, body.attachmentId)
                : Utilities.base64DecodeWebSafe(body.data + '='.repeat((4 - (body.data.length % 4)) % 4));
            const invoice = parseClaudeInvoice_(PdfText.extractLines(bytes));
            if (invoice) {
                invoices.push(invoice);
            }
        }
        (part.parts || []).forEach(visit);
    };
    visit(message.payload);
    return invoices;
}


/** Parse the text lines of an Anthropic (Stripe generated) invoice PDF.
 *
 * Expected layout (cells of the lines as returned by PdfText.extractLines):
 *
 *   Invoice
 *   Invoice number | ABCD1234-0001
 *   Date of issue  | September 30, 2026
 *   ...
 *   Description                          | Qty | Unit price   | Tax | Amount
 *   Auto recharge extra usage, Team plan | 1   | €450.8172472 | 0%  | €450.82
 *   (optional detail lines, like the service period)
 *   Subtotal | €450.82
 *   Total    | €450.82
 *
 * @param {Array<{cells: Array<string>}|Array<string>>} lines The text lines (or their cells) of the PDF.
 *
 * @return {null|Object} The invoice {number, date, currency, total, items: [{description, kind, tier, quantity,
 *         unitPrice, amount}]}, or null if the document is no Anthropic invoice (for example a receipt).
 */
function parseClaudeInvoice_(lines) {
    const rows = lines.map(function(line) { return Array.isArray(line) ? line : line.cells; });
    const isTitle = function(title) {
        return rows.some(function(cells) { return cells.length === 1 && cells[0] === title; });
    };
    if (!isTitle('Invoice') || isTitle('Receipt')
        || !rows.some(function(cells) { return cells.some(function(c) { return /Anthropic/.test(c); }); })) {
        return null;
    }

    const valueOf = function(label) {
        const row = rows.find(function(cells) { return cells[0] === label && cells.length > 1; });
        return row ? row[row.length - 1] : null;
    };

    // Stripe invoice numbers are PREFIX-NNNN, the PDF's fonts lack a character mapping for the hyphen
    const number = (valueOf('Invoice number') || '').replace(/\uFFFD/g, '-');
    const date = parseInvoiceDate_(valueOf('Date of issue') || '');
    const total = parseMoney_(valueOf('Total') || '');
    if (!number || !date || !total) {
        throw new Error('Invoice without number, date of issue or total: ' + JSON.stringify(rows.slice(0, 8)));
    }

    const header = rows.findIndex(function(cells) {
        return cells[0] === 'Description' && cells[cells.length - 1] === 'Amount';
    });
    if (header < 0) {
        throw new Error('Invoice ' + number + ' without line item table');
    }

    const items = [];
    for (let i = header + 1; i < rows.length; i++) {
        const cells = rows[i];
        if (/^(Subtotal|Total|Amount due)/.test(cells[0])) {
            break;
        }
        if (cells.every(function(c) { return /^[\[\uFFFD]\d+[\]\uFFFD]$/.test(c); })) {
            continue;  // footnote marker (e.g. "[1]" above the tax rate)
        }

        let k = cells.length - 1;
        const amount = cells.length > 1 ? parseMoney_(cells[k]) : null;
        if (!amount) {
            // detail line of the previous item (e.g. the service period)
            if (items.length) {
                items[items.length - 1].details.push(cells.join(' '));
            }
            continue;
        }
        k--;
        if (k > 0 && /%/.test(cells[k])) k--;  // tax rate
        const unitPrice = k > 0 ? parseMoney_(cells[k]) : null;
        if (unitPrice) k--;
        const quantity = k > 0 && /^\d+$/.test(cells[k]) ? +cells[k--] : 1;
        const description = cells.slice(0, k + 1).join(' ');
        const kind = classifyClaudeInvoiceItem_(description);
        items.push({
            description: description,
            details: [],
            kind: kind.kind,
            tier: kind.tier,
            quantity: quantity,
            unitPrice: unitPrice ? unitPrice.amount : amount.amount / quantity,
            amount: amount.amount,
        });
    }

    // the items must add up to the subtotal (before tax), otherwise the layout wasn't understood
    const subtotal = parseMoney_(valueOf('Subtotal') || '') || total;
    const sum = items.reduce(function(acc, item) { return acc + item.amount; }, 0);
    if (!items.length || Math.abs(sum - subtotal.amount) > 0.005 + 0.005 * items.length) {
        throw new Error('Invoice ' + number + ': line items sum up to ' + sum + ' instead of the subtotal '
            + subtotal.amount + ': ' + JSON.stringify(rows.slice(header)));
    }

    items.forEach(function(item) {
        if (item.details.length) {
            item.description += ' (' + item.details.join(', ') + ')';
        }
        delete item.details;
    });

    return {number: number, date: date, currency: total.currency, total: total.amount, items: items};
}


/** Classify an invoice line item by its description.
 *
 * @return {{kind: string, tier: string}} kind is one of 'extra_usage', 'seat_adjustment' (prorated seat changes),
 *         'seats' (with tier 'Standard' or 'Premium') or 'other'.
 */
function classifyClaudeInvoiceItem_(description) {
    const tier = /\bpremium\b/i.test(description) ? 'Premium' : /\bstandard\b/i.test(description) ? 'Standard' : '';
    if (/extra usage/i.test(description)) {
        return {kind: 'extra_usage', tier: ''};
    }
    if (/remaining time|unused time|prorat/i.test(description)) {
        return {kind: 'seat_adjustment', tier: tier};
    }
    if (tier && /seat|team/i.test(description)) {
        return {kind: 'seats', tier: tier};
    }
    return {kind: 'other', tier: tier};
}


/** Parse an invoice date like "September 30, 2026" (or "30 September 2026") into YYYY-MM-DD (null if invalid). */
function parseInvoiceDate_(value) {
    const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
    const text = String(value || '').trim();
    let m = /^([A-Za-z]+)\.? (\d{1,2}), (\d{4})$/.exec(text);
    let month, day, year;
    if (m) {
        month = months.indexOf(m[1].slice(0, 3).toLowerCase());
        day = +m[2];
        year = +m[3];
    } else if ((m = /^(\d{1,2}) ([A-Za-z]+)\.? (\d{4})$/.exec(text)) !== null) {
        month = months.indexOf(m[2].slice(0, 3).toLowerCase());
        day = +m[1];
        year = +m[3];
    } else {
        return null;
    }
    const d = new Date(Date.UTC(year, month, day));
    return month >= 0 && d.getUTCDate() === day ? d.toISOString().slice(0, 10) : null;
}


/** Parse an amount like "€450.82", "-€1,234.5" or "$12.00" into {amount, currency} (null if no amount). */
function parseMoney_(value) {
    const symbols = {'€': 'EUR', '$': 'USD', '£': 'GBP'};
    const m = /^([-−–]?)\s*([€$£])\s*([-−–]?)\s*(\d[\d,]*(?:\.\d+)?)$/.exec(String(value || '').trim());
    if (!m) {
        return null;
    }
    const sign = m[1] || m[3] ? -1 : 1;
    return {amount: sign * parseFloat(m[4].replace(/,/g, '')), currency: symbols[m[2]]};
}


/** Read the line items of the Invoices-<year> sheet (empty if the sheet doesn't exist). */
function readInvoiceLedger_(spreadsheet, year) {
    const sheet = spreadsheet.getSheetByName(INVOICE_SHEET_PREFIX + year);
    if (!sheet || sheet.getLastRow() < 2) {
        return [];
    }
    const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    const header = values[0].map(function(h) { return String(h).trim(); });
    const tz = spreadsheet.getSpreadsheetTimeZone();
    return values.slice(1).map(function(r) {
        const item = {};
        header.forEach(function(h, i) { item[h] = r[i]; });
        item.date = item.date instanceof Date ? Utilities.formatDate(item.date, tz, 'yyyy-MM-dd') : String(item.date);
        item.invoice = String(item.invoice);
        item.message_id = String(item.message_id);
        return item;
    });
}


/** Append line items to the Invoices-<year> sheet (created with a header if missing). */
function appendInvoiceLedger_(spreadsheet, year, items) {
    if (!items.length) return;
    const sheet = SheetUtil.ensureSheet(spreadsheet, INVOICE_SHEET_PREFIX + year);
    if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, INVOICE_COLUMNS.length).setValues([INVOICE_COLUMNS]);
    }
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
    const values = items.map(function(item) {
        return header.map(function(h) { return item[h] !== undefined ? item[h] : ''; });
    });
    const firstRow = sheet.getLastRow() + 1;
    // keep dates (YYYY-MM-DD) and invoice numbers as text
    for (const name of ['date', 'invoice']) {
        if (header.indexOf(name) >= 0) {
            sheet.getRange(firstRow, header.indexOf(name) + 1, values.length, 1).setNumberFormat('@');
        }
    }
    sheet.getRange(firstRow, 1, values.length, header.length).setValues(values);
}


/** Calculate the static data rows of one month from its invoice line items.
 *
 * @param {Array<Object>} items Invoice line items (of any month, as in the Invoices-<year> sheet).
 * @param {number} month The month (1-12) to calculate.
 *
 * @return {Array<{type: string, count: number, cost: number, note: string}>} The rows: one per seat tier
 *         (quantity and EUR unit price of the latest invoice of the month billing that tier), plus the sums of
 *         extra usage and of seat adjustments (prorated charges and credits) as user_count 1 rows.
 */
function staticRowsFromInvoices_(items, month) {
    const monthItems = items.filter(function(item) { return +String(item.date).slice(5, 7) === month; });
    const rows = [];

    const latestByTier = {};
    for (const item of monthItems) {
        if (item.kind === 'seats' && item.tier && (!latestByTier[item.tier] || latestByTier[item.tier].date <= item.date)) {
            latestByTier[item.tier] = item;
        }
    }
    for (const tier of Object.keys(latestByTier).sort()) {
        const item = latestByTier[tier];
        const quantity = +item.quantity || 0;
        rows.push({
            type: tier,
            count: quantity,
            cost: quantity ? SheetUtil.round(+item.amount_eur / quantity, 2) : 0,
            note: STATIC_AUTO_NOTE_PREFIX + ' from invoice ' + item.invoice + ' of ' + item.date + '.',
        });
    }

    const sums = [['extra_usage', STATIC_EXTRA_USAGE_TYPE], ['seat_adjustment', STATIC_SEAT_ADJUSTMENT_TYPE]];
    for (const [kind, type] of sums) {
        const matching = monthItems.filter(function(item) { return item.kind === kind; });
        if (!matching.length) continue;
        const invoices = Array.from(new Set(matching.map(function(item) { return item.invoice; })));
        const last = matching[matching.length - 1];
        rows.push({
            type: type,
            count: 1,
            cost: SheetUtil.round(matching.reduce(function(acc, item) { return acc + (+item.amount_eur || 0); }, 0), 2),
            note: STATIC_AUTO_NOTE_PREFIX + ' sum of ' + invoices.length + ' invoice(s), the last one ' + last.invoice
                + ' of ' + last.date + '.',
        });
    }
    return rows;
}


/** Write the static data rows calculated from invoices into Static-Data-<year>.
 *
 * Existing rows of the month (matched by provider and type) are updated, missing ones are appended.
 * Rows whose user_count or cost cell carries a note not starting with STATIC_AUTO_NOTE_PREFIX are manual and left alone.
 */
function updateStaticDataFromInvoices_(spreadsheet, year, month, items) {
    const computed = staticRowsFromInvoices_(items, month);
    if (!computed.length) return;

    const sheet = SheetUtil.ensureSheet(spreadsheet, STATIC_SHEET_PREFIX + year);
    if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, STATIC_COLUMNS.length).setValues([STATIC_COLUMNS]);
    }
    const table = readStaticTable_(sheet);
    const monthRows = table.rows.filter(function(r) { return r.provider === STATIC_PROVIDER && r.month === month; });
    let rate = null;

    for (const c of computed) {
        const existing = monthRows.find(function(r) { return String(r.type).trim().toLowerCase() === c.type.toLowerCase(); });
        if (existing) {
            const countCell = sheet.getRange(existing.rowNumber, table.col.count + 1);
            const costCell = sheet.getRange(existing.rowNumber, table.col.cost + 1);
            const isManual = [countCell.getNote(), costCell.getNote()].some(function(note) {
                return note && note.indexOf(STATIC_AUTO_NOTE_PREFIX) !== 0;
            });
            if (isManual) {
                Logger.log('  Static data %s-%s %s: manual values kept (invoices: %s x %s)', year, month, c.type, c.count, c.cost);
                continue;
            }
            if (+existing.count !== c.count || SheetUtil.round(+existing.cost, 2) !== c.cost || countCell.getNote() !== c.note) {
                countCell.setValue(c.count).setNote(c.note);
                costCell.setValue(c.cost).setNote(c.note);
                Logger.log('  Static data %s-%s %s: %s x %s', year, month, c.type, c.count, c.cost);
            }
        } else {
            if (rate === null) {
                rate = monthRows.length ? monthRows[0].rate : SheetUtil.round(Currency.convertCurrency('EUR', 'USD', 1), 4);
            }
            const row = new Array(table.header.length).fill('');
            row[table.col.month] = month;
            row[table.col.provider] = STATIC_PROVIDER;
            row[table.col.count] = c.count;
            row[table.col.type] = c.type;
            row[table.col.cost] = c.cost;
            row[table.col.rate] = rate;
            const rowNumber = sheet.getLastRow() + 1;
            sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
            sheet.getRange(rowNumber, table.col.month + 1).setNumberFormat('0');
            sheet.getRange(rowNumber, table.col.count + 1).setNumberFormat('0').setNote(c.note);
            sheet.getRange(rowNumber, table.col.cost + 1).setNumberFormat('#,##0.00"€"').setNote(c.note);
            sheet.getRange(rowNumber, table.col.rate + 1).setNumberFormat('0.0000');
            Logger.log('  Static data %s-%s %s: added %s x %s', year, month, c.type, c.count, c.cost);
        }
    }
}
