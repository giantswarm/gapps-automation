/**
 * Fetch AI usage and cost statistics from LLM provider APIs and write to a Google Sheet.
 *
 * Supported sources: anthropic, claude-code, openai
 *
 * All costs are written in EUR (we are billed in EUR). Provider APIs report USD, which is
 * converted using the ECB reference rate of the usage day (see Currency.convertCurrency).
 *
 * Seat based (subscription) costs live in the manually maintained Static-Data-<year> sheet.
 * The script keeps the current month's rows present (carrying seat counts and EUR prices forward
 * from the previous month) and fills in the informational EUR->USD reference rate. Seat tiers of
 * the claude.ai Team plan are not exposed by any API, so user_count must be corrected by hand
 * from the claude.ai members export.
 *
 * Script Properties:
 *   AiCost.anthropicAdminKey  Anthropic admin API key (for anthropic + claude-code sources)
 *   AiCost.openaiAdminKey     OpenAI admin API key (for openai source)
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

const ANTHROPIC_BASE = 'https://api.anthropic.com';
const OPENAI_BASE = 'https://api.openai.com';

/** Anthropic standard-tier pricing per token (USD). Used to estimate per-model costs
 *  from the usage report, since the cost report has no model breakdown.
 *  Cache reads are 0.1x base input (0.025x on Fable 5.1 / Mythos 5.1), 5m cache writes
 *  are 1.25x, 1h cache writes are 2x. The 1M context window is billed at standard rates.
 *  Source: https://platform.claude.com/docs/en/about-claude/pricing (checked 2026-09-04)
 */
const ANTHROPIC_PRICING = {
    'claude-fable-5-1':  { input: 10.00 / 1e6, cache_read: 0.25 / 1e6, cache_5m: 12.50 / 1e6, cache_1h: 20.00 / 1e6, output: 50.00 / 1e6 },
    'claude-mythos-5-1': { input: 10.00 / 1e6, cache_read: 0.25 / 1e6, cache_5m: 12.50 / 1e6, cache_1h: 20.00 / 1e6, output: 50.00 / 1e6 },
    'claude-fable-5':    { input: 10.00 / 1e6, cache_read: 1.00 / 1e6, cache_5m: 12.50 / 1e6, cache_1h: 20.00 / 1e6, output: 50.00 / 1e6 },
    'claude-mythos-5':   { input: 10.00 / 1e6, cache_read: 1.00 / 1e6, cache_5m: 12.50 / 1e6, cache_1h: 20.00 / 1e6, output: 50.00 / 1e6 },
    'claude-opus-5':     { input:  5.00 / 1e6, cache_read: 0.50 / 1e6, cache_5m:  6.25 / 1e6, cache_1h: 10.00 / 1e6, output: 25.00 / 1e6 },
    'claude-opus-4-8':   { input:  5.00 / 1e6, cache_read: 0.50 / 1e6, cache_5m:  6.25 / 1e6, cache_1h: 10.00 / 1e6, output: 25.00 / 1e6 },
    'claude-opus-4-7':   { input:  5.00 / 1e6, cache_read: 0.50 / 1e6, cache_5m:  6.25 / 1e6, cache_1h: 10.00 / 1e6, output: 25.00 / 1e6 },
    'claude-opus-4-6':   { input:  5.00 / 1e6, cache_read: 0.50 / 1e6, cache_5m:  6.25 / 1e6, cache_1h: 10.00 / 1e6, output: 25.00 / 1e6 },
    'claude-opus-4-5':   { input:  5.00 / 1e6, cache_read: 0.50 / 1e6, cache_5m:  6.25 / 1e6, cache_1h: 10.00 / 1e6, output: 25.00 / 1e6 },
    'claude-opus-4-1':   { input: 15.00 / 1e6, cache_read: 1.50 / 1e6, cache_5m: 18.75 / 1e6, cache_1h: 30.00 / 1e6, output: 75.00 / 1e6 },
    'claude-opus-4':     { input: 15.00 / 1e6, cache_read: 1.50 / 1e6, cache_5m: 18.75 / 1e6, cache_1h: 30.00 / 1e6, output: 75.00 / 1e6 },
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

    Logger.log('Done. %s total new rows.', rows.length);

    if (firstError) {
        throw firstError;
    }
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

function fetchAnthropicCosts_(apiKey, startDate, endDate) {
    const url = ANTHROPIC_BASE + '/v1/organizations/cost_report'
        + '?starting_at=' + toIso8601_(startDate)
        + '&ending_at=' + toIso8601_(endDate)
        + '&bucket_width=1d';

    return fetchAllPages_(url, anthropicHeaders_(apiKey),
        function(json) {
            const rows = [];
            for (const bucket of json.data || []) {
                const date = (bucket.starting_at || '').slice(0, 10);
                for (const result of bucket.results || []) {
                    rows.push(makeRow_({
                        date: date, source: 'anthropic', record_type: 'cost',
                        model: result.model || '',
                        cost_eur: toCostCurrency_(parseFloat(result.amount || '0') / 100, result.currency, date),
                        cost_type: result.cost_type || '',
                    }));
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
 * - If no rows for (current month, STATIC_PROVIDER) exist, they are created by copying the rows of the most
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

    const current = table.rows.filter(function(r) { return r.provider === STATIC_PROVIDER && r.month === month; });
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


/** Rows of STATIC_PROVIDER for the latest month strictly before the given month (empty if none). */
function latestMonthRows_(rows, beforeMonth) {
    const candidates = rows.filter(function(r) { return r.provider === STATIC_PROVIDER && r.month < beforeMonth; });
    if (candidates.length === 0) return [];
    const latest = Math.max.apply(null, candidates.map(function(r) { return r.month; }));
    return candidates.filter(function(r) { return r.month === latest; });
}
