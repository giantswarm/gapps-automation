/**
 * Fetch AI usage and cost statistics from LLM provider APIs and write to a Google Sheet.
 *
 * Supported sources: anthropic, claude-code, openai
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
    'num_requests', 'cost_usd', 'cost_type', 'sessions', 'metadata',
];

const ANTHROPIC_BASE = 'https://api.anthropic.com';
const OPENAI_BASE = 'https://api.openai.com';

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
    // Start one day before today so that all sources (including cost endpoints
    // that only report completed days) cover the same date range for dedup.
    const d = new Date(toIso8601_(defaultStartDate_()));
    d.setUTCDate(d.getUTCDate() - 1);
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

    // Collect dates present in fetched data
    const fetchedDates = new Set();
    for (const row of rows) {
        fetchedDates.add(row.date);
    }

    if (rows.length > 0) {
        rows.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
        appendToSheet_(spreadsheet, rows, fetchedDates);
    }

    Logger.log('Done. %s total new rows.', rows.length);

    if (firstError) {
        throw firstError;
    }
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
        num_requests: 0, cost_usd: 0, cost_type: '', sessions: 0, metadata: '',
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


/** Estimate USD cost from token counts using the OPENAI_PRICING table.
 *  Returns 0 if the model is not in the table.
 *  Model IDs like "gpt-5.1-2025-11-13" are matched by stripping the date suffix.
 */
function estimateOpenaiCost_(model, inputTokens, cachedTokens, outputTokens) {
    const base = (model || '').replace(/-\d{4}-\d{2}-\d{2}$/, '');
    const p = OPENAI_PRICING[base];
    if (!p) return 0;
    return (inputTokens * p.input) + (cachedTokens * p.cached) + (outputTokens * p.output);
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
                    const cacheCreation = (result.cache_creation?.ephemeral_5m_input_tokens || 0)
                        + (result.cache_creation?.ephemeral_1h_input_tokens || 0);
                    rows.push(makeRow_({
                        date: date, source: 'anthropic', record_type: 'usage',
                        model: result.model || '',
                        input_tokens: result.uncached_input_tokens || 0,
                        output_tokens: result.output_tokens || 0,
                        cache_read_tokens: result.cache_read_input_tokens || 0,
                        cache_creation_tokens: cacheCreation,
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
                        cost_usd: parseFloat(result.amount || '0') / 100,
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
                        rows.push(makeRow_({
                            date: day, source: 'claude-code', record_type: 'usage',
                            model: mb.model || '', actor: actor,
                            input_tokens: tokens.input || 0,
                            output_tokens: tokens.output || 0,
                            cache_read_tokens: tokens.cache_read || 0,
                            cache_creation_tokens: tokens.cache_creation || 0,
                            cost_usd: costCents / 100, cost_type: 'tokens',
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
                    rows.push(makeRow_({
                        date: new Date((bucket.start_time || 0) * 1000).toISOString().slice(0, 10),
                        source: 'openai', record_type: 'usage',
                        model: result.model || '',
                        input_tokens: inputTok,
                        output_tokens: outputTok,
                        cache_read_tokens: cachedTok,
                        num_requests: result.num_model_requests || 0,
                        cost_usd: estimateOpenaiCost_(result.model, inputTok, cachedTok, outputTok),
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
                        cost_usd: result.amount?.value || 0,
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
 * Removes any existing rows whose date is in fetchedDates, then appends all new rows.
 */
function appendToSheet_(spreadsheet, newRows, fetchedDates) {
    const sheet = SheetUtil.ensureSheet(spreadsheet, 'Data-' + new Date().getUTCFullYear());

    let lastRow = sheet.getLastRow();

    // Ensure header row exists
    if (lastRow === 0) {
        sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
        lastRow = 1;
    }

    // Remove existing rows for the fetched dates (contiguous ranges, bottom-to-top)
    if (lastRow > 1 && fetchedDates.size > 0) {
        const tz = spreadsheet.getSpreadsheetTimeZone();
        const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        let i = dates.length - 1;
        while (i >= 0) {
            const val = dates[i][0];
            const dateStr = val instanceof Date
                ? Utilities.formatDate(val, tz, 'yyyy-MM-dd') : String(val);
            if (fetchedDates.has(dateStr)) {
                const rangeEnd = i;
                while (i > 0) {
                    const prev = dates[i - 1][0];
                    const prevStr = prev instanceof Date
                        ? Utilities.formatDate(prev, tz, 'yyyy-MM-dd') : String(prev);
                    if (!fetchedDates.has(prevStr)) break;
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
