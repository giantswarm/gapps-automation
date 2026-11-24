/** Base URL of the Frankfurter currency rates API. */
const CURRENCY_RATES_API_URL = 'https://api.frankfurter.dev';
/** Script cache key prefix for currency rates (suffixed with the rate date or 'latest'). */
const CURRENCY_RATES_CACHE_KEY_PREFIX = 'Currency.rates.';
/** How long to cache currency rates (rates are only updated once per working day). Max. script cache TTL is 6h. */
const CURRENCY_RATES_CACHE_SECONDS = 6 * 60 * 60;


/** Currency conversion using daily reference rates from https://frankfurter.dev. */
class Currency {

    /** Convert an amount between two currencies at the reference rate of a given day.
     *
     * All EUR based rates for the day are fetched once from https://api.frankfurter.dev/v2/rates and cached
     * (per execution in memory and across executions in the script cache) for CURRENCY_RATES_CACHE_SECONDS.
     * Arbitrary pairs are derived as cross rates: rate(base, quote) = rate(EUR, quote) / rate(EUR, base).
     *
     * @param {string} base ISO 4217 code of the source currency (e.g. 'USD').
     * @param {string} quote ISO 4217 code of the target currency (e.g. 'EUR').
     * @param {number} amount The amount in the base currency to convert.
     * @param {string|Date} [date] Day whose reference rate to use ('YYYY-MM-DD' or Date), defaults to the latest rate.
     *                             Days without a published rate (weekends, holidays) use the last available rate.
     * @return {number} The converted amount in the quote currency (not rounded).
     */
    static async convertCurrency(base, quote, amount, date) {
        const rates = await Currency.getRates(date);
        const baseCode = String(base || '').trim().toUpperCase();
        const quoteCode = String(quote || '').trim().toUpperCase();

        const baseRate = rates[baseCode];
        const quoteRate = rates[quoteCode];
        if (!(baseRate > 0)) {
            throw new Error(`Unknown base currency: ${base}`);
        }
        if (!(quoteRate > 0)) {
            throw new Error(`Unknown quote currency: ${quote}`);
        }

        return (+amount || 0) * (quoteRate / baseRate);
    }


    /** Get EUR based currency rates of a day as a map of ISO 4217 code to rate (EUR itself has rate 1.0).
     *
     * Rates are fetched from https://api.frankfurter.dev/v2/rates and cached in memory (per execution)
     * and in the script cache (across executions) for CURRENCY_RATES_CACHE_SECONDS.
     *
     * @param {string|Date} [date] Day whose rates to get ('YYYY-MM-DD' or Date), defaults to the latest rates.
     * @param {boolean} [forceRefresh] If true, ignore cached rates and fetch fresh ones.
     * @return {Object<string, number>} Map of currency code to EUR based rate.
     */
    static async getRates(date, forceRefresh = false) {
        const day = Currency._normalizeDate(date);
        const cacheKey = CURRENCY_RATES_CACHE_KEY_PREFIX + (day || 'latest');
        Currency._rates = Currency._rates || {};

        if (!forceRefresh) {
            if (Currency._rates[cacheKey]) {
                return Currency._rates[cacheKey];
            }

            const cached = Currency._cache()?.get(cacheKey);
            if (cached) {
                try {
                    Currency._rates[cacheKey] = JSON.parse(cached);
                    return Currency._rates[cacheKey];
                } catch (e) {
                    // corrupt cache entry, fall through to re-fetch
                }
            }
        }

        const url = `${CURRENCY_RATES_API_URL}/v2/rates` + (day ? `?date=${day}` : '');
        const response = await UrlFetchApp.fetch(url, {
            method: 'get',
            muteHttpExceptions: true
        });

        const code = response.getResponseCode();
        if (code !== 200) {
            throw new Error(`Failed to fetch currency rates from ${url} (HTTP ${code})`);
        }

        const rows = JSON.parse(response.getContentText());
        if (!Array.isArray(rows) || !rows.length) {
            throw new Error(`Unexpected currency rates response from ${url}`);
        }

        // v2/rates without base parameter is EUR based, but respect whatever base was returned
        const rates = {[rows[0].base || 'EUR']: 1.0};
        for (const row of rows) {
            if (row && row.quote && row.rate > 0) {
                rates[row.quote] = row.rate;
            }
        }

        Currency._rates[cacheKey] = rates;
        try {
            Currency._cache()?.put(cacheKey, JSON.stringify(rates), CURRENCY_RATES_CACHE_SECONDS);
        } catch (e) {
            // caching is best effort (value too large, cache unavailable, ...)
        }

        return rates;
    }


    /** Normalize a date argument to 'YYYY-MM-DD' (UTC) or '' if not given.
     *
     * @param {string|Date|undefined} date The date to normalize.
     * @return {string} The normalized date string or '' for "latest".
     */
    static _normalizeDate(date) {
        if (date === undefined || date === null || date === '') {
            return '';
        }
        if (date instanceof Date) {
            if (isNaN(date.getTime())) {
                throw new Error('Invalid rate date');
            }
            return date.toISOString().slice(0, 10);
        }
        const day = String(date).trim();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(day)) {
            throw new Error(`Invalid rate date (expected YYYY-MM-DD): ${date}`);
        }
        return day;
    }


    /** Get the script cache used for currency rates or undefined if not available. */
    static _cache() {
        try {
            return CacheService.getScriptCache();
        } catch (e) {
            return undefined;
        }
    }
}
