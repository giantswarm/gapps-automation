import assert from 'node:assert/strict';

// mock fetch() for api.frankfurter.dev: EUR based rates per day
const RATES = {
    latest: {date: '2026-09-07', USD: 1.1631, GBP: 0.8613},
    '2026-09-04': {date: '2026-09-04', USD: 1.1622, GBP: 0.8620},
    '2026-09-05': {date: '2026-09-04', USD: 1.1629, GBP: 0.8620}, // weekend, carries last rate
};
const requested = [];
globalThis.fetch = async (url) => {
    requested.push(url);
    const m = /[?&]date=(\d{4}-\d{2}-\d{2})/.exec(url);
    const day = m ? m[1] : 'latest';
    const r = RATES[day];
    if (!r) {
        return {status: 404, text: async () => '{"status":404}', headers: {}};
    }
    const rows = Object.keys(r).filter(k => k !== 'date')
        .map(quote => ({date: r.date, base: 'EUR', quote: quote, rate: r[quote]}));
    return {status: 200, text: async () => JSON.stringify(rows), headers: {}};
};

const {Currency} = (await import('../lib-output/lib.js')).default;

// latest rate
const eur = await Currency.convertCurrency('USD', 'EUR', 116.31);
assert.ok(Math.abs(eur - 100) < 1e-9, `USD->EUR at latest rate failed: ${eur}`);
assert.equal(typeof eur, 'number', 'convertCurrency must return a number');

// per day rate, string and Date arguments hit the same cache entry
const eur1 = await Currency.convertCurrency('USD', 'EUR', 1.1622, '2026-09-04');
assert.ok(Math.abs(eur1 - 1) < 1e-9, `USD->EUR at 2026-09-04 failed: ${eur1}`);
const eur2 = await Currency.convertCurrency('usd', ' eur ', 1.1622, new Date('2026-09-04T10:00:00Z'));
assert.equal(eur1, eur2, 'Date and string arguments must yield the same result');

// cross rate (USD -> GBP) and EUR identity
const gbp = await Currency.convertCurrency('USD', 'GBP', 1.1622, '2026-09-04');
assert.ok(Math.abs(gbp - 0.8620) < 1e-9, `USD->GBP cross rate failed: ${gbp}`);
assert.equal(await Currency.convertCurrency('EUR', 'EUR', 42, '2026-09-04'), 42, 'EUR->EUR must be identity');
assert.equal(await Currency.convertCurrency('USD', 'EUR', 0, '2026-09-05'), 0, 'zero amount must convert to zero');

// caching: one fetch per distinct day (latest, 2026-09-04, 2026-09-05)
assert.equal(requested.length, 3, `expected 3 fetches, got ${requested.length}: ${requested.join(', ')}`);
assert.ok(requested[0].endsWith('/v2/rates'), 'latest rates must be requested without date');
assert.ok(requested[1].endsWith('/v2/rates?date=2026-09-04'), 'day rates must be requested with date');

// errors
await assert.rejects(Currency.convertCurrency('XXX', 'EUR', 1), /Unknown base currency/);
await assert.rejects(Currency.convertCurrency('USD', 'YYY', 1), /Unknown quote currency/);
await assert.rejects(Currency.convertCurrency('USD', 'EUR', 1, 'yesterday'), /Invalid rate date/);
await assert.rejects(Currency.convertCurrency('USD', 'EUR', 1, '1999-01-01'), /HTTP 404/);

console.log('Currency tests passed');
