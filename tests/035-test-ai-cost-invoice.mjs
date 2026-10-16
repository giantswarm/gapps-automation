import assert from 'node:assert/strict';
import fs from 'fs';
import zlib from 'zlib';

const {PdfText, SheetUtil} = (await import('../lib-output/lib.js')).default;

/** Call some function in the code under test which doesn't export anything, without polluting our scope. */
function createTestFunctionWrapper(fileName, functionName) {
    return Function(`"use strict"; ${fs.readFileSync(fileName)}; return ${functionName}(...arguments);`);
}

globalThis.SheetUtil = SheetUtil;

const COST = 'ai-cost/Cost.js';
const parseClaudeInvoice_ = createTestFunctionWrapper(COST, 'parseClaudeInvoice_');
const staticRowsFromInvoices_ = createTestFunctionWrapper(COST, 'staticRowsFromInvoices_');
const parseInvoiceDate_ = createTestFunctionWrapper(COST, 'parseInvoiceDate_');
const parseMoney_ = createTestFunctionWrapper(COST, 'parseMoney_');


// inflate: zlib and raw streams, stored, fixed and dynamic Huffman blocks
const sample = Buffer.from(Array.from({length: 5000}, (_, i) => 'line ' + (i % 37) + ' of text ' + (i * 7919 % 1000) + '\n').join(''));
for (const [name, compressed] of [
    ['zlib default', zlib.deflateSync(sample)],
    ['zlib stored', zlib.deflateSync(sample, {level: 0})],
    ['raw best', zlib.deflateRawSync(sample, {level: 9})],
    ['zlib fixed', zlib.deflateSync(Buffer.from('short'))],
    ['zlib empty', zlib.deflateSync(Buffer.alloc(0))],
]) {
    assert.deepEqual(Buffer.from(PdfText.inflate(new Uint8Array(compressed))), name === 'zlib fixed'
        ? Buffer.from('short') : name === 'zlib empty' ? Buffer.alloc(0) : sample, `inflate ${name} failed`);
}
assert.throws(() => PdfText.inflate(new Uint8Array(zlib.deflateSync(sample)).subarray(0, 100)), /Unexpected end/);


/** Build a single page PDF like the Stripe invoices: one Type0/Identity-H font, glyphs positioned one by one.
 *
 * '-' is mapped to U+0000 in /ToUnicode (like the hyphen in the real invoices).
 *
 * @param {Array<[number, number, string]>} texts The texts to show as [x, y from top, text].
 */
function buildPdf(texts) {
    const chars = [...new Set(texts.map(t => t[2]).join(''))];
    const cid = ch => 0x100 + chars.indexOf(ch);
    const hex = n => n.toString(16).toUpperCase().padStart(4, '0');

    // y axis flipped like in the real invoices
    let content = '.24 0 0 -.24 0 792 cm\nq\n4.1666665 0 0 4.1666665 0 0 cm\n';
    for (const [x, y, text] of texts) {
        content += `BT\n/F4 7.5 Tf\n1 0 0 -1 ${x} ${y} Tm\n`;
        content += [...text].map(ch => `<${hex(cid(ch))}> Tj`).join(' 4.5 0 Td ') + '\nET\n';
    }
    content += 'Q\n';

    const cmap = '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n1 begincodespacerange\n<0000> <FFFF>\n'
        + `endcodespacerange\n${chars.length} beginbfchar\n`
        + chars.map(ch => `<${hex(cid(ch))}> <${ch === '-' ? '0000' : hex(ch.charCodeAt(0))}>`).join('\n')
        + '\nendbfchar\nendcmap\nend\nend\n';

    const objects = [
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Count 1 /Kids [3 0 R] >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F4 4 0 R >> >> /Contents 5 0 R >>',
        '<< /Type /Font /Subtype /Type0 /BaseFont /Test /Encoding /Identity-H /DescendantFonts [6 0 R] /ToUnicode 7 0 R >>',
        {dict: '<< /Length 0 R /Filter /FlateDecode >>', data: zlib.deflateSync(Buffer.from(content, 'latin1'))},
        `<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Test /CIDToGIDMap /Identity /DW 0 /W [256 ${255 + chars.length} 600] >>`,
        {dict: '<< /Length 8 0 R >>', data: Buffer.from(cmap, 'latin1')},
        null,  // length of object 7, set below
    ];
    objects[7] = String(objects[6].data.length);
    objects[4].dict = `<< /Length ${objects[4].data.length} /Filter /FlateDecode >>`;

    const parts = [Buffer.from('%PDF-1.4\n%\xf6\xe4\xfc\xdf\n', 'latin1')];
    objects.forEach((obj, i) => {
        if (typeof obj === 'string') {
            parts.push(Buffer.from(`${i + 1} 0 obj\n${obj}\nendobj\n`, 'latin1'));
        } else {
            parts.push(Buffer.from(`${i + 1} 0 obj\n${obj.dict}\nstream\n`, 'latin1'), obj.data, Buffer.from('\nendstream\nendobj\n'));
        }
    });
    parts.push(Buffer.from(`trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\n%%EOF\n`));
    return Buffer.concat(parts);
}

// an invoice laid out like the real ones (glyph advance 4.5 per char at 7.5pt, cells far apart)
const invoicePdf = buildPdf([
    [30, 60, 'Invoice'],
    [30, 80, 'Invoice number'], [130, 80, 'ABCD1234-0007'],
    [30, 90, 'Date of issue'], [130, 90, 'October 2, 2026'],
    [30, 120, 'Anthropic, PBC'], [300, 120, 'Bill to'],
    [30, 200, 'Description'], [350, 200, 'Qty'], [420, 200, 'Unit price'], [500, 200, 'Tax'], [550, 200, 'Amount'],
    [500, 208, '[1]'],
    [30, 215, 'Team plan Standard seat'], [350, 215, '34'], [420, 215, '€21.04'], [500, 215, '0%'], [550, 215, '€715.36'],
    [30, 223, 'Oct 2 2026 to Nov 2 2026'],
    [30, 235, 'Remaining time on 1 Premium seat'], [350, 235, '1'], [420, 235, '€50.00'], [500, 235, '0%'], [550, 235, '€50.00'],
    [30, 255, 'Auto recharge extra usage, Team plan'], [350, 255, '1'], [420, 255, '€450.8172472'], [500, 255, '0%'], [550, 255, '€450.82'],
    [300, 280, 'Subtotal'], [550, 280, '€1,216.18'],
    [300, 290, 'Total'], [550, 290, '€1,216.18'],
    [300, 300, 'Amount due'], [550, 300, '€1,216.18'],
]);

const lines = PdfText.extractLines(invoicePdf);
assert.deepEqual(lines[0], {page: 1, cells: ['Invoice'], text: 'Invoice'});
assert.deepEqual(lines[1].cells, ['Invoice number', 'ABCD1234�0007'], 'glyphs mapped to U+0000 must become U+FFFD');
assert.deepEqual(lines.find(l => l.cells[0] === 'Description').cells, ['Description', 'Qty', 'Unit price', 'Tax', 'Amount']);
assert.deepEqual(lines.find(l => l.cells[0].startsWith('Auto')).cells,
    ['Auto recharge extra usage, Team plan', '1', '€450.8172472', '0%', '€450.82'], 'words must be joined, columns split');

const invoice = parseClaudeInvoice_(lines);
assert.deepEqual(invoice, {
    number: 'ABCD1234-0007', date: '2026-10-02', currency: 'EUR', total: 1216.18, items: [
        {description: 'Team plan Standard seat (Oct 2 2026 to Nov 2 2026)', kind: 'seats', tier: 'Standard',
            quantity: 34, unitPrice: 21.04, amount: 715.36},
        {description: 'Remaining time on 1 Premium seat', kind: 'seat_adjustment', tier: 'Premium',
            quantity: 1, unitPrice: 50, amount: 50},
        {description: 'Auto recharge extra usage, Team plan', kind: 'extra_usage', tier: '',
            quantity: 1, unitPrice: 450.8172472, amount: 450.82},
    ]
});

// receipts (attached to the same mails) and foreign documents are ignored
assert.equal(parseClaudeInvoice_([['Receipt'], ['Invoice'], ['Anthropic, PBC']]), null);
assert.equal(parseClaudeInvoice_([['Invoice'], ['Other Corp']]), null);

// line items not adding up to the subtotal indicate a changed layout
const broken = lines.map(l => l.cells).filter(c => !c[0].startsWith('Remaining'));
assert.throws(() => parseClaudeInvoice_(broken), /line items sum up to/);

assert.equal(parseInvoiceDate_('September 30, 2026'), '2026-09-30');
assert.equal(parseInvoiceDate_('1 Oct 2026'), '2026-10-01');
assert.equal(parseInvoiceDate_('February 30, 2026'), null);
assert.deepEqual(parseMoney_('-€1,234.5'), {amount: -1234.5, currency: 'EUR'});
assert.deepEqual(parseMoney_('$12.00'), {amount: 12, currency: 'USD'});
assert.equal(parseMoney_('0%'), null);


// static data rows of a month from invoice line items
const items = [
    {date: '2026-09-30', invoice: 'A-1', kind: 'extra_usage', tier: '', quantity: 1, amount_eur: 450.82},
    {date: '2026-10-01', invoice: 'A-2', kind: 'seats', tier: 'Standard', quantity: 34, amount_eur: 715.36},
    {date: '2026-10-01', invoice: 'A-2', kind: 'seats', tier: 'Premium', quantity: 16, amount_eur: 1683.68},
    {date: '2026-10-03', invoice: 'A-3', kind: 'extra_usage', tier: '', quantity: 1, amount_eur: 100.10},
    {date: '2026-10-03', invoice: 'A-4', kind: 'extra_usage', tier: '', quantity: 1, amount_eur: 200.20},
    {date: '2026-10-05', invoice: 'A-5', kind: 'seat_adjustment', tier: 'Premium', quantity: 1, amount_eur: 50},
    {date: '2026-10-09', invoice: 'A-6', kind: 'seats', tier: 'Premium', quantity: 17, amount_eur: 1788.91},
];
assert.deepEqual(staticRowsFromInvoices_(items, 10), [
    {type: 'Premium', count: 17, cost: 105.23, note: 'Auto: from invoice A-6 of 2026-10-09.'},
    {type: 'Standard', count: 34, cost: 21.04, note: 'Auto: from invoice A-2 of 2026-10-01.'},
    {type: 'Extra Usage', count: 1, cost: 300.3, note: 'Auto: sum of 2 invoice(s), the last one A-4 of 2026-10-03.'},
    {type: 'Seat Adjustment', count: 1, cost: 50, note: 'Auto: sum of 1 invoice(s), the last one A-5 of 2026-10-05.'},
]);
assert.deepEqual(staticRowsFromInvoices_(items, 8), []);

// the backfill mailbox must only be set in the deployed script, never in the (public) repository
assert.match(fs.readFileSync(COST, 'utf8'), /^const INVOICE_BACKFILL_MAILBOX = 'mailbox@example\.com';$/m,
    'INVOICE_BACKFILL_MAILBOX must stay the placeholder in the repository');

console.log('AI cost invoice tests passed');
