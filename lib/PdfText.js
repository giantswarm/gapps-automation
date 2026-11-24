/** Huffman code length order of the code length alphabet (RFC 1951, 3.2.7). */
const PDF_TEXT_CLEN_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
/** Base lengths and extra bits of the length codes 257..285 (RFC 1951, 3.2.5). */
const PDF_TEXT_LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const PDF_TEXT_LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
/** Base distances and extra bits of the distance codes 0..29 (RFC 1951, 3.2.5). */
const PDF_TEXT_DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073,
    4097, 6145, 8193, 12289, 16385, 24577];
const PDF_TEXT_DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];

/** PDF white-space and delimiter characters (PDF 32000-1:2008, 7.2.2). */
const PDF_TEXT_WHITESPACE = '\0\t\n\f\r ';
const PDF_TEXT_DELIMITERS = '()<>[]{}/%';


/**
 * Minimal PDF text extractor for machine generated documents (like invoices).
 *
 * Supports uncompressed and FlateDecode compressed streams, simple and Type0 (Identity-H) fonts with ToUnicode CMaps
 * and the text positioning and showing operators. Text is reassembled into visual lines and cells (runs separated
 * by wide gaps, typically table columns) from the glyph positions.
 *
 * Not supported: encrypted documents, object streams (compressed cross-reference sections of PDF 1.5+),
 * other stream filters, text inside form XObjects and vertical writing.
 */
class PdfText {

    /** Extract the text of all pages as visual lines.
     *
     * @param {Array<number>|Uint8Array} bytes The PDF file content (signed bytes, as returned by GAS Blob.getBytes(), are accepted).
     *
     * @return {Array<{page: number, cells: Array<string>, text: string}>} The lines, page by page and top to bottom.
     *         The cells are the runs of text of a line (left to right) separated by gaps wider than about one em,
     *         text is the cells joined by two spaces. Glyphs without known character are returned as U+FFFD.
     */
    static extractLines(bytes) {
        const doc = PdfText.openDocument_(bytes);
        const lines = [];
        PdfText.getPages_(doc).forEach((page, i) => {
            const glyphs = [];
            for (const content of PdfText.getContentRefs_(doc, page.dict.Contents)) {
                const data = PdfText.getStreamData_(doc, content);
                if (data) {
                    PdfText.runContent_(doc, PdfText.toLatin1_(data), page.resources, glyphs);
                }
            }
            for (const cells of PdfText.groupLines_(glyphs)) {
                lines.push({page: i + 1, cells: cells, text: cells.join('  ')});
            }
        });

        return lines;
    }


    /** Decompress a zlib (RFC 1950) or raw deflate (RFC 1951) stream.
     *
     * @param {Uint8Array} data The compressed data.
     *
     * @return {Uint8Array} The decompressed data.
     */
    static inflate(data) {
        // skip the zlib header if present
        let pos = ((data[0] & 0x0f) === 8 && (((data[0] << 8) | data[1]) % 31) === 0) ? 2 : 0;
        let bitBuf = 0;
        let bitCnt = 0;
        let out = new Uint8Array(Math.max(1024, data.length * 4));
        let outLen = 0;

        const bits = n => {
            while (bitCnt < n) {
                if (pos >= data.length) {
                    throw new Error('Unexpected end of deflate data');
                }
                bitBuf |= data[pos++] << bitCnt;
                bitCnt += 8;
            }
            const value = bitBuf & ((1 << n) - 1);
            bitBuf >>>= n;
            bitCnt -= n;
            return value;
        };

        const ensure = n => {
            if (outLen + n > out.length) {
                const grown = new Uint8Array(Math.max(out.length * 2, outLen + n));
                grown.set(out.subarray(0, outLen));
                out = grown;
            }
        };

        // canonical Huffman decoding, as in zlib's puff.c
        const buildHuffman = lengths => {
            const count = new Array(16).fill(0);
            for (const len of lengths) {
                count[len]++;
            }
            count[0] = 0;
            const offs = new Array(16).fill(0);
            for (let len = 1; len < 15; len++) {
                offs[len + 1] = offs[len] + count[len];
            }
            const symbol = new Array(lengths.length);
            lengths.forEach((len, sym) => {
                if (len) {
                    symbol[offs[len]++] = sym;
                }
            });
            return {count: count, symbol: symbol};
        };

        const decode = h => {
            let code = 0, first = 0, index = 0;
            for (let len = 1; len < 16; len++) {
                code |= bits(1);
                const count = h.count[len];
                if (code - count < first) {
                    return h.symbol[index + (code - first)];
                }
                index += count;
                first += count;
                first <<= 1;
                code <<= 1;
            }
            throw new Error('Invalid Huffman code in deflate data');
        };

        let fixed = null;
        let last;
        do {
            last = bits(1);
            const type = bits(2);
            if (type === 0) {
                // stored block
                bitBuf = 0;
                bitCnt = 0;
                if (pos + 4 > data.length) {
                    throw new Error('Unexpected end of deflate data');
                }
                const len = data[pos] | (data[pos + 1] << 8);
                pos += 4;
                if (pos + len > data.length) {
                    throw new Error('Unexpected end of deflate data');
                }
                ensure(len);
                out.set(data.subarray(pos, pos + len), outLen);
                outLen += len;
                pos += len;
                continue;
            }

            let lencode, distcode;
            if (type === 1) {
                if (!fixed) {
                    const lengths = new Array(288);
                    for (let i = 0; i < 288; i++) {
                        lengths[i] = i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8;
                    }
                    fixed = {lencode: buildHuffman(lengths), distcode: buildHuffman(new Array(30).fill(5))};
                }
                lencode = fixed.lencode;
                distcode = fixed.distcode;
            } else if (type === 2) {
                const nlen = bits(5) + 257;
                const ndist = bits(5) + 1;
                const ncode = bits(4) + 4;
                const clens = new Array(19).fill(0);
                for (let i = 0; i < ncode; i++) {
                    clens[PDF_TEXT_CLEN_ORDER[i]] = bits(3);
                }
                const clencode = buildHuffman(clens);
                const lengths = [];
                while (lengths.length < nlen + ndist) {
                    const sym = decode(clencode);
                    if (sym < 16) {
                        lengths.push(sym);
                    } else {
                        let len = 0, repeat;
                        if (sym === 16) {
                            if (!lengths.length) {
                                throw new Error('Invalid code length repeat in deflate data');
                            }
                            len = lengths[lengths.length - 1];
                            repeat = 3 + bits(2);
                        } else if (sym === 17) {
                            repeat = 3 + bits(3);
                        } else {
                            repeat = 11 + bits(7);
                        }
                        while (repeat--) {
                            lengths.push(len);
                        }
                    }
                }
                lencode = buildHuffman(lengths.slice(0, nlen));
                distcode = buildHuffman(lengths.slice(nlen, nlen + ndist));
            } else {
                throw new Error('Invalid deflate block type');
            }

            for (; ;) {
                let sym = decode(lencode);
                if (sym < 256) {
                    ensure(1);
                    out[outLen++] = sym;
                } else if (sym === 256) {
                    break;
                } else {
                    sym -= 257;
                    if (sym >= 29) {
                        throw new Error('Invalid length code in deflate data');
                    }
                    const len = PDF_TEXT_LEN_BASE[sym] + bits(PDF_TEXT_LEN_EXTRA[sym]);
                    const dsym = decode(distcode);
                    if (dsym >= 30) {
                        throw new Error('Invalid distance code in deflate data');
                    }
                    const dist = PDF_TEXT_DIST_BASE[dsym] + bits(PDF_TEXT_DIST_EXTRA[dsym]);
                    if (dist > outLen) {
                        throw new Error('Invalid distance in deflate data');
                    }
                    ensure(len);
                    for (let i = 0; i < len; i++, outLen++) {
                        out[outLen] = out[outLen - dist];
                    }
                }
            }
        } while (!last);

        return out.subarray(0, outLen);
    }


    // --- Document structure ---

    static openDocument_(bytes) {
        const data = bytes instanceof Uint8Array ? bytes : Uint8Array.from(bytes, b => b & 0xff);
        const text = PdfText.toLatin1_(data);
        if (text.indexOf('%PDF-') < 0) {
            throw new Error('Not a PDF document');
        }

        // later definitions of an object (incremental updates) override earlier ones
        const offsets = {};
        const objRe = /(\d+)\s+(\d+)\s+obj\b/g;
        let m;
        while ((m = objRe.exec(text)) !== null) {
            offsets[m[1]] = objRe.lastIndex;
        }

        let trailer = null;
        const trailerAt = text.lastIndexOf('trailer');
        if (trailerAt >= 0) {
            try {
                trailer = PdfText.parseValue_(PdfText.lexer_(text, trailerAt + 7));
            } catch (e) {
                trailer = null;  // fall back to scanning for pages
            }
        }

        return {data: data, text: text, offsets: offsets, cache: {}, trailer: trailer};
    }


    /** Get an indirect object as {value, stream: null|{start, end}}. */
    static getObject_(doc, num) {
        if (doc.cache[num]) {
            return doc.cache[num];
        }
        const offset = doc.offsets[num];
        let obj = {value: null, stream: null};
        if (offset !== undefined) {
            const lexer = PdfText.lexer_(doc.text, offset);
            const value = PdfText.parseValue_(lexer);
            obj = {value: value, stream: null};
            const save = lexer.pos;
            const token = lexer.next();
            if (token && token.type === 'kw' && token.value === 'stream') {
                let start = lexer.pos;
                if (doc.text[start] === '\r') ++start;
                if (doc.text[start] === '\n') ++start;
                // the dictionary must be cached before resolving /Length (it may be an indirect reference)
                doc.cache[num] = obj;
                let end = -1;
                const length = PdfText.resolve_(doc, (value || {}).Length);
                if (typeof length === 'number' && /^\s*endstream/.test(doc.text.substr(start + length, 32))) {
                    end = start + length;
                } else {
                    end = doc.text.indexOf('endstream', start);
                    if (end < 0) {
                        throw new Error('Unterminated stream in object ' + num);
                    }
                    if (doc.text[end - 1] === '\n') --end;
                    if (doc.text[end - 1] === '\r') --end;
                }
                obj.stream = {start: start, end: end};
            } else {
                lexer.pos = save;
            }
        }
        doc.cache[num] = obj;
        return obj;
    }


    static resolve_(doc, value) {
        for (let depth = 0; value && value.ref !== undefined && depth < 32; ++depth) {
            value = PdfText.getObject_(doc, value.ref).value;
        }
        return value;
    }


    /** Get the stream references of a page's /Contents (a stream reference or an array of them). */
    static getContentRefs_(doc, contents) {
        if (contents && contents.ref !== undefined && Array.isArray(PdfText.resolve_(doc, contents))) {
            return PdfText.resolve_(doc, contents);
        }
        return PdfText.asArray_(contents);
    }


    /** Get the decoded data of a referenced stream object, null if it isn't a stream or uses an unsupported filter. */
    static getStreamData_(doc, ref) {
        const obj = ref && ref.ref !== undefined ? PdfText.getObject_(doc, ref.ref) : null;
        if (!obj || !obj.stream) {
            return null;
        }

        let data = doc.data.subarray(obj.stream.start, obj.stream.end);
        for (const filter of PdfText.asArray_(PdfText.resolve_(doc, obj.value.Filter))) {
            const name = (filter || {}).name;
            if (name === 'FlateDecode' || name === 'Fl') {
                data = PdfText.inflate(data);
            } else {
                return null;
            }
        }
        return data;
    }


    /** Get the pages in document order as [{dict, resources}] (resources are inherited from parent nodes). */
    static getPages_(doc) {
        const pages = [];
        const walk = (node, resources, depth) => {
            const dict = PdfText.resolve_(doc, node);
            if (!dict || depth > 64) {
                return;
            }
            const inherited = PdfText.resolve_(doc, dict.Resources) || resources;
            if ((dict.Type || {}).name === 'Pages' || dict.Kids) {
                for (const kid of PdfText.asArray_(PdfText.resolve_(doc, dict.Kids))) {
                    walk(kid, inherited, depth + 1);
                }
            } else {
                pages.push({dict: dict, resources: inherited || {}});
            }
        };

        const root = PdfText.resolve_(doc, (doc.trailer || {}).Root);
        if (root && root.Pages) {
            walk(root.Pages, null, 0);
        }

        if (!pages.length) {
            // no usable trailer (e.g. cross-reference stream): scan all objects in object number order
            for (const num of Object.keys(doc.offsets).sort((a, b) => a - b)) {
                const value = PdfText.getObject_(doc, num).value;
                if (value && (value.Type || {}).name === 'Page') {
                    pages.push({dict: value, resources: PdfText.resolve_(doc, value.Resources) || {}});
                }
            }
        }

        return pages;
    }


    // --- Fonts ---

    static getFont_(doc, resources, name) {
        const fonts = PdfText.resolve_(doc, (resources || {}).Font) || {};
        const ref = fonts[name];
        const key = ref && ref.ref !== undefined ? 'font:' + ref.ref : null;
        if (key && doc.cache[key]) {
            return doc.cache[key];
        }

        const dict = PdfText.resolve_(doc, ref) || {};
        const twoByte = (dict.Subtype || {}).name === 'Type0';
        const font = {twoByte: twoByte, toUnicode: {}, widths: {}, defaultWidth: twoByte ? 1000 : 500};

        if (dict.ToUnicode) {
            const data = PdfText.getStreamData_(doc, dict.ToUnicode);
            if (data) {
                font.toUnicode = PdfText.parseToUnicode_(PdfText.toLatin1_(data));
            }
        }

        if (twoByte) {
            const descendant = PdfText.resolve_(doc, PdfText.asArray_(PdfText.resolve_(doc, dict.DescendantFonts))[0]) || {};
            if (typeof PdfText.resolve_(doc, descendant.DW) === 'number') {
                font.defaultWidth = PdfText.resolve_(doc, descendant.DW);
            }
            const w = PdfText.asArray_(PdfText.resolve_(doc, descendant.W)).map(v => PdfText.resolve_(doc, v));
            for (let i = 0; i < w.length;) {
                const first = w[i];
                if (Array.isArray(w[i + 1])) {
                    w[i + 1].forEach((width, j) => font.widths[first + j] = PdfText.resolve_(doc, width));
                    i += 2;
                } else {
                    for (let code = first; code <= w[i + 1]; code++) {
                        font.widths[code] = w[i + 2];
                    }
                    i += 3;
                }
            }

            // Some producers map glyphs to U+0000 (or not at all) in /ToUnicode, the embedded TrueType font's cmap
            // has the actual characters (CIDs are glyph IDs with /CIDToGIDMap /Identity).
            const cidToGid = PdfText.resolve_(doc, descendant.CIDToGIDMap);
            const descriptor = PdfText.resolve_(doc, descendant.FontDescriptor) || {};
            if ((!cidToGid || cidToGid.name === 'Identity') && descriptor.FontFile2) {
                try {
                    const data = PdfText.getStreamData_(doc, descriptor.FontFile2);
                    const glyphs = data ? PdfText.parseTrueTypeCmap_(data) : {};
                    for (const gid of Object.keys(glyphs)) {
                        const mapped = font.toUnicode[gid];
                        if (mapped === undefined || /^\0*$/.test(mapped)) {
                            font.toUnicode[gid] = glyphs[gid];
                        }
                    }
                } catch (e) {
                    // the font file is only a fallback, ignore broken ones
                }
            }
        } else {
            const firstChar = PdfText.resolve_(doc, dict.FirstChar) || 0;
            PdfText.asArray_(PdfText.resolve_(doc, dict.Widths))
                .forEach((width, j) => font.widths[firstChar + j] = PdfText.resolve_(doc, width));
        }

        if (key) {
            doc.cache[key] = font;
        }
        return font;
    }


    /** Invert the Unicode cmap (format 4 or 12) of a TrueType font file into {glyphId: string}. */
    static parseTrueTypeCmap_(data) {
        const u16 = p => (data[p] << 8) | data[p + 1];
        const u32 = p => ((data[p] << 24) >>> 0) + ((data[p + 1] << 16) | (data[p + 2] << 8) | data[p + 3]);
        const glyphs = {};
        const add = (gid, codePoint) => {
            if (gid && glyphs[gid] === undefined && codePoint > 0) {
                glyphs[gid] = String.fromCodePoint(codePoint);
            }
        };

        let cmap = -1;
        for (let i = 0, n = u16(4); i < n; i++) {
            const rec = 12 + 16 * i;
            if (String.fromCharCode(data[rec], data[rec + 1], data[rec + 2], data[rec + 3]) === 'cmap') {
                cmap = u32(rec + 8);
            }
        }
        if (cmap < 0) {
            return glyphs;
        }

        // prefer full Unicode (3,10) over BMP (3,1 / 0,x) subtables
        let best = -1, bestRank = 0;
        for (let i = 0, n = u16(cmap + 2); i < n; i++) {
            const rec = cmap + 4 + 8 * i;
            const platform = u16(rec), encoding = u16(rec + 2);
            const rank = platform === 3 && encoding === 10 ? 3 : platform === 3 && encoding === 1 ? 2 : platform === 0 ? 1 : 0;
            if (rank > bestRank) {
                best = cmap + u32(rec + 4);
                bestRank = rank;
            }
        }
        if (best < 0) {
            return glyphs;
        }

        const format = u16(best);
        if (format === 4) {
            const segX2 = u16(best + 6);
            const ends = best + 14, starts = ends + segX2 + 2, deltas = starts + segX2, ranges = deltas + segX2;
            for (let i = 0; i < segX2; i += 2) {
                const start = u16(starts + i), end = u16(ends + i), delta = u16(deltas + i), range = u16(ranges + i);
                for (let c = start; c <= end && c !== 0xffff; c++) {
                    if (range === 0) {
                        add((c + delta) & 0xffff, c);
                    } else {
                        const g = u16(ranges + i + range + 2 * (c - start));
                        add(g ? (g + delta) & 0xffff : 0, c);
                    }
                }
            }
        } else if (format === 12) {
            for (let i = 0, n = u32(best + 12); i < n; i++) {
                const group = best + 16 + 12 * i;
                const start = u32(group), end = u32(group + 4), gid = u32(group + 8);
                for (let c = start; c <= end && c - start < 65536; c++) {
                    add(gid + c - start, c);
                }
            }
        }
        return glyphs;
    }


    /** Parse the bfchar and bfrange mappings of a ToUnicode CMap into {code: string}. */
    static parseToUnicode_(text) {
        const map = {};
        const toCode = s => {
            let code = 0;
            for (let i = 0; i < s.length; i++) {
                code = code * 256 + s.charCodeAt(i);
            }
            return code;
        };
        const toUnicode = s => {
            let result = '';
            for (let i = 0; i + 1 < s.length; i += 2) {
                result += String.fromCharCode((s.charCodeAt(i) << 8) | s.charCodeAt(i + 1));
            }
            return result;
        };

        const lexer = PdfText.lexer_(text, 0);
        const operands = [];
        let mode = null;
        for (; ;) {
            const value = PdfText.parseValue_(lexer, true);
            if (value === undefined) {
                break;
            }
            if (value && value.kw !== undefined) {
                if (value.kw === 'beginbfchar' || value.kw === 'beginbfrange') {
                    mode = value.kw;
                } else if (value.kw === 'endbfchar' || value.kw === 'endbfrange') {
                    mode = null;
                }
                operands.length = 0;
                continue;
            }
            if (!mode) {
                continue;
            }
            operands.push(value);
            if (mode === 'beginbfchar' && operands.length === 2) {
                map[toCode(operands[0].str)] = toUnicode(operands[1].str);
                operands.length = 0;
            } else if (mode === 'beginbfrange' && operands.length === 3) {
                const lo = toCode(operands[0].str);
                const hi = toCode(operands[1].str);
                const dst = operands[2];
                for (let code = lo; code <= hi && code - lo < 65536; code++) {
                    if (Array.isArray(dst)) {
                        if (dst[code - lo]) {
                            map[code] = toUnicode(dst[code - lo].str);
                        }
                    } else {
                        const base = toUnicode(dst.str);
                        map[code] = base.slice(0, -1) + String.fromCharCode(base.charCodeAt(base.length - 1) + code - lo);
                    }
                }
                operands.length = 0;
            }
        }
        return map;
    }


    // --- Content streams ---

    /** Interpret a content stream, collecting the shown glyphs as {x, y, x2, size, ch} in default user space. */
    static runContent_(doc, text, resources, glyphs) {
        const mul = (a, b) => [
            a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
            a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
            a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]
        ];
        const translate = (tx, ty, m) => mul([1, 0, 0, 1, tx, ty], m);

        let gs = {ctm: [1, 0, 0, 1, 0, 0], font: null, size: 0, tc: 0, tw: 0, tz: 1, tl: 0, rise: 0};
        const stack = [];
        let tm = [1, 0, 0, 1, 0, 0];
        let tlm = tm;

        const show = str => {
            const font = gs.font;
            if (!font || !str) {
                return;
            }
            const step = font.twoByte ? 2 : 1;
            for (let i = 0; i + step <= str.length; i += step) {
                const code = step === 2 ? (str.charCodeAt(i) << 8) | str.charCodeAt(i + 1) : str.charCodeAt(i);
                let ch = font.toUnicode[code];
                if (ch === undefined) {
                    ch = font.twoByte ? '\uFFFD' : String.fromCharCode(code);
                } else if (/^\0*$/.test(ch)) {
                    ch = '\uFFFD';  // glyph without known character (e.g. a stylistic alternate)
                }
                const width = font.widths[code] !== undefined ? font.widths[code] : font.defaultWidth;
                const tx = (width / 1000 * gs.size + gs.tc + (!font.twoByte && code === 32 ? gs.tw : 0)) * gs.tz;
                const trm = mul([gs.size * gs.tz, 0, 0, gs.size, 0, gs.rise], mul(tm, gs.ctm));
                const end = mul([1, 0, 0, 1, tx, gs.rise], mul(tm, gs.ctm));
                if (ch) {
                    glyphs.push({x: trm[4], y: trm[5], x2: end[4], size: Math.hypot(trm[2], trm[3]), ch: ch});
                }
                tm = translate(tx, 0, tm);
            }
        };

        const lexer = PdfText.lexer_(text, 0);
        const operands = [];
        for (; ;) {
            const value = PdfText.parseValue_(lexer, true);
            if (value === undefined) {
                break;
            }
            if (!value || value.kw === undefined) {
                operands.push(value);
                continue;
            }

            const n = i => +operands[i] || 0;
            switch (value.kw) {
                case 'q':
                    stack.push(gs);
                    gs = {...gs};
                    break;
                case 'Q':
                    gs = stack.pop() || gs;
                    break;
                case 'cm':
                    gs.ctm = mul([n(0), n(1), n(2), n(3), n(4), n(5)], gs.ctm);
                    break;
                case 'BT':
                    tm = tlm = [1, 0, 0, 1, 0, 0];
                    break;
                case 'Tf':
                    gs.font = PdfText.getFont_(doc, resources, (operands[0] || {}).name);
                    gs.size = n(1);
                    break;
                case 'Tc':
                    gs.tc = n(0);
                    break;
                case 'Tw':
                    gs.tw = n(0);
                    break;
                case 'Tz':
                    gs.tz = n(0) / 100;
                    break;
                case 'TL':
                    gs.tl = n(0);
                    break;
                case 'Ts':
                    gs.rise = n(0);
                    break;
                case 'Td':
                    tm = tlm = translate(n(0), n(1), tlm);
                    break;
                case 'TD':
                    gs.tl = -n(1);
                    tm = tlm = translate(n(0), n(1), tlm);
                    break;
                case 'Tm':
                    tm = tlm = [n(0), n(1), n(2), n(3), n(4), n(5)];
                    break;
                case 'T*':
                    tm = tlm = translate(0, -gs.tl, tlm);
                    break;
                case 'Tj':
                    show((operands[0] || {}).str);
                    break;
                case '\'':
                    tm = tlm = translate(0, -gs.tl, tlm);
                    show((operands[0] || {}).str);
                    break;
                case '"':
                    gs.tw = n(0);
                    gs.tc = n(1);
                    tm = tlm = translate(0, -gs.tl, tlm);
                    show((operands[2] || {}).str);
                    break;
                case 'TJ':
                    for (const item of PdfText.asArray_(operands[0])) {
                        if (typeof item === 'number') {
                            tm = translate(-item / 1000 * gs.size * gs.tz, 0, tm);
                        } else {
                            show((item || {}).str);
                        }
                    }
                    break;
                case 'BI': {
                    // skip inline image data
                    const end = text.slice(lexer.pos).search(/\sEI(?=[\s\/\[<(]|$)/);
                    lexer.pos = end < 0 ? text.length : lexer.pos + end + 3;
                    break;
                }
            }
            operands.length = 0;
        }
    }


    /** Group glyphs into lines (top to bottom), each an array of cell strings (left to right). */
    static groupLines_(glyphs) {
        const sorted = glyphs.slice().sort((a, b) => (b.y - a.y) || (a.x - b.x));
        const lines = [];
        for (const g of sorted) {
            const line = lines.length ? lines[lines.length - 1] : null;
            if (line && Math.abs(line.y - g.y) <= 0.3 * Math.max(1, Math.min(line.size, g.size))) {
                line.glyphs.push(g);
            } else {
                lines.push({y: g.y, size: g.size, glyphs: [g]});
            }
        }

        const result = [];
        for (const line of lines) {
            line.glyphs.sort((a, b) => a.x - b.x);
            const cells = [];
            let cell = '';
            let prev = null;
            for (const g of line.glyphs) {
                if (prev) {
                    const gap = g.x - prev.x2;
                    const em = Math.max(1, Math.min(prev.size, g.size));
                    if (gap > 1.2 * em) {
                        cells.push(cell);
                        cell = '';
                    } else if (gap > 0.15 * em && !/\s$/.test(cell) && !/^\s/.test(g.ch)) {
                        cell += ' ';
                    }
                }
                cell += g.ch;
                prev = g;
            }
            cells.push(cell);
            const trimmed = cells.map(c => c.replace(/\s+/g, ' ').trim()).filter(c => c);
            if (trimmed.length) {
                result.push(trimmed);
            }
        }
        return result;
    }


    // --- Lexer and parser (PDF 32000-1:2008, 7.2 and 7.3) ---

    static lexer_(text, pos) {
        const lexer = {text: text, pos: pos};
        lexer.next = () => PdfText.nextToken_(lexer);
        return lexer;
    }


    static nextToken_(lexer) {
        const s = lexer.text;
        let p = lexer.pos;
        for (; ;) {
            while (p < s.length && PDF_TEXT_WHITESPACE.indexOf(s[p]) >= 0) {
                ++p;
            }
            if (s[p] === '%') {
                while (p < s.length && s[p] !== '\n' && s[p] !== '\r') {
                    ++p;
                }
                continue;
            }
            break;
        }
        if (p >= s.length) {
            lexer.pos = p;
            return null;
        }

        const c = s[p];
        let token;
        if (c === '[' || c === ']' || c === '{' || c === '}') {
            token = {type: c, pos: p + 1};
        } else if (c === '<' && s[p + 1] === '<') {
            token = {type: '<<', pos: p + 2};
        } else if (c === '>' && s[p + 1] === '>') {
            token = {type: '>>', pos: p + 2};
        } else if (c === '<') {
            const end = s.indexOf('>', p);
            const hex = s.slice(p + 1, end < 0 ? s.length : end).replace(/[^0-9a-fA-F]/g, '');
            let str = '';
            for (let i = 0; i < hex.length; i += 2) {
                str += String.fromCharCode(parseInt((hex[i] + (hex[i + 1] || '0')), 16));
            }
            token = {type: 'str', value: str, pos: end < 0 ? s.length : end + 1};
        } else if (c === '(') {
            let depth = 1, str = '', i = p + 1;
            for (; i < s.length && depth > 0; i++) {
                const ch = s[i];
                if (ch === '\\') {
                    const e = s[++i];
                    const escapes = {n: '\n', r: '\r', t: '\t', b: '\b', f: '\f'};
                    if (escapes[e] !== undefined) {
                        str += escapes[e];
                    } else if (e >= '0' && e <= '7') {
                        let oct = e;
                        while (oct.length < 3 && s[i + 1] >= '0' && s[i + 1] <= '7') {
                            oct += s[++i];
                        }
                        str += String.fromCharCode(parseInt(oct, 8) & 0xff);
                    } else if (e === '\r') {
                        if (s[i + 1] === '\n') ++i;
                    } else if (e !== '\n') {
                        str += e;
                    }
                } else {
                    if (ch === '(') {
                        ++depth;
                    } else if (ch === ')' && --depth === 0) {
                        break;
                    }
                    str += ch;
                }
            }
            token = {type: 'str', value: str, pos: i + 1};
        } else if (c === '/') {
            let i = p + 1;
            while (i < s.length && PDF_TEXT_WHITESPACE.indexOf(s[i]) < 0 && PDF_TEXT_DELIMITERS.indexOf(s[i]) < 0) {
                ++i;
            }
            const name = s.slice(p + 1, i).replace(/#([0-9a-fA-F]{2})/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
            token = {type: 'name', value: name, pos: i};
        } else {
            let i = p;
            while (i < s.length && PDF_TEXT_WHITESPACE.indexOf(s[i]) < 0 && PDF_TEXT_DELIMITERS.indexOf(s[i]) < 0) {
                ++i;
            }
            if (i === p) {
                i = p + 1;  // stray delimiter like ')' or '>'
            }
            const word = s.slice(p, i);
            token = /^[+-]?(\d+\.?\d*|\.\d+)$/.test(word)
                ? {type: 'num', value: parseFloat(word), pos: i}
                : {type: 'kw', value: word, pos: i};
        }

        lexer.pos = token.pos;
        return token;
    }


    /** Parse the next value. Returns undefined at the end of input.
     *
     * Keywords are returned as {kw} if allowKeywords is set (content streams), otherwise as null.
     * Names are returned as {name}, strings as {str} (bytes as latin1 string), references as {ref}.
     */
    static parseValue_(lexer, allowKeywords) {
        const token = lexer.next();
        if (!token) {
            return undefined;
        }

        switch (token.type) {
            case 'num': {
                if (!allowKeywords && Number.isInteger(token.value)) {
                    // possibly an indirect reference "num gen R"
                    const save = lexer.pos;
                    const gen = lexer.next();
                    const r = gen && gen.type === 'num' ? lexer.next() : null;
                    if (r && r.type === 'kw' && r.value === 'R') {
                        return {ref: token.value};
                    }
                    lexer.pos = save;
                }
                return token.value;
            }
            case 'str':
                return {str: token.value};
            case 'name':
                return {name: token.value};
            case '[': {
                const array = [];
                for (; ;) {
                    const save = lexer.pos;
                    const t = lexer.next();
                    if (!t || t.type === ']') {
                        return array;
                    }
                    lexer.pos = save;
                    array.push(PdfText.parseValue_(lexer, false));
                }
            }
            case '<<': {
                const dict = {};
                for (; ;) {
                    const t = lexer.next();
                    if (!t || t.type === '>>') {
                        return dict;
                    }
                    if (t.type === 'name') {
                        dict[t.value] = PdfText.parseValue_(lexer, false);
                    }
                }
            }
            case 'kw':
                if (token.value === 'true' || token.value === 'false') {
                    return token.value === 'true';
                }
                return allowKeywords ? {kw: token.value} : null;
            default:
                return allowKeywords ? {kw: token.type} : null;
        }
    }


    // --- Helpers ---

    static asArray_(value) {
        return Array.isArray(value) ? value : (value === undefined || value === null ? [] : [value]);
    }


    static toLatin1_(data) {
        let s = '';
        for (let i = 0; i < data.length; i += 8192) {
            s += String.fromCharCode.apply(null, data.subarray(i, Math.min(data.length, i + 8192)));
        }
        return s;
    }
}
