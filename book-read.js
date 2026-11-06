// Excel のブック（xlsx・xlsm・xls）を読む部品（画面の部品を持たない）。画面（Edge・Chrome・Safari）と試験（Node）で同じ物を使う。
//   readBook(bytes) で、ブックの中身を「読んだブックの形」にして返す。形は tools/dump_book.py の先頭の説明と同じ
//   （行と列は 1 から・cells のキーは 'r,c'・v/f/e/font/fill/runs/x・シートの x0・載せるセルは値か塗りか書式のあるセル）。
//   形の見分けは、ファイルの名前でなく中身の先頭で行う（PK…＝xlsx/xlsm・D0 CF 11 E0…＝xls）。
//   外の部品・通信は使わない。xlsx の圧縮は、ブラウザと Node にある DecompressionStream('deflate-raw') で解く。
//   信頼境界: ここへ来るのは、人が選んだファイルの中身（社内の書類だが、壊れている・別の形のこともある）。
//     式は計算しない・中の文字を実行しない。読めない形・壊れた物は、日本語の文の Error で止める。
//     読めたが一部を戻せなかったとき（式の字など）は止めずに、book.notes（JSON には出ない）に件数と理由を書く。
(function (root) {
  'use strict';

  // ---------------------------------------------------------------- 止めるときの文
  const MSG_NOT_EXCEL = 'Excel のファイルとして読めませんでした（xlsx・xlsm・xls のどれでもありません）。';
  const MSG_BROKEN = 'Excel のファイルとして読めませんでした（ファイルが壊れているか、途中で切れています）。';
  const MSG_OLD = 'Excel のファイルとして読めませんでした（Excel 95 までの古い形式です）。Excel で開き、「Excel ブック（*.xlsx）」で保存し直してください。';
  const MSG_PASSWORD = 'パスワードの付いたブックは読めません。Excel でパスワードを外して保存し直してください。';
  const MSG_XLSB = 'Excel バイナリ ブック（*.xlsb）は読めません。Excel で開き、「Excel ブック（*.xlsx）」で保存し直してください。';
  const MSG_NO_INFLATE = 'このブラウザでは xlsx を開けません（圧縮を解く仕組みがありません）。Edge・Chrome・Safari の新しい版で開いてください。';
  function fail(msg, cause) {
    const e = new Error(msg);
    e.bookError = true;
    if (cause) e.cause = cause;
    return e;
  }

  // ---------------------------------------------------------------- バイトを読む小道具（どれもリトルエンディアン）
  const u16 = (b, p) => b[p] | (b[p + 1] << 8);
  const i16 = (b, p) => (u16(b, p) << 16) >> 16;
  const u32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16)) + b[p + 3] * 0x1000000;
  function f64(b, p) { return new DataView(b.buffer, b.byteOffset + p, 8).getFloat64(0, true); }
  let decUtf16 = null, decUtf8 = null;
  function utf16(b, p, n) {
    if (n <= 0) return '';
    if (!decUtf16) decUtf16 = new TextDecoder('utf-16le');
    return decUtf16.decode(b.subarray(p, p + n * 2));
  }
  function utf8(b) {
    if (!decUtf8) decUtf8 = new TextDecoder('utf-8');
    return decUtf8.decode(b);
  }
  // xls の「詰めた字」（上の 1 バイトが 0 の字）。0x80〜0x9F も、そのままの字にする（windows-1252 に読み替えない）
  function latin(b, p, n) {
    let s = '';
    for (let i = 0; i < n; i += 4096) s += String.fromCharCode.apply(null, b.subarray(p + i, p + Math.min(n, i + 4096)));
    return s;
  }
  function toU8(bytes) {
    if (bytes instanceof Uint8Array) return bytes;
    if (bytes instanceof ArrayBuffer) return new Uint8Array(bytes);
    if (bytes && bytes.buffer instanceof ArrayBuffer) return new Uint8Array(bytes.buffer, bytes.byteOffset || 0, bytes.byteLength);
    throw fail(MSG_NOT_EXCEL);
  }
  const hex2 = (x) => (x < 16 ? '0' : '') + x.toString(16).toUpperCase();

  // ---------------------------------------------------------------- 色（tools/souzai_style.py の Palette と同じ結果）
  // 色番号の既定（openpyxl の COLOR_INDEX と同じ。0〜7 は固定の8色・8〜63 はブックが変えられる56色）
  const COLOR_INDEX = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
  // テーマの色の番号の並び（Excel の番号: 0＝lt1・1＝dk1・2＝lt2・3＝dk2…）
  const THEME_ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  // xls にテーマが入っていないとき（既定のテーマ「Office 2007 - 2010」）
  const THEME_2007 = { 0: 'FFFFFF', 1: '000000', 2: 'EEECE1', 3: '1F497D', 4: '4F81BD', 5: 'C0504D', 6: '9BBB59', 7: '8064A2', 8: '4BACC6', 9: 'F79646', 10: '0000FF', 11: '800080' };

  /** theme1.xml の文字 → {番号: 'RRGGBB'}（souzai_style.Palette と同じ読み方） */
  function themeOfXml(s) {
    const out = {};
    const m = /<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(s || '');
    if (!m) return out;
    THEME_ORDER.forEach((tag, i) => {
      const mm = new RegExp('<a:' + tag + '>([\\s\\S]*?)</a:' + tag + '>').exec(m[1]);
      if (!mm) return;
      const hexes = [];
      const re = /(?:lastClr|val)="([0-9A-Fa-f]{6})"/g;
      let h;
      while ((h = re.exec(mm[1]))) hexes.push(h[1]);
      if (hexes.length) out[i] = (mm[1].indexOf('lastClr') >= 0 ? hexes[hexes.length - 1] : hexes[0]).toUpperCase();
    });
    return out;
  }

  // Python の round（ちょうど半分は偶数へ）・%（割る数の符号にそろえる）と同じ計算
  function pyRound(x) {
    const f = Math.floor(x), d = x - f;
    if (d > 0.5) return f + 1;
    if (d < 0.5) return f;
    return f % 2 === 0 ? f : f + 1;
  }
  function pyMod(a, b) {
    const m = a % b;
    return m !== 0 && (m < 0) !== (b < 0) ? m + b : m;
  }
  // Python の colorsys と同じ順の計算（明るさを変えるため）
  function rgbToHls(r, g, b) {
    const maxc = Math.max(r, g, b), minc = Math.min(r, g, b);
    const sumc = maxc + minc, rangec = maxc - minc, l = sumc / 2.0;
    if (minc === maxc) return [0.0, l, 0.0];
    const s = l <= 0.5 ? rangec / sumc : rangec / (2.0 - maxc - minc);
    const rc = (maxc - r) / rangec, gc = (maxc - g) / rangec, bc = (maxc - b) / rangec;
    let h;
    if (r === maxc) h = bc - gc;
    else if (g === maxc) h = 2.0 + rc - bc;
    else h = 4.0 + gc - rc;
    h = pyMod(h / 6.0, 1.0);
    return [h, l, s];
  }
  function hueV(m1, m2, hue) {
    hue = pyMod(hue, 1.0);
    if (hue < 1.0 / 6.0) return m1 + (m2 - m1) * hue * 6.0;
    if (hue < 0.5) return m2;
    if (hue < 2.0 / 3.0) return m1 + (m2 - m1) * (2.0 / 3.0 - hue) * 6.0;
    return m1;
  }
  function hlsToRgb(h, l, s) {
    if (s === 0.0) return [l, l, l];
    const m2 = l <= 0.5 ? l * (1.0 + s) : l + s - (l * s);
    const m1 = 2.0 * l - m2;
    return [hueV(m1, m2, h + 1.0 / 3.0), hueV(m1, m2, h), hueV(m1, m2, h - 1.0 / 3.0)];
  }
  /** テーマの色に明るさ（tint・-1〜1）を当てる（souzai_style._tint と同じ） */
  function tintHex(hex6, tint) {
    const r = parseInt(hex6.slice(0, 2), 16) / 255.0, g = parseInt(hex6.slice(2, 4), 16) / 255.0, b = parseInt(hex6.slice(4, 6), 16) / 255.0;
    const hls = rgbToHls(r, g, b);
    let l = hls[1];
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    const rgb = hlsToRgb(hls[0], Math.max(0.0, Math.min(1.0, l)), hls[2]);
    return hex2(pyRound(rgb[0] * 255)) + hex2(pyRound(rgb[1] * 255)) + hex2(pyRound(rgb[2] * 255));
  }
  /** 色の決まり（テーマ・色番号）→ 色を 'RRGGBB' にする関数（自動・無しは ''） */
  function makePalette(theme, indexed) {
    const cache = new Map();
    return function rgb(c) {
      if (!c) return '';
      if (c.indexed !== undefined) {
        const i = c.indexed;
        if (!(i >= 0) || i >= 64) return ''; // 64＝システムの字の色・65＝背景（自動と同じ扱い）
        const v = i < indexed.length ? indexed[i] : '';
        return v ? v.slice(-6).toUpperCase() : '';
      }
      if (c.theme !== undefined) {
        const base = theme[c.theme];
        if (!base) return '';
        const t = c.tint || 0;
        const key = base + ':' + t;
        let v = cache.get(key);
        if (v === undefined) { v = tintHex(base, t); cache.set(key, v); }
        return v;
      }
      if (c.auto) return '';
      const v = typeof c.rgb === 'string' ? c.rgb : '';
      return v ? v.slice(-6).toUpperCase() : '';
    };
  }
  /** 字の飾り → {c, s, b, i, u, f}（dump_book.py の font_of・make_seed.py の runs_of と同じ。自動の色は書かない） */
  function fontDict(fs, rgb) {
    const d = {};
    if (!fs) return d;
    if (fs.color) { const c = rgb(fs.color); if (c) d.c = c; }
    if (fs.sz !== null && fs.sz !== undefined && !Number.isNaN(fs.sz)) d.s = fs.sz;
    if (fs.b) d.b = 1;
    if (fs.i) d.i = 1;
    if (fs.u) d.u = 1;
    if (fs.name) d.f = fs.name;
    return d;
  }
  const isEmptyObj = (o) => { for (const k in o) return false; return true; };

  // ---------------------------------------------------------------- 表示形式（決まった形・日付かどうか・openpyxl と同じ）
  // Excel の決まった表示形式（openpyxl の BUILTIN_FORMATS と同じ。ここに無い番号〈国ごとの形〉は 'General' として扱う）
  const BUILTIN_FORMATS = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 5: '"$"#,##0_);("$"#,##0)', 6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)', 8: '"$"#,##0.00_);[Red]("$"#,##0.00)', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm',
    37: '#,##0_);(#,##0)', 38: '#,##0_);[Red](#,##0)', 39: '#,##0.00_);(#,##0.00)', 40: '#,##0.00_);[Red](#,##0.00)',
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_)_("$"* \\(#,##0.00\\)_("$"* "-"??_)_(@_)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@' };
  const BUILTIN_TEXTS = new Set(Object.values(BUILTIN_FORMATS));
  // 罫線の種類（xls の番号 → xlsx の名前）・横と縦の位置
  const BORDER_NAMES = ['', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot', '', ''];
  const H_NAMES = ['', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
  const V_NAMES = ['top', 'center', '', 'justify', 'distributed', '', '', ''];
  /** 書式の残り x（罫線・位置・折り返し・縮小・表示形式。dump_book.py の rest_of と同じ形） */
  function restOf(bl, br, bt, bb, h, v, wrap, shrink, n) {
    const x = { bl, br, bt, bb, h, v, n: n || 'General' };
    if (wrap) x.w = 1;
    if (shrink) x.k = 1;
    return x;
  }
  const DEFAULT_REST = Object.freeze(restOf('', '', '', '', '', '', false, false, 'General'));
  function isDateFormat(fmt) {
    if (fmt === null || fmt === undefined) return false;
    let s = String(fmt).split(';')[0];
    s = s.replace(/".*?"|\[(?!hh?\]|mm?\]|ss?\])[^\]]*\]/g, '');
    for (let i = 0; i < s.length; i++) {
      if ('dmhysDMHYS'.indexOf(s[i]) >= 0 && (i === 0 || (s[i - 1] !== '_' && s[i - 1] !== '\\'))) return true;
    }
    return false;
  }
  function isTimedeltaFormat(fmt) {
    if (fmt === null || fmt === undefined) return false;
    return /\[hh?\](:mm(:ss(\.0*)?)?)?|\[mm?\](:ss(\.0*)?)?|\[ss?\](\.0*)?/i.test(String(fmt).split(';')[0]);
  }
  const pad = (n, w) => { let s = String(n); while (s.length < w) s = '0' + s; return s; };
  /** Excel の日付の数 → ISO の文字（openpyxl の from_excel と Python の isoformat と同じ） */
  function excelDateIso(value, date1904) {
    const day = Math.floor(value), fraction = value - day;
    const ms = pyRound(fraction * 86400 * 1000);
    const usec = (ms % 1000) * 1000;
    const tail = usec ? '.' + pad(usec, 6) : '';
    if (value >= 0 && value < 1 && ms < 86400000) {
      const secs = Math.floor(ms / 1000);
      return pad(Math.floor(secs / 3600), 2) + ':' + pad(Math.floor(secs / 60) % 60, 2) + ':' + pad(secs % 60, 2) + tail;
    }
    let d = day;
    if (!date1904 && value > 0 && value < 60) d += 1; // Excel の「1900年2月29日」の決まり
    const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    const t = new Date(epoch + d * 86400000 + ms);
    const y = t.getUTCFullYear();
    if (!(y >= 1 && y <= 9999)) return '#VALUE!';
    return pad(y, 4) + '-' + pad(t.getUTCMonth() + 1, 2) + '-' + pad(t.getUTCDate(), 2) + 'T' +
      pad(t.getUTCHours(), 2) + ':' + pad(t.getUTCMinutes(), 2) + ':' + pad(t.getUTCSeconds(), 2) + tail;
  }
  const ERROR_TEXTS = new Set(['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!']);

  // ---------------------------------------------------------------- 小さな XML の読み手（Node に DOMParser が無いため）
  //   要素 = { n: 名前（接頭辞を外す）, a: 属性, c: 子の要素, t: 直下の字 }
  const ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
  function decodeEnt(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#[xX][0-9a-fA-F]+|#[0-9]+|[A-Za-z]+);/g, (m, e) => {
      if (e[0] === '#') {
        const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return cp >= 0 && cp <= 0x10FFFF ? String.fromCodePoint(cp) : m;
      }
      return ENTITIES[e] !== undefined ? ENTITIES[e] : m;
    });
  }
  const isSpace = (ch) => ch === 32 || ch === 10 || ch === 9 || ch === 13;
  function parseXml(src) {
    if (src.charCodeAt(0) === 0xFEFF) src = src.slice(1);
    if (src.indexOf('\r') >= 0) src = src.replace(/\r\n?/g, '\n'); // XML の行の終わりのそろえ（Excel・openpyxl と同じ）
    const top = { n: '', a: {}, c: [], t: '' };
    const stack = [top];
    let cur = top, i = 0;
    const n = src.length;
    while (i < n) {
      const lt = src.indexOf('<', i);
      if (lt < 0) { cur.t += decodeEnt(src.slice(i)); break; }
      if (lt > i) cur.t += decodeEnt(src.slice(i, lt));
      const c1 = src.charCodeAt(lt + 1);
      if (c1 === 47) { // </名前>
        const gt = src.indexOf('>', lt);
        if (gt < 0) throw fail(MSG_BROKEN);
        if (stack.length > 1) { stack.pop(); cur = stack[stack.length - 1]; }
        i = gt + 1;
        continue;
      }
      if (c1 === 63) { const e = src.indexOf('?>', lt); i = e < 0 ? n : e + 2; continue; } // <?xml …?>
      if (c1 === 33) { // <!-- -->・<![CDATA[ ]]>・<!DOCTYPE>
        if (src.startsWith('<!--', lt)) { const e = src.indexOf('-->', lt + 4); i = e < 0 ? n : e + 3; continue; }
        if (src.startsWith('<![CDATA[', lt)) {
          const e = src.indexOf(']]>', lt + 9), end = e < 0 ? n : e;
          cur.t += src.slice(lt + 9, end);
          i = end + 3;
          continue;
        }
        let j = lt + 2, depth = 0;
        for (; j < n; j++) {
          const ch = src.charCodeAt(j);
          if (ch === 91) depth++;
          else if (ch === 93) depth--;
          else if (ch === 62 && depth <= 0) break;
        }
        i = j + 1;
        continue;
      }
      let j = lt + 1;
      while (j < n) { const ch = src.charCodeAt(j); if (isSpace(ch) || ch === 47 || ch === 62) break; j++; }
      let name = src.slice(lt + 1, j);
      const colon = name.indexOf(':');
      if (colon >= 0) name = name.slice(colon + 1);
      const el = { n: name, a: {}, c: [], t: '' };
      let selfClose = false;
      for (;;) {
        while (j < n && isSpace(src.charCodeAt(j))) j++;
        if (j >= n) throw fail(MSG_BROKEN);
        const ch = src.charCodeAt(j);
        if (ch === 62) { j++; break; }
        if (ch === 47) { selfClose = true; const gt = src.indexOf('>', j); if (gt < 0) throw fail(MSG_BROKEN); j = gt + 1; break; }
        const eq = src.indexOf('=', j);
        if (eq < 0) throw fail(MSG_BROKEN);
        const an = src.slice(j, eq).trim();
        let q = eq + 1;
        while (q < n && isSpace(src.charCodeAt(q))) q++;
        const qc = src[q];
        if (qc !== '"' && qc !== "'") throw fail(MSG_BROKEN);
        const qe = src.indexOf(qc, q + 1);
        if (qe < 0) throw fail(MSG_BROKEN);
        let av = src.slice(q + 1, qe);
        if (av.indexOf('\n') >= 0 || av.indexOf('\t') >= 0) av = av.replace(/[\n\t]/g, ' ');
        el.a[an] = decodeEnt(av);
        j = qe + 1;
      }
      cur.c.push(el);
      if (!selfClose) { stack.push(el); cur = el; }
      i = j;
    }
    return top;
  }
  const kid = (el, name) => { if (el) for (const k of el.c) if (k.n === name) return k; return null; };
  const kids = (el, name) => (el ? el.c.filter((k) => k.n === name) : []);
  const rootEl = (doc) => { for (const k of doc.c) return k; return { n: '', a: {}, c: [], t: '' }; };
  // 真偽の属性（'1'・'true' が真。無ければ既定）
  function boolAttr(v, dflt) {
    if (v === undefined || v === null) return dflt;
    return !(v === '0' || v === 'false' || v === 'False' || v === 'f');
  }

  // ---------------------------------------------------------------- zip（xlsx の入れ物）
  async function inflateRaw(data) {
    if (typeof DecompressionStream === 'undefined') throw fail(MSG_NO_INFLATE);
    let ds;
    try { ds = new DecompressionStream('deflate-raw'); } catch (e) { throw fail(MSG_NO_INFLATE, e); }
    const w = ds.writable.getWriter();
    w.write(data).catch(() => {}); // 失敗は、下の読む側で受ける
    w.close().catch(() => {});
    const r = ds.readable.getReader();
    const parts = [];
    let total = 0;
    try {
      for (;;) {
        const x = await r.read();
        if (x.done) break;
        parts.push(x.value);
        total += x.value.length;
      }
    } catch (e) { throw fail(MSG_BROKEN, e); }
    if (parts.length === 1) return parts[0];
    const out = new Uint8Array(total);
    let p = 0;
    for (const x of parts) { out.set(x, p); p += x.length; }
    return out;
  }
  function openZip(u8) {
    const n = u8.length;
    let eocd = -1;
    for (let p = n - 22; p >= Math.max(0, n - 65557); p--) {
      if (u8[p] === 0x50 && u8[p + 1] === 0x4B && u8[p + 2] === 0x05 && u8[p + 3] === 0x06) { eocd = p; break; }
    }
    if (eocd < 0) throw fail(MSG_BROKEN);
    let count = u16(u8, eocd + 10), cdOff = u32(u8, eocd + 16);
    if ((count === 0xFFFF || cdOff === 0xFFFFFFFF) && eocd >= 20 && u32(u8, eocd - 20) === 0x07064b50) { // zip64
      const z = u32(u8, eocd - 12) + u32(u8, eocd - 8) * 0x100000000;
      if (u32(u8, z) === 0x06064b50) { count = u32(u8, z + 32); cdOff = u32(u8, z + 48) + u32(u8, z + 52) * 0x100000000; }
    }
    const entries = new Map(), lower = new Map();
    let p = cdOff;
    for (let k = 0; k < count; k++) {
      if (p + 46 > n || u32(u8, p) !== 0x02014b50) throw fail(MSG_BROKEN);
      const flags = u16(u8, p + 8), method = u16(u8, p + 10);
      let csize = u32(u8, p + 20), usize = u32(u8, p + 24), lho = u32(u8, p + 42);
      const nl = u16(u8, p + 28), xl = u16(u8, p + 30), cl = u16(u8, p + 32);
      const name = utf8(u8.subarray(p + 46, p + 46 + nl)).replace(/\\/g, '/');
      if (usize === 0xFFFFFFFF || csize === 0xFFFFFFFF || lho === 0xFFFFFFFF) { // zip64 の大きさ
        let q = p + 46 + nl;
        const qe = q + xl;
        while (q + 4 <= qe) {
          const id = u16(u8, q), sz = u16(u8, q + 2);
          if (id === 0x0001) {
            let r = q + 4;
            if (usize === 0xFFFFFFFF) { usize = u32(u8, r) + u32(u8, r + 4) * 0x100000000; r += 8; }
            if (csize === 0xFFFFFFFF) { csize = u32(u8, r) + u32(u8, r + 4) * 0x100000000; r += 8; }
            if (lho === 0xFFFFFFFF) { lho = u32(u8, r) + u32(u8, r + 4) * 0x100000000; }
          }
          q += 4 + sz;
        }
      }
      const e = { name, flags, method, csize, usize, lho };
      entries.set(name, e);
      lower.set(name.toLowerCase(), e);
      p += 46 + nl + xl + cl;
    }
    const find = (path) => {
      const s = String(path || '').replace(/^\/+/, '');
      return entries.get(s) || lower.get(s.toLowerCase()) || null;
    };
    async function bytes(path) {
      const e = find(path);
      if (!e) return null;
      if (e.flags & 1) throw fail(MSG_PASSWORD);
      const h = e.lho;
      if (u32(u8, h) !== 0x04034b50) throw fail(MSG_BROKEN);
      const start = h + 30 + u16(u8, h + 26) + u16(u8, h + 28);
      if (start + e.csize > n) throw fail(MSG_BROKEN);
      const data = u8.subarray(start, start + e.csize);
      if (e.method === 0) return data;
      if (e.method === 8) return inflateRaw(data);
      throw fail('Excel のファイルとして読めませんでした（xlsx の中の圧縮の形 ' + e.method + ' は読めません）。');
    }
    async function text(path) {
      const b = await bytes(path);
      if (!b) return null;
      if (b[0] === 0xFF && b[1] === 0xFE) return new TextDecoder('utf-16le').decode(b.subarray(2));
      if (b[0] === 0xFE && b[1] === 0xFF) return new TextDecoder('utf-16be').decode(b.subarray(2));
      return utf8(b);
    }
    return { has: (path) => !!find(path), names: () => Array.from(entries.keys()), bytes, text };
  }
  const dirOf = (path) => { const i = path.lastIndexOf('/'); return i < 0 ? '' : path.slice(0, i + 1); };
  function joinPath(base, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    const out = [];
    for (const s of (base + target).split('/')) {
      if (s === '..') out.pop();
      else if (s !== '.' && s !== '') out.push(s);
    }
    return out.join('/');
  }
  const relsPathOf = (path) => dirOf(path) + '_rels/' + path.slice(dirOf(path).length) + '.rels';
  /** .rels の文字 → Map(Id → {type, target}) */
  function relsOf(xml, base) {
    const out = new Map();
    if (!xml) return out;
    for (const r of rootEl(parseXml(xml)).c) {
      if (r.n !== 'Relationship' || r.a.TargetMode === 'External') continue;
      out.set(r.a.Id, { type: r.a.Type || '', target: joinPath(base, r.a.Target || '') });
    }
    return out;
  }
  const relOfType = (rels, suffix) => { for (const r of rels.values()) if (r.type.slice(-suffix.length) === suffix) return r.target; return null; };

  // ---------------------------------------------------------------- セルの番地
  function colName(c) { // 1 から
    let s = '';
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function colIndex(letters) {
    let n = 0;
    const s = letters.toUpperCase();
    for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  }
  /** 'A1' → [行, 列]（1 から）。読めなければ null */
  function cellOf(ref) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref || '');
    return m ? [parseInt(m[2], 10), colIndex(m[1])] : null;
  }
  /** 'A1:B5'・'$A:$C'・'1:3' → [r1, c1, r2, c2]（無い所は null・openpyxl の range_boundaries と同じ） */
  function rangeOf(ref) {
    const m = /^\$?([A-Za-z]{1,3})?\$?(\d+)?(:\$?([A-Za-z]{1,3})?\$?(\d+)?)?$/.exec(String(ref || '').replace(/\s/g, ''));
    if (!m || (!m[1] && !m[2])) return null;
    const c1 = m[1] ? colIndex(m[1]) : null, r1 = m[2] ? parseInt(m[2], 10) : null;
    const c2 = m[4] ? colIndex(m[4]) : c1, r2 = m[5] ? parseInt(m[5], 10) : r1;
    return [r1, c1, r2, c2];
  }

  // ---------------------------------------------------------------- 共有の式を、ほかのセルへずらす（openpyxl の Translator と同じ）
  const ERROR_CODES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];
  const TOKEN_ENDERS = ',;}) +-*/^&=><%';
  const isPyFloat = (s) => /^\s*[+-]?((\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?|inf|infinity|nan)\s*$/i.test(s);
  function shiftRow(s, dr) {
    if (s.charAt(0) === '$') return s;
    const v = parseInt(s, 10) + dr;
    if (v <= 0) throw new Error('out');
    return String(v);
  }
  function shiftCol(s, dc) {
    if (s.charAt(0) === '$') return s;
    const v = colIndex(s) + dc;
    if (v < 1 || v > 18278) throw new Error('out');
    return colName(v);
  }
  function shiftRange(s, dr, dc) {
    let ws = '', rng = s;
    const bang = s.lastIndexOf('!');
    if (bang >= 0) { ws = s.slice(0, bang + 1); rng = s.slice(bang + 1); }
    let m = /^(\$?[1-9][0-9]{0,6}):(\$?[1-9][0-9]{0,6})$/.exec(rng);
    if (m) return ws + shiftRow(m[1], dr) + ':' + shiftRow(m[2], dr);
    m = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})$/.exec(rng);
    if (m) return ws + shiftCol(m[1], dc) + ':' + shiftCol(m[2], dc);
    if (rng.indexOf(':') >= 0) return ws + rng.split(':').map((x) => shiftRange(x, dr, dc)).join(':');
    m = /^(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$/.exec(rng);
    if (!m) return rng; // 名前（openpyxl と同じく、シートの名前は落とす）
    return ws + shiftCol(m[1], dc) + shiftRow(m[2], dr);
  }
  /** '=…' の式を、dr 行・dc 列ずらした式にする（文字・エラー・数・関数の名前はそのまま。$ の付いた所は動かさない） */
  function translateFormula(formula, dr, dc) {
    const f = String(formula || '');
    if (f.charAt(0) !== '=') return f;
    const n = f.length;
    let out = '=', tok = '', i = 1;
    const operand = (v) => {
      if (v.charAt(0) === '#' || v === 'TRUE' || v === 'FALSE' || isPyFloat(v)) return v;
      try { return shiftRange(v, dr, dc); } catch (e) { return '#REF!'; } // 表の外へ出る参照（Excel も #REF! にする）
    };
    const save = () => { if (tok) { out += operand(tok); tok = ''; } };
    while (i < n) {
      const ch = f[i];
      if ((ch === '+' || ch === '-') && tok && /^[1-9](\.[0-9]+)?[Ee]$/.test(tok)) { tok += ch; i++; continue; } // 1E+5
      if (TOKEN_ENDERS.indexOf(ch) >= 0) save();
      if (ch === '"' || ch === "'") {
        let j = i + 1;
        for (;;) {
          const k = f.indexOf(ch, j);
          if (k < 0) { j = n; break; }
          if (f[k + 1] === ch) { j = k + 2; continue; }
          j = k + 1;
          break;
        }
        const s = f.slice(i, j);
        if (ch === '"') out += s; else tok += s;
        i = j;
      } else if (ch === '[') {
        let depth = 0, j = i;
        for (; j < n; j++) { if (f[j] === '[') depth++; else if (f[j] === ']') { depth--; if (depth === 0) { j++; break; } } }
        tok += f.slice(i, j);
        i = j;
      } else if (ch === '#') {
        const rest = f.slice(i);
        const err = ERROR_CODES.find((e) => rest.startsWith(e));
        if (err) { out += tok + err; tok = ''; i += err.length; } else { tok += ch; i++; }
      } else if (ch === ' ' || ch === '\n') {
        out += ch; // openpyxl は続く空白を1つにする
        while (i < n && (f[i] === ' ' || f[i] === '\n')) i++;
      } else if ('+-*/^&=><%'.indexOf(ch) >= 0) {
        const two = f.slice(i, i + 2);
        if (two === '>=' || two === '<=' || two === '<>') { out += two; i += 2; continue; }
        out += ch;
        i++;
      } else if (ch === '(' || ch === '{') {
        if (ch === '(' && tok) { out += tok + '('; tok = ''; } else out += ch;
        i++;
      } else if (ch === ')' || ch === '}' || ch === ';' || ch === ',') {
        out += ch;
        i++;
      } else { tok += ch; i++; }
    }
    save();
    return out;
  }

  // ---------------------------------------------------------------- xlsx（xlsm）を読む
  function colorSpec(el) {
    if (!el) return null;
    const a = el.a, c = {};
    if (a.indexed !== undefined) c.indexed = parseInt(a.indexed, 10);
    else if (a.theme !== undefined) c.theme = parseInt(a.theme, 10);
    else if (a.auto !== undefined) c.auto = true;
    else c.rgb = a.rgb !== undefined ? a.rgb : '00000000'; // 中身の無い色は黒（openpyxl と同じ）
    if (a.tint !== undefined) c.tint = parseFloat(a.tint);
    return c;
  }
  function fontSpec(el, nameTag) {
    const f = { color: null, sz: null, b: false, i: false, u: false, name: null };
    for (const k of el.c) {
      const v = k.a.val;
      if (k.n === 'b') f.b = boolAttr(v, true);
      else if (k.n === 'i') f.i = boolAttr(v, true);
      else if (k.n === 'u') f.u = v !== 'none';
      else if (k.n === 'sz') f.sz = v !== undefined ? parseFloat(v) : null;
      else if (k.n === 'color') f.color = colorSpec(k);
      else if (k.n === nameTag) f.name = v !== undefined ? v : null;
    }
    return f;
  }
  // Excel の _xHHHH_ の書き方を字に戻す（_x005F_ は「_」そのもの）
  const unescapeX = (s) => (s.indexOf('_x') < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16))));
  /** <si> や <is> → {text, parts}（parts は字ごとの飾りのある文字のとき [{text, font}]・openpyxl の CellRichText と同じ分け方） */
  function richOf(el) {
    const t = kid(el, 't');
    if (t && t.t) return { text: unescapeX(t.t), parts: null };
    const parts = [];
    for (const r of kids(el, 'r')) {
      const rt = kid(r, 't'), rp = kid(r, 'rPr');
      parts.push({ text: rt ? unescapeX(rt.t) : '', font: rp ? fontSpec(rp, 'rFont') : null });
    }
    const text = parts.map((x) => x.text).join('');
    if (parts.length === 0 || (parts.length === 1 && !parts[0].font)) return { text, parts: null };
    return { text, parts };
  }
  /** 字ごとの飾り [[from, to, {c, s, b, i, u, f}], …]（make_seed.py の runs_of と同じ・飾りの無い部分は書かない） */
  function runsOfParts(parts, rgb) {
    const out = [];
    let at = 0;
    for (const p of parts) {
      const len = p.text.length;
      if (p.font) out.push([at, at + len, fontDict(p.font, rgb)]);
      at += len;
    }
    return out;
  }
  // 位置の既定と違うか（openpyxl の Alignment() と比べる。値の書いてある属性は、既定と同じ字でも「違う」になる）
  function alignNotDefault(al) {
    if (!al) return false;
    for (const k in al.a) {
      if (k === 'textRotation' || k === 'indent' || k === 'relativeIndent' || k === 'readingOrder') { if (parseFloat(al.a[k]) !== 0) return true; } else return true;
    }
    return false;
  }
  function protNotDefault(pr) {
    if (!pr) return false;
    return !boolAttr(pr.a.locked, true) || boolAttr(pr.a.hidden, false);
  }
  function readStyles(xml) {
    const st = { numFmts: {}, fonts: [], fills: [], borders: [], xfs: [], indexed: null };
    if (!xml) return st;
    for (const sec of rootEl(parseXml(xml)).c) {
      if (sec.n === 'numFmts') {
        for (const f of kids(sec, 'numFmt')) st.numFmts[parseInt(f.a.numFmtId, 10)] = f.a.formatCode || '';
      } else if (sec.n === 'fonts') {
        for (const f of kids(sec, 'font')) st.fonts.push(fontSpec(f, 'name'));
      } else if (sec.n === 'fills') {
        for (const f of kids(sec, 'fill')) {
          const pf = kid(f, 'patternFill');
          st.fills.push(pf ? { type: pf.a.patternType || null, fg: colorSpec(kid(pf, 'fgColor')) || { rgb: '00000000' } } : { type: 'gradient', fg: null });
        }
      } else if (sec.n === 'borders') {
        for (const bd of kids(sec, 'border')) {
          const side = (name) => { const e = kid(bd, name); return e && e.a.style ? e.a.style : ''; };
          st.borders.push({ l: side('left'), r: side('right'), t: side('top'), b: side('bottom') });
        }
      } else if (sec.n === 'cellXfs') {
        for (const x of kids(sec, 'xf')) {
          const al = kid(x, 'alignment');
          st.xfs.push({
            font: parseInt(x.a.fontId || '0', 10), fill: parseInt(x.a.fillId || '0', 10), fmt: parseInt(x.a.numFmtId || '0', 10),
            border: parseInt(x.a.borderId || '0', 10), xfId: parseInt(x.a.xfId || '0', 10),
            flag: boolAttr(x.a.pivotButton, false) || boolAttr(x.a.quotePrefix, false) || alignNotDefault(al) || protNotDefault(kid(x, 'protection')),
            h: al && al.a.horizontal ? al.a.horizontal : '', v: al && al.a.vertical ? al.a.vertical : '',
            wrap: al ? boolAttr(al.a.wrapText, false) : false, shrink: al ? boolAttr(al.a.shrinkToFit, false) : false
          });
        }
      } else if (sec.n === 'colors') {
        const ic = kid(sec, 'indexedColors');
        if (ic) st.indexed = kids(ic, 'rgbColor').map((c) => c.a.rgb || '');
      }
    }
    return st;
  }
  /** xlsx の書式の番号 → { x（書式の残り）, has（書式があるか＝openpyxl の has_style） } */
  function xlsxRest(styles, s) {
    const xf = styles.xfs[s];
    const bd0 = styles.borders[0] || { l: '', r: '', t: '', b: '' };
    if (!xf) return { x: restOf(bd0.l, bd0.r, bd0.t, bd0.b, '', '', false, false, 'General'), has: false };
    const bd = styles.borders[xf.border] || bd0;
    const custom = styles.numFmts[xf.fmt];
    const n = custom !== undefined ? custom : (BUILTIN_FORMATS[xf.fmt] || 'General');
    // 表示形式の番号: ブックの形でも、決まった形と同じ字なら決まった番号になる（openpyxl と同じ）
    const fmtId = custom !== undefined ? (BUILTIN_TEXTS.has(custom) ? (custom === 'General' ? 0 : 1) : 164) : xf.fmt;
    const has = !!(xf.font || xf.fill || xf.border || fmtId || xf.xfId || xf.flag);
    return { x: restOf(bd.l, bd.r, bd.t, bd.b, xf.h, xf.v, xf.wrap, xf.shrink, n), has };
  }
  /** 範囲をいくつか → openpyxl と同じく、列・行の小さい順で先頭の1つ */
  function firstArea(list) {
    if (!list.length) return null;
    const k = (a) => [a[1] || 0, a[0] || 0, a[3] || 0, a[2] || 0];
    list.sort((x, y) => { const a = k(x), b = k(y); for (let i = 0; i < 4; i++) if (a[i] !== b[i]) return a[i] - b[i]; return 0; });
    return list[0];
  }
  /** 印刷範囲の文字（'Sheet'!$A$1:$I$70,…）→ [r1, c1, r2, c2] */
  function printAreaOf(text) {
    const parts = [];
    let cur = '', q = false;
    for (const ch of String(text || '')) {
      if (ch === "'") q = !q;
      if (ch === ',' && !q) { parts.push(cur); cur = ''; } else cur += ch;
    }
    parts.push(cur);
    const areas = [];
    for (const p of parts) {
      const a = rangeOf(p.slice(p.lastIndexOf('!') + 1).replace(/\$/g, ''));
      if (a) areas.push(a);
    }
    return firstArea(areas);
  }
  /** 結合の内側（左上より後ろ）のセルを落とし、左上のセルの罫線を、結合の縁のセルへ届ける（openpyxl と同じ。
   *  Excel も、結合したところは左上のセルの値・字・塗りだけを見せ、縁の罫線で囲む）。base: 何も無いセルの {font, x} */
  function applyMerges(cells, merges, base) {
    if (!merges.length) return;
    const blank = () => { const o = { v: null, f: null }; if (!isEmptyObj(base.font)) o.font = base.font; o.x = Object.assign({}, base.x); return o; };
    // 左上のセルは、右下のセルの右と下の罫線も受け取る（openpyxl と同じ。右下のセルは、このあと消える）
    for (const m of merges) {
      if (m[0] === m[2] && m[1] === m[3]) continue;
      const end = cells[m[2] + ',' + m[3]];
      if (!end || !end.x || (!end.x.br && !end.x.bb)) continue;
      const sk = m[0] + ',' + m[1];
      const st = cells[sk] || (cells[sk] = blank());
      st.x = Object.assign({}, st.x || base.x);
      if (!st.x.br) st.x.br = end.x.br;
      if (!st.x.bb) st.x.bb = end.x.bb;
    }
    for (const key of Object.keys(cells)) {
      const comma = key.indexOf(',');
      const r = +key.slice(0, comma), c = +key.slice(comma + 1);
      for (const m of merges) {
        if (r >= m[0] && r <= m[2] && c >= m[1] && c <= m[3] && (r !== m[0] || c !== m[1])) { delete cells[key]; break; }
      }
    }
    for (const m of merges) {
      const st = cells[m[0] + ',' + m[1]];
      const sx = st && st.x ? st.x : base.x;
      const edges = [['bt', m[0], m[1], m[0], m[3]], ['bl', m[0], m[1], m[2], m[1]], ['br', m[0], m[3], m[2], m[3]], ['bb', m[2], m[1], m[2], m[3]]];
      for (const [k, r1, c1, r2, c2] of edges) {
        const style = sx[k];
        if (!style) continue;
        for (let r = r1; r <= r2; r++) {
          for (let c = c1; c <= c2; c++) {
            if (r === m[0] && c === m[1]) continue;
            const key = r + ',' + c;
            let o = cells[key];
            if (!o) {
              o = { v: null, f: null };
              if (!isEmptyObj(base.font)) o.font = base.font;
              o.x = Object.assign({}, base.x);
              cells[key] = o;
            }
            if (!o.x[k]) o.x[k] = style;
          }
        }
      }
    }
  }
  function newSheet(name, state) {
    return { name, state, tabColor: '', printArea: null, merges: [], hiddenRows: [], hiddenCols: [], colWidths: {}, rowHeights: {}, defaultRowHeight: 15, maxRow: 1, maxCol: 1, x0: null, cells: {} };
  }
  function readSheetXml(xml, sh, ctx) {
    const ws = rootEl(parseXml(xml));
    let maxR = 0, maxC = 0;
    const shared = new Map();
    for (const sec of ws.c) {
      if (sec.n === 'sheetPr') {
        const tc = kid(sec, 'tabColor');
        if (tc) sh.tabColor = ctx.rgb(colorSpec(tc));
      } else if (sec.n === 'sheetFormatPr') {
        if (sec.a.defaultRowHeight !== undefined) sh.defaultRowHeight = parseFloat(sec.a.defaultRowHeight);
      } else if (sec.n === 'cols') {
        const hid = new Set();
        for (const c of kids(sec, 'col')) {
          const lo = parseInt(c.a.min, 10), hi = parseInt(c.a.max, 10);
          if (!(lo >= 1) || !(hi >= lo)) continue;
          const w = c.a.width !== undefined ? parseFloat(c.a.width) : 13;
          const h = boolAttr(c.a.hidden, false);
          for (let k = lo; k <= hi; k++) {
            if (w) sh.colWidths[String(k)] = w;
            if (h) hid.add(k);
          }
        }
        sh.hiddenCols = Array.from(hid).sort((a, b) => a - b);
      } else if (sec.n === 'sheetData') {
        let rowNo = 0;
        const hiddenRows = [];
        for (const row of sec.c) {
          if (row.n !== 'row') continue;
          if (row.a.r !== undefined) rowNo = Math.round(parseFloat(row.a.r)); else rowNo += 1;
          let colNo = 0, dims = false;
          for (const an in row.a) if (an !== 'r' && an !== 'spans' && an.indexOf(':') < 0) { dims = true; break; }
          if (dims) {
            if (boolAttr(row.a.hidden, false)) hiddenRows.push(rowNo);
            const ht = row.a.ht !== undefined ? parseFloat(row.a.ht) : 0;
            if (ht) sh.rowHeights[String(rowNo)] = ht;
          }
          for (const c of row.c) {
            if (c.n !== 'c') continue;
            let r = rowNo, col;
            const rc = c.a.r ? cellOf(c.a.r) : null;
            if (rc) { r = rc[0]; col = rc[1]; } else col = colNo + 1;
            colNo = col;
            if (r > maxR) maxR = r;
            if (col > maxC) maxC = col;
            const o = cellXml(c, r, col, ctx, shared);
            if (o) sh.cells[r + ',' + col] = o;
          }
        }
        sh.hiddenRows = hiddenRows.sort((a, b) => a - b);
      } else if (sec.n === 'mergeCells') {
        for (const m of kids(sec, 'mergeCell')) {
          const a = rangeOf(m.a.ref);
          if (a && a[0] && a[1]) sh.merges.push(a);
        }
      }
    }
    for (const m of sh.merges) { if (m[2] > maxR) maxR = m[2]; if (m[3] > maxC) maxC = m[3]; }
    sh.maxRow = maxR || 1;
    sh.maxCol = maxC || 1;
    applyMerges(sh.cells, sh.merges, ctx.base);
    sh.x0 = { font: ctx.base.font, fill: '', x: ctx.base.x };
  }
  function cellXml(c, r, col, ctx, shared) {
    const t = c.a.t || 'n';
    const s = c.a.s ? parseInt(c.a.s, 10) || 0 : 0;
    const fEl = kid(c, 'f'), vEl = kid(c, 'v');
    const vText = t === 'inlineStr' ? null : (vEl && vEl.t ? vEl.t : null);
    let formula = null;
    if (fEl) {
      formula = '=' + fEl.t;
      if (fEl.a.t === 'shared') { // 共有の式: 元のセルの式を、相対の位置だけずらす
        const si = fEl.a.si;
        const base = shared.get(si);
        if (base) formula = translateFormula(base.f, r - base.r, col - base.c);
        else if (formula !== '=') shared.set(si, { f: formula, r, c: col });
      }
    }
    let v = null, rich = null;
    if (vText !== null) {
      if (t === 'n') {
        v = Number(vText);
        if (ctx.dateXf(s)) v = ctx.timedeltaXf(s) ? v : excelDateIso(v, ctx.date1904);
      } else if (t === 's') {
        const ss = ctx.sst[parseInt(vText, 10)];
        if (ss) { v = ss.text; rich = ss.parts; } else v = '';
      } else if (t === 'b') v = parseInt(vText, 10) !== 0;
      else v = vText; // str・e・d
    } else if (t === 'inlineStr') {
      const is = kid(c, 'is');
      if (is) { const rr = richOf(is); v = rr.text; rich = rr.parts; }
    }
    const lk = ctx.look(s);
    if (formula === null && v === null && !lk.fill && !lk.has) return null; // 値も塗りも書式も無いセル
    const o = { v, f: formula };
    if (typeof v === 'string' && ERROR_TEXTS.has(v)) o.e = 1;
    if (lk.font) o.font = lk.font;
    if (lk.fill) o.fill = lk.fill;
    if (rich && formula === null) {
      const runs = runsOfParts(rich, ctx.rgb);
      if (runs.length) o.runs = runs;
    }
    o.x = lk.x;
    return o;
  }
  async function readXlsx(u8) {
    const zip = openZip(u8);
    let wbPath = 'xl/workbook.xml';
    const rootRels = relsOf(await zip.text('_rels/.rels'), '');
    const od = relOfType(rootRels, '/officeDocument');
    if (od && zip.has(od)) wbPath = od;
    if (/\.bin$/i.test(wbPath) || (!zip.has(wbPath) && zip.has('xl/workbook.bin'))) throw fail(MSG_XLSB);
    if (!zip.has(wbPath)) throw fail(MSG_NOT_EXCEL);
    const [wbXml, relXml] = await Promise.all([zip.text(wbPath), zip.text(relsPathOf(wbPath))]);
    const rels = relsOf(relXml, dirOf(wbPath));
    const wb = rootEl(parseXml(wbXml));
    const pr = kid(wb, 'workbookPr');
    const date1904 = pr ? boolAttr(pr.a.date1904, false) : false;
    const list = [];
    for (const s of kids(kid(wb, 'sheets'), 'sheet')) {
      let rid = s.a['r:id'];
      if (rid === undefined) for (const an in s.a) if (/:id$/.test(an)) rid = s.a[an];
      list.push({ name: s.a.name || '', state: s.a.state || 'visible', rel: rels.get(rid) || null });
    }
    const areas = new Map();
    for (const d of kids(kid(wb, 'definedNames'), 'definedName')) {
      if (d.a.name === '_xlnm.Print_Area' && d.a.localSheetId !== undefined) areas.set(parseInt(d.a.localSheetId, 10), d.t);
    }
    const stylePath = relOfType(rels, '/styles'), sstPath = relOfType(rels, '/sharedStrings');
    const themePath = relOfType(rels, '/theme') || 'xl/theme/theme1.xml';
    // 要る部品だけを、まとめて同時に解く（図・印刷の設定などは解かない）
    const sheetJobs = list.map((s) => (s.rel && /\/worksheet$/.test(s.rel.type) && zip.has(s.rel.target) ? zip.text(s.rel.target) : Promise.resolve(null)));
    const [stXml, sstXml, thXml, ...sheetXmls] = await Promise.all([
      stylePath ? zip.text(stylePath) : null, sstPath ? zip.text(sstPath) : null, zip.text(themePath), ...sheetJobs]);
    const styles = readStyles(stXml);
    const rgb = makePalette(themeOfXml(thXml), styles.indexed && styles.indexed.length ? styles.indexed : COLOR_INDEX);
    const sst = [];
    if (sstXml) for (const si of rootEl(parseXml(sstXml)).c) if (si.n === 'si') sst.push(richOf(si));
    const dateCache = new Map(), tdCache = new Map(), lookCache = new Map();
    const fmtOfXf = (s) => { const xf = styles.xfs[s]; if (!xf) return null; return styles.numFmts[xf.fmt] !== undefined ? styles.numFmts[xf.fmt] : (BUILTIN_FORMATS[xf.fmt] || null); };
    const font0 = fontDict(styles.fonts[0], rgb);
    const ctx = {
      styles, rgb, sst, date1904,
      // 何も無いセルの見た目（書式の番号 0 ではなく、どの書式も当たっていない状態＝openpyxl の新しいセル）
      base: { font: font0, x: xlsxRest(styles, -1).x },
      dateXf: (s) => { let v = dateCache.get(s); if (v === undefined) { v = isDateFormat(fmtOfXf(s)); dateCache.set(s, v); } return v; },
      timedeltaXf: (s) => { let v = tdCache.get(s); if (v === undefined) { v = isTimedeltaFormat(fmtOfXf(s)); tdCache.set(s, v); } return v; },
      look: (s) => {
        let v = lookCache.get(s);
        if (v) return v;
        const xf = styles.xfs[s] || null;
        let fill = '', font = null;
        if (xf) {
          const fl = styles.fills[xf.fill];
          if (fl && fl.type === 'solid') fill = rgb(fl.fg);
          const fd = fontDict(styles.fonts[xf.font], rgb);
          if (!isEmptyObj(fd)) font = fd;
        }
        const rest = xlsxRest(styles, xf ? s : -1);
        v = { font, fill, x: rest.x, has: rest.has };
        lookCache.set(s, v);
        return v;
      }
    };
    const book = { format: 'xlsx', sheets: [] };
    const notes = [];
    list.forEach((s, idx) => {
      const xml = sheetXmls[idx];
      if (xml === null) { notes.push('「' + s.name + '」は表のシートではない（グラフなど）ため、載せていません。'); return; }
      const sh = newSheet(s.name, s.state);
      readSheetXml(xml, sh, ctx);
      if (areas.has(idx)) sh.printArea = printAreaOf(areas.get(idx));
      book.sheets.push(sh);
    });
    return { book, notes, stats: { formulas: 0, formulasUnknown: 0 } };
  }

  // ---------------------------------------------------------------- xls（BIFF8）の入れ物（複合ファイル＝CFB）
  function openCfb(u8) {
    if (u8.length < 512) throw fail(MSG_BROKEN);
    const shift = u16(u8, 0x1E);
    if (shift !== 9 && shift !== 12) throw fail(MSG_BROKEN);
    const ssz = 1 << shift, per = ssz / 4;
    const msz = 1 << u16(u8, 0x20);
    const nFat = u32(u8, 0x2C), dirStart = u32(u8, 0x30), cutoff = u32(u8, 0x38) || 4096;
    const miniFatStart = u32(u8, 0x3C), nMiniFat = u32(u8, 0x40), difStart = u32(u8, 0x44);
    const nSec = Math.ceil((u8.length - ssz) / ssz);
    const off = (s) => (s + 1) * ssz;
    const fatSecs = [];
    for (let k = 0; k < 109 && fatSecs.length < nFat; k++) fatSecs.push(u32(u8, 0x4C + k * 4));
    let d = difStart, guard = 0;
    while (fatSecs.length < nFat && d < 0xFFFFFFFA && guard++ <= nSec) {
      const o = off(d);
      if (o + ssz > u8.length) throw fail(MSG_BROKEN);
      for (let k = 0; k < per - 1 && fatSecs.length < nFat; k++) fatSecs.push(u32(u8, o + k * 4));
      d = u32(u8, o + ssz - 4);
    }
    const fat = new Uint32Array(fatSecs.length * per);
    fatSecs.forEach((s, k) => {
      const o = off(s);
      for (let m = 0; m < per; m++) fat[k * per + m] = o + m * 4 + 4 <= u8.length ? u32(u8, o + m * 4) : 0xFFFFFFFE;
    });
    function chain(start, table) {
      const out = [];
      let s = start;
      while (s < 0xFFFFFFFA) {
        if (s >= table.length || out.length > table.length) throw fail(MSG_BROKEN); // 輪になった鎖・表の外
        out.push(s);
        s = table[s];
      }
      return out;
    }
    function readChain(start, size) {
      const secs = chain(start, fat);
      const buf = new Uint8Array(secs.length * ssz);
      secs.forEach((s, k) => { const o = off(s); buf.set(u8.subarray(o, Math.min(u8.length, o + ssz)), k * ssz); });
      return size === undefined ? buf : buf.subarray(0, Math.min(size, buf.length));
    }
    const dir = readChain(dirStart);
    const entries = [];
    for (let p = 0; p + 128 <= dir.length; p += 128) {
      const type = dir[p + 66];
      if (type !== 1 && type !== 2 && type !== 5) continue;
      const nl = u16(dir, p + 64);
      entries.push({ name: nl >= 2 ? utf16(dir, p, Math.min(31, (nl - 2) / 2)) : '', type, start: u32(dir, p + 116), size: u32(dir, p + 120) });
    }
    const rootE = entries.find((e) => e.type === 5);
    let mini = null, miniFat = null;
    function stream(name) {
      const e = entries.find((x) => x.type === 2 && x.name.toLowerCase() === name.toLowerCase());
      if (!e) return null;
      if (e.size >= cutoff) return readChain(e.start, e.size);
      if (!mini) { // 小さな流れ（4096 バイトより小さい物）は、まとめ置きの中にある
        mini = rootE ? readChain(rootE.start, rootE.size) : new Uint8Array(0);
        const mf = readChain(miniFatStart, nMiniFat * ssz);
        miniFat = new Uint32Array(mf.length / 4);
        for (let k = 0; k < miniFat.length; k++) miniFat[k] = u32(mf, k * 4);
      }
      const secs = chain(e.start, miniFat);
      const buf = new Uint8Array(secs.length * msz);
      secs.forEach((s, k) => buf.set(mini.subarray(s * msz, s * msz + msz), k * msz));
      return buf.subarray(0, Math.min(e.size, buf.length));
    }
    return { has: (name) => entries.some((x) => x.name.toLowerCase() === name.toLowerCase()), stream };
  }

  // ---------------------------------------------------------------- xls の続きの記録（CONTINUE）をまたいで字を読む
  //   字の並びが記録の境目で切れたときは、次の記録の先頭に「字の幅」の 1 バイトが入る（字以外は、そのまま続く）
  function contReader(parts) {
    let total = 0;
    for (const x of parts) total += x.length;
    const buf = new Uint8Array(total), bounds = [];
    let p = 0;
    parts.forEach((x, k) => { if (k) bounds.push(p); buf.set(x, p); p += x.length; });
    return { buf, pos: 0, bounds, bi: 0 };
  }
  function contChars(st, n, hi) {
    let out = '', need = n;
    for (;;) {
      while (st.bi < st.bounds.length && st.bounds[st.bi] < st.pos) st.bi++;
      const lim = st.bi < st.bounds.length ? st.bounds[st.bi] : st.buf.length;
      const size = hi ? 2 : 1;
      const avail = Math.min(need, Math.floor((lim - st.pos) / size));
      if (avail > 0) {
        out += hi ? utf16(st.buf, st.pos, avail) : latin(st.buf, st.pos, avail);
        st.pos += avail * size;
        need -= avail;
      }
      if (need <= 0) return out;
      if (st.pos >= st.buf.length) throw fail(MSG_BROKEN);
      hi = st.buf[st.pos] & 1;
      st.pos += 1;
    }
  }
  // XLUnicodeString（字の数 2 バイト・幅 1 バイト・字）
  function contString(st) {
    const cch = u16(st.buf, st.pos), hi = st.buf[st.pos + 2] & 1;
    st.pos += 3;
    return contChars(st, cch, hi);
  }
  // 記録の中の字（字の数の大きさ cchSize＝1 か 2）
  function recString(b, p, cchSize) {
    const cch = cchSize === 1 ? b[p] : u16(b, p);
    const q = p + cchSize, hi = b[q] & 1;
    const s = hi ? utf16(b, q + 1, cch) : latin(b, q + 1, cch);
    return { s, end: q + 1 + cch * (hi ? 2 : 1) };
  }

  // ---------------------------------------------------------------- xls の式（Ptg の並び）を Excel と同じ字に戻す
  const BIN_OPS = { 0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<', 0x0A: '<=', 0x0B: '=', 0x0C: '>=', 0x0D: '>', 0x0E: '<>', 0x0F: ' ', 0x10: ',', 0x11: ':' };
  const ERR_CODES = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A', 0x2B: '#GETTING_DATA' };
  // 関数の番号 → [名前, 引数の数（-1 は数が変わる関数）]
  const FUNCS = {};
  ('0 COUNT -1|1 IF -1|2 ISNA 1|3 ISERROR 1|4 SUM -1|5 AVERAGE -1|6 MIN -1|7 MAX -1|8 ROW -1|9 COLUMN -1|10 NA 0|11 NPV -1|12 STDEV -1|' +
    '13 DOLLAR -1|14 FIXED -1|15 SIN 1|16 COS 1|17 TAN 1|18 ATAN 1|19 PI 0|20 SQRT 1|21 EXP 1|22 LN 1|23 LOG10 1|24 ABS 1|25 INT 1|' +
    '26 SIGN 1|27 ROUND 2|28 LOOKUP -1|29 INDEX -1|30 REPT 2|31 MID 3|32 LEN 1|33 VALUE 1|34 TRUE 0|35 FALSE 0|36 AND -1|37 OR -1|38 NOT 1|' +
    '39 MOD 2|40 DCOUNT 3|41 DSUM 3|42 DAVERAGE 3|43 DMIN 3|44 DMAX 3|45 DSTDEV 3|46 VAR -1|47 DVAR 3|48 TEXT 2|49 LINEST -1|50 TREND -1|' +
    '51 LOGEST -1|52 GROWTH -1|56 PV -1|57 FV -1|58 NPER -1|59 PMT -1|60 RATE -1|61 MIRR 3|62 IRR -1|63 RAND 0|64 MATCH -1|65 DATE 3|' +
    '66 TIME 3|67 DAY 1|68 MONTH 1|69 YEAR 1|70 WEEKDAY -1|71 HOUR 1|72 MINUTE 1|73 SECOND 1|74 NOW 0|75 AREAS 1|76 ROWS 1|77 COLUMNS 1|' +
    '78 OFFSET -1|82 SEARCH -1|83 TRANSPOSE 1|86 TYPE 1|97 ATAN2 2|98 ASIN 1|99 ACOS 1|100 CHOOSE -1|101 HLOOKUP -1|102 VLOOKUP -1|' +
    '105 ISREF 1|109 LOG -1|111 CHAR 1|112 LOWER 1|113 UPPER 1|114 PROPER 1|115 LEFT -1|116 RIGHT -1|117 EXACT 2|118 TRIM 1|119 REPLACE 4|' +
    '120 SUBSTITUTE -1|121 CODE 1|124 FIND -1|125 CELL -1|126 ISERR 1|127 ISTEXT 1|128 ISNUMBER 1|129 ISBLANK 1|130 T 1|131 N 1|' +
    '140 DATEVALUE 1|141 TIMEVALUE 1|142 SLN 3|143 SYD 4|144 DDB -1|148 INDIRECT -1|162 CLEAN 1|163 MDETERM 1|164 MINVERSE 1|165 MMULT 2|' +
    '167 IPMT -1|168 PPMT -1|169 COUNTA -1|183 PRODUCT -1|184 FACT 1|189 DPRODUCT 3|190 ISNONTEXT 1|193 STDEVP -1|194 VARP -1|195 DSTDEVP 3|' +
    '196 DVARP 3|197 TRUNC -1|198 ISLOGICAL 1|199 DCOUNTA 3|204 USDOLLAR -1|205 FINDB -1|206 SEARCHB -1|207 REPLACEB 4|208 LEFTB -1|' +
    '209 RIGHTB -1|210 MIDB 3|211 LENB 1|212 ROUNDUP 2|213 ROUNDDOWN 2|214 ASC 1|215 JIS 1|216 RANK -1|219 ADDRESS -1|220 DAYS360 -1|' +
    '221 TODAY 0|222 VDB -1|227 MEDIAN -1|228 SUMPRODUCT -1|229 SINH 1|230 COSH 1|231 TANH 1|232 ASINH 1|233 ACOSH 1|234 ATANH 1|235 DGET 3|' +
    '244 INFO 1|247 DB -1|252 FREQUENCY 2|261 ERROR.TYPE 1|269 AVEDEV -1|270 BETADIST -1|271 GAMMALN 1|272 BETAINV -1|273 BINOMDIST 4|' +
    '274 CHIDIST 2|275 CHIINV 2|276 COMBIN 2|277 CONFIDENCE 3|278 CRITBINOM 3|279 EVEN 1|280 EXPONDIST 3|281 FDIST 3|282 FINV 3|283 FISHER 1|' +
    '284 FISHERINV 1|285 FLOOR 2|286 GAMMADIST 4|287 GAMMAINV 3|288 CEILING 2|289 HYPGEOMDIST 4|290 LOGNORMDIST 3|291 LOGINV 3|' +
    '292 NEGBINOMDIST 3|293 NORMDIST 4|294 NORMSDIST 1|295 NORMINV 3|296 NORMSINV 1|297 STANDARDIZE 3|298 ODD 1|299 PERMUT 2|300 POISSON 3|' +
    '301 TDIST 3|302 WEIBULL 4|303 SUMXMY2 2|304 SUMX2MY2 2|305 SUMX2PY2 2|306 CHITEST 2|307 CORREL 2|308 COVAR 2|309 FORECAST 3|310 FTEST 2|' +
    '311 INTERCEPT 2|312 PEARSON 2|313 RSQ 2|314 STEYX 2|315 SLOPE 2|316 TTEST 4|317 PROB -1|318 DEVSQ -1|319 GEOMEAN -1|320 HARMEAN -1|' +
    '321 SUMSQ -1|322 KURT -1|323 SKEW -1|324 ZTEST -1|325 LARGE 2|326 SMALL 2|327 QUARTILE 2|328 PERCENTILE 2|329 PERCENTRANK -1|330 MODE -1|' +
    '331 TRIMMEAN 2|332 TINV 2|336 CONCATENATE -1|337 POWER 2|342 RADIANS 1|343 DEGREES 1|344 SUBTOTAL -1|345 SUMIF -1|346 COUNTIF 2|' +
    '347 COUNTBLANK 1|350 ISPMT 4|351 DATEDIF 3|352 DATESTRING 1|353 NUMBERSTRING 2|354 ROMAN -1|358 GETPIVOTDATA -1|359 HYPERLINK -1|' +
    '360 PHONETIC 1|361 AVERAGEA -1|362 MAXA -1|363 MINA -1|364 STDEVPA -1|365 VARPA -1|366 STDEVA -1|367 VARA -1')
    .split('|').forEach((s) => { const x = s.split(' '); FUNCS[+x[0]] = [x[1], +x[2]]; });
  const BUILTIN_NAMES = ['Consolidate_Area', 'Auto_Open', 'Auto_Close', 'Extract', 'Database', 'Criteria', 'Print_Area', 'Print_Titles', 'Recorder', 'Data_Form', 'Auto_Activate', 'Auto_Deactivate', 'Sheet_Title', '_FilterDatabase'];
  // シートの名前を式に書くとき、要るなら ' で囲む（Excel と同じ: 字・数・_ と . だけで、字から始まり、番地に見えないものは囲まない）
  const SAFE_SHEET = /^[\p{L}_][\p{L}\p{N}_.]*$/u;
  function quoteSheet(name) {
    if (SAFE_SHEET.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^(R\d*C\d*|R|C|TRUE|FALSE)$/i.test(name)) return name;
    return "'" + name.replace(/'/g, "''") + "'";
  }
  /** 式の中の数を Excel と同じ字に（有効な桁は 15 まで・指数は 1E+20 の形） */
  function numText(x) {
    if (!isFinite(x)) return '#NUM!';
    let s = String(Number(x.toPrecision(15)));
    const m = /^(-?[\d.]+)e([+-])(\d+)$/.exec(s);
    if (m) s = m[1] + 'E' + m[2] + (m[3].length < 2 ? '0' + m[3] : m[3]);
    return s;
  }
  function refParts(rw, cf, baseRow, baseCol, relN) {
    const colRel = (cf & 0x4000) !== 0, rowRel = (cf & 0x8000) !== 0;
    let c = cf & 0xFF, r = rw;
    if (relN) { // 共有の式: 相対の所は、このセルからのずれ
      if (rowRel) r = (baseRow + rw) & 0xFFFF;
      if (colRel) c = (baseCol + c) & 0xFF;
    }
    const ct = (colRel ? '' : '$') + colName(c + 1), rt = (rowRel ? '' : '$') + (r + 1);
    return { r, c, ct, rt, t: ct + rt };
  }
  function areaText(b, p, baseRow, baseCol, relN) {
    const a = refParts(u16(b, p), u16(b, p + 4), baseRow, baseCol, relN);
    const z = refParts(u16(b, p + 2), u16(b, p + 6), baseRow, baseCol, relN);
    if (a.r === 0 && z.r === 0xFFFF) return a.ct + ':' + z.ct; // 列まるごと（A:A）
    if (a.c === 0 && z.c === 0xFF) return a.rt + ':' + z.rt;   // 行まるごと（1:1）
    return a.t + ':' + z.t;
  }
  /** EXTERNSHEET の番号 → 'シート!' の字（読めなければ null） */
  function sheetPrefix(g, ixti) {
    const x = g.xti[ixti];
    if (!x) return null;
    const sb = g.supbooks[x.book];
    if (!sb) return null;
    if (sb.self) {
      if (x.first === 0xFFFE) return '';
      if (x.first === 0xFFFF || x.first >= g.sheetNames.length) return '#REF!'; // 消されたシート
      let name = g.sheetNames[x.first];
      if (x.last !== x.first && x.last < g.sheetNames.length) name += ':' + g.sheetNames[x.last];
      return quoteSheet(name) + '!';
    }
    if (sb.ext) { // ほかのブック（[1]シート!）
      if (x.first >= sb.sheets.length) return '[' + sb.ext + ']!';
      const name = sb.sheets[x.first];
      return (quoteSheet(name) === name ? '[' + sb.ext + ']' + name : "'[" + sb.ext + ']' + name.replace(/'/g, "''") + "'") + '!';
    }
    return null;
  }
  function nameXText(g, ixti, idx) {
    const x = g.xti[ixti];
    const sb = x ? g.supbooks[x.book] : null;
    if (!sb) return null;
    if (sb.self) { const nm = g.names[idx - 1]; return nm ? nm.name : null; }
    const en = sb.names[idx - 1];
    if (en === undefined) return null;
    if (sb.addin) return en;
    return sb.ext ? '[' + sb.ext + ']!' + en : null;
  }
  function arrayText(b, cb) { // 式の中の配列の定数（{1,2;3,4}）
    if (cb + 3 > b.length) return null;
    const cols = b[cb] + 1, rows = u16(b, cb + 1) + 1;
    let p = cb + 3;
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const t = b[p];
        if (t === 0x01) { row.push(numText(f64(b, p + 1))); p += 9; }
        else if (t === 0x02) { const s = recString(b, p + 1, 2); row.push('"' + s.s.replace(/"/g, '""') + '"'); p = s.end; }
        else if (t === 0x04) { row.push(b[p + 1] ? 'TRUE' : 'FALSE'); p += 9; }
        else if (t === 0x10) { row.push(ERR_CODES[b[p + 1]] || '#N/A'); p += 9; }
        else if (t === 0x00) { row.push(''); p += 9; }
        else return null;
      }
      out.push(row.join(','));
    }
    return { t: '{' + out.join(';') + '}', end: p };
  }
  /** Ptg の並び（後置の記法）→ '=…'。戻せない Ptg に当たったら null。row・col はこの式のセル（0 から） */
  function decompile(rgce, rgcb, row, col, g) {
    const st = [];
    const n = rgce.length;
    let p = 0, cb = 0, sp = '', spOpen = '', spClose = '';
    const push = (s) => { st.push(sp + s); sp = ''; };
    while (p < n) {
      const ptg = rgce[p++];
      const id = ptg < 0x20 ? ptg : (ptg & 0x1F) | 0x20;
      const op = BIN_OPS[id];
      if (op !== undefined) {
        if (st.length < 2) return null;
        const b = st.pop(), a = st.pop();
        st.push(a + sp + op + b);
        sp = '';
        continue;
      }
      switch (id) {
        case 0x12: case 0x13: if (!st.length) return null; st.push(sp + (id === 0x12 ? '+' : '-') + st.pop()); sp = ''; break;
        case 0x14: if (!st.length) return null; st.push(st.pop() + sp + '%'); sp = ''; break;
        case 0x15: if (!st.length) return null; st.push(sp + spOpen + '(' + st.pop() + spClose + ')'); sp = spOpen = spClose = ''; break;
        case 0x16: push(''); break;
        case 0x17: {
          const cch = rgce[p], hi = rgce[p + 1] & 1;
          const s = hi ? utf16(rgce, p + 2, cch) : latin(rgce, p + 2, cch);
          p += 2 + cch * (hi ? 2 : 1);
          push('"' + s.replace(/"/g, '""') + '"');
          break;
        }
        case 0x19: { // ptgAttr
          const at = rgce[p];
          if (at & 0x04) { p += 3 + 2 * (u16(rgce, p + 1) + 1); break; } // CHOOSE の飛び先
          if (at & 0x10) { p += 3; if (!st.length) return null; st.push('SUM(' + st.pop() + ')'); break; } // 引数が1つの SUM
          if (at & 0x40) { // 空白・改行
            const kind = rgce[p + 1], cnt = rgce[p + 2];
            p += 3;
            const s = (kind & 1 ? '\n' : ' ').repeat(cnt);
            if (kind === 0 || kind === 1 || kind === 6) sp += s;
            else if (kind === 2 || kind === 3) spOpen += s;
            else if (kind === 4 || kind === 5) spClose += s;
            break;
          }
          p += 3; // IF・飛び先・毎回計算の印（字にしない）
          break;
        }
        case 0x1C: { const e = ERR_CODES[rgce[p]]; if (!e) return null; push(e); p += 1; break; }
        case 0x1D: push(rgce[p] ? 'TRUE' : 'FALSE'); p += 1; break;
        case 0x1E: push(String(u16(rgce, p))); p += 2; break;
        case 0x1F: push(numText(f64(rgce, p))); p += 8; break;
        case 0x20: { p += 7; const a = arrayText(rgcb, cb); if (!a) return null; cb = a.end; push(a.t); break; }
        case 0x21: { // 引数の数が決まった関数
          const info = FUNCS[u16(rgce, p)];
          p += 2;
          if (!info || info[1] < 0 || st.length < info[1]) return null;
          const args = info[1] ? st.splice(st.length - info[1], info[1]) : [];
          push(info[0] + '(' + args.join(',') + ')');
          break;
        }
        case 0x22: { // 引数の数が変わる関数
          const argc = rgce[p] & 0x7F, tab = u16(rgce, p + 1) & 0x7FFF;
          p += 3;
          if (st.length < argc) return null;
          const args = argc ? st.splice(st.length - argc, argc) : [];
          let name;
          if (tab === 255) { if (!args.length) return null; name = args.shift().replace(/^_xlfn\./, ''); } // 名前で呼ぶ関数（IFERROR など）
          else { const info = FUNCS[tab]; if (!info) return null; name = info[0]; }
          push(name + '(' + args.join(',') + ')');
          break;
        }
        case 0x23: { const nm = g.names[u32(rgce, p) - 1]; p += 4; if (!nm) return null; push(nm.name); break; }
        case 0x24: push(refParts(u16(rgce, p), u16(rgce, p + 2), row, col, false).t); p += 4; break;
        case 0x25: push(areaText(rgce, p, row, col, false)); p += 8; break;
        case 0x26: p += 6; if (rgcb.length >= cb + 2) cb += 2 + 8 * u16(rgcb, cb); break; // 範囲の覚え（式の字は後ろの Ptg が持つ）
        case 0x27: case 0x28: p += 6; break;
        case 0x29: case 0x2E: case 0x2F: p += 2; break;
        case 0x2A: p += 4; push('#REF!'); break;
        case 0x2B: p += 8; push('#REF!'); break;
        case 0x2C: push(refParts(u16(rgce, p), u16(rgce, p + 2), row, col, true).t); p += 4; break;
        case 0x2D: push(areaText(rgce, p, row, col, true)); p += 8; break;
        case 0x39: { const t = nameXText(g, u16(rgce, p), u32(rgce, p + 2)); p += 6; if (t === null) return null; push(t); break; }
        case 0x3A: { const pre = sheetPrefix(g, u16(rgce, p)); if (pre === null) return null; push(pre + refParts(u16(rgce, p + 2), u16(rgce, p + 4), row, col, false).t); p += 6; break; }
        case 0x3B: { const pre = sheetPrefix(g, u16(rgce, p)); if (pre === null) return null; push(pre + areaText(rgce, p + 2, row, col, false)); p += 10; break; }
        case 0x3C: case 0x3D: { const pre = sheetPrefix(g, u16(rgce, p)); if (pre === null) return null; push(pre === '#REF!' ? '#REF!' : pre + '#REF!'); p += id === 0x3C ? 6 : 10; break; }
        default: return null;
      }
      if (p > n) return null;
    }
    if (st.length !== 1) return null;
    return '=' + st[0];
  }

  // ---------------------------------------------------------------- xls（BIFF8）を読む
  function rkValue(rk) {
    let v;
    if (rk & 2) v = (rk | 0) >> 2;
    else {
      const dv = new DataView(new ArrayBuffer(8));
      dv.setUint32(4, (rk & 0xFFFFFFFC) >>> 0, true);
      v = dv.getFloat64(0, true);
    }
    return rk & 1 ? v / 100 : v;
  }
  function extColor(type, val, tint) { // FullColorExt・CFColor → 色
    if (type === 0) return { auto: true };
    if (type === 1) return { indexed: val };
    if (type === 2) return { rgb: 'FF' + hex2(val & 0xFF) + hex2((val >>> 8) & 0xFF) + hex2((val >>> 16) & 0xFF) };
    if (type === 3) return { theme: val, tint };
    return null;
  }
  // 記録を1つ（種類・中身・次の位置）。続き（CONTINUE）は呼ぶ側で集める
  function recAt(S, p) {
    if (p + 4 > S.length) return null;
    const len = u16(S, p + 2);
    if (p + 4 + len > S.length) throw fail(MSG_BROKEN);
    return { rt: u16(S, p), body: S.subarray(p + 4, p + 4 + len), next: p + 4 + len };
  }
  function contsAfter(S, rec, type) {
    const parts = [rec.body];
    let p = rec.next;
    for (;;) {
      const r = recAt(S, p);
      if (!r || r.rt !== type) break;
      parts.push(r.body);
      p = r.next;
    }
    return { parts, next: p };
  }
  async function readXls(u8) {
    const cfb = openCfb(u8);
    if (cfb.has('EncryptedPackage') || cfb.has('EncryptionInfo')) throw fail(MSG_PASSWORD);
    const S = cfb.stream('Workbook') || cfb.stream('Book');
    if (!S) throw fail(MSG_NOT_EXCEL);
    const first = recAt(S, 0);
    if (!first || (first.rt & 0xFF) !== 0x09) throw fail(MSG_NOT_EXCEL);
    if (first.rt !== 0x0809 || u16(first.body, 0) !== 0x0600) throw fail(MSG_OLD); // BIFF8（Excel 97〜2003）だけを読む
    const notes = [];
    const g = { fonts: [], fmts: {}, xfs: [], xfExt: new Map(), indexed: COLOR_INDEX.slice(), sheets: [], sheetNames: [], sst: [],
      names: [], xti: [], supbooks: [], date1904: false, themeZip: null };
    let p = first.next, nExt = 0, extNameMiss = 0;
    for (;;) {
      const rec = recAt(S, p);
      if (!rec) break;
      p = rec.next;
      const b = rec.body;
      if (rec.rt === 0x000A) break; // EOF
      switch (rec.rt) {
        case 0x002F: throw fail(MSG_PASSWORD); // FILEPASS
        case 0x0022: g.date1904 = u16(b, 0) === 1; break;
        case 0x0031: g.fonts.push({ sz: u16(b, 0) / 20, i: (u16(b, 2) & 2) !== 0, icv: u16(b, 4), bls: u16(b, 6), uls: b[10], name: recString(b, 14, 1).s }); break;
        case 0x041E: g.fmts[u16(b, 0)] = recString(b, 2, 2).s; break;
        case 0x00E0: g.xfs.push({ font: u16(b, 0), fmt: u16(b, 2), prot: u16(b, 4), al: b[6], rot: b[7], ind: b[8], bd: u32(b, 10), bd2: u32(b, 14), fls: (u32(b, 14) >>> 26) & 0x3F, fore: u16(b, 18) & 0x7F }); break;
        case 0x087D: { // XFEXT（Excel 2007 からの色・テーマの色）
          const ixfe = u16(b, 14), cexts = u16(b, 18);
          const ext = {};
          let q = 20;
          for (let k = 0; k < cexts && q + 4 <= b.length; k++) {
            const et = u16(b, q), cbx = u16(b, q + 2);
            if (cbx < 4) break;
            if (et >= 4 && et <= 13 && et !== 6 && q + 12 <= b.length) {
              const c = extColor(u16(b, q + 4), u32(b, q + 8), i16(b, q + 6) / 32767);
              if (c) ext[et] = c;
            }
            q += cbx;
          }
          g.xfExt.set(ixfe, ext);
          break;
        }
        case 0x0092: { const n = u16(b, 0); for (let k = 0; k < n && 8 + k < 64; k++) g.indexed[8 + k] = hex2(b[2 + 4 * k]) + hex2(b[3 + 4 * k]) + hex2(b[4 + 4 * k]); break; } // PALETTE
        case 0x0085: { const nm = recString(b, 6, 1).s; g.sheets.push({ pos: u32(b, 0), state: ['visible', 'hidden', 'veryHidden', 'veryHidden'][b[4] & 3], dt: b[5], name: nm }); g.sheetNames.push(nm); break; }
        case 0x00FC: { // SST（文字の表）
          const c = contsAfter(S, rec, 0x003C);
          p = c.next;
          const st = contReader(c.parts);
          const count = u32(st.buf, 4);
          st.pos = 8;
          for (let k = 0; k < count && st.pos < st.buf.length; k++) {
            const cch = u16(st.buf, st.pos), fl = st.buf[st.pos + 2];
            st.pos += 3;
            let cRun = 0, cbExt = 0;
            if (fl & 8) { cRun = u16(st.buf, st.pos); st.pos += 2; }
            if (fl & 4) { cbExt = u32(st.buf, st.pos); st.pos += 4; }
            const text = contChars(st, cch, fl & 1);
            let runs = null;
            if (cRun) {
              runs = [];
              for (let m = 0; m < cRun; m++) { runs.push([u16(st.buf, st.pos), u16(st.buf, st.pos + 2)]); st.pos += 4; }
            }
            st.pos += cbExt; // ふりがな（読まない）
            g.sst.push({ text, runs });
          }
          break;
        }
        case 0x0017: { // EXTERNSHEET（シートの参照の表）
          const c = contsAfter(S, rec, 0x003C);
          p = c.next;
          const bb = contReader(c.parts).buf;
          const n = u16(bb, 0);
          for (let k = 0; k < n && 2 + 6 * k + 6 <= bb.length; k++) g.xti.push({ book: u16(bb, 2 + 6 * k), first: u16(bb, 4 + 6 * k), last: u16(bb, 6 + 6 * k) });
          break;
        }
        case 0x01AE: { // SUPBOOK（参照するブック）
          const c = contsAfter(S, rec, 0x003C);
          p = c.next;
          const ctab = u16(b, 0), cch = u16(b, 2);
          if (cch === 0x0401) g.supbooks.push({ self: true, names: [] });
          else if (cch === 0x3A01) g.supbooks.push({ addin: true, names: [] });
          else {
            const st = contReader(c.parts);
            st.pos = 4;
            const sheets = [];
            try {
              contChars(st, cch, st.buf[st.pos++] & 1);
              for (let k = 0; k < ctab; k++) sheets.push(contString(st));
            } catch (e) { extNameMiss++; }
            g.supbooks.push({ ext: ++nExt, sheets, names: [] });
          }
          break;
        }
        case 0x0023: { const sb = g.supbooks[g.supbooks.length - 1]; if (sb) sb.names.push(b.length > 7 ? recString(b, 6, 1).s : ''); break; } // EXTERNNAME
        case 0x0018: { // NAME（名前・印刷範囲）
          const c = contsAfter(S, rec, 0x003C);
          p = c.next;
          const nb = contReader(c.parts).buf;
          const fl = u16(nb, 0), cch = nb[3], cce = u16(nb, 4), itab = u16(nb, 8);
          const hi = nb[14] & 1;
          let name = hi ? utf16(nb, 15, cch) : latin(nb, 15, cch);
          const q = 15 + cch * (hi ? 2 : 1);
          if (fl & 0x20) name = BUILTIN_NAMES[name.charCodeAt(0)] || name;
          g.names.push({ name, builtin: (fl & 0x20) !== 0, itab, rgce: nb.subarray(q, q + cce) });
          break;
        }
        case 0x0896: { // THEME（ブックのテーマ。版が 0 のとき、zip のテーマが入っている）
          const ver = u32(b, 12);
          if (ver === 0 && b.length > 16) {
            const parts = [b.subarray(16)];
            let q = p;
            for (;;) {
              const r = recAt(S, q);
              if (!r || (r.rt !== 0x0812 && r.rt !== 0x087F)) break;
              parts.push(r.body.subarray(r.rt === 0x087F ? 12 : 4));
              q = r.next;
            }
            p = q;
            g.themeZip = contReader(parts).buf;
          }
          break;
        }
        default: break;
      }
    }
    if (extNameMiss) notes.push('ほかのブックを参照する所のうち ' + extNameMiss + ' 件は、シートの名前が読めませんでした（その参照の式は空にしました）。');
    // テーマと色
    let theme = THEME_2007;
    if (g.themeZip) {
      try {
        const z = openZip(g.themeZip);
        const tn = z.names().find((x) => /theme\/theme\/theme\d*\.xml$/i.test(x)) || z.names().find((x) => /theme\d*\.xml$/i.test(x));
        const t = tn ? themeOfXml(await z.text(tn)) : {};
        if (Object.keys(t).length) theme = t; else notes.push('ブックのテーマの色が見つからなかったため、既定のテーマ（Office 2007 - 2010）の色で読みました。');
      } catch (e) {
        if (e && e.message === MSG_NO_INFLATE) throw e;
        notes.push('ブックのテーマの色が読めなかったため、既定のテーマ（Office 2007 - 2010）の色で読みました。');
      }
    }
    const rgb = makePalette(theme, g.indexed);
    const fontAt = (i) => g.fonts[i < 4 ? i : i - 1] || g.fonts[0] || null; // 字の番号 4 は無い決まり
    const xfFont = (f, colorOverride) => {
      if (!f) return {};
      return fontDict({ color: colorOverride || (f.icv < 64 ? { indexed: f.icv } : null), sz: f.sz, b: f.bls >= 700, i: f.i, u: f.uls !== 0, name: f.name }, rgb);
    };
    const xfCache = new Map();
    const xfInfo = (ix) => {
      let v = xfCache.get(ix);
      if (v) return v;
      const xf = g.xfs[ix] || g.xfs[15] || g.xfs[0] || { font: 0, fmt: 0, prot: 1, al: 0x20, rot: 0, ind: 0, bd: 0, bd2: 0, fls: 0, fore: 64 };
      const ext = g.xfExt.get(ix) || {};
      const fd = xfFont(fontAt(xf.font), ext[13]);
      const fill = xf.fls === 1 ? rgb(ext[4] || { indexed: xf.fore }) : '';
      // 表示形式: 163 までは Excel の決まった形（xlsx に写しても番号だけが残る）・164 からはブックの形
      const fmt = xf.fmt < 164 ? (BUILTIN_FORMATS[xf.fmt] || null) : (g.fmts[xf.fmt] !== undefined ? g.fmts[xf.fmt] : null);
      // 書式の残り（罫線・位置・折り返し・縮小・表示形式）。縦の位置は「下」が既定（xlsx に書かれない）
      const bd = xf.bd, alc = xf.al & 7, alcV = (xf.al >> 4) & 7;
      const x = restOf(BORDER_NAMES[bd & 15], BORDER_NAMES[(bd >>> 4) & 15], BORDER_NAMES[(bd >>> 8) & 15], BORDER_NAMES[(bd >>> 12) & 15],
        H_NAMES[alc], V_NAMES[alcV], (xf.al & 8) !== 0, (xf.ind & 0x10) !== 0, fmt);
      // 書式があるか（Excel が xlsx に写したとき、どれかの書式の番号が 0 でなくなるか＝openpyxl の has_style）
      const has = xf.font !== 0 || xf.fmt !== 0 || xf.fls !== 0 || (bd & 0xFFFF) !== 0 || alc !== 0 || alcV !== 2 || (xf.al & 0x88) !== 0 ||
        xf.rot !== 0 || (xf.ind & 0xDF) !== 0 || (xf.prot & 0x0B) !== 0x01 || (xf.prot >>> 4) !== 0 || ext[13] !== undefined;
      v = { font: isEmptyObj(fd) ? null : fd, fill, x, has, date: isDateFormat(fmt), td: isTimedeltaFormat(fmt) };
      xfCache.set(ix, v);
      return v;
    };
    const runFontCache = new Map();
    const runFont = (ifnt) => { let v = runFontCache.get(ifnt); if (!v) { v = xfFont(fontAt(ifnt), null); runFontCache.set(ifnt, v); } return v; };
    const runsOfSst = (s) => {
      if (!s.runs) return null;
      const out = [], len = s.text.length;
      for (let k = 0; k < s.runs.length; k++) {
        const from = s.runs[k][0], to = k + 1 < s.runs.length ? Math.min(len, s.runs[k + 1][0]) : len;
        if (from < to) out.push([from, to, runFont(s.runs[k][1])]);
      }
      return out.length ? out : null;
    };
    // 印刷範囲（組み込みの名前 Print_Area。ptgArea3d と EXTERNSHEET でシートを決める）
    const areaBySheet = new Map();
    for (const nm of g.names) {
      if (!nm.builtin || nm.name !== 'Print_Area') continue;
      const rg = nm.rgce, list = [];
      let q = 0, sheet = -1;
      while (q < rg.length) {
        const t = rg[q], id = t < 0x20 ? t : (t & 0x1F) | 0x20;
        if (id === 0x3B || id === 0x3A) {
          const x = g.xti[u16(rg, q + 1)];
          if (x && g.supbooks[x.book] && g.supbooks[x.book].self && x.first < g.sheets.length) sheet = x.first;
          if (id === 0x3B) {
            const r1 = u16(rg, q + 3), r2 = u16(rg, q + 5), c1 = u16(rg, q + 7) & 0xFF, c2 = u16(rg, q + 9) & 0xFF;
            const wcol = r1 === 0 && r2 === 0xFFFF, wrow = c1 === 0 && c2 === 0xFF;
            list.push([wcol ? null : r1 + 1, wrow ? null : c1 + 1, wcol ? null : r2 + 1, wrow ? null : c2 + 1]);
            q += 11;
          } else {
            const r = u16(rg, q + 3) + 1, c = (u16(rg, q + 5) & 0xFF) + 1;
            list.push([r, c, r, c]);
            q += 7;
          }
        } else if (id === 0x29) q += 3;
        else if (id === 0x10 || id === 0x15) q += 1;
        else break;
      }
      if (sheet < 0 && nm.itab > 0) sheet = nm.itab - 1;
      if (sheet >= 0 && list.length) areaBySheet.set(sheet, firstArea(list));
    }
    const book = { format: 'xls', sheets: [] };
    const stats = { formulas: 0, formulasUnknown: 0 };
    const base = { font: xfFont(fontAt(0), null), x: DEFAULT_REST }; // 何も無いセルの見た目（字は字の番号 0）
    g.sheets.forEach((bs, idx) => {
      if (bs.dt !== 0) { notes.push('「' + bs.name + '」は表のシートではない（グラフ・マクロなど）ため、載せていません。'); return; }
      const sh = newSheet(bs.name, bs.state);
      readXlsSheet(S, bs.pos, sh, g, { rgb, xfInfo, runsOfSst, stats, base });
      if (areaBySheet.has(idx)) sh.printArea = areaBySheet.get(idx);
      book.sheets.push(sh);
    });
    if (stats.formulasUnknown) notes.push('式を字に戻せなかったセルが ' + stats.formulasUnknown + ' 件あります（値は読めています。そのセルの式は空にして fUnknown の印を付けました）。');
    return { book, notes, stats };
  }
  function readXlsSheet(S, pos, sh, g, ctx) {
    const head = recAt(S, pos);
    if (!head || head.rt !== 0x0809) throw fail(MSG_BROKEN);
    let p = head.next, depth = 0, pending = null, maxR = 0, maxC = 0;
    const cells = sh.cells, shared = [], arrays = [], expCells = [], hiddenRows = [], hidCols = new Set();
    const rowInfo = [], rowXf = new Map(), colXf = [], phon = [], seen = new Set();
    // セルを1つ置く（r, c は 0 から）。値も塗りも書式も無いセルは載せない
    const put = (r, c, ixfe, v, isValue) => {
      if (r + 1 > maxR) maxR = r + 1;
      if (c + 1 > maxC) maxC = c + 1;
      seen.add(r * 256 + c);
      const info = ctx.xfInfo(ixfe);
      if (!isValue && !info.fill && !info.has) return null; // 値も塗りも書式も無いセル
      if (typeof v === 'number' && info.date && !info.td) v = excelDateIso(v, g.date1904);
      const o = { v, f: null };
      if (typeof v === 'string' && ERROR_TEXTS.has(v)) o.e = 1;
      if (info.font) o.font = info.font;
      if (info.fill) o.fill = info.fill;
      o.x = info.x;
      cells[(r + 1) + ',' + (c + 1)] = o;
      return o;
    };
    for (;;) {
      const rec = recAt(S, p);
      if (!rec) break;
      p = rec.next;
      const b = rec.body;
      if (rec.rt === 0x0809) { depth++; continue; } // シートの中のグラフ（BOF〜EOF）は飛ばす
      if (rec.rt === 0x000A) { if (depth === 0) break; depth--; continue; }
      if (depth > 0) continue;
      switch (rec.rt) {
        case 0x00FD: { // LABELSST
          const s = g.sst[u32(b, 6)] || { text: '', runs: null };
          const o = put(u16(b, 0), u16(b, 2), u16(b, 4), s.text, true);
          const runs = ctx.runsOfSst(s);
          if (runs) o.runs = runs;
          break;
        }
        case 0x0204: put(u16(b, 0), u16(b, 2), u16(b, 4), recString(b, 6, 2).s, true); break; // LABEL
        case 0x00D6: { // RSTRING
          const s = recString(b, 6, 2);
          const n = u16(b, s.end), runs = [];
          for (let k = 0; k < n; k++) runs.push([u16(b, s.end + 2 + 4 * k), u16(b, s.end + 4 + 4 * k)]);
          const o = put(u16(b, 0), u16(b, 2), u16(b, 4), s.s, true);
          const rr = ctx.runsOfSst({ text: s.s, runs: runs.length ? runs : null });
          if (rr) o.runs = rr;
          break;
        }
        case 0x0203: put(u16(b, 0), u16(b, 2), u16(b, 4), f64(b, 6), true); break; // NUMBER
        case 0x027E: put(u16(b, 0), u16(b, 2), u16(b, 4), rkValue(u32(b, 6)), true); break; // RK
        case 0x00BD: { // MULRK
          const r = u16(b, 0), c0 = u16(b, 2), n = (b.length - 6) / 6;
          for (let k = 0; k < n; k++) put(r, c0 + k, u16(b, 4 + 6 * k), rkValue(u32(b, 6 + 6 * k)), true);
          break;
        }
        case 0x0205: put(u16(b, 0), u16(b, 2), u16(b, 4), b[7] ? (ERR_CODES[b[6]] || '#N/A') : b[6] !== 0, true); break; // BOOLERR
        case 0x0201: put(u16(b, 0), u16(b, 2), u16(b, 4), null, false); break; // BLANK
        case 0x00BE: { // MULBLANK
          const r = u16(b, 0), c0 = u16(b, 2), n = (b.length - 6) / 2;
          for (let k = 0; k < n; k++) put(r, c0 + k, u16(b, 4 + 2 * k), null, false);
          break;
        }
        case 0x0006: { // FORMULA
          const r = u16(b, 0), c = u16(b, 2), ixfe = u16(b, 4);
          let v, str = false;
          if (b[12] === 0xFF && b[13] === 0xFF) {
            const t = b[6];
            if (t === 0) { v = null; str = true; } // 文字（後ろの STRING の記録にある）
            else if (t === 1) v = b[8] !== 0;
            else if (t === 2) v = ERR_CODES[b[8]] || '#N/A';
            else v = null; // 空の文字（openpyxl と同じく null）
          } else v = f64(b, 6);
          const cce = u16(b, 20), rgce = b.subarray(22, 22 + cce), rgcb = b.subarray(22 + cce);
          const o = put(r, c, ixfe, v, true);
          ctx.stats.formulas++;
          if (cce >= 5 && (rgce[0] === 0x01 || rgce[0] === 0x02)) expCells.push({ o, r, c, mr: u16(rgce, 1), mc: u16(rgce, 3), tbl: rgce[0] === 0x02 });
          else {
            o.f = decompile(rgce, rgcb, r, c, g);
            if (o.f === null) { o.fUnknown = 1; ctx.stats.formulasUnknown++; }
          }
          pending = str ? o : null;
          break;
        }
        case 0x0207: { // STRING（直前の式の、文字の結果）
          const cc = contsAfter(S, rec, 0x003C);
          p = cc.next;
          if (pending) {
            const s = contString(contReader(cc.parts));
            pending.v = s === '' ? null : s;
            if (ERROR_TEXTS.has(s)) pending.e = 1;
          }
          pending = null;
          break;
        }
        case 0x04BC: { const cce = u16(b, 8); shared.push({ r1: u16(b, 0), r2: u16(b, 2), c1: b[4], c2: b[5], rgce: b.subarray(10, 10 + cce), rgcb: b.subarray(10 + cce) }); break; } // SHRFMLA
        case 0x0221: { const cce = u16(b, 12); arrays.push({ r1: u16(b, 0), c1: b[4], rgce: b.subarray(14, 14 + cce), rgcb: b.subarray(14 + cce) }); break; } // ARRAY
        case 0x00E5: { // MERGEDCELLS
          const n = u16(b, 0);
          for (let k = 0; k < n && 2 + 8 * k + 8 <= b.length; k++) sh.merges.push([u16(b, 2 + 8 * k) + 1, u16(b, 6 + 8 * k) + 1, u16(b, 4 + 8 * k) + 1, u16(b, 8 + 8 * k) + 1]);
          break;
        }
        case 0x007D: { // COLINFO（列の幅・隠し）
          // 最後の列（IV）まで届く列の設定は、Excel が xlsx に写すとき、xlsx の最後の列（XFD）まで伸ばす
          const lo = u16(b, 0), hi = u16(b, 2) >= 255 ? 16383 : u16(b, 2), w = u16(b, 4) / 256, hid = (u16(b, 8) & 1) !== 0, ixfe = u16(b, 6);
          for (let k = lo; k <= hi; k++) { if (w) sh.colWidths[String(k + 1)] = w; if (hid) hidCols.add(k + 1); }
          colXf.push([lo, Math.min(255, hi), ixfe]);
          break;
        }
        case 0x0208: { // ROW（行の高さ・隠し）
          const r = u16(b, 0) + 1, fl = u16(b, 12);
          if (fl & 0x20) hiddenRows.push(r);
          rowInfo.push([r, (u16(b, 6) & 0x7FFF) / 20, (fl & 0x40) !== 0]);
          if (fl & 0x80) rowXf.set(r - 1, u16(b, 14) & 0x0FFF); // 行の書式
          break;
        }
        case 0x00EF: { // PHONETICINFO（ふりがなを見せるセルの範囲。Excel は xlsx に写すとき、そこにセルを作る）
          const n = u16(b, 4);
          for (let k = 0; k < n && 6 + 8 * k + 8 <= b.length; k++) phon.push([u16(b, 6 + 8 * k), u16(b, 8 + 8 * k), u16(b, 10 + 8 * k), u16(b, 12 + 8 * k)]);
          break;
        }
        case 0x0225: sh.defaultRowHeight = u16(b, 2) / 20; break; // DEFAULTROWHEIGHT
        case 0x0862: { // SHEETEXT（シートの見出しの色）
          const cb = u32(b, 12);
          let c = null;
          if (cb >= 0x28 && b.length >= 40) c = extColor(u32(b, 24), u32(b, 28), f64(b, 32));
          if (!c) { const icv = u32(b, 16) & 0x7F; c = icv < 64 ? { indexed: icv } : null; }
          sh.tabColor = ctx.rgb(c);
          break;
        }
        default: break;
      }
    }
    // 行の高さ: 高さを決めた行と、既定の高さと違う行（Excel が xlsx に書く行と同じ）
    for (const [r, h, custom] of rowInfo) if (h && (custom || h !== sh.defaultRowHeight)) sh.rowHeights[String(r)] = h;
    // ふりがなの範囲で、記録の無いセル（書式は 行 → 列 → 既定 の順で決まる）
    for (const [r1, r2, c1, c2] of phon) {
      if ((r2 - r1 + 1) * (c2 - c1 + 1) > 100000) continue; // 行・列まるごとの範囲（Excel もセルを作らない）
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
        if (seen.has(r * 256 + c)) continue;
        let ixfe = rowXf.get(r);
        if (ixfe === undefined) { const ci = colXf.find((x) => c >= x[0] && c <= x[1]); ixfe = ci ? ci[2] : 15; }
        put(r, c, ixfe, null, false);
      }
    }
    for (const x of expCells) { // 共有の式・配列の式（元の式は SHRFMLA・ARRAY の記録にある）
      let f = null;
      if (!x.tbl) {
        const s = shared.find((y) => y.r1 === x.mr && y.c1 === x.mc) || shared.find((y) => x.r >= y.r1 && x.r <= y.r2 && x.c >= y.c1 && x.c <= y.c2);
        if (s) f = decompile(s.rgce, s.rgcb, x.r, x.c, g);
        else {
          const a = arrays.find((y) => y.r1 === x.mr && y.c1 === x.mc);
          if (a && x.r === a.r1 && x.c === a.c1) f = decompile(a.rgce, a.rgcb, x.r, x.c, g);
          else if (a) continue; // 配列の式の左上以外（Excel も式を左上にだけ持つ）
        }
      }
      x.o.f = f;
      if (f === null) { x.o.fUnknown = 1; ctx.stats.formulasUnknown++; }
    }
    for (const m of sh.merges) { if (m[2] > maxR) maxR = m[2]; if (m[3] > maxC) maxC = m[3]; }
    sh.maxRow = maxR || 1;
    sh.maxCol = maxC || 1;
    sh.hiddenRows = hiddenRows.sort((a, b2) => a - b2);
    sh.hiddenCols = Array.from(hidCols).sort((a, b2) => a - b2);
    applyMerges(cells, sh.merges, ctx.base);
    sh.x0 = { font: ctx.base.font, fill: '', x: ctx.base.x };
  }

  // ---------------------------------------------------------------- 入口
  /** 中身の先頭から形を見分ける → 'xlsx' | 'xls' | 'old'（Excel 2〜4 の生の BIFF）| ''（Excel ではない） */
  function detectFormat(bytes) {
    const b = toU8(bytes);
    if (b.length >= 4 && b[0] === 0x50 && b[1] === 0x4B && b[2] === 0x03 && b[3] === 0x04) return 'xlsx';
    const sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    if (b.length >= 8 && sig.every((x, i) => b[i] === x)) return 'xls';
    if (b.length >= 8 && b[0] === 0x09 && (b[1] === 0x00 || b[1] === 0x02 || b[1] === 0x04) && u16(b, 2) >= 4 && u16(b, 2) <= 16) return 'old';
    return '';
  }
  /** Excel のブックの中身（Uint8Array か ArrayBuffer）→ Promise<読んだブック>
   *  book.notes（JSON に出ない）: 読めたが戻せなかった物の件数と理由。book.stats（JSON に出ない）: 式の数・戻せなかった式の数 */
  async function readBook(bytes) {
    const u8 = toU8(bytes);
    const kind = detectFormat(u8);
    if (kind === 'old') throw fail(MSG_OLD);
    if (!kind) throw fail(MSG_NOT_EXCEL);
    let r;
    try {
      r = kind === 'xlsx' ? await readXlsx(u8) : await readXls(u8);
    } catch (e) {
      if (e && e.bookError) throw e;
      throw fail(MSG_BROKEN, e);
    }
    Object.defineProperty(r.book, 'notes', { value: r.notes, enumerable: false });
    Object.defineProperty(r.book, 'stats', { value: r.stats, enumerable: false });
    return r.book;
  }
  /** セルの字ごとの見え方（字の飾りがあればそれ・無ければセルの字の飾り）→ [{c, s, b, i, u, f}, …]（字の位置は UTF-16） */
  function charLooks(cell) {
    if (!cell || cell.v === null || cell.v === undefined) return [];
    const text = String(cell.v);
    const out = new Array(text.length).fill(cell.font || {});
    if (cell.runs) for (const r of cell.runs) for (let k = r[0]; k < r[1] && k < out.length; k++) out[k] = r[2];
    return out;
  }

  root.SouzaiBook = { readBook, detectFormat, charLooks, translateFormula, tintHex, isDateFormat, excelDateIso, numText, quoteSheet };
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
