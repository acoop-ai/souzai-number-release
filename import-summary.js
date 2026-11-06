// サマリー（その月の販売計画の Excel）から、月間特売の商品と価格を読む部品（画面の部品を持たない）。画面と試験で同じ物を使う。
//   サマリーから番号表へ直接写るのは「月間厳選特」の行だけ（月間商品の上半分＝月間特売・番号表のピンクの字・特売の価格）。
//   ① readSummaryBook: 読んだブック（frontend/book-read.js の形）から、その行の商品・規格・価格を読む
//   ② matchItems: 読んだ商品に、マスターの商品の候補（上位3つ・点数・理由）を付ける
//   ③ toMonthlyRows: 月間商品の「月間特売」の部の行と、番号表の行に当てる変更の案を作る（確かでない物は「人が選ぶ」へ）
//   信頼境界: ブックの中身は BY が手で打った文字（社内の資料・書き方はゆれる）。式は計算しない（保存されていた値だけを読む）。外へは何も送らない。
//   名前の字の直し方は frontend/search-norm.js（SearchNorm）があれば使う（fold と、読みの表を読んでいれば漢字の読み）。無いときは同じ決まりの、ここの直し方で動く。
//   前提（2026年8〜11月のサマリーで確かめた形。合わなければ problems で知らせる）:
//     見出し＝「月間厳選特」で始まるセル（月間厳選特売・月間厳選特価商品）。読む行＝見出しのつないだセルの行（つないでいなければ、見出しの列が空の行が続く所まで）。
//     商品の文＝その行の、見出しより右のセル（左から）。商品の区切り＝価格（数＋円）。数だけのセルは価格（円）として読む。
(function (root) {
  'use strict';

  // ---------------------------------------------------------------- 字をそろえる
  const nk = (v) => String(v === null || v === undefined ? '' : v).normalize('NFKC');
  const isBlank = (v) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  const COMBINE = /^[ﾞﾟ゙゚]$/; // 半角カナの濁点・半濁点（前の字とまとめてそろえる）

  /** 1字ずつ NFKC にそろえ（空白・改行・タブは半角の空白1字に）、元の字の位置を覚える。名前は元の字のまま返すため */
  function normMap(text) {
    const src = String(text === null || text === undefined ? '' : text);
    let s = '';
    const from = [], to = [];
    for (let i = 0; i < src.length;) {
      let w = src.codePointAt(i) > 0xFFFF ? 2 : 1;
      if (i + w < src.length && COMBINE.test(src.charAt(i + w))) w++;
      const n = src.substr(i, w).normalize('NFKC').replace(/\s/g, ' ');
      for (let k = 0; k < n.length; k++) { s += n.charAt(k); from.push(i); to.push(i + w); }
      i += w;
    }
    return { src, s, from, to };
  }
  /** そろえた字の [a, b) にあたる、元の字 */
  const origOf = (m, a, b) => (a >= b ? '' : m.src.slice(m.from[a], m.to[b - 1]));

  // 名前を比べる形（search-norm.js の fold と同じ決まり: NFKC・小文字・カタカナ → ひらがな・小書き → ふつうの字・記号と空白を外す）
  function foldLocal(v) {
    let s = nk(v).toLowerCase();
    s = s.replace(/[ァ-ヶヽヾ]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0x60));
    s = s.replace(/ゔ([ぁぃぇぉ])/g, (m, x) => 'ばびべぼ'.charAt('ぁぃぇぉ'.indexOf(x))).replace(/ゔ/g, 'ぶ');
    s = s.replace(/ぢ/g, 'じ').replace(/づ/g, 'ず').replace(/を/g, 'お');
    s = s.replace(/[ぁぃぅぇぉゃゅょっゎゕゖ]/g, (c) => 'あいうえおやゆよつわかけ'.charAt('ぁぃぅぇぉゃゅょっゎゕゖ'.indexOf(c)));
    return s.replace(/[^0-9a-zぁ-ゖ㐀-䶿一-鿿々〆]/g, '');
  }
  const searchNorm = () => (root.SearchNorm && typeof root.SearchNorm.fold === 'function' ? root.SearchNorm : null);
  /** 濁点・半濁点を外した形（ドッグ と ドック を同じに） */
  const loosen = (f) => f.normalize('NFD').replace(/[゙゚]/g, '').normalize('NFC');

  // ---------------------------------------------------------------- 規格と価格の形（NFKC にそろえた字で見る）
  const UNIT = '(?:個|ヶ|コ|本|枚|切れ|切|尾|貫|パック|P|p|g|kg|ml|等分|袋|杯|食|人前|串|粒|玉|箱|缶|束|丁|カップ|セット|皿|入り|入)';
  // 頭の「各種・各」は数にくっついているときだけ規格（「各１パック」）。空白をはさむ「…各種　２個」の各種は名前の側
  const QTY = '(?:各種(?=\\d)|各(?=\\d)|よりどり\\s*|約)?\\d+(?:\\.\\d+)?\\s*' + UNIT + '(?:入り|入)?(?:\\s*\\+\\s*\\d+\\s*' + UNIT + '?\\s*増量)?(?:\\s*\\([^()]*\\))?';
  const RE_SPEC_END = new RegExp('(' + QTY + ')\\s*$');
  const RE_QTY_IN = new RegExp(QTY);
  const RE_PRICE = /(\d{1,3}(?:,\d{3})+|\d+)\s*円/g;
  const RE_MORE_END = /\s*(\d+(?:\.\d+)?\s*%\s*増量)\s*$/;
  const RE_MORE_HEAD = /^\s*(\d+(?:\.\d+)?\s*%\s*増量)\s*/;
  const SIZE_WORD = /^(?:小|中|大|特大|ミニ|ハーフ|小パック|大パック|徳用|お徳用)$/;
  const LIST_SEP = /[・･、]/;

  /** 規格をそろえる: NFKC・数と単位の間の空白を外す・2つ以上の規格は空白1つで並べる */
  function normSpec(v) {
    return nk(v).replace(/\s+/g, ' ').trim().replace(/(\d)\s+(?=\D)/g, '$1').replace(/\s*\+\s*/g, '+').replace(/\s*\(\s*/g, '(').replace(/\s*\)/g, ')');
  }
  /** 規格を比べる形（1P と 1パック・3切 と 3切れ・6個 と 6個入 を同じに） */
  function specKey(v) {
    return normSpec(v).toLowerCase().replace(/\s/g, '').replace(/パック|p/g, 'パ').replace(/切れ/g, '切').replace(/入り|入/g, '').replace(/ヶ|コ/g, '個');
  }

  // ---------------------------------------------------------------- 1セルの長い文を、商品・規格・価格に分ける
  /**
   * 文を分ける。返す: { items: [{ text, name, specs: [{ 規格, 本体価格 }], 増量?, 各種?, 各?, variants? }], leftovers: [読めなかった文] }
   *   区切り＝価格（数＋円）。価格の前の字が「名前＋規格」。名前の無い「規格＋価格」は、前の商品の2つ目の規格（３個 ２８０円）。
   *   名前の中の2つ以上続く空白は商品の区切りなので、その前の字は価格の無い文として leftovers へ。最後の価格のあとの字も leftovers へ。
   */
  function splitText(text) {
    const m = normMap(text), s = m.s, items = [], leftovers = [];
    let pos = 0, cur = null, hit;
    const trimL = (a, b) => { while (a < b && s.charAt(a) === ' ') a++; return a; };
    const trimR = (a, b) => { while (b > a && s.charAt(b - 1) === ' ') b--; return b; };
    const left = (a, b) => { a = trimL(a, b); b = trimR(a, b); if (a < b) leftovers.push(origOf(m, a, b)); };
    const re = new RegExp(RE_PRICE.source, 'g');
    while ((hit = re.exec(s))) {
      const price = Number(hit[1].replace(/,/g, '')), pe = hit.index + hit[0].length;
      let a = trimL(pos, hit.index);
      const b = trimR(a, hit.index);
      pos = pe;
      if (a === b) { // 価格だけ（前の商品の、規格の無い2つ目の価格）
        if (cur) { cur.specs.push({ '規格': '', '本体価格': price }); cur._end = pe; } else left(a, pe);
        continue;
      }
      // 終わりの規格（空白で区切った規格がもう1つあれば、それも）
      const specs = [];
      let nameEnd = b, mm = RE_SPEC_END.exec(s.slice(a, b));
      if (mm) {
        specs.unshift(mm[1]); nameEnd = trimR(a, a + mm.index);
        if (nameEnd < a + mm.index || /^\s/.test(mm[1])) {
          const m2 = RE_SPEC_END.exec(s.slice(a, nameEnd));
          if (m2 && (m2.index === 0 || s.charAt(a + m2.index - 1) === ' ' || /^\s/.test(m2[1]))) { specs.unshift(m2[1]); nameEnd = trimR(a, a + m2.index); }
        }
      }
      // ％増量（名前の終わりか頭）
      let more = '', start = a, z = RE_MORE_END.exec(s.slice(a, nameEnd));
      if (z) { more = z[1]; nameEnd = trimR(a, a + z.index); } else if ((z = RE_MORE_HEAD.exec(s.slice(a, nameEnd)))) { more = z[1]; a = trimL(a + z[0].length, nameEnd); }
      if (!specs.length && more && (mm = RE_SPEC_END.exec(s.slice(a, nameEnd)))) { specs.unshift(mm[1]); nameEnd = trimR(a, a + mm.index); }
      // 名前の中の、2つ以上続く空白（その前は価格の無い文）
      const gap = /\s{2,}(?=\S)(?!.*\s{2,}\S)/.exec(s.slice(a, nameEnd));
      if (gap) { left(a, a + gap.index); a = trimL(a + gap.index, nameEnd); start = a; }
      const spec = normSpec(specs.join(' '));
      const nameN = s.slice(a, nameEnd);
      if (cur && !more && (a === nameEnd || SIZE_WORD.test(nameN))) { // 名前の無い「規格＋価格」→ 前の商品の続き
        cur.specs.push({ '規格': normSpec((a === nameEnd ? '' : nameN + ' ') + spec), '本体価格': price }); cur._end = pe;
        continue;
      }
      if (a === nameEnd) { left(start, pe); cur = null; continue; }
      cur = { _start: start, _end: pe, name: origOf(m, a, nameEnd).trim(), specs: [{ '規格': spec, '本体価格': price }] };
      if (more) cur['増量'] = normSpec(more);
      items.push(cur);
    }
    left(pos, s.length);
    return {
      items: items.map((it) => {
        const o = { text: origOf(m, it._start, it._end).trim(), name: it.name, specs: it.specs };
        if (it['増量']) o['増量'] = it['増量'];
        // 各種＝まとめて1つ（月間商品にも1行）。各＝並んだ種類の「それぞれ」（各１パック。月間商品は種類ごとの行）
        if (/各種/.test(it.name) || it.specs.some((x) => /^各種/.test(x['規格']))) o['各種'] = true;
        else if (it.specs.some((x) => /^各/.test(x['規格']))) o['各'] = true;
        const v = variantsOf(it.name); if (v) o.variants = v;
        return o;
      }),
      leftovers
    };
  }

  /** 名前に並んだ種類（A（x・y）各種／A・B）を、1つずつの名前に分ける。分けられなければ null。返す: [{ name, base, part }] */
  function variantsOf(name) {
    const t = String(name).replace(/各種/g, '').trim();
    const top = [];
    let depth = 0, from = 0;
    for (let i = 0; i < t.length; i++) {
      const c = t.charAt(i);
      if ('（(【「'.indexOf(c) >= 0) depth++;
      else if ('）)】」'.indexOf(c) >= 0) depth = Math.max(0, depth - 1);
      else if (depth === 0 && LIST_SEP.test(c)) { top.push(t.slice(from, i)); from = i + 1; }
    }
    if (top.length) {
      top.push(t.slice(from));
      const parts = top.map((x) => x.trim()).filter(Boolean);
      return parts.length >= 2 ? parts.map((p) => ({ name: p, base: '', part: p })) : null;
    }
    const g = /[（(]([^（）()]*[・･、][^（）()]*)[）)]/.exec(t);
    if (!g) return null;
    const base = (t.slice(0, g.index) + t.slice(g.index + g[0].length)).replace(/\s+/g, ' ').trim();
    const parts = g[1].split(LIST_SEP).map((x) => x.trim()).filter(Boolean);
    if (!base || parts.length < 2) return null;
    return parts.map((p) => ({ name: base + '（' + p + '）', base, part: p }));
  }

  // ---------------------------------------------------------------- ブックから「月間厳選特」の行を読む
  const colName = (c) => { let s = ''; for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  const RE_HEAD = /^[^\p{L}\p{N}]*月間厳選特/u; // 頭の記号（★・【 など）はあってよい
  /** 見出しの文から、期間（月/日～月/日）。無ければ null */
  function periodOf(text) {
    const t = nk(text).replace(/\s/g, '');
    const p = /(\d{1,2})[\/月](\d{1,2})日?[~〜\-ー―－](?:(\d{1,2})[\/月])?(\d{1,2})日?/.exec(t);
    if (!p) return null;
    const m1 = Number(p[1]), d1 = Number(p[2]), m2 = p[3] ? Number(p[3]) : m1, d2 = Number(p[4]);
    if (m1 < 1 || m1 > 12 || m2 < 1 || m2 > 12) return null;
    const y = /((?:19|20)\d\d)/.exec(t.slice(0, p.index));
    return { m1, d1, m2, d2, year: y ? Number(y[1]) : null, text: m1 + '/' + d1 + '～' + m2 + '/' + d2 };
  }
  /** ファイルの名前から、年と月（2026.10サマリー → 2026・10）。無ければ null */
  function fileMonthOf(fileName) {
    const t = nk(fileName || '');
    const p = /((?:19|20)\d\d)\s*[.\-_年\/]?\s*(\d{1,2})(?!\d)/.exec(t);
    if (!p) { const y = /((?:19|20)\d\d)/.exec(t); return y ? { year: Number(y[1]), month: null } : null; }
    const mo = Number(p[2]);
    return { year: Number(p[1]), month: mo >= 1 && mo <= 12 ? mo : null };
  }

  /**
   * 読んだブックから、月間特売の商品を読む。opts: { fileName（年と月の手がかり）, year（年。ファイルの名前に無いとき） }
   *   読むのは、見えるシートの、印刷範囲の中の、隠し行・隠し列でない「月間厳選特」の見出しの行（とつないだセルの行）だけ。
   * 返す: { month: 'YYYY-MM' | null, period: '11/1～11/30' | '', where: { sheet, cell, rows }, items, leftovers, problems: [人が読む文] }
   */
  function readSummaryBook(book, opts) {
    opts = opts || {};
    const problems = [], found = [];
    const sheets = (book && Array.isArray(book.sheets)) ? book.sheets : [];
    if (!sheets.length) problems.push('ブックにシートがありません');
    sheets.forEach((sh) => {
      if (sh.state && sh.state !== 'visible') return;
      const pa = sh.printArea || [1, 1, sh.maxRow || 1e6, sh.maxCol || 1e4];
      const hiddenR = new Set(sh.hiddenRows || []), hiddenC = new Set(sh.hiddenCols || []);
      const inArea = (r, c) => r >= pa[0] && r <= pa[2] && c >= pa[1] && c <= pa[3] && !hiddenR.has(r) && !hiddenC.has(c);
      const byRow = new Map();
      Object.keys(sh.cells || {}).forEach((k) => {
        const i = k.indexOf(','), r = Number(k.slice(0, i)), c = Number(k.slice(i + 1));
        const cell = sh.cells[k]; if (!cell || isBlank(cell.v) || cell.e || !inArea(r, c)) return;
        if (!byRow.has(r)) byRow.set(r, []);
        byRow.get(r).push([c, cell.v]);
      });
      byRow.forEach((cs) => cs.sort((x, y) => x[0] - y[0]));
      byRow.forEach((cs, r) => cs.forEach(([c, v]) => {
        if (typeof v !== 'string' || !RE_HEAD.test(nk(v).replace(/\s/g, ''))) return;
        found.push({ sh, r, c, v, byRow, merge: (sh.merges || []).find((x) => x[0] === r && x[1] === c) || null, pa });
      }));
    });
    found.sort((x, y) => sheets.indexOf(x.sh) - sheets.indexOf(y.sh) || x.r - y.r || x.c - y.c);
    const out = { month: null, period: '', where: null, items: [], leftovers: [], problems };
    if (!found.length) { if (sheets.length) problems.push('「月間厳選特」の見出しが見つかりません（見えるシートの、印刷範囲の中・隠し行でない所）'); return out; }
    if (found.length > 1) problems.push('「月間厳選特」の見出しが ' + found.length + ' つあります（' + found.map((f) => f.sh.name + '!' + colName(f.c) + f.r).join('・') + '）。いちばん上の物だけ読みました');
    const h = found[0];
    // 読む行: 見出しのつないだセルの行。つないでいなければ、見出しの列が空の行が続く所まで
    const lastCol = h.merge ? h.merge[3] : h.c;
    const rows = [];
    if (h.merge) { for (let r = h.merge[0]; r <= h.merge[2]; r++) rows.push(r); } else {
      rows.push(h.r);
      for (let r = h.r + 1; r <= h.pa[2] && rows.length < 10; r++) {
        const cs = h.byRow.get(r) || [];
        if (!cs.length || cs.some(([c]) => c <= lastCol)) break;
        rows.push(r);
      }
    }
    const texts = [];
    rows.forEach((r) => {
      const cs = (h.byRow.get(r) || []).filter(([c]) => c > lastCol);
      const t = cs.map(([, v]) => (typeof v === 'number' ? (Number.isFinite(v) ? v + '円' : '') : String(v))).filter((x) => !isBlank(x)).join('　　');
      if (!isBlank(t)) texts.push(t);
    });
    out.where = { sheet: h.sh.name, cell: colName(h.c) + h.r, rows: rows.filter((r) => h.byRow.has(r)) };
    const part = splitText(texts.join('　　'));
    out.items = part.items; out.leftovers = part.leftovers;
    if (!texts.length) problems.push('「月間厳選特」の行に、商品の文がありません（' + out.where.sheet + '!' + out.where.cell + '）');
    else if (!part.items.length) problems.push('「月間厳選特」の行から、価格（〇〇円）の付いた商品を読めませんでした');
    if (part.leftovers.length) problems.push('価格の付いていない文が ' + part.leftovers.length + ' つあります（読めなかった文）');
    // 月: 見出しの期間と、年（見出し → ファイルの名前 → opts.year）
    const p = periodOf(h.v), fm = fileMonthOf(opts.fileName);
    if (!p) problems.push('見出しから期間（月/日～月/日）を読めません（' + out.where.cell + '）');
    else {
      out.period = p.text;
      if (p.m1 !== p.m2) problems.push('期間が2つの月にまたがっています（' + p.text + '）。はじめの月にしました');
      const year = p.year || (fm && fm.year) || (Number.isInteger(opts.year) ? opts.year : null);
      if (!year) problems.push('年が分かりません（ファイルの名前にも見出しにも年が無い）');
      else out.month = year + '-' + String(p.m1).padStart(2, '0');
      if (fm && fm.month && fm.month !== p.m1) problems.push('ファイルの名前の月（' + fm.month + '月）と、見出しの月（' + p.m1 + '月）が違います');
    }
    return out;
  }

  // ---------------------------------------------------------------- 名前を比べる
  const rowsOf = (x) => {
    if (!x) return [];
    if (Array.isArray(x)) return x;
    if (typeof Map !== 'undefined' && x instanceof Map) return Array.from(x.values());
    if (typeof Set !== 'undefined' && x instanceof Set) return Array.from(x);
    if (Array.isArray(x.c) && Array.isArray(x.r)) return x.r.map((r) => { const o = {}; x.c.forEach((c, i) => { o[c] = r[i]; }); return o; });
    return [];
  };
  const RE_BRACKET = /[(【\[]([^()【】\[\]]*)[)】\]]/g;
  const RE_QTY_PRICE = new RegExp('\\d\\s*(?:' + UNIT.slice(3, -1) + '|円)');
  /** 名前から、入数・価格の括弧（（６本入）（２本299円））と、各種・％増量を外す（NFKC の字） */
  function stripName(v) {
    return nk(v).replace(/\s+/g, ' ').replace(RE_BRACKET, (all, inner) => (RE_QTY_PRICE.test(inner) ? ' ' : all))
      .replace(/各種/g, '').replace(/\d+(?:\.\d+)?\s*%\s*増量/g, '').replace(/\+\s*\d+\s*\S{0,3}増量/g, '').trim();
  }
  /** 入数（名前の括弧や終わりの規格から）。無ければ '' */
  function qtyOfName(v) {
    const t = nk(v);
    let q = '';
    t.replace(RE_BRACKET, (all, inner) => { if (!q) { const m = RE_QTY_IN.exec(inner); if (m) q = m[0]; } return all; });
    if (!q) { const m = RE_SPEC_END.exec(t); if (m) q = m[1]; }
    return q ? specKey(q.replace(/\d+\s*円.*$/, '')) : '';
  }
  /** 別の書き方（読みの表の alias）から、直した字の置き換えの表。words: { alias: [[語, 別の書き方…]] } か、読みの表の行 { 種類: 'alias', 語, 値: 'a/b' } */
  function aliasTable(words, fold) {
    const groups = [];
    rowsOf(words && words.alias ? words.alias : words).forEach((g) => {
      if (Array.isArray(g)) groups.push(g);
      else if (g && g['種類'] === 'alias' && !isBlank(g['語'])) groups.push([g['語']].concat(String(g['値'] || '').split('/')));
    });
    const pairs = [];
    groups.forEach((g) => {
      const to = fold(g[0]); if (!to) return;
      g.slice(1).forEach((w) => { const f = fold(w); if (f && f !== to && f.length >= 2) pairs.push([f, to]); });
    });
    pairs.sort((x, y) => y[0].length - x[0].length);
    return pairs;
  }
  function grams(f) { const g = new Map(); if (f.length === 1) g.set(f, 1); for (let i = 0; i + 1 < f.length; i++) { const x = f.substr(i, 2); g.set(x, (g.get(x) || 0) + 1); } return g; }
  // 名前ごとの比べる形は、同じ直し方（SearchNorm の有無と表の数・別の書き方の表）の間は覚えて使い回す（商品の名前は呼ぶたびに同じ。2回目からは作らない）
  const KEY_CACHE = { gen: '', map: new Map() };
  /** 名前の比べ方の道具。名前ごとに: full（入数の括弧を外して直した字）・core（括弧の中も外す）・loose（濁点も外す）・qty（入数）。
   *  漢字の読み（yomi）は、名前が近いときだけ作る（SearchNorm の読みの表を引くのが重いので。yomiOf） */
  function makeKeyer(opts) {
    const SN = searchNorm(), fold = SN ? SN.fold : foldLocal;
    const pairs = aliasTable(opts && opts.words, fold);
    const canon = (f) => { for (let i = 0; i < pairs.length; i++) if (f.indexOf(pairs[i][0]) >= 0) f = f.split(pairs[i][0]).join(pairs[i][1]); return f; };
    const count = SN && typeof SN.count === 'function' ? SN.count() : null;
    const hasYomi = Boolean(SN && typeof SN.maps === 'function' && count && count.yomi);
    const gen = (SN ? 'sn' + JSON.stringify(count) : 'local') + '|' + pairs.map((p) => p.join('=')).join(',');
    if (KEY_CACHE.gen !== gen || KEY_CACHE.map.size > 20000) { KEY_CACHE.gen = gen; KEY_CACHE.map = new Map(); }
    const cache = KEY_CACHE.map;
    const key = (v) => {
      const s = String(v === null || v === undefined ? '' : v);
      let k = cache.get(s);
      if (k) return k;
      const t = stripName(s);
      const full = canon(fold(t)), core = canon(fold(t.replace(RE_BRACKET, ' ')));
      k = { t, full, core, loose: loosen(full), yomi: undefined, qty: qtyOfName(s), grams: grams(full), cgrams: grams(core) };
      cache.set(s, k);
      return k;
    };
    /** 漢字を読みにした字（表に無ければ、直した字と同じなので ''）。1度作ったら覚える */
    const yomiOf = (k) => {
      if (k.yomi !== undefined) return k.yomi;
      let y = '';
      if (hasYomi && /[㐀-䶿一-鿿々〆]/.test(k.t)) { try { const f = SN.maps(k.t); y = f.length > 2 ? canon(f[2].k) : ''; } catch (e) { y = ''; } }
      k.yomi = y && y !== k.full ? y : '';
      return k.yomi;
    };
    return { key, yomiOf, fold: (v) => canon(fold(v)) };
  }
  function dice(ga, gb) {
    let na = 0, nb = 0, both = 0;
    ga.forEach((n) => { na += n; }); gb.forEach((n) => { nb += n; });
    if (!na || !nb) return 0;
    ga.forEach((n, k) => { const m = gb.get(k); if (m) both += Math.min(n, m); });
    return (2 * both) / (na + nb);
  }
  function levRatio(a, b) {
    if (a === b) return 1;
    const n = a.length, m = b.length; if (!n || !m) return 0;
    let prev = new Array(m + 1), cur = new Array(m + 1);
    for (let j = 0; j <= m; j++) prev[j] = j;
    for (let i = 1; i <= n; i++) {
      cur[0] = i;
      for (let j = 1; j <= m; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a.charCodeAt(i - 1) === b.charCodeAt(j - 1) ? 0 : 1));
      const t = prev; prev = cur; cur = t;
    }
    return 1 - prev[m] / Math.max(n, m);
  }
  function isSubseq(a, b) { let j = 0; for (let i = 0; i < b.length && j < a.length; i++) if (a.charAt(j) === b.charAt(i)) j++; return j === a.length; }
  const round1 = (x) => Math.round(x * 10) / 10;

  /** 2つの名前の点数（0〜100）と理由。a・b は名前ごとの比べる形。v＝種類に分けた名前のときの { base, part }（直した字）。yomiOf＝読みを作る道具（無ければ読みを使わない） */
  function scoreKeys(a, b, v, yomiOf) {
    if (!a.full || !b.full) return { score: 0, why: '' };
    if (a.full === b.full) return { score: 100, why: '名前が同じ' };
    let best = 0, why = '';
    const take = (s, w) => { if (s > best) { best = s; why = w; } };
    if (a.loose === b.loose) take(94, '濁点の違いだけ');
    if (a.core && a.core === b.core && a.core.length >= 3) take(92, '括弧の外が同じ');
    if (v && v.part && b.full.indexOf(v.part) >= 0 && (!v.base || b.full.indexOf(v.base) >= 0)) take(v.base ? 88 : 70 + 15 * Math.min(1, v.part.length / b.full.length), v.base ? '元の名前と種類の語が入っている' : '種類の語が入っている');
    const d = Math.max(dice(a.grams, b.grams), 0.97 * dice(a.cgrams, b.cgrams));
    if (d >= 0.35) { // 字の組が3分の1より多く同じときだけ、細かく比べる（読みもここで作る）
      let l = Math.max(levRatio(a.full, b.full), 0.97 * levRatio(a.core, b.core), levRatio(a.loose, b.loose));
      const ya = yomiOf ? yomiOf(a) : '', yb = yomiOf ? yomiOf(b) : '';
      if (ya || yb) {
        const ra = ya || a.full, rb = yb || b.full;
        if (ra === rb) take(93, '読みが同じ');
        l = Math.max(l, 0.97 * levRatio(ra, rb));
      }
      take(90 * Math.max(d, l), 'よく似ている');
    }
    // 片方の字が、もう片方の中に順に入っている。サマリーは名前を縮めて書く（国産小麦の を落とす など）ので、a（読んだ名前）が短い向きを高く。
    //   b（商品）の方が短いときは、読んだ物が別の商品（その商品の名前を含む、新しい商品）のこともあるので低く
    if (a.full.length <= b.full.length) {
      if (a.full.length >= 4 && isSubseq(a.full, b.full)) take(72 + 18 * (a.full.length / b.full.length), b.full.indexOf(a.full) >= 0 ? '名前の一部が同じ' : '間に字が足されているだけ');
      else if (a.core.length >= 4 && b.full.indexOf(a.core) >= 0) take(68 + 18 * (a.core.length / b.full.length), '名前の一部が同じ');
    } else if (b.full.length >= 4 && isSubseq(b.full, a.full)) take(68 + 18 * (b.full.length / a.full.length), '商品の名前が、読んだ名前の中にある');
    else if (b.core.length >= 4 && a.full.indexOf(b.core) >= 0) take(64 + 18 * (b.core.length / a.full.length), '商品の名前の一部が、読んだ名前の中にある');
    return { score: round1(best), why };
  }

  // ---------------------------------------------------------------- 商品の候補を付ける
  const SURE_MIN = 85, SURE_GAP = 8, HINT_MIN = 88;
  /**
   * 読んだ商品に、マスターの商品の候補を付ける。products: 商品の行（配列・Map・{c, r} の表。商品ID・名前・呼出番・表の種類）。
   * opts: { monthlyRows: 前の月の月間商品の行（同じ名前の行の 商品ID・呼出番 を手がかりにする）, words: 読みの表（{ alias } か行。別の書き方をそろえる）, limit: 候補の数（3） }
   * 返す（items と同じ並び）: [{ kind: 'one'|'family', candidates: [{ 商品ID, 名前, 呼出番, 表の種類, 点数, 理由 }], best, sure, why, parts: [{ name, candidates, best, sure, why }] | null }]
   *   確か（sure）＝1位が 85 点以上で、2位と 8 点以上の差。「各種」で種類の分からない物と、％増量の物（アウトパックが多い）は、確かにしない。
   *   family＝名前に種類が並んだ物（A（x・y）各種・A・B）。種類ごとに候補を付ける（parts）。
   */
  function matchItems(items, products, opts) {
    opts = opts || {};
    const limit = opts.limit || 3, K = makeKeyer(opts);
    const prods = rowsOf(products).filter((p) => p && !isBlank(p['商品ID']) && !isBlank(p['名前'])).map((p) => ({ p, k: K.key(p['名前']), call: nk(p['呼出番']).trim() }));
    const byId = new Map(prods.map((x) => [String(x.p['商品ID']), x]));
    const byCall = new Map(); prods.forEach((x) => { if (/^\d{4}$/.test(x.call)) { if (!byCall.has(x.call)) byCall.set(x.call, []); byCall.get(x.call).push(x); } });
    const prev = rowsOf(opts.monthlyRows).filter((r) => r && (isBlank(r['行の種類']) || r['行の種類'] === '商品') && !isBlank(r['商品名'])).map((r) => ({ r, k: K.key(r['商品名']) }));

    function rank(name, v, specs) {
      const k = K.key(name);
      const vk = v ? { base: v.base ? K.key(v.base).core : '', part: K.fold(stripName(v.part)) } : null;
      const qty = specs && specs[0] ? specKey(specs[0]['規格']) : '';
      const score = new Map();
      prods.forEach((x) => {
        const s = scoreKeys(k, x.k, vk, K.yomiOf);
        if (s.score < 40) return;
        let sc = s.score; const why = [s.why];
        if (qty && x.k.qty && qty === x.k.qty) { sc += 4; why.push('入数が同じ'); }
        score.set(x, { sc, why });
      });
      // 前の月の月間商品の、同じ名前の行
      let hint = null;
      prev.forEach((y) => { const s = scoreKeys(k, y.k, null, K.yomiOf).score; if (s >= HINT_MIN && (!hint || s > hint.s)) hint = { s, r: y.r }; });
      if (hint) {
        const id = String(hint.r['商品ID'] || ''), call = nk(hint.r['呼出番']).trim();
        const x = byId.get(id);
        if (x) { const o = score.get(x) || { sc: 0, why: [] }; if (o.sc < 96) o.sc = 96; o.why.push('前の月の月間商品に同じ名前の行（商品ID）'); score.set(x, o); }
        (byCall.get(call) || []).forEach((y) => { if (y === x) return; const o = score.get(y) || { sc: 50, why: ['名前は離れている'] }; o.sc += 6; o.why.push('前の月の月間商品の呼出番が同じ'); score.set(y, o); });
      }
      const list = Array.from(score.entries()).map(([x, o]) => ({ '商品ID': String(x.p['商品ID']), '名前': x.p['名前'], '呼出番': x.call, '表の種類': x.p['表の種類'] || '', '点数': round1(Math.min(100, o.sc)), '理由': o.why.filter(Boolean).join('・') }))
        .sort((a, b) => b['点数'] - a['点数'] || (a['商品ID'] < b['商品ID'] ? -1 : 1));
      const top = list[0] || null, second = list[1] || null;
      const sure = Boolean(top && top['点数'] >= SURE_MIN && (!second || top['点数'] - second['点数'] >= SURE_GAP));
      const why = !top ? '当たる商品がありません' : (sure ? '' : (top['点数'] < SURE_MIN ? '名前が離れています' : '候補が近いです'));
      return { candidates: list.slice(0, limit), best: top, sure, why };
    }

    return (items || []).map((it) => {
      const whole = rank(it.name, null, it.specs);
      // ％増量の物はアウトパック（番号表に無い物）が多いので、名前が同じでも確かにしない
      if (it['増量']) return Object.assign({ kind: 'one', parts: null }, whole, { sure: false, why: '％増量の物（アウトパックが多い）なので、番号表の商品と同じか確かめてください' });
      const one = () => (it['各種'] && !it.variants ? Object.assign({ kind: 'one', parts: null }, whole, { sure: false, why: '「各種」なので、当てる商品を選んでください' }) : Object.assign({ kind: 'one', parts: null }, whole));
      if (!it.variants) return one();
      const parts = it.variants.map((v) => Object.assign({ name: v.name }, rank(v.name, v, it.specs)));
      const avg = parts.reduce((t, x) => t + (x.best ? x.best['点数'] : 0), 0) / parts.length;
      const wholeTop = whole.best ? whole.best['点数'] : 0;
      if (!it['各種'] && !it['各'] && wholeTop >= 90 && whole.sure) return one();
      if (it['各種'] || it['各'] || avg > wholeTop) {
        const sure = parts.every((x) => x.sure);
        return { kind: 'family', candidates: whole.candidates, best: whole.best, sure, why: sure ? '' : '種類のうち、確かでない物があります', parts };
      }
      return one();
    });
  }

  // ---------------------------------------------------------------- 月間商品の行と、番号表の変更の案
  const MONTHLY_COLUMNS = ['行ID', '部', '順', '行の種類', '商品名', '呼出番', '規格', '本体価格', '小パックの規格', '小パックの本体価格', '商品ID', '見た目'];
  const SECTION_DEFAULT = { '惣菜': '≪惣　菜≫', '米飯': '≪米　飯≫', 'アウトパック': 'アウトパック商品' };
  const SECTION_ORDER = ['惣菜', '米飯', 'アウトパック'];
  /** 節の見出しの文から、節の種類（惣菜・米飯・アウトパック。どれでもなければ見出しの文） */
  function sectionKeyOf(title) { const t = nk(title).replace(/\s/g, ''); return /惣菜/.test(t) ? '惣菜' : /米飯/.test(t) ? '米飯' : /アウトパック/.test(t) ? 'アウトパック' : t; }
  /** 番号表の名前から、月間商品に書く名前（入数・価格の括弧と、余分な空白を外す。元の字のまま） */
  function displayName(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[（(]([^（）()]*)[）)]/g, (all, inner) => (RE_QTY_PRICE.test(nk(inner)) ? '' : all)).replace(/[\s　]+/g, ' ').trim();
  }

  /**
   * 月間商品の「月間特売」の部の行と、番号表の行に当てる変更の案を作る。
   * opts: { products（商品の行）, monthlyRows（前の月の月間商品の行: 節の見出し・置き場・小パック）, words（読みの表）, idPrefix（'G'）, firstNo（1）, usedIds（使っている行ID）, carrySmallPack（true） }
   * 返す: { rows: [MONTHLY_COLUMNS の形], changes: [{ 商品ID, 販売形態: '月間特売', 特売価格 }], toPick: [{ 行ID, 名前, 文, 規格, 本体価格, 理由, 候補 }], notes: [文] }
   *   特売価格＝1つ目の規格の価格。変更の案に入れるのは確かな候補だけ（同じ商品に2つの価格があれば、案から外して「人が選ぶ」へ）。
   *   2つ目からの規格は「続きの行」。小パックは、前の月の同じ商品の行で、特売の価格が同じときだけ写す（notes に出す）。
   *   節: 前の月の同じ商品・同じ名前の行の節 → 商品の表の種類（焼き物＝惣菜・寿司弁当＝米飯）→ ％増量の物はアウトパック → 前の商品の節。
   *   「各種」は1行のまま（種類ごとの商品は変更の案だけ）。種類を並べた物（A・B／A（x・y）各１パック）は、種類ごとに1行。
   *   ％増量の物（アウトパック）は、呼出番の欄に「30%増量」と書く（元の月間商品と同じ書き方）。
   */
  function toMonthlyRows(items, matches, opts) {
    opts = opts || {};
    const K = makeKeyer(opts);
    const prods = new Map(rowsOf(opts.products).map((p) => [String(p['商品ID']), p]));
    const prevAll = rowsOf(opts.monthlyRows).filter((r) => r && (isBlank(r['部']) || r['部'] === '月間特売'));
    // 前の月の節（見出しの文と、その節の行）
    const titles = {}, order = [], prevRows = [];
    let sec = '';
    prevAll.forEach((r) => {
      if (r['行の種類'] === '節の見出し') { sec = sectionKeyOf(r['商品名']); if (!titles[sec]) { titles[sec] = r['商品名']; order.push(sec); } return; }
      if (r['行の種類'] === '商品' && !isBlank(r['商品名'])) prevRows.push({ r, sec, k: K.key(r['商品名']) });
    });
    SECTION_ORDER.forEach((s) => { if (order.indexOf(s) < 0) { order.push(s); titles[s] = SECTION_DEFAULT[s]; } });
    const prevById = new Map(); prevRows.forEach((y) => { const id = String(y.r['商品ID'] || ''); if (id && !prevById.has(id)) prevById.set(id, y); });
    const prevByName = (name) => { const k = K.key(name); let best = null; prevRows.forEach((y) => { const s = scoreKeys(k, y.k, null, K.yomiOf).score; if (s >= HINT_MIN && (!best || s > best.s)) best = { s, y }; }); return best ? best.y : null; };
    const kindSec = (p) => (p && p['表の種類'] === '焼き物' ? '惣菜' : p && p['表の種類'] === '寿司弁当' ? '米飯' : '');

    const used = new Set(rowsOf(opts.usedIds).map(String));
    const prefix = opts.idPrefix || 'G';
    let no = opts.firstNo || 1;
    const nextId = () => { let id; do { id = prefix + String(no++).padStart(3, '0'); } while (used.has(id)); used.add(id); return id; };
    const carry = opts.carrySmallPack !== false;
    const notes = [], toPick = [], picked = new Map(); // 商品ID → { price, conflict }
    const blank = (kind) => ({ '行ID': '', '部': '月間特売', '順': 0, '行の種類': kind, '商品名': '', '呼出番': '', '規格': '', '本体価格': '', '小パックの規格': '', '小パックの本体価格': '', '商品ID': '', '見た目': '' });
    const groups = new Map(); // 節 → 行の並び（行ID・順は最後に付ける）
    const put = (s, rows) => { if (!groups.has(s)) groups.set(s, []); groups.get(s).push(...rows); };
    const contRows = (specs) => specs.map((x) => Object.assign(blank('続きの行'), { '規格': x['規格'], '本体価格': x['本体価格'] }));
    const pick = (row, it, name, m, why) => toPick.push({ row, '名前': name, '文': it.text, '規格': it.specs[0]['規格'], '本体価格': it.specs[0]['本体価格'], '理由': why, '候補': m ? m.candidates : [] });
    const sure = (row, it, x, pid) => {
      const p = prods.get(pid) || {};
      row['商品名'] = displayName(p['名前'] || x['名前']); row['呼出番'] = nk(p['呼出番'] || x['呼出番']).trim(); row['商品ID'] = pid;
      const y = prevById.get(pid);
      if (carry && y && !isBlank(y.r['小パックの規格']) && String(y.r['本体価格']) === String(it.specs[0]['本体価格'])) {
        row['小パックの規格'] = y.r['小パックの規格']; row['小パックの本体価格'] = y.r['小パックの本体価格'];
        notes.push(row['商品名'] + ': 小パックは前の月の行から写しました（特売の価格が前の月と同じ）');
      }
      const had = picked.get(pid), price = it.specs[0]['本体価格'];
      if (had && had.price !== price) { had.conflict = true; pick(row, it, row['商品名'], null, '同じ商品に、2つの価格があります（' + had.price + '円・' + price + '円）'); }
      else if (!had) picked.set(pid, { price, conflict: false });
      const q = p['対の商品ID'];
      if (!isBlank(q)) notes.push(row['商品名'] + ': 対の商品（' + q + '）にも月間特売を付けるか、確かめてください');
    };
    const pidOf = (x) => (x && x.sure && x.best ? x.best['商品ID'] : '');
    let lastSec = '';
    (items || []).forEach((it, i) => {
      const m = (matches || [])[i] || null;
      const firstPid = m ? (pidOf(m) || (m.parts || []).map(pidOf).find(Boolean) || '') : '';
      const py = (firstPid && prevById.get(firstPid)) || prevByName(it.name);
      let s = py && py.sec ? py.sec : '';
      if (!s && firstPid) s = kindSec(prods.get(firstPid));
      if (!s && it['増量']) s = 'アウトパック';
      if (!s && m) { // 確かでなくても、近い候補（70点から）の表の種類
        const near = [m.best].concat((m.parts || []).map((x) => x.best)).filter((x) => x && x['点数'] >= 70).sort((x, y) => y['点数'] - x['点数'])[0];
        if (near) s = kindSec(prods.get(near['商品ID']));
      }
      if (!s) s = lastSec || '惣菜';
      lastSec = s;
      const more = contRows(it.specs.slice(1));
      const first = { '規格': it.specs[0]['規格'], '本体価格': it.specs[0]['本体価格'] };
      if (m && m.kind === 'family' && !it['各種']) { // 種類ごとに1行（2つ目からの規格は、最後の行の下に。「各１パック」の各は外す）
        const rows = m.parts.map((x) => {
          const row = Object.assign(blank('商品'), first, { '商品名': x.name, '規格': String(first['規格']).replace(/^各(?=\d)/, '') });
          if (pidOf(x)) sure(row, it, x.best, pidOf(x)); else pick(row, it, x.name, x, x.why || '確かめてください');
          return row;
        });
        put(s, rows.concat(more));
        return;
      }
      const head = Object.assign(blank('商品'), first, { '商品名': it.name });
      if (m && m.kind === 'family') { // 「各種」は1行のまま。種類ごとの商品は、変更の案だけ
        m.parts.forEach((x) => { if (pidOf(x)) sure(blank('商品'), it, x.best, pidOf(x)); else pick(head, it, x.name, x, x.why || '確かめてください'); });
      } else if (pidOf(m)) sure(head, it, m.best, pidOf(m));
      else {
        if (it['増量']) head['呼出番'] = it['増量']; // アウトパックは呼出番の欄に「30%増量」（元の月間商品と同じ書き方。商品を選んだら呼出番に替わる）
        pick(head, it, it.name, m, m && m.best ? (m.why || '確かめてください') : '番号表の商品に当たりません（アウトパックなど番号表に無い物なら、そのままでよい）');
      }
      put(s, [head].concat(more));
    });
    // 節の順に並べ、見出し・空き行・行ID・順を付ける
    const rows = [];
    const secs = order.filter((x) => groups.has(x)).concat(Array.from(groups.keys()).filter((x) => order.indexOf(x) < 0));
    secs.forEach((x, n) => {
      if (n) rows.push(blank('空き行'));
      rows.push(Object.assign(blank('節の見出し'), { '商品名': titles[x] || x }));
      groups.get(x).forEach((r) => rows.push(r));
    });
    rows.forEach((r, n) => { r['行ID'] = nextId(); r['順'] = (n + 1) * 10; });
    const changes = [];
    picked.forEach((v, pid) => { if (!v.conflict) changes.push({ '商品ID': pid, '販売形態': '月間特売', '特売価格': v.price }); });
    return { rows, changes, toPick: toPick.map((x) => ({ '行ID': x.row['行ID'], '名前': x['名前'], '文': x['文'], '規格': x['規格'], '本体価格': x['本体価格'], '理由': x['理由'], '候補': x['候補'] })), notes };
  }

  root.SouzaiImportSummary = {
    readSummaryBook, matchItems, toMonthlyRows, splitText, variantsOf, periodOf, fileMonthOf, normSpec, specKey, stripName, displayName, sectionKeyOf, MONTHLY_COLUMNS, SURE_MIN, SURE_GAP,
    /** 2つの名前の点数と理由（試験・画面の説明用） */
    nameScore: (a, b, opts) => { const K = makeKeyer(opts); return scoreKeys(K.key(a), K.key(b), null, K.yomiOf); }
  };
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
