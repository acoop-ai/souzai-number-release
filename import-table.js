// 手作りの「商品呼出番号表」のブック（Book）から、その月の表の束（月の束）を作る部品（画面の部品を持たない）。ブラウザと Node の両方で動く。
//   受け取る Book の形: frontend/book-read.js が読んだ形（tools/dump_book.py の先頭の説明と同じ）。
//     あれば使う（無くても読める）: セルの x＝書式の残り { bl, br, bt, bb: 罫線の種類, h, v: 横・縦の位置, w: 折り返し 1, k: 縮小 1, n: 表示形式 }・
//     書式だけのセル（値も塗りも無い）も載せる・シートの x0＝一度も触れていないセルの書式 { font, fill, x }。これがあると見た目（書式キー）まで元と同じになる。
//   作る表: 商品・原価・区分・月の状態・月の行_<月>・月の置き場_<月>・月間商品_<月>・疑問。data/種_<月>.json と同じ列・同じ形（{c:[列], r:[[値]]}）・同じ ID の決め方。
//   決まりは Python の道具と同じ（読む所＝tools/souzai_read.py・souzai_common.py／マスターにする所＝import_master.py／種にする所＝make_seed.py）。
//   信頼境界: ブックは BY が手で作った物（社内・ログインした人がアップロード）。中身は文字として運ぶだけで、式は計算しない・実行しない。外へは何も送らない。外の部品は使わない。
(function (root) {
  'use strict';

  // ================================================================ 決まった名前（Python の道具と同じ・変わりうる文字はここだけ）
  const TABLE_KEYS = ['首都圏', '北関東', '焼き物'];               // 番号表の3枚（この順で読む）
  const ORDER_FOR_ID = ['首都圏', '焼き物', '北関東'];             // 商品ID を振る順
  // 表の種類（商品を見分ける鍵の1つ目。首都圏と北関東は同じ種類＝同じ商品を1行にまとめる）
  const TABLE_KIND = { '首都圏': '寿司弁当', '北関東': '寿司弁当', '焼き物': '焼き物' };
  // 見えているシートの中から、名前（NFKC 後）で1枚に決める
  const SHEET_RULE = { '首都圏': (n) => n.indexOf('(首都圏)') >= 0, '北関東': (n) => n.indexOf('(北関東)') >= 0, '焼き物': (n) => n.indexOf('焼き物') === 0 };
  const MONTHLY_SHEET = '月間商品', PROCESSING_SHEET = '加工基本表';
  const SECTION_MARKS = ['≪', '＜', '<', '《'];
  const CONT_MARKS = ['(', '（', '※'];
  const PINK_FONTS = ['FF00FF', 'FF3399', 'FF33CC'];              // 名前の字がピンク＝月間特売（焼き物の表の凡例）
  const KIND_PRODUCT = '商品', KIND_BLANK = '空き行', KIND_CONT = '続きの行', KIND_SECTION = '節の見出し';
  const ROLE_JA = { name: '名前', tray: '指定トレー', tray_change: '変更トレー', call: '呼出番', cost: '原価', gross_before: '粗利（包材値上前）', price_before: '定番価格（包材値上前）', gross_after: '粗利（包材変更後）', price_calc: '定番価格（包材変更後）', price_visible: '見える価格', small: '小パック' };
  const MONTHLY_ROLE_JA = { name: '商品名', call: '呼出番', size: '規格', price: '本体価格', tax: '税込価格', s_size: '小パックの規格', s_price: '小パックの本体価格', s_tax: '小パックの税込価格' };
  const KIND_STYLE_COL = { product: '書式（商品）', section: '書式（区分の表題）', blank: '書式（空き）' };
  const LEVEL_STOP = '止める', LEVEL_CHECK = '確かめる';
  const EMPTY_HEADER = '（空）';
  const COST_COLS = ['原価', '粗利（包材値上前）', '定番価格（包材値上前）', '定番価格（包材変更後）'];

  // 見出しの並び（指紋）の、確かめた形。tools/known_layouts.json の写し（番号表と月間商品だけ。見出しの名前だけで、商品の中身は入っていない）。
  //   形を足すときは、元の書類を確かめてから tools/layouts.py で登録し、ここにも同じ物を写す（試験 tests/test_import_table.js が同じかを見る）。
  const H10 = ['指定トレー', '呼出番', '原価', '粗利', '定番価格', '粗利', '定番価格', '定番価格', '小パック'];
  const H89 = ['指定トレー', '変更トレー', '呼出番', '原価', '粗利', '定番価格', '粗利', '定番価格', '定番価格', '小パック'];
  const swapFirst = (a, v) => [v].concat(a.slice(1));
  const sushiForm = (h) => [['≪生寿司≫'].concat(h), ['≪細巻寿司≫'].concat(h), ['≪和菓子≫'].concat(swapFirst(h, '旧トレー')), ['≪弁当(おこわ・ご飯商材)≫'].concat(h), ['≪丼ぶり(肉商材)≫'].concat(h.map((x) => (x === '指定トレー' || x === '変更トレー' ? EMPTY_HEADER : x)))];
  const yakiForm = (h) => [['≪焼き鳥・焼き物≫'].concat(h), ['≪天ぷら・バラ≫'].concat(h), ['≪フライ・唐揚≫'].concat(h), ['≪揚げ物バラ定番品≫'].concat(h.slice(0, -1)), ['≪均一セール≫', '呼出番', '原価', '粗利', '定番価格', '定番価格']];
  const NAME10 = '2026-10の形（変更トレーの列なし）', NAME89 = '2026-08〜09の形（変更トレーの列あり）';
  const MONTHLY_FORM = [['商品名', '呼出番', '規格', '本体価格', '税込価格', '規格', '小パック', EMPTY_HEADER, EMPTY_HEADER], [EMPTY_HEADER, EMPTY_HEADER, EMPTY_HEADER, EMPTY_HEADER, EMPTY_HEADER, EMPTY_HEADER, '本体価格', '税込価格', EMPTY_HEADER]];
  const LAYOUTS = {
    '番号表': {
      '首都圏': [{ '名前': NAME10, '指紋': sushiForm(H10) }, { '名前': NAME89, '指紋': sushiForm(H89) }],
      '北関東': [{ '名前': NAME10, '指紋': sushiForm(H10) }, { '名前': NAME89, '指紋': sushiForm(H89) }],
      '焼き物': [{ '名前': NAME10, '指紋': yakiForm(H10) }, { '名前': NAME89, '指紋': yakiForm(H89) }]
    },
    '月間商品': { '月間特売': [{ '名前': '2026-08〜10の形', '指紋': MONTHLY_FORM }], '新規商品': [{ '名前': '2026-08〜10の形', '指紋': MONTHLY_FORM }] }
  };

  // ================================================================ 文字と値の決まり（Python と同じ動き）
  // Python の空白（str.strip・正規表現の \s）と同じ字の組
  const WS_CHARS = '\\t\\n\\v\\f\\r\\x1c-\\x1f \\x85\\xa0\\u1680\\u2000-\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000';
  const RE_STRIP = new RegExp('^[' + WS_CHARS + ']+|[' + WS_CHARS + ']+$', 'g');
  const RE_WS = new RegExp('[' + WS_CHARS + ']+', 'g');
  const pyStrip = (s) => String(s).replace(RE_STRIP, '');
  /** Python の str(数)（整数は整数の形・小数は短い形・とても小さい／大きい数は 1e-05 の形） */
  function pyNum(n) {
    if (Number.isNaN(n)) return 'nan';
    if (!Number.isFinite(n)) return n > 0 ? 'inf' : '-inf';
    if (Number.isInteger(n) && Math.abs(n) < 1e16) return String(n);
    const a = Math.abs(n);
    if (a !== 0 && (a < 1e-4 || a >= 1e16)) {
      const m = /^(-?[0-9.]+)e([+-])([0-9]+)$/.exec(n.toExponential());
      if (m) return m[1] + 'e' + m[2] + (m[3].length < 2 ? '0' + m[3] : m[3]);
    }
    return String(n);
  }
  /** Python の '%s' と同じ字（None・True・False も Python の書き方） */
  function pyStr(v) {
    if (v === null || v === undefined) return 'None';
    if (v === true) return 'True';
    if (v === false) return 'False';
    if (typeof v === 'number') return pyNum(v);
    return String(v);
  }
  /** NFKC にそろえ、改行を空白にし、前後の空白を外す（null は ''） */
  function nk(v) {
    if (v === null || v === undefined) return '';
    return pyStrip(pyStr(v).normalize('NFKC').replace(/[\r\n]/g, ' '));
  }
  /** 突き合わせ用: NFKC＋空白をすべて外す */
  const squash = (v) => nk(v).replace(RE_WS, '');
  const isBlank = (v) => v === null || v === undefined || (typeof v === 'string' && pyStrip(v) === '');
  const isFormula = (v) => typeof v === 'string' && v.charAt(0) === '=';
  const startsWithAny = (s, marks) => marks.some((m) => s.indexOf(m) === 0);
  const isSection = (name) => startsWithAny(nk(name), SECTION_MARKS);
  /** 呼出番を半角4桁にそろえる。空なら ''。4桁の数字でなければ null */
  function halfCall(v) {
    if (typeof v === 'number' && Number.isFinite(v) && Number.isInteger(v)) v = String(v);
    const s = nk(v);
    if (s === '') return '';
    return /^[0-9]{4}$/.test(s) ? s : null;
  }
  const RE_FLOAT = /^[+-]?(?:[0-9](?:_?[0-9])*(?:\.(?:[0-9](?:_?[0-9])*)?)?|\.[0-9](?:_?[0-9])*)(?:[eE][+-]?[0-9](?:_?[0-9])*)?$/;
  /** 数か、数として読める文字なら数。読めなければ null（Python の float() と同じ書き方を受ける） */
  function num(v) {
    if (typeof v === 'boolean') return null;
    if (typeof v === 'number') return v;
    const s = nk(v);
    if (s === '') return null;
    if (RE_FLOAT.test(s)) return Number(s.replace(/_/g, ''));
    const m = /^([+-]?)(inf|infinity|nan)$/i.exec(s);
    if (m) return m[2].toLowerCase() === 'nan' ? NaN : (m[1] === '-' ? -Infinity : Infinity);
    return null;
  }
  /** 小数は 10 けたで丸める（Python の round(v, 10)） */
  const cleanNum = (v) => (typeof v === 'number' && Number.isFinite(v) && !Number.isInteger(v) ? Number(v.toFixed(10)) : (v === undefined ? null : v));
  /** 数か数の文字なら数。それ以外はそのまま */
  function asNumber(v) {
    if (typeof v === 'string' && !isFormula(v) && num(v) !== null) return cleanNum(num(v));
    if (typeof v === 'number') return cleanNum(v);
    return v === undefined ? null : v;
  }
  /** Python の ==（null と undefined は同じ・真偽は 1 と 0） */
  function pyEq(a, b) {
    if (a === undefined) a = null;
    if (b === undefined) b = null;
    if (typeof a === 'boolean') a = a ? 1 : 0;
    if (typeof b === 'boolean') b = b ? 1 : 0;
    return a === b;
  }
  /** Python の「値 or 既定」（null・''・0・False・空の並び・空の辞書は既定） */
  function pyOr(v, d) {
    const falsy = v === null || v === undefined || v === '' || v === 0 || v === false || (Array.isArray(v) && !v.length) || (typeof v === 'object' && v !== null && !Array.isArray(v) && !Object.keys(v).length);
    return falsy ? d : v;
  }
  /** 中まで同じか（辞書は並びを見ない・並びは順も見る） */
  function deepEq(a, b) {
    if (a === null || a === undefined || b === null || b === undefined) return (a === null || a === undefined) && (b === null || b === undefined);
    if (Array.isArray(a) || Array.isArray(b)) return Array.isArray(a) && Array.isArray(b) && a.length === b.length && a.every((x, i) => deepEq(x, b[i]));
    if (typeof a === 'object' || typeof b === 'object') {
      if (typeof a !== 'object' || typeof b !== 'object') return false;
      const ka = Object.keys(a), kb = Object.keys(b);
      return ka.length === kb.length && ka.every((k) => Object.prototype.hasOwnProperty.call(b, k) && deepEq(a[k], b[k]));
    }
    return pyEq(a, b);
  }
  /** Python の json.dumps(…, ensure_ascii=False) と同じ字（区切りは「, 」「: 」・字の飾りの大きさ s は小数の形 16.0） */
  function pyJson(v, key) {
    if (v === null || v === undefined) return 'null';
    if (typeof v === 'boolean') return v ? 'true' : 'false';
    if (typeof v === 'number') return key === 's' && Number.isInteger(v) && Math.abs(v) < 1e16 ? v + '.0' : pyNum(v);
    if (typeof v === 'string') return JSON.stringify(v);
    if (Array.isArray(v)) return '[' + v.map((x) => pyJson(x, '')).join(', ') + ']';
    return '{' + Object.keys(v).map((k) => JSON.stringify(k) + ': ' + pyJson(v[k], k)).join(', ') + '}';
  }
  const nfkcStr = (v) => (v === null || v === undefined ? '' : pyStr(v)).normalize('NFKC');
  const RE_HYPHENS = new RegExp('[\\u2010-\\u2015\\u2212\\uFE58\\uFE63\\uFF0D]', 'g');   // ハイフンのゆれ
  const RE_LONG_BAR = new RegExp('\\u30FC', 'g');                                       // 長音「ー」
  const RE_KANA = new RegExp('[\\u3041-\\u3096\\u30A1-\\u30FA\\u30FC]');                 // かな（長音の前がかななら、長音のまま）
  /** 包材（指定トレー）の字をそろえる（frontend/core.js の normTray・tools/make_seed.py の norm_tray と同じ決まり。試験で同じかを見る） */
  function normTray(v) {
    return nfkcStr(v).replace(/[\r\n\t]/g, ' ').replace(RE_HYPHENS, '-')
      .replace(RE_LONG_BAR, (m, at, all) => (at > 0 && RE_KANA.test(all[at - 1]) ? m : '-')).trim();
  }
  const pad = (n, w) => { const s = String(n); return s.length >= w ? s : '0'.repeat(w - s.length) + s; };
  function colLetter(c) { let s = ''; let n = c; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  const range = (a, b) => { const out = []; for (let i = a; i <= b; i++) out.push(i); return out; };
  const hdrName = (v) => { const s = squash(v); return s ? s : EMPTY_HEADER; };
  const sameJson = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  /** いちばん多い値（同じ数なら先に出た物） */
  function mostCommon(list) { const m = new Map(); list.forEach((v) => m.set(v, (m.get(v) || 0) + 1)); let best, n = 0; m.forEach((c, v) => { if (c > n) { best = v; n = c; } }); return n ? best : undefined; }

  // ---------------------------------------------------------------- 名前から入数の括弧を外した鍵（突き合わせ用・souzai_common.split_name → name_key と同じ）
  const WS_CLASS = '[' + WS_CHARS + ']*';
  const RE_UNIT = new RegExp('[0-9]+(?:\\.[0-9]+)?' + WS_CLASS + '(?:個|貫|本|枚|切れ|切|尾|等分|袋|P|p|g|kg|人前|食|串|杯|パック)');
  const RE_TRAIL_BRACKET = new RegExp('\\(([^()]*)\\)' + WS_CLASS + '([^()]*)$');
  const RE_TRAIL_COUNT = new RegExp('(?<=\\))' + WS_CLASS + '([0-9]+(?:本|個|枚|貫))$');
  function baseName(raw) {
    const s = pyStrip(nk(raw).replace(RE_WS, ' '));
    let base = s;
    const m = RE_TRAIL_COUNT.exec(s);
    if (m) base = pyStrip(s.slice(0, m.index));
    else {
      const b = RE_TRAIL_BRACKET.exec(s);
      if (b && RE_UNIT.test(b[1])) base = pyStrip(s.slice(0, b.index));
    }
    return base ? base : s;
  }
  const nameKey = (raw) => squash(baseName(raw));

  // ---------------------------------------------------------------- 2つの名前の似ている度合い（Python の difflib.SequenceMatcher(None, a, b).ratio() と同じ）
  function similar(aStr, bStr) {
    const a = Array.from(aStr), b = Array.from(bStr);
    const la = a.length, lb = b.length;
    if (!la && !lb) return 1;
    const b2j = new Map();
    b.forEach((ch, j) => { if (!b2j.has(ch)) b2j.set(ch, []); b2j.get(ch).push(j); });
    if (lb >= 200) { const ntest = Math.floor(lb / 100) + 1; Array.from(b2j.keys()).forEach((ch) => { if (b2j.get(ch).length > ntest) b2j.delete(ch); }); }
    function longest(alo, ahi, blo, bhi) {
      let besti = alo, bestj = blo, bestsize = 0;
      let j2len = new Map();
      for (let i = alo; i < ahi; i++) {
        const next = new Map();
        const js = b2j.get(a[i]) || [];
        for (let t = 0; t < js.length; t++) {
          const j = js[t];
          if (j < blo) continue;
          if (j >= bhi) break;
          const k = (j2len.get(j - 1) || 0) + 1;
          next.set(j, k);
          if (k > bestsize) { besti = i - k + 1; bestj = j - k + 1; bestsize = k; }
        }
        j2len = next;
      }
      while (besti > alo && bestj > blo && a[besti - 1] === b[bestj - 1]) { besti--; bestj--; bestsize++; }
      while (besti + bestsize < ahi && bestj + bestsize < bhi && a[besti + bestsize] === b[bestj + bestsize]) bestsize++;
      return [besti, bestj, bestsize];
    }
    let matches = 0;
    const queue = [[0, la, 0, lb]];
    while (queue.length) {
      const [alo, ahi, blo, bhi] = queue.pop();
      const [i, j, k] = longest(alo, ahi, blo, bhi);
      if (k) {
        matches += k;
        if (alo < i && blo < j) queue.push([alo, i, blo, j]);
        if (i + k < ahi && j + k < bhi) queue.push([i + k, ahi, j + k, bhi]);
      }
    }
    return 2 * matches / (la + lb);
  }

  // ---------------------------------------------------------------- 色の意味（ピンクの字＝月間特売・名前のセルの塗り〈白・灰色以外〉＝新商品）
  const isSaleFont = (rgb) => PINK_FONTS.indexOf(rgb) >= 0;
  const isNewFill = (rgb) => Boolean(rgb) && !/^([0-9A-Fa-f]{2})\1\1$/.test(rgb);   // 白・灰色は R=G=B
  // 名前の一部だけをピンクにした行もある（末尾の量のカッコは赤など）。量のカッコを除いた商品名の部分で、空白を除いた字の半分を超えてピンクなら
  // 月間特売と読む（2026-11-06・10月の5商品。Python の道具と種はこの直しの前）
  function pinkMostly(runs, text) {
    const s = String(text === null || text === undefined ? '' : text);
    const end = (s.replace(/\s*[（(][^（）()]*[)）]\s*$/, '') || s).length;
    const total = s.slice(0, end).replace(/\s/g, '').length;
    if (!total || !runs || !runs.length) return false;
    let pink = 0;
    runs.forEach((r) => { if (r[2] && isSaleFont(r[2].c) && r[0] < end) pink += s.slice(r[0], Math.min(r[1], end)).replace(/\s/g, '').length; });
    return pink * 2 > total;
  }
  // ピンクは販売形態の印なので、商品の名前の飾りには残さない（core.js の dropSaleColor と同じ決まり）
  function dropPink(runs) {
    const out = [];
    (runs || []).forEach((r) => { const p = Object.assign({}, r[2] || {}); if (isSaleFont(p.c)) delete p.c; if (Object.keys(p).length) out.push([r[0], r[1], p]); });
    return out;
  }

  // ================================================================ 止める・確かめる
  class StopRead extends Error {
    constructor(items) { super(items.map((x) => x.what).join(' / ')); this.items = items; }
  }
  const stopItem = (what, where) => ({ level: LEVEL_STOP, what, where: where || '' });
  const stop = (what, where) => new StopRead([stopItem(what, where)]);

  // ================================================================ ブックの1枚を見る道具
  function sheetView(sh) {
    const cells = sh.cells || {};
    return {
      name: sh.name, state: sh.state, raw: sh, cells,
      cell: (r, c) => cells[r + ',' + c],
      /** 式なら式の文字・ほかは入っている値（Python で式のまま開いたときの値） */
      f: (r, c) => { const x = cells[r + ',' + c]; if (!x) return null; if (x.f !== null && x.f !== undefined) return x.f; return x.v === undefined ? null : x.v; },
      /** 保存されていた値（式なら計算の結果） */
      v: (r, c) => { const x = cells[r + ',' + c]; return x && x.v !== undefined ? x.v : null; },
      maxRow: sh.maxRow || 0, maxCol: sh.maxCol || 0,
      hiddenRows: new Set(sh.hiddenRows || []), hiddenCols: new Set(sh.hiddenCols || []),
      printArea: sh.printArea || null, merges: sh.merges || []
    };
  }
  const sheetWhere = (ws) => 'シート「' + ws.name + '」';
  const cellWhere = (ws, r, c) => sheetWhere(ws) + ' の ' + colLetter(c) + r;

  // ================================================================ 見出しの並び（指紋）を確かめる
  /** 確かめた形のどれかと同じなら、その形の名前。違えば、違う見出しごとに場所を付けて止める */
  function checkLayout(kind, key, fp, label, cellOf, partLabel) {
    const forms = (LAYOUTS[kind] || {})[key] || [];
    for (const form of forms) if (sameJson(form['指紋'], fp)) return form['名前'];
    const tail = '。元の書類の見出しを確かめてください（正しい形なら、形を登録してから読み直します）';
    if (!forms.length) throw stop(label + 'の見出しの並びの、確かめた形がありません' + tail, '');
    const score = (form) => { let n = 0; const a = form['指紋']; for (let i = 0; i < Math.min(a.length, fp.length); i++) if (sameJson(a[i], fp[i])) n++; return n - Math.abs(a.length - fp.length); };
    let best = forms[0];
    forms.forEach((f) => { if (score(f) > score(best)) best = f; });
    const items = [];
    const a = best['指紋'];
    for (let i = 0; i < Math.max(a.length, fp.length); i++) {
      const x = a[i], y = fp[i];
      if (sameJson(x, y)) continue;
      if (x === undefined) { items.push(stopItem(label + 'に、確かめた形（' + best['名前'] + '）に無い' + partLabel(i, y) + 'があります' + tail, cellOf(i, 0))); continue; }
      if (y === undefined) { items.push(stopItem(label + 'に、確かめた形（' + best['名前'] + '）の' + partLabel(i, x) + 'がありません' + tail, cellOf(i, 0))); continue; }
      for (let j = 0; j < Math.max(x.length, y.length); j++) {
        if (x[j] === y[j]) continue;
        items.push(stopItem(label + 'の' + partLabel(i, y) + 'の ' + (j + 1) + ' 番目の見出しが「' + (y[j] === undefined ? '（無し）' : y[j]) + '」です（確かめた形「' + best['名前'] + '」では「' + (x[j] === undefined ? '（無し）' : x[j]) + '」）' + tail, cellOf(i, j)));
      }
    }
    throw new StopRead(items.length ? items : [stopItem(label + 'の見出しの並びが、確かめた形と違います' + tail, '')]);
  }

  // ================================================================ 番号表（首都圏・北関東・焼き物）を読む（souzai_read.read_number_table と同じ）
  function findHeaderRow(ws) {
    for (let r = 1; r <= 6; r++) {
      let n = 0;
      for (let c = 1; c <= ws.maxCol; c++) if (squash(ws.f(r, c)) === '呼出番') n++;
      if (n >= 2) return r;
    }
    throw stop('1〜6行目に「呼出番」の見出しが2つ以上ある行がありません', sheetWhere(ws));
  }

  function resolveBlocks(ws, hr) {
    const maxc = ws.maxCol;
    const hdr = {}, sq = {};
    for (let c = 1; c <= maxc + 1; c++) { hdr[c] = ws.f(hr, c); sq[c] = squash(hdr[c]); }
    const isHead = (s) => s.indexOf('≪') === 0 || s.indexOf('《') === 0;
    const starts = [];
    range(1, maxc).filter((c) => sq[c] === '呼出番').forEach((cc) => {
      let nc = null;
      for (let c = cc - 1; c >= 1; c--) {
        if (isHead(sq[c])) { nc = c; break; }
        if (sq[c] === '呼出番') break;
      }
      if (nc === null) throw stop('呼出番（' + colLetter(cc) + '列）の左に ≪…≫ のまとまりの名前の見出しがありません', cellWhere(ws, hr, cc));
      starts.push([nc, cc]);
    });
    const list = (cols) => cols.map((c) => colLetter(c) + '=' + hdrName(hdr[c])).join('、');
    return starts.map(([nc, cc], i) => {
      let end;
      if (i + 1 < starts.length) end = starts[i + 1][0] - 1;
      else { end = cc; while (end + 1 <= maxc && sq[end + 1] !== '' && !isHead(sq[end + 1])) end++; }
      const cols = range(nc, end);
      const roles = { name: nc, call: cc };
      const mid = range(nc + 1, cc - 1);
      const trays = mid.filter((c) => sq[c] === '指定トレー' || sq[c] === '旧トレー');
      const changes = mid.filter((c) => sq[c] === '変更トレー');
      const blanks = mid.filter((c) => sq[c] === '');
      const others = mid.filter((c) => ['指定トレー', '旧トレー', '変更トレー', ''].indexOf(sq[c]) < 0);
      if (trays.length > 1 || changes.length > 1 || others.length) {
        const bad = others[0] || trays[1] || changes[1];
        throw stop('まとまり「' + nk(hdr[nc]) + '」の名前と呼出番の間の見出しが想定と違います: ' + list(mid), cellWhere(ws, hr, bad));
      }
      if (trays.length) roles.tray = trays[0];
      else if (blanks.length) roles.tray = blanks.shift();            // 前提 A1（丼ぶりのまとまりは、見出しの無い左の列がトレー）
      if (changes.length) roles.tray_change = changes[0];
      blanks.forEach((c, k) => { roles['blank' + (k + 1)] = c; });
      const after = range(cc + 1, end);
      const cost = after.filter((c) => sq[c] === '原価'), gross = after.filter((c) => sq[c] === '粗利');
      const price = after.filter((c) => sq[c] === '定番価格'), small = after.filter((c) => sq[c] === '小パック');
      const unknown = after.filter((c) => cost.concat(gross, price, small).indexOf(c) < 0);
      if (cost.length !== 1 || gross.length < 1 || gross.length > 2 || price.length < 2 || price.length > 3 || small.length > 1 || unknown.length) {
        throw stop('まとまり「' + nk(hdr[nc]) + '」の呼出番の右の見出しが想定と違います: ' + list(after), cellWhere(ws, hr, unknown[0] || cc));
      }
      roles.cost = cost[0];
      roles.gross_after = gross[gross.length - 1];
      if (gross.length === 2) roles.gross_before = gross[0];
      roles.price_visible = price[price.length - 1];
      roles.price_calc = price[price.length - 2];
      if (price.length === 3) roles.price_before = price[0];
      if (small.length) roles.small = small[0];
      // 見える価格は見えている列・計算用の価格は隠し列（1回でも違えば止める）
      if (ws.hiddenCols.has(roles.price_visible) || !ws.hiddenCols.has(roles.price_calc)) {
        throw stop('まとまり「' + nk(hdr[nc]) + '」の定番価格の列の見え方が想定と違います（見える価格=' + colLetter(roles.price_visible) + '列・計算用=' + colLetter(roles.price_calc) + '列）', cellWhere(ws, hr, roles.price_visible));
      }
      return { index: i, name: nk(hdr[nc]), name_raw: hdr[nc], first: nc, last: end, roles, fingerprint: cols.map((c) => hdrName(hdr[c])) };
    });
  }

  function readNumberTable(book, key) {
    const label = '番号表（' + key + '）';
    const hits = book.sheets.filter((s) => s.state === 'visible' && SHEET_RULE[key](nk(s.name)));
    if (hits.length !== 1) {
      throw stop(label + 'のシートが ' + hits.length + ' 枚見つかりました（1枚の前提）。見えているシート: ' + book.sheets.filter((s) => s.state === 'visible').map((s) => s.name).join('、'), 'ブック');
    }
    const ws = sheetView(hits[0]);
    if (!ws.printArea) throw stop('印刷範囲がありません。印刷範囲の中を「今月並べる商品」として読むため、止めました', sheetWhere(ws));
    const [r1, c1, r2, c2] = ws.printArea;
    const hr = findHeaderRow(ws);
    const blocks = resolveBlocks(ws, hr);
    const layoutName = checkLayout('番号表', key, blocks.map((b) => b.fingerprint), label + '「' + ws.name + '」',
      (i, j) => (blocks[i] ? cellWhere(ws, hr, blocks[i].first + j) : sheetWhere(ws) + ' の ' + hr + '行目'),
      (i, fpBlock) => ' ' + (i + 1) + ' つ目のまとまり「' + (fpBlock && fpBlock[0]) + '」');
    const place = (r, c) => (ws.hiddenRows.has(r) ? '隠し行' : (r >= r1 && r <= r2 && c >= c1 && c <= c2 ? '印刷範囲の中' : '印刷範囲の外'));
    const lines = [];
    blocks.forEach((b) => {
      const roles = Object.keys(b.roles);
      for (let r = hr + 1; r <= ws.maxRow; r++) {
        const cells = {};
        let any = false;
        roles.forEach((role) => { const f = ws.f(r, b.roles[role]); if (!isBlank(f)) { cells[role] = { f, v: ws.v(r, b.roles[role]) }; any = true; } });
        if (!any) continue;
        const name = cells.name ? cells.name.f : null;
        if (!isBlank(name) && isFormula(name)) throw stop('名前が式になっています（文字の前提）', cellWhere(ws, r, b.roles.name));
        const kind = !isBlank(name) && isSection(name) ? 'section' : (!isBlank(name) ? 'product' : 'cont');
        const nm = ws.cell(r, b.roles.name) || {};
        lines.push({ block: b.index, row: r, kind, place: place(r, b.roles.name), cells, nameFont: (nm.font && nm.font.c) || '', nameRuns: nm.runs || null, nameFill: nm.fill || '', addr: colLetter(b.roles.name) + r });
      }
    });
    return { sheet: ws.name, ws, table: key, header_row: hr, print: [c1, r1, c2, r2], blocks, layoutName, lines };
  }

  // ================================================================ 月間商品を読む（souzai_read.read_monthly と同じ）
  function readMonthly(book) {
    const sh = book.sheets.find((s) => s.name === MONTHLY_SHEET);
    if (!sh) throw stop('「' + MONTHLY_SHEET + '」のシートがありません', 'ブック');
    const ws = sheetView(sh);
    if (!ws.printArea) throw stop('印刷範囲がありません', sheetWhere(ws));
    const [r1, c1, r2, c2] = ws.printArea;
    const cols = range(c1, c2);
    const hdrRows = range(r1, r2).filter((r) => { const vals = cols.map((c) => squash(ws.f(r, c))); return vals.indexOf('商品名') >= 0 && vals.indexOf('呼出番') >= 0; });
    if (hdrRows.length !== 2) throw stop('「商品名」「呼出番」の見出しの行が ' + hdrRows.length + ' 行あります（月間特売と新規商品の2行の前提）', sheetWhere(ws));
    const layoutErrors = [];
    const parts = hdrRows.map((h) => {
      let t = null;
      for (let r = h - 1; r >= r1; r--) { if (cols.some((c) => !isBlank(ws.f(r, c)) && !isFormula(ws.f(r, c)))) { t = r; break; } }
      if (t === null) throw stop(h + '行目の見出しの上に題（「N月月間特売」など）がありません', sheetWhere(ws));
      const title = nk(ws.f(t, c1));
      let kind;
      if (title.indexOf('月間特売') >= 0) kind = '月間特売';
      else if (title.indexOf('新規') >= 0) kind = '新規商品';
      else throw stop(t + '行目の題「' + title + '」が、月間特売とも新規商品とも読めません', cellWhere(ws, t, c1));
      const row1 = {}, row2 = {};
      cols.forEach((c) => { row1[c] = squash(ws.f(h, c)); row2[c] = squash(ws.f(h + 1, c)); });
      const fp = [cols.map((c) => hdrName(ws.f(h, c))), cols.map((c) => hdrName(ws.f(h + 1, c)))];
      try {
        checkLayout('月間商品', kind, fp, '月間商品（' + kind + '）', (i, j) => cellWhere(ws, h + i, c1 + j), (i) => '見出しの ' + (i + 1) + ' 行目');
      } catch (e) { if (e instanceof StopRead) { layoutErrors.push.apply(layoutErrors, e.items); return null; } throw e; }
      const one = (rowmap, label, nth) => { const cs = cols.filter((c) => rowmap[c] === label); if (cs.length <= nth) throw stop(h + '行目に見出し「' + label + '」が見つかりません', sheetWhere(ws)); return cs[nth]; };
      const roles = { name: one(row1, '商品名', 0), call: one(row1, '呼出番', 0), size: one(row1, '規格', 0), price: one(row1, '本体価格', 0), tax: one(row1, '税込価格', 0), s_size: one(row1, '規格', 1), s_price: one(row2, '本体価格', 0), s_tax: one(row2, '税込価格', 0) };
      return { kind, title_row: t, title: ws.f(t, c1), header_row: h, roles };
    });
    if (layoutErrors.length) throw new StopRead(layoutErrors);
    const lines = [];
    parts.forEach((p, i) => {
      const end = i + 1 < parts.length ? parts[i + 1].title_row - 1 : r2;
      for (let r = p.header_row + 2; r <= end; r++) {
        const cells = {};
        Object.keys(p.roles).forEach((role) => { cells[role] = { f: ws.f(r, p.roles[role]), v: ws.v(r, p.roles[role]) }; });
        const name = cells.name.f;
        const hasData = ['call', 'size', 'price', 's_size', 's_price'].some((k) => !isBlank(cells[k].f));
        let kind;
        if (isBlank(name) && !hasData) kind = 'empty';
        else if (!isBlank(name) && !hasData) kind = 'section';
        else if (!isBlank(name) && !startsWithAny(nk(name), CONT_MARKS)) kind = 'product';
        else kind = 'cont';
        lines.push({ part: p.kind, row: r, kind, cells });
      }
    });
    return { sheet: MONTHLY_SHEET, ws, print: [c1, r1, c2, r2], parts, lines };
  }

  // ================================================================ 月をブックの中から決める（月間商品の題・シート名・加工基本表の日付）
  const RE_ISO_DATE = /^([0-9]{4})-([0-9]{2})-([0-9]{2})(?:[T ]|$)/;
  /** Excel の日の通し番号（日付が数で来たとき）→ [年, 月] */
  function serialToYm(n) { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000); return [d.getUTCFullYear(), d.getUTCMonth() + 1]; }
  function monthOf(book, monthly, wanted) {
    const found = [];                           // { m, y, where }
    monthly.parts.forEach((p) => { const m = /^[ ]*([0-9]{1,2})月/.exec(nk(p.title)); if (m) found.push({ m: Number(m[1]), y: null, where: 'シート「' + MONTHLY_SHEET + '」の題（' + p.title_row + '行目）' }); });
    book.sheets.forEach((s) => { if (s.state !== 'visible') return; const m = /([0-9]{1,2})月$/.exec(nk(s.name)); if (m) found.push({ m: Number(m[1]), y: null, where: 'シート「' + s.name + '」の名前' }); });
    const proc = book.sheets.find((s) => s.name === PROCESSING_SHEET);
    if (proc) {
      const ws = sheetView(proc);
      let hr = null;
      for (let r = 1; r <= 5 && hr === null; r++) { let n = 0; for (let c = 1; c <= ws.maxCol; c++) if (squash(ws.f(r, c)).indexOf('品名') === 0) n++; if (n >= 2) hr = r; }
      let hit = null;
      for (let r = 1; hr !== null && r < hr && !hit; r++) {
        for (let c = 1; c <= ws.maxCol; c++) {
          const v = ws.v(r, c);
          const iso = typeof v === 'string' ? RE_ISO_DATE.exec(v) : null;
          if (iso) { hit = { y: Number(iso[1]), m: Number(iso[2]), where: cellWhere(ws, r, c) + ' の日付' }; break; }
          if (typeof v === 'number' && v >= 36526 && v < 73051) { const ym = serialToYm(v); hit = { y: ym[0], m: ym[1], where: cellWhere(ws, r, c) + ' の日付' }; break; }
        }
      }
      if (hit) found.push(hit);
    }
    const months = Array.from(new Set(found.map((x) => x.m)));
    if (!months.length) throw stop('ブックの中から月が決められません（月間商品の題「N月月間特売」・シート名「惣菜N月」・加工基本表の日付のどれもありません）', 'ブック');
    if (months.length > 1) throw stop('ブックの中で月が食い違っています: ' + found.map((x) => x.where + '＝' + x.m + '月').join('、'), 'ブック');
    const m = months[0];
    const years = Array.from(new Set(found.filter((x) => x.y).map((x) => x.y)));
    if (years.length > 1) throw stop('ブックの中で年が食い違っています: ' + found.filter((x) => x.y).map((x) => x.where + '＝' + x.y + '年').join('、'), 'ブック');
    let y = years.length === 1 ? years[0] : null;
    const wm = wanted ? /^([0-9]{4})-([0-9]{2})$/.exec(String(wanted)) : null;
    if (wanted && !wm) throw stop('作る月「' + wanted + '」の書き方が違います（YYYY-MM）', '');
    if (y === null) {
      if (!wm) throw stop('ブックの中から年が決められません（加工基本表の日付がありません）。作る月（YYYY-MM）を渡してください', 'ブック');
      y = Number(wm[1]);
    }
    const month = y + '-' + pad(m, 2);
    if (wm && month !== wanted) throw stop('ブックの月（' + month + '・' + found.map((x) => x.where).join('、') + 'から）が、作る月（' + wanted + '）と違います', 'ブック');
    return { month, mNum: m };
  }

  // ================================================================ 見た目（書式キー）: アプリの中にある「書式」「表の型_列」に当てる
  //   ブックには、セルの字（名前・大きさ・太字・斜体・色）と塗りしか無い。書式キーは、アプリの「書式」の表のうち、この6つが同じ物から、
  //   罫線・位置・折り返し・縮小・表示形式が型にいちばん近い物を選ぶ。型と違うかは、色・大きさ・太字で見る（折り返し・縮小は型どおりとみなす）。
  const REST_KEYS = ['罫線（左）', '罫線（右）', '罫線（上）', '罫線（下）', '横の位置', '縦の位置', '折り返し', '縮小して全体を表示', '表示形式'];
  const blankStr = (v) => (v === null || v === undefined ? '' : String(v));
  const sizeKey = (v) => (v === null || v === undefined || v === '' ? '' : String(Number(v)));
  const visKey = (s) => [blankStr(s['字']), sizeKey(s['大きさ']), blankStr(s['太字']), blankStr(s['斜体']), blankStr(s['字の色']), blankStr(s['塗り'])].join('|');
  function rowsOf(t) { if (!t || !t.c) return []; return (t.r || []).map((r) => { const o = {}; t.c.forEach((c, i) => { o[c] = r[i]; }); return o; }); }
  function styleContext(typeTables) {
    if (!typeTables || !typeTables['書式'] || !typeTables['表の型_列']) return null;
    const byKey = new Map(), byVis = new Map(), byFull = new Map();
    rowsOf(typeTables['書式']).forEach((s) => {
      const key = blankStr(s['書式キー']);
      if (!key) return;
      s.key = key;
      byKey.set(key, s);
      const k = visKey(s);
      if (!byVis.has(k)) byVis.set(k, []);
      byVis.get(k).push(s);
      const fk = fullKey(s);
      if (!byFull.has(fk)) byFull.set(fk, s);
    });
    const colType = new Map();
    rowsOf(typeTables['表の型_列']).forEach((r) => { if (!isBlank(r['役割'])) colType.set(r['表'] + '|' + blankStr(r['まとまり']) + '|' + r['役割'], r); });
    const shape = new Map(); rowsOf(typeTables['表の型']).forEach((r) => shape.set(r['表'], r));
    const blockType = new Map(); rowsOf(typeTables['表の型_まとまり']).forEach((r) => blockType.set(r['表'] + '|' + blankStr(r['まとまり']), r));
    return { byKey, byVis, byFull, colType, shape, blockType };
  }
  const FULL_KEYS = ['字', '大きさ', '太字', '斜体', '字の色', '塗り'].concat(REST_KEYS);
  const fullKey = (s) => FULL_KEYS.map((k) => (k === '大きさ' ? sizeKey(s[k]) : blankStr(s[k]))).join('|');
  /** ブックのセルの見える所（字・塗り） */
  function bookLook(x) {
    const f = (x && x.font) || {};
    return { '字': f.f || '', '大きさ': f.s === undefined || f.s === null ? '' : f.s, '太字': f.b ? '○' : '', '斜体': f.i ? '○' : '', '字の色': f.c || '', '塗り': (x && x.fill) || '' };
  }
  /** セルの書式。Book に書式の残り（x: 罫線・位置・折り返し・縮小・表示形式）があれば全部（full）、無ければ字と塗りだけ。値も塗りも無いセルは null */
  function lookAt(ws, r, c) {
    const x = ws.cell(r, c);
    const rest = (look, e) => Object.assign(look, { '罫線（左）': e.bl || '', '罫線（右）': e.br || '', '罫線（上）': e.bt || '', '罫線（下）': e.bb || '', '横の位置': e.h || '', '縦の位置': e.v || '', '折り返し': e.w ? '○' : '', '縮小して全体を表示': e.k ? '○' : '', '表示形式': e.n || 'General', full: true });
    if (x && x.x) return rest(bookLook(x), x.x);
    if (!x && ws.raw.x0 && ws.raw.x0.x) return rest(bookLook(ws.raw.x0), ws.raw.x0.x);   // 一度も触れていないセルの書式
    return x ? bookLook(x) : null;
  }
  /** 見える所が同じ書式のキー（書式の全部が分かれば、全部が同じ物。候補が2つ以上なら、hints の書式に近い物）。無ければ null */
  function keyOf(ctx, look, hints) {
    if (look.full) { const hit = ctx.byFull.get(fullKey(look)); if (hit) return hit.key; }
    const cands = ctx.byVis.get(visKey(look)) || [];
    if (!cands.length) return null;
    for (const h of hints) if (h && cands.some((s) => s.key === h)) return h;
    const base = hints.map((h) => ctx.byKey.get(blankStr(h))).find(Boolean);
    if (!base) return cands[0].key;
    let best = cands[0], bestN = -1;
    cands.forEach((s) => { const n = REST_KEYS.filter((k) => blankStr(s[k]) === blankStr(base[k])).length; if (n > bestN) { best = s; bestN = n; } });
    return best.key;
  }
  /** 見た目に効く所（字の色・塗り・大きさ・太字・折り返し・縮小）。色を渡すとその色で見る（月間特売・新商品） */
  function lookTuple(d, fc, fl) {
    if (!d) return ['', '', null, '', '', ''];
    const f = fc !== null && fc !== undefined ? fc : blankStr(d['字の色']);
    const l = fl !== null && fl !== undefined ? fl : blankStr(d['塗り']);
    const s = d['大きさ'] === '' || d['大きさ'] === null || d['大きさ'] === undefined ? null : Number(d['大きさ']);
    return [f || '000000', l || 'FFFFFF', s, blankStr(d['太字']), blankStr(d['折り返し']), blankStr(d['縮小して全体を表示'])];
  }
  function sameColor(a, b) {
    if (a === b) return true;
    if (a.length !== 6 || b.length !== 6) return false;
    for (let i = 0; i < 6; i += 2) { const x = parseInt(a.slice(i, i + 2), 16), y = parseInt(b.slice(i, i + 2), 16); if (!(Math.abs(x - y) <= 3)) return false; }
    return true;
  }
  const sameLook = (a, b) => sameColor(a[0], b[0]) && sameColor(a[1], b[1]) && a[2] === b[2] && a[3] === b[3] && a[4] === b[4] && a[5] === b[5];

  // ================================================================ 取り込み（import_master.build と同じ。この部品に要る所だけ）
  function raw(ln, role, which) { const c = ln.cells[role]; return c === undefined ? null : c[which || 'f']; }
  /** 続きの行のうち、意味のある値がある行（隠し列の式・#DIV/0!・0 だけの行は空き行） */
  function meaningfulCont(ln) {
    if (['call', 'price_visible', 'small', 'cost'].some((role) => !isBlank(raw(ln, role)))) return true;
    return ['price_calc', 'price_before'].some((role) => { const v = raw(ln, role); if (isBlank(v)) return false; const n = num(v); return n !== null && n !== 0; });
  }
  /** 列の役割の名前。8・9月の丼ぶりの、見出しの無い2つ目の列は「変更トレー」 */
  const roleJa = (b, role) => (role === 'blank1' && !b.roles.tray_change ? ROLE_JA.tray_change : (ROLE_JA[role] || role));

  function buildMaster(tables, monthly, mNum, ctx, problems) {
    const doubts = [];
    const D = (where, what, hint) => doubts.push({ '番号': doubts.length + 1, 'どこ': where, '何が': what, '手がかり': hint || '', '済み': '' });
    const styleMiss = [], notInType = new Set();
    const keyOrMiss = (look, hints, where) => { const k = keyOf(ctx, look, hints); if (k === null) styleMiss.push(where); return k; };

    // ---- 月間特売（ふだんの価格を見分けるため）
    const sale = monthly.lines.filter((ln) => ln.part === '月間特売' && ln.kind === 'product').map((ln) => ({ call: halfCall(ln.cells.call.f) || '', nkey: nameKey(ln.cells.name.f), price: num(ln.cells.price.f) }));

    // ---- 商品（印刷範囲の中だけ・首都圏→焼き物→北関東の順に SZ0001 から。表・呼出番・空白を外した名前が同じなら同じ商品）
    const prods = new Map(), occ = new Map();
    ORDER_FOR_ID.forEach((tkey) => tables[tkey].lines.forEach((ln) => {
      if (ln.kind !== 'product' || ln.place !== '印刷範囲の中') return;
      let hc = halfCall(raw(ln, 'call'));
      if (hc === null) hc = nk(raw(ln, 'call'));
      const key = [TABLE_KIND[tkey], hc, squash(raw(ln, 'name'))];
      const ks = key.join('\u0000');
      if (!prods.has(ks)) prods.set(ks, { id: 'SZ' + pad(prods.size + 1, 4), key, first: [tkey, ln] });
      if (!occ.has(ks)) occ.set(ks, []);
      occ.get(ks).push([tkey, ln]);
    }));
    const byLine = new Map();
    const special = [];
    prods.forEach((p, ks) => {
      const [tkey, ln] = p.first;
      const key = p.key;
      const tabs = new Set(occ.get(ks).map((x) => x[0]));
      let region;
      if (TABLE_KIND[tkey] !== TABLE_KIND['首都圏']) region = '共通';
      else if (tabs.has('首都圏') && tabs.has('北関東')) region = '共通';
      else region = tabs.has('首都圏') ? '首都圏だけ' : '北関東だけ';
      const printed = asNumber(raw(ln, 'price_visible', 'v'));
      const calc = asNumber(raw(ln, 'price_calc', 'v'));
      const nkey = nameKey(raw(ln, 'name'));
      const isSale = sale.some((s) => ((s.call && s.call === key[1]) || (!s.call && s.nkey === nkey)) && pyEq(s.price, printed));
      let regular = printed;
      if (isSale && typeof calc === 'number' && typeof printed === 'number' && calc !== printed) { regular = calc; special.push([p, printed, calc]); }
      let cost = raw(ln, 'cost');
      if (!isFormula(cost)) cost = asNumber(raw(ln, 'cost', 'v'));
      p.row = {
        '商品ID': p.id, '表': key[0], '地域': region, '呼出番': key[1], '表示名（印刷どおり）': raw(ln, 'name'), '指定トレー': raw(ln, 'tray'),
        '見える価格': regular, '小パック': raw(ln, 'small'), 'メモ': null, '原価': cost, '粗利（包材値上前）': cleanNum(raw(ln, 'gross_before', 'v')),
        '定番価格（包材値上前）': asNumber(raw(ln, 'price_before', 'v')), '粗利（包材変更後）': null, '定番価格（包材変更後）': calc
      };
      occ.get(ks).forEach(([t, l]) => byLine.set(t + ',' + l.block + ',' + l.row, p));
      // 同じ商品（首都圏と北関東）で値が違う欄
      occ.get(ks).slice(1).forEach(([t, l]) => ['name', 'tray', 'cost', 'gross_before', 'price_before', 'price_calc', 'price_visible', 'small'].forEach((role) => {
        const a = raw(ln, role, 'v'), b = raw(l, role, 'v');
        if (!(isBlank(a) && isBlank(b)) && !pyEq(a, b)) D(t + ' ' + l.addr, '同じ商品（' + key[1] + ' ' + nk(raw(ln, 'name')) + '）で、' + role + ' の値が ' + tkey + ' と違う（' + pyStr(b) + '）', '台帳は ' + tkey + ' の値で入れた');
      }));
    });
    const prodList = Array.from(prods.values());
    if (special.length) {
      D('首都圏・北関東・焼き物（月間特売）',
        '月間特売で、見える価格が月間商品の特売の価格と同じ商品 ' + special.length + ' 件。商品の「見える価格」には、隠しの定番価格（包材変更後）の値を「ふだんの価格」として入れ、' + mNum + '月の並びの「その月の見える価格」に特売の価格を入れた（読み）',
        special.map(([p, a, b]) => p.row['呼出番'] + ' ' + nk(p.row['表示名（印刷どおり）']) + '（特売 ' + pyStr(a) + '・ふだん ' + pyStr(b) + '）').join('／'));
    }

    // ---- 区分・並び（番号表3枚。表題 ≪…≫ の行から次の表題の手前までが、その区分）
    const sections = [], layout = [], redCells = [];
    TABLE_KEYS.forEach((tkey) => {
      const t = tables[tkey];
      const ws = t.ws, hr = t.header_row;
      const lines = new Map();
      t.lines.forEach((l) => { if (l.place === '印刷範囲の中') lines.set(l.block + ',' + l.row, l); });
      const nameCols = new Set(t.blocks.map((b) => b.roles.name));
      const vmerge = new Set();
      ws.merges.forEach((m) => { if (m[0] > hr && m[1] === m[3] && nameCols.has(m[1]) && m[2] === m[0] + 1) vmerge.add(m[1] + ',' + m[2]); });
      let seqSec = 0;
      const rowrefs = [];
      const titleKey = (r, b, place) => {
        if (!ctx) return null;
        const bt = ctx.blockType.get(tkey + '|' + (b.index + 1)) || {};
        const ct = ctx.colType.get(tkey + '|' + (b.index + 1) + '|名前') || {};
        return keyOrMiss(lookAt(ws, r, b.roles.name) || bookLook(null), [place === '見出しの行' ? bt['表題の書式（見出しの行）'] : bt['表題の書式（表の中）'], ct['書式（区分の表題）']], cellWhere(ws, r, b.roles.name));
      };
      t.blocks.forEach((b) => {
        const nc = b.roles.name;
        let last = hr;
        lines.forEach((l) => { if (l.block === b.index && (l.kind === 'product' || l.kind === 'section' || (l.kind === 'cont' && meaningfulCont(l)))) last = Math.max(last, l.row); });
        seqSec++;
        let secOrder = 1;
        let cur = { '区分ID': tkey + '-' + pad(seqSec, 2), '表': tkey, '列のまとまり': b.index + 1, 'まとまりの中の順': 1, '表題（印刷どおり）': b.name_raw, '表題の置き場': '見出しの行', '表題の続き（呼出番の欄）': null, '表題の書式': titleKey(hr, b, '見出しの行'), '表題の行の書式を変える欄': null };
        sections.push(cur);
        let item = 0;
        for (let r = hr + 1; r <= last; r++) {
          const ln = lines.get(b.index + ',' + r);
          if (ln && ln.kind === 'section') {
            seqSec++;
            secOrder++;
            const extra = raw(ln, 'call');
            cur = { '区分ID': tkey + '-' + pad(seqSec, 2), '表': tkey, '列のまとまり': b.index + 1, 'まとまりの中の順': secOrder, '表題（印刷どおり）': raw(ln, 'name'), '表題の置き場': '表の中', '表題の続き（呼出番の欄）': isBlank(extra) ? null : extra, '表題の書式': titleKey(r, b, '表の中'), '表題の行の書式を変える欄': null };
            sections.push(cur);
            rowrefs.push([r, b, 'section', cur]);
            item = 0;
            continue;
          }
          item++;
          const row = { '表': tkey, '区分ID': cur['区分ID'], '何番目': item, '行の種類': null, '商品ID': null, '月間特売': null, '新商品': null, 'その月の見える価格': null, '名前のセルを上とつなぐ': null, '続き: 呼出番の欄': null, '続き: 見える価格': null, '続き: 小パック': null, '続き: 原価': null, '続き: 粗利（包材値上前）': null, '続き: 定番価格（包材値上前）': null, '続き: 定番価格（包材変更後）': null, '書式を変える欄': null };
          if (vmerge.has(nc + ',' + r)) row['名前のセルを上とつなぐ'] = '○';
          if (ln && ln.kind === 'product') {
            const p = byLine.get(tkey + ',' + b.index + ',' + r);
            row['行の種類'] = KIND_PRODUCT;
            row['商品ID'] = p.id;
            row['月間特売'] = isSaleFont(ln.nameFont) || pinkMostly(ln.nameRuns, raw(ln, 'name', 'v')) ? '○' : null;
            row['新商品'] = isNewFill(ln.nameFill) ? '○' : null;
            const pv = asNumber(raw(ln, 'price_visible', 'v'));
            if (!pyEq(pv, p.row['見える価格'])) row['その月の見える価格'] = pv;
            // 隠しの粗利の式（式が無い行は、番号表で決まった形の式を入れるので疑問にしない）
            const ga = raw(ln, 'gross_after');
            const L = (role) => colLetter(b.roles[role]);
            const hint = '番号表では決まった形の式で作る。商品 ' + p.row['呼出番'] + ' ' + nk(p.row['表示名（印刷どおり）']);
            if (!isBlank(ga) && !isFormula(ga)) {
              D(tkey + ' ' + L('gross_after') + r, '隠しの粗利（包材変更後）の欄が、式でなく値「' + nk(ga) + '」になっている', hint);
            } else if (!isBlank(ga)) {
              const std = '=+((' + L('price_calc') + r + '-' + L('cost') + r + ')/' + L('price_calc') + r + ')';
              if (String(ga).replace(/ /g, '') !== std) D(tkey + ' ' + L('gross_after') + r, '隠しの粗利の式 ' + ga + ' が、ほかの欄を見ている（決まった形は ' + std + '）', hint);
            }
          } else if (ln && ln.kind === 'cont' && meaningfulCont(ln)) {
            row['行の種類'] = KIND_CONT;
            row['続き: 呼出番の欄'] = raw(ln, 'call');
            row['続き: 見える価格'] = asNumber(raw(ln, 'price_visible', 'v'));
            row['続き: 小パック'] = raw(ln, 'small');
            const cost = raw(ln, 'cost');
            row['続き: 原価'] = isFormula(cost) ? cost : asNumber(raw(ln, 'cost', 'v'));
            row['続き: 粗利（包材値上前）'] = cleanNum(raw(ln, 'gross_before', 'v'));
            row['続き: 定番価格（包材値上前）'] = asNumber(raw(ln, 'price_before', 'v'));
            row['続き: 定番価格（包材変更後）'] = asNumber(raw(ln, 'price_calc', 'v'));
          } else {
            row['行の種類'] = KIND_BLANK;
          }
          layout.push(row);
          rowrefs.push([r, b, row['行の種類'] === KIND_BLANK ? 'blank' : 'product', row]);
        }
      });
      // 月間特売の字の色・新商品の塗り（アプリの表の型にあればそれ、無ければこのブックでいちばん多い物）
      const inside = Array.from(lines.values()).filter((l) => l.kind === 'product');
      const shape = ctx ? (ctx.shape.get(tkey) || {}) : {};
      const saleRgb = blankStr(shape['月間特売の字の色']) || mostCommon(inside.filter((l) => isSaleFont(l.nameFont)).map((l) => l.nameFont)) || 'FF00FF';
      const newRgb = blankStr(shape['新商品の塗り']) || mostCommon(inside.filter((l) => isNewFill(l.nameFill)).map((l) => l.nameFill)) || null;
      // 色・大きさ・太字が型と違うセルは「書式を変える欄」に（赤い字・橙の塗り・区分の行の色など）。赤い字のセルは疑問に数える
      rowrefs.forEach(([r, b, kind, d]) => {
        const over = [];
        Object.keys(b.roles).map((role) => [role, b.roles[role]]).sort((x, y) => x[1] - y[1]).forEach(([role, c]) => {
          if (kind === 'section' && role === 'name') return;
          const look = lookAt(ws, r, c);
          if (look && look['字の色'] === 'FF0000' && !isBlank(ws.f(r, c)) && !ws.hiddenCols.has(c)) redCells.push(tkey + ' ' + colLetter(c) + r + '「' + nk(ws.f(r, c)) + '」');
          if (!ctx || !look) return;                       // 書式の分からないセル（Book に無い・値も塗りも無い）は、型どおりとみなす
          const ct = ctx.colType.get(tkey + '|' + (b.index + 1) + '|' + roleJa(b, role));
          const expKey = ct ? blankStr(ct[KIND_STYLE_COL[kind]]) : '';
          const exp = ctx.byKey.get(expKey);
          if (!exp) { notInType.add(tkey + '「' + roleJa(b, role) + '」'); return; }   // アプリの表の型に無い列（8・9月の変更トレーなど）
          let fc = null, fl = null;
          if (kind === 'product' && role === 'name' && d['行の種類'] === KIND_PRODUCT) {
            fc = d['月間特売'] === '○' ? saleRgb : null;
            fl = d['新商品'] === '○' ? newRgb : null;
          }
          const act = look.full ? look : Object.assign({}, look, { '折り返し': exp['折り返し'], '縮小して全体を表示': exp['縮小して全体を表示'] });
          if (sameLook(lookTuple(act), lookTuple(exp, fc, fl))) return;
          const k = keyOrMiss(look, [expKey], cellWhere(ws, r, c));
          if (k !== null) over.push(roleJa(b, role) + ':' + k);
        });
        if (over.length) d[kind === 'section' ? '表題の行の書式を変える欄' : '書式を変える欄'] = over.join('／');
      });
    });
    if (redCells.length) {
      D('首都圏・北関東・焼き物', '字が赤のセル ' + redCells.length + ' か所（凡例はピンク字＝月間特売・塗り＝新商品だけで、赤の意味が書かれていない）',
        '見た目は「書式を変える欄」で元のとおりに出している。意味が分かれば、列（例: 地域の目印）で持つ形にできる。' + redCells.slice(0, 30).join('／'));
    }

    // ---- 月間商品（行。部ごとに、中身の最後の行まで）
    const mws = monthly.ws;
    const mrows = [], mrefs = [];
    const KIND_OF_MONTHLY = { product: KIND_PRODUCT, section: KIND_SECTION, cont: KIND_CONT, empty: KIND_BLANK };
    monthly.parts.forEach((p) => {
      const plines = monthly.lines.filter((l) => l.part === p.kind);
      const content = plines.filter((l) => l.kind !== 'empty').map((l) => l.row);
      const last = content.length ? Math.max.apply(null, content) : p.header_row + 1;
      let seq = 0;
      plines.forEach((l) => {
        if (l.row > last) return;
        seq++;
        const c = l.cells;
        const hc = halfCall(c.call.f);
        mrefs.push([l.row, { product: 'product', cont: 'product', section: 'section' }[l.kind] || 'blank', p.roles]);
        mrows.push({ '部': p.kind, '何番目': seq, '行の種類': KIND_OF_MONTHLY[l.kind], '商品名（この表の書き方）': c.name.f, '呼出番（この表に書く番号）': hc ? hc : c.call.f, '規格': c.size.f, '本体価格': asNumber(c.price.f), '小パックの規格': c.s_size.f, '小パックの本体価格': asNumber(c.s_price.f), '商品ID': null, '書式を変える欄': null });
      });
    });
    if (ctx) {
      mrefs.forEach(([r, kind, roles], n) => {
        const over = [];
        Object.keys(roles).map((role) => [role, roles[role]]).sort((x, y) => x[1] - y[1]).forEach(([role, c]) => {
          const look = lookAt(mws, r, c);
          if (!look) return;
          const ct = ctx.colType.get(MONTHLY_SHEET + '||' + MONTHLY_ROLE_JA[role]);
          const expKey = ct ? blankStr(ct[KIND_STYLE_COL[kind]]) : '';
          const exp = ctx.byKey.get(expKey);
          if (!exp) { notInType.add(MONTHLY_SHEET + '「' + MONTHLY_ROLE_JA[role] + '」'); return; }
          const act = look.full ? look : Object.assign({}, look, { '折り返し': exp['折り返し'], '縮小して全体を表示': exp['縮小して全体を表示'] });
          if (sameLook(lookTuple(act), lookTuple(exp))) return;
          const k = keyOrMiss(look, [expKey], cellWhere(mws, r, c));
          if (k !== null) over.push(MONTHLY_ROLE_JA[role] + ':' + k);
        });
        if (over.length) mrows[n]['書式を変える欄'] = over.join('／');
      });
    }

    // ---- 月間商品の行を商品に結ぶ（呼出番 → 呼出番と名前 → 名前）
    mrows.forEach((r) => { if (r['行の種類'] === KIND_PRODUCT) r['商品ID'] = matchProduct(prodList, r['呼出番（この表に書く番号）'], r['商品名（この表の書き方）']); });

    // ---- 地域の対（首都圏だけ ↔ 北関東だけ、同じ位置。無ければ同じまとまりで名前が似ている物）
    const pos = new Map();
    byLine.forEach((p, k) => { const [t, bi, row] = k.split(','); const pk = p.id + '|' + t; if (!pos.has(pk)) pos.set(pk, [Number(bi), Number(row)]); });
    const onlyS = prodList.filter((p) => p.row['地域'] === '首都圏だけ'), onlyK = prodList.filter((p) => p.row['地域'] === '北関東だけ');
    const used = new Set(), pairs = [];
    onlyS.forEach((a) => {
      const pa = pos.get(a.id + '|首都圏');
      let hit = onlyK.filter((b) => !used.has(b.id) && sameJson(pos.get(b.id + '|北関東'), pa));
      if (hit.length !== 1) {
        const cands = onlyK.filter((b) => !used.has(b.id) && (pos.get(b.id + '|北関東') || [null])[0] === pa[0])
          .map((b) => [similar(nameKey(a.row['表示名（印刷どおり）']), nameKey(b.row['表示名（印刷どおり）'])), b.id, b])
          .sort((x, y) => (y[0] - x[0]) || (y[1] < x[1] ? -1 : (y[1] > x[1] ? 1 : 0)));
        hit = cands.length && cands[0][0] >= 0.5 ? [cands[0][2]] : [];
      }
      if (hit.length) { used.add(hit[0].id); pairs.push({ '首都圏の商品ID': a.id, '北関東の商品ID': hit[0].id }); }
    });
    const paired = new Set();
    pairs.forEach((x) => { paired.add(x['首都圏の商品ID']); paired.add(x['北関東の商品ID']); });
    onlyS.concat(onlyK).forEach((p) => {
      if (!paired.has(p.id)) D('商品 ' + p.id, p.row['地域'] + ' の商品 ' + p.row['呼出番'] + ' ' + nk(p.row['表示名（印刷どおり）']) + ' の、もう一方の地域の組が見つからない', '');
    });

    if (styleMiss.length) {
      problems.push({ level: LEVEL_CHECK, what: 'アプリの書式に無い見た目（字・大きさ・色・塗りの組）のセルが ' + styleMiss.length + ' か所あります。そのセルは型の見た目で出ます', where: styleMiss.slice(0, 10).join('、') + (styleMiss.length > 10 ? ' ほか' : '') });
    }
    if (notInType.size) {
      problems.push({ level: LEVEL_CHECK, what: 'アプリの表の型に無い列なので、見た目を見ていない列があります（番号表には出ない列）: ' + Array.from(notInType).join('、'), where: '' });
    }
    return { products: prodList.map((p) => p.row), sections, layout, monthly: mrows, pairs, doubts };
  }

  function matchProduct(prodList, hc, name) {
    if (hc) {
      const hits = prodList.filter((p) => p.row['呼出番'] === hc);
      if (hits.length === 1) return hits[0].id;
      if (hits.length > 1) { const k = nameKey(name); const hh = hits.filter((p) => nameKey(p.row['表示名（印刷どおり）']) === k); if (hh.length === 1) return hh[0].id; }
    }
    const k = nameKey(name);
    const hits = prodList.filter((p) => nameKey(p.row['表示名（印刷どおり）']) === k);
    return hits.length === 1 ? hits[0].id : null;
  }

  // ================================================================ 月の束にする（make_seed.build と同じ。この部品に要る表だけ）
  function toMonthTables(M, tables, month) {
    const prods = M.products, secs = M.sections, lay = M.layout, mon = M.monthly;
    const pmap = new Map(prods.map((p) => [p['商品ID'], p]));
    const secmap = new Map(secs.map((s) => [s['区分ID'], s]));
    const idTail = (id) => { const s = String(id); return s.slice(s.indexOf('-') + 1); };
    const layBySec = new Map();
    lay.forEach((x) => { if (!layBySec.has(x['区分ID'])) layBySec.set(x['区分ID'], []); layBySec.get(x['区分ID']).push(x); });
    layBySec.forEach((list) => list.sort((a, b) => a['何番目'] - b['何番目']));

    // ---- 位置を歩いて、セルの字の飾り（字ごとの色・大きさ・太さ）を、商品・区分・続きの行に結ぶ
    const prodRich = new Map(), secRich = new Map(), rowRich = new Map(), usedCells = new Set();
    const getOr = (m, k) => { if (!m.has(k)) m.set(k, {}); return m.get(k); };
    TABLE_KEYS.forEach((t) => {
      const T = tables[t];
      const ws = T.ws, hr = T.header_row, c1 = T.print[0], c2 = T.print[2];
      const runsAt = (r, c) => { const x = ws.cell(r, c); return x && x.runs && x.runs.length ? x.runs : null; };
      T.blocks.forEach((b) => {
        // 役割の列（印刷範囲の中の列だけ・列の順）
        const roles = Object.keys(b.roles).map((role) => [roleJa(b, role), b.roles[role]]).filter((x) => x[1] >= c1 && x[1] <= c2).sort((x, y) => x[1] - y[1]);
        const nameCol = (roles.find((x) => x[0] === '名前') || [])[1];
        let r = hr + 1;
        secs.filter((s) => s['表'] === t && s['列のまとまり'] === b.index + 1).sort((x, y) => x['まとまりの中の順'] - y['まとまりの中の順']).forEach((s) => {
          if (s['表題の置き場'] === '見出しの行') {
            if (nameCol === undefined) throw stop('まとまり「' + b.name + '」の名前の列が印刷範囲の外です', sheetWhere(ws));
            usedCells.add(t + ',' + hr + ',' + nameCol);
            const rr = runsAt(hr, nameCol);
            if (rr) getOr(secRich, s['区分ID'])['名前'] = rr;
          } else {
            roles.forEach(([role, c]) => { usedCells.add(t + ',' + r + ',' + c); const rr = runsAt(r, c); if (rr) getOr(secRich, s['区分ID'])[role] = rr; });
            r++;
          }
          (layBySec.get(s['区分ID']) || []).forEach((it) => {
            roles.forEach(([role, c]) => {
              usedCells.add(t + ',' + r + ',' + c);
              const rr = runsAt(r, c);
              if (!rr) return;
              if (it['行の種類'] === KIND_PRODUCT) { const kept = role === '名前' ? dropPink(rr) : rr; if (kept.length) getOr(prodRich, it['商品ID'] + '|' + t)[role] = kept; }
              else getOr(rowRich, t + '|' + it['区分ID'] + '|' + it['何番目'])[role] = rr;
            });
            r++;
          });
        });
      });
    });
    // 数えて確かめる（結べなかった字の飾りのセルがあれば止める。見出しより上の決まったセルは表の型の物なので数えない）
    const lost = [];
    TABLE_KEYS.forEach((t) => {
      const T = tables[t];
      const ws = T.ws, hr = T.header_row, [c1, r1, c2, r2] = T.print;
      const headNames = new Set(T.blocks.map((b) => b.roles.name));
      Object.keys(ws.cells).forEach((k) => {
        const x = ws.cells[k];
        if (!x.runs || !x.runs.length) return;
        const [r, c] = k.split(',').map(Number);
        if (r < r1 || r > r2 || c < c1 || c > c2) return;
        if (r < hr || (r === hr && !headNames.has(c))) return;
        if (usedCells.has(t + ',' + r + ',' + c)) return;
        lost.push(stopItem('字の飾りのセルを、商品・区分・続きの行に結べませんでした', cellWhere(ws, r, c)));
      });
    });
    if (lost.length) throw new StopRead(lost);

    // ---- 商品（1商品1行）と原価（別の表。はじめの読み込みに入れないため）
    const pairOf = new Map();
    M.pairs.forEach((p) => { pairOf.set(p['首都圏の商品ID'], p['北関東の商品ID']); pairOf.set(p['北関東の商品ID'], p['首都圏の商品ID']); });
    const P_COLS = ['商品ID', '表の種類', '地域', '対の商品ID', '呼出番', '名前', '字の飾り', '指定トレー', '定番価格', '小パック', 'メモ', '元の商品ID'];
    const C_COLS = ['商品ID'].concat(COST_COLS);
    const pRows = [], cRows = [];
    prods.forEach((p) => {
      const pid = p['商品ID'];
      const kind = String(p['表']).indexOf('焼き物') === 0 ? '焼き物' : '寿司弁当';
      const baseT = kind === '焼き物' ? '焼き物' : (p['地域'] === '北関東だけ' ? '北関東' : '首都圏');
      const deco = Object.assign({}, prodRich.get(pid + '|' + baseT) || {});
      if (kind === '寿司弁当' && p['地域'] === '共通') {
        const other = prodRich.get(pid + '|北関東') || {};
        if (!deepEq(other, prodRich.get(pid + '|首都圏') || {})) deco['北関東'] = other;
      }
      pRows.push({ '商品ID': pid, '表の種類': kind, '地域': p['地域'], '対の商品ID': pairOf.has(pid) ? pairOf.get(pid) : '', '呼出番': p['呼出番'], '名前': p['表示名（印刷どおり）'], '字の飾り': Object.keys(deco).length ? pyJson(deco) : '', '指定トレー': normTray(p['指定トレー']), '定番価格': p['見える価格'], '小パック': p['小パック'], 'メモ': p['メモ'] });
      const cr = { '商品ID': pid };
      COST_COLS.forEach((k) => { cr[k] = p[k]; });
      cRows.push(cr);
    });

    // ---- 月の行（1商品1行。首都圏と北関東で違う所だけ「北関東の違い」に）
    const byTable = {};
    lay.forEach((r) => { const t = r['表'], k = idTail(r['区分ID']); byTable[t] = byTable[t] || {}; (byTable[t][k] = byTable[t][k] || []).push(r); });
    Object.keys(byTable).forEach((t) => Object.keys(byTable[t]).forEach((k) => byTable[t][k].sort((a, b) => a['何番目'] - b['何番目'])));
    const R_COLS = ['行ID', '区分ID', '順', '商品ID', '地域', '販売形態', '特売価格', 'つなぐ', '続き', '見た目', '北関東の違い', '元の区分ID', '元の順'];
    const S_COLS = ['区分ID', '表の種類', '表題', '表題の続き', '表題の書式', '表題の置き場（元）', '見た目', '字の飾り', '北関東の違い'];
    const mRows = [], sRows = [], placeRows = [];
    let seq = 0;
    const formOf = (it) => {
      const sale = it['月間特売'] === '○', fresh = it['新商品'] === '○';
      if (sale || fresh) return sale && fresh ? '月間特売・新商品' : (sale ? '月間特売' : '新商品');
      // ≪カット商品≫ の区分にある定番の商品は、販売形態「カット商品」（PO 2026-11-06）
      const sec = secmap.get(it['区分ID']);
      return sec && squash(sec['表題（印刷どおり）']).indexOf('カット商品') >= 0 ? 'カット商品' : '定番';
    };
    const regionOf = (it) => (it['行の種類'] === KIND_PRODUCT ? pmap.get(it['商品ID'])['地域'] : '');
    function emit(secId, order, it, region, other) {
      seq++;
      const rid = 'R' + pad(seq, 4);
      const isProduct = it['行の種類'] === KIND_PRODUCT;
      const row = { '行ID': rid, '区分ID': secId, '順': order, '商品ID': pyOr(it['商品ID'], ''), '地域': isProduct || region === '共通' ? '' : region,
        '販売形態': isProduct ? formOf(it) : '', '特売価格': it['その月の見える価格'], 'つなぐ': pyOr(it['名前のセルを上とつなぐ'], ''), '見た目': pyOr(it['書式を変える欄'], ''),
        '元の区分ID': secId, '元の順': order };
      const diff = {};
      if (other) {
        [['書式を変える欄', '見た目'], ['その月の見える価格', '特売価格']].forEach(([k, key]) => { const a = pyOr(it[k], ''), b = pyOr(other[k], ''); if (!pyEq(a, b)) diff[key] = b; });
        if (isProduct && formOf(it) !== formOf(other)) diff['販売形態'] = formOf(other);
      }
      if (Object.keys(diff).length) row['北関東の違い'] = pyJson(diff);
      if (it['行の種類'] === KIND_CONT) {
        const v = it['続き: 見える価格'];
        const cont = { '呼出番の欄': pyOr(it['続き: 呼出番の欄'], ''), '見える価格': v !== null && v !== undefined ? v : '', '小パック': pyOr(it['続き: 小パック'], '') };
        cRows.push({ '商品ID': rid, '原価': it['続き: 原価'], '粗利（包材値上前）': it['続き: 粗利（包材値上前）'], '定番価格（包材値上前）': it['続き: 定番価格（包材値上前）'], '定番価格（包材変更後）': it['続き: 定番価格（包材変更後）'] });
        ['焼き物', '首都圏', '北関東'].forEach((tt) => { const rr = rowRich.get(tt + '|' + it['区分ID'] + '|' + it['何番目']); if (rr) cont['字の飾り'] = rr; });
        row['続き'] = pyJson(cont);
      }
      mRows.push(row);
    }
    [['首都圏', '寿司弁当'], ['焼き物', '焼き物']].forEach(([t, kindKey]) => {
      const ss = secs.filter((s) => s['表'] === t).sort((x, y) => (x['列のまとまり'] - y['列のまとまり']) || (x['まとまりの中の順'] - y['まとまりの中の順']));
      const seqBlock = {};
      ss.forEach((s) => {
        const numId = idTail(s['区分ID']);
        const sid = kindKey + '-' + numId;
        const sdiff = {};
        if (t === '首都圏') {
          const o = secmap.get('北関東-' + numId);
          if (!o || nk(o['表題（印刷どおり）']) !== nk(s['表題（印刷どおり）']) || o['列のまとまり'] !== s['列のまとまり'] || o['まとまりの中の順'] !== s['まとまりの中の順']) {
            throw stop('首都圏と北関東で区分の置き場が違います: ' + s['区分ID'] + '「' + nk(s['表題（印刷どおり）']) + '」', '番号表（首都圏・北関東）');
          }
          [['表題（印刷どおり）', '表題'], ['表題の書式', '表題の書式'], ['表題の行の書式を変える欄', '見た目'], ['表題の続き（呼出番の欄）', '表題の続き']].forEach(([k, key]) => {
            if (!pyEq(pyOr(o[k], ''), pyOr(s[k], ''))) sdiff[key] = pyOr(o[k], '');
          });
          const ra = secRich.get(s['区分ID']), rb = secRich.get(o['区分ID']);
          if (!deepEq(ra, rb)) sdiff['字の飾り'] = rb || [];
        }
        const deco = secRich.get(s['区分ID']);
        sRows.push({ '区分ID': sid, '表の種類': kindKey, '表題': s['表題（印刷どおり）'], '表題の続き': s['表題の続き（呼出番の欄）'], '表題の書式': s['表題の書式'], '表題の置き場（元）': s['表題の置き場'], '見た目': s['表題の行の書式を変える欄'], '字の飾り': deco && Object.keys(deco).length ? pyJson(deco) : '', '北関東の違い': Object.keys(sdiff).length ? pyJson(sdiff) : '' });
        seqBlock[s['列のまとまり']] = (seqBlock[s['列のまとまり']] || 0) + 10;
        placeRows.push({ '区分ID': sid, 'まとまり': s['列のまとまり'], '順': seqBlock[s['列のまとまり']], '元のまとまり': s['列のまとまり'], '元の順': seqBlock[s['列のまとまり']] });
        const a = (byTable[t] || {})[numId] || [];
        const b = t === '首都圏' ? ((byTable['北関東'] || {})[numId] || []) : [];
        let order = 0;
        if (t === '焼き物') { a.forEach((it) => { order += 10; emit(sid, order, it, '共通'); }); return; }
        let i = 0, j = 0;
        while (i < a.length || j < b.length) {
          const x = i < a.length ? a[i] : null, y = j < b.length ? b[j] : null;
          order += 10;
          if (x && y && x['行の種類'] === y['行の種類'] && pyOr(x['商品ID'], '') === pyOr(y['商品ID'], '')) { emit(sid, order, x, '共通', y); i++; j++; }
          else if (x && regionOf(x) === '首都圏だけ') { emit(sid, order, x, '首都圏だけ'); i++; }
          else if (y && regionOf(y) === '北関東だけ') { emit(sid, order, y, '北関東だけ'); j++; }
          else if (x && x['行の種類'] === KIND_BLANK) { emit(sid, order, x, '首都圏だけ'); i++; }
          else if (y && y['行の種類'] === KIND_BLANK) { emit(sid, order, y, '北関東だけ'); j++; }
          else throw stop('首都圏と北関東の並びを1つにできません: ' + s['区分ID'] + ' の ' + (i + 1) + ' 番目', '番号表（首都圏・北関東）');
        }
      });
    });

    // ---- 月間商品（商品ID で結ぶ）
    const MON_COLS = ['行ID', '部', '順', '行の種類', '商品名', '呼出番', '規格', '本体価格', '小パックの規格', '小パックの本体価格', '商品ID', '見た目'];
    const monRows = mon.map((g, i) => ({ '行ID': 'G' + pad(i + 1, 3), '部': g['部'], '順': g['何番目'] * 10, '行の種類': g['行の種類'], '商品名': g['商品名（この表の書き方）'], '呼出番': g['呼出番（この表に書く番号）'], '規格': g['規格'], '本体価格': g['本体価格'], '小パックの規格': g['小パックの規格'], '小パックの本体価格': g['小パックの本体価格'], '商品ID': g['商品ID'], '見た目': g['書式を変える欄'] }));

    const table = (cols, rows) => ({ c: cols, r: rows.map((r) => cols.map((c) => (r[c] === null || r[c] === undefined ? '' : r[c]))) });
    const out = {};
    out['商品'] = table(P_COLS, pRows);
    out['原価'] = table(C_COLS, cRows);
    out['区分'] = table(S_COLS, sRows);
    out['月の行_' + month] = table(R_COLS, mRows);
    out['月の置き場_' + month] = table(['区分ID', 'まとまり', '順', '元のまとまり', '元の順'], placeRows);
    out['月の状態'] = table(['月', '表の種類', '形'], ['寿司弁当', '焼き物'].map((k) => ({ '月': month, '表の種類': k, '形': '自動' })));
    out['月間商品_' + month] = table(MON_COLS, monRows);
    out['疑問'] = table(['番号', 'どこ', '何が', '手がかり', '済み'], M.doubts);
    return out;
  }

  // ================================================================ 入口
  /**
   * 手作りの商品呼出番号表のブック → 月の束。
   * @param book  Book（frontend/book-read.js の形）
   * @param opts  { month: 'YYYY-MM'（あれば、ブックの月と合うか確かめる。ブックに年が無いときは、ここから年を取る）,
   *                typeTables: アプリにいまある表の型の表 { '書式', '表の型_列', '表の型', '表の型_まとまり' }（{c, r} の形）。
   *                            見た目（区分の表題の書式・見た目）の書式キーは、この「書式」の中から選ぶ。無ければ見た目は空 }
   * @returns { month, tables（止めたときは null）, problems: [{ level: '止める'|'確かめる', what, where }], counts（読んだ数） }
   */
  function readTableBook(book, opts) {
    opts = opts || {};
    const problems = [];
    let month = null;
    try {
      if (!book || !Array.isArray(book.sheets)) throw stop('ブックの形が違います（sheets がありません）', '');
      // 見出しの並びは4枚とも確かめてから止める（違う所をまとめて返す）
      const tables = {}, layoutErrors = [];
      TABLE_KEYS.forEach((k) => { try { tables[k] = readNumberTable(book, k); } catch (e) { if (e instanceof StopRead) layoutErrors.push.apply(layoutErrors, e.items); else throw e; } });
      let monthly = null;
      try { monthly = readMonthly(book); } catch (e) { if (e instanceof StopRead) layoutErrors.push.apply(layoutErrors, e.items); else throw e; }
      if (layoutErrors.length) throw new StopRead(layoutErrors);
      const mo = monthOf(book, monthly, opts.month);
      month = mo.month;
      const ctx = styleContext(opts.typeTables);
      if (!ctx) problems.push({ level: LEVEL_CHECK, what: 'アプリの表の型（書式・表の型_列）が渡されていないので、見た目（区分の表題の書式・見た目）を空にしました', where: '' });
      else if (!TABLE_KEYS.every((k) => tables[k].ws.raw.x0) || !monthly.ws.raw.x0) {
        problems.push({ level: LEVEL_CHECK, what: 'ブックに罫線・位置・折り返し・表示形式が無いので、見た目（書式キー）は字と塗りだけで当てました（値も塗りも無いセルは型どおり）。元と細かい所が違うことがあります', where: '' });
      }
      const M = buildMaster(tables, monthly, mo.mNum, ctx, problems);
      const out = toMonthTables(M, tables, month);
      const counts = { '商品の行': {}, '区分': {}, '形': {} };
      TABLE_KEYS.forEach((k) => {
        counts['商品の行'][k] = tables[k].lines.filter((l) => l.kind === 'product' && l.place === '印刷範囲の中').length;
        counts['区分'][k] = M.sections.filter((s) => s['表'] === k).length;
        counts['形'][k] = tables[k].layoutName;
      });
      return { month, tables: out, problems, counts };
    } catch (e) {
      if (!(e instanceof StopRead)) throw e;
      return { month, tables: null, problems: problems.concat(e.items), counts: null };
    }
  }

  root.SouzaiImportTable = { readTableBook, normTray, LAYOUTS, LEVEL_STOP, LEVEL_CHECK };
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
