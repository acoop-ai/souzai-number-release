// 商品呼出番号表の「決まり」の部品（画面の部品を持たない）。画面・窓口（gas/Core.gs に同じ物を写す）・試験で同じ物を使う。
//   置き場の計算（自動の決め直し）・動かす（ドラッグも「ここに移す」もここだけ）・アラート・入数と字の飾り・原価の式。
//   信頼境界: ここへ来る値は、人が打った文字。式は「数字と + − × ÷ と括弧」だけ計算する（eval を使わない）。
(function (root) {
  'use strict';
  const TABLES = { '首都圏': { kind: '寿司弁当', region: '首都圏' }, '北関東': { kind: '寿司弁当', region: '北関東' }, '焼き物': { kind: '焼き物', region: '' } };
  const TABLE_KEYS = ['首都圏', '北関東', '焼き物'];
  const KIND_TABLES = { '寿司弁当': ['首都圏', '北関東'], '焼き物': ['焼き物'] };
  const KINDS = ['寿司弁当', '焼き物'];
  const KIND_PRODUCT = '商品', KIND_BLANK = '空き行', KIND_CONT = '続きの行', KIND_SECTION = '節の見出し';
  // 販売形態（PO 2026-11-05）: 1つの欄で選ぶ6つ。定番・月間特売・新商品は表に載る。休み・終売・計画は載らない（番号の管理のために一覧には残す）。
  //   番号表の色（ピンクの字・水色のセル）もここから付く。月間特売と新商品だけ両方選べる（Q9・10月は12商品）。欄は1つのまま、両方のときの値が「月間特売・新商品」
  //   カット商品（PO 2026-11-06）: 表に載る。行は ≪カット商品≫ の区分へ動く（表の種類をまたぐ）。1つだけ選ぶ。次の月を作ると終売になり、元の区分へ戻る
  const FORM_REGULAR = '定番', FORM_SALE = '月間特売', FORM_NEW = '新商品', FORM_SALE_NEW = '月間特売・新商品', FORM_CUT = 'カット商品', FORM_REST = '休み', FORM_END = '終売', FORM_PLAN = '計画';
  const FORMS = [FORM_REGULAR, FORM_SALE, FORM_NEW, FORM_CUT, FORM_REST, FORM_END, FORM_PLAN];
  const FORMS_ALL = [FORM_REGULAR, FORM_SALE, FORM_NEW, FORM_SALE_NEW, FORM_CUT, FORM_REST, FORM_END, FORM_PLAN];
  const isOffForm = (f) => f === FORM_REST || f === FORM_END || f === FORM_PLAN;
  const isSaleForm = (f) => f === FORM_SALE || f === FORM_SALE_NEW;
  const isNewForm = (f) => f === FORM_NEW || f === FORM_SALE_NEW;
  const isCutForm = (f) => f === FORM_CUT;
  /** その項目に印が付いているか（両方のときは 月間特売 と 新商品 の2つに付く） */
  const formHas = (current, item) => current === item || (current === FORM_SALE_NEW && (item === FORM_SALE || item === FORM_NEW));
  /** 押したあとの値: 月間特売と新商品は付け外しで、両方にできる。定番・休み・終売・計画は1つだけで、選ぶとほかは外れる。全部外れたら定番 */
  function toggleForm(current, picked) {
    if (picked !== FORM_SALE && picked !== FORM_NEW) return picked;
    let sale = isSaleForm(current), fresh = isNewForm(current);
    if (picked === FORM_SALE) sale = !sale; else fresh = !fresh;
    return sale && fresh ? FORM_SALE_NEW : (sale ? FORM_SALE : (fresh ? FORM_NEW : FORM_REGULAR));
  }
  const COST_KEYS = ['原価', '粗利（包材値上前）', '定番価格（包材値上前）', '定番価格（包材変更後）'];
  const GROSS_LOW = 0.25, GROSS_HIGH = 0.70; // 粗利（包材変更後）の目安（読み・設定で変えられる）
  const LEVEL_FIX = '直す', LEVEL_CHECK = '確かめる';
  const ORDER_STEP = 10;

  const isBlank = (v) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  const nfkc = (v) => String(v === null || v === undefined ? '' : v).normalize('NFKC');
  const toFull = (s) => String(s).replace(/[0-9]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0xFEE0));
  /** 呼出番を半角4桁にそろえる。空なら ''。4桁の数字でなければ null */
  function halfCall(v) {
    if (typeof v === 'number' && Number.isInteger(v)) v = String(v);
    const s = nfkc(v).replace(/[\r\n]/g, ' ').trim();
    if (s === '') return '';
    return /^\d{4}$/.test(s) ? s : null;
  }
  /** マスターは半角・番号表は全角（元の書類と同じ見た目）。4桁でない文字は、そのまま */
  function fullDigits(v) { if (v === null || v === undefined) return null; const s = String(v); return halfCall(s) ? toFull(s) : s; }
  function num(v) {
    if (typeof v === 'boolean') return null;
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    const s = nfkc(v).trim(); if (s === '') return null;
    const n = Number(s); return Number.isFinite(n) ? n : null;
  }

  /** 包材（指定トレー）の字をそろえる部品（ここ1か所。画面と窓口と種で同じ決まり）: 英数字・記号・空白は半角へ（全角空白も）・前後の空白は取る・ハイフンのゆれは - へ・半角カナは全角へ。
   *  長音の「ー」は、かなの続きならそのまま（トレー）。かなの続きでなければハイフンのつもりなので - へ（19ー15 → 19-15） */
  function normTray(v) {
    return nfkc(v).replace(/[\r\n\t]/g, ' ').replace(/[\u2010-\u2015\u2212\uFE58\uFE63\uFF0D]/g, '-')
      .replace(/\u30FC/g, (m, at, all) => (at > 0 && /[\u3041-\u3096\u30A1-\u30FA\u30FC]/.test(all[at - 1]) ? m : '-')).trim();
  }
  /** 足した商品の印: 元の商品ID に元の ID（写したとき）か、この印（新しく足したとき）が入る。最初の取り込みの商品は空 */
  const ADDED_MARK = '新規';
  const isAddedProduct = (p) => Boolean(p) && !isBlank(p['元の商品ID']);

  // ---------------------------------------------------------------- 原価の式（数字と + − × ÷ と括弧だけ）
  /** 数か、=数字の式 なら値。それ以外は null */
  function evalArith(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    const s = String(v).trim();
    if (s[0] !== '=') return num(s);
    const src = s.slice(1).replace(/\s+/g, '');
    if (!/^[0-9.+\-*/()]+$/.test(src) || src.length > 200) return null;
    let i = 0;
    function expr() { let a = term(); while (src[i] === '+' || src[i] === '-') { const o = src[i++]; const b = term(); a = o === '+' ? a + b : a - b; } return a; }
    function term() { let a = unary(); while (src[i] === '*' || src[i] === '/') { const o = src[i++]; const b = unary(); if (o === '/' && b === 0) throw new Error('div0'); a = o === '*' ? a * b : a / b; } return a; }
    function unary() { if (src[i] === '+') { i++; return unary(); } if (src[i] === '-') { i++; return -unary(); } if (src[i] === '(') { i++; const a = expr(); if (src[i] !== ')') throw new Error('paren'); i++; return a; } const m = /^[0-9]*\.?[0-9]+/.exec(src.slice(i)); if (!m) throw new Error('num'); i += m[0].length; return Number(m[0]); }
    try { const out = expr(); return i === src.length && Number.isFinite(out) ? out : null; } catch (e) { return null; }
  }
  /** 選んでいないセルに見せる字（Excel と同じ）: 数字の式は計算した数（小数2けたまで）。読めない式と、それ以外はそのまま */
  function showCell(v) { if (v === null || v === undefined) return ''; if (typeof v === 'string' && v.trim()[0] === '=') { const n = evalArith(v); return n === null ? v : String(Math.round(n * 100) / 100); } return String(v); }
  /** 原価の欄に受け付ける形か（空・数・数字の式） */
  function costFormOk(v) { return isBlank(v) || evalArith(v) !== null; }
  /** 粗利（包材変更後）の割合（計算できなければ null） */
  function grossAfter(cost, price) { const c = evalArith(cost), p = evalArith(price); return c !== null && p ? (p - c) / p : null; }

  // ---------------------------------------------------------------- 字の飾り（[[from, to, {c,s,b,i,u,f}], …]。to は含まない）
  function parseDeco(text) { if (!text) return {}; if (typeof text === 'object') return text; try { const o = JSON.parse(text); return o && typeof o === 'object' ? o : {}; } catch (e) { return {}; } }
  /** 文字を打ち直したとき、飾りを字についていかせる（同じ頭と同じ尾を残して、間だけ変わったとみなす） */
  function editRuns(runs, oldText, newText) {
    if (!runs || !runs.length) return [];
    let a = 0; const oldLen = oldText.length, newLen = newText.length;
    while (a < oldLen && a < newLen && oldText[a] === newText[a]) a++;
    let b = 0; while (b < oldLen - a && b < newLen - a && oldText[oldLen - 1 - b] === newText[newLen - 1 - b]) b++;
    const delta = newLen - oldLen, oldEnd = oldLen - b;
    const mapPos = (p, isEnd) => { if (p <= a) return p; if (p >= oldEnd) return p + delta; return isEnd ? a + (newLen - a - b) : a; };
    const out = [];
    runs.forEach((r) => { const f = mapPos(r[0], false), t = mapPos(r[1], true); if (t > f) out.push([f, t, r[2]]); });
    return out;
  }
  /** from〜to に飾りを足す（重なる所は分けて上書き）。props=null で飾りを外す */
  function paintRuns(runs, from, to, props) {
    const out = [];
    (runs || []).forEach((r) => {
      if (r[1] <= from || r[0] >= to) { out.push(r); return; }
      if (r[0] < from) out.push([r[0], from, r[2]]);
      if (r[1] > to) out.push([to, r[1], r[2]]);
    });
    if (props && to > from) out.push([from, to, props]);
    return out.sort((x, y) => x[0] - y[0]);
  }

  // ---------------------------------------------------------------- 入数（名前の中のカッコと連動。列には持たない＝二重にしない）
  const QTY_RE = /([(（])\s*([0-9０-９]+)\s*([^\d０-９()（）]{1,4}?)\s*([)）])\s*$/;
  function parseQty(name) {
    const s = String(name === null || name === undefined ? '' : name);
    const m = QTY_RE.exec(s);
    if (!m) return { has: false, count: '', unit: '' };
    return { has: true, count: nfkc(m[2]), unit: m[3], start: m.index, end: s.length };
  }
  /** 入数を変えた名前を返す（全角のカッコ・全角の数字が決まった形。数も単位も空ならカッコを外す） */
  function setQty(name, count, unit) {
    const s = String(name === null || name === undefined ? '' : name);
    const q = parseQty(s);
    const c = nfkc(count).trim(), u = String(unit === null || unit === undefined ? '' : unit).trim();
    const part = c && u ? '（' + toFull(c) + u + '）' : '';
    if (q.has) return s.slice(0, q.start) + part;
    return s + part;
  }

  // ---------------------------------------------------------------- データの形（窓口の返事 → 手元の入れ物）
  function rowsOf(t) { if (!t) return []; if (Array.isArray(t)) return t; const c = t.c || [], out = []; (t.r || []).forEach((r) => { const o = {}; c.forEach((k, i) => { o[k] = r[i]; }); out.push(o); }); return out; }
  /** 列に持たない「行の種類」は、中身から決める（商品ID があれば商品・続きがあれば続きの行・ほかは空き行） */
  function prepRow(r) { r._diff = parseDeco(r['北関東の違い']); r._cont = parseDeco(r['続き']); r._deco = (r._cont && r._cont['字の飾り']) || {}; r['行の種類'] = r['商品ID'] ? KIND_PRODUCT : (r['続き'] ? KIND_CONT : KIND_BLANK); return r; }
  /** tables: { 商品, 区分, 月の行, 月の置き場, 月の状態, 月間商品, 選べる値, 過去の番号, 疑問, 確かめ済み }（compact か、行の配列）。type: 表の型 */
  function makeDb(tables, type, month) {
    const db = { month: month || '', products: new Map(), rows: new Map(), sections: new Map(), places: new Map(), status: {}, monthly: [], options: {}, past: [], doubts: [], checked: new Set(), type: null, costs: null, costInfo: { empty: [], grossOut: [], badForm: [] } };
    rowsOf(tables['商品']).forEach((p) => { p._deco = parseDeco(p['字の飾り']); db.products.set(String(p['商品ID']), p); });
    rowsOf(tables['月の行']).forEach((r) => { prepRow(r); db.rows.set(String(r['行ID']), r); });
    rowsOf(tables['区分']).forEach((s) => { s._diff = parseDeco(s['北関東の違い']); s._deco = parseDeco(s['字の飾り']); db.sections.set(String(s['区分ID']), s); });
    rowsOf(tables['月の置き場']).forEach((p) => { db.places.set(String(p['区分ID']), p); });
    rowsOf(tables['月の状態']).forEach((s) => { if (!month || String(s['月']) === month) db.status[s['表の種類']] = s['形'] || '自動'; });
    db.monthly = rowsOf(tables['月間商品']);
    rowsOf(tables['選べる値']).forEach((o) => { (db.options[o['項目']] = db.options[o['項目']] || []).push(o); });
    db.past = rowsOf(tables['過去の番号']);
    db.doubts = rowsOf(tables['疑問']);
    rowsOf(tables['確かめ済み']).forEach((c) => { db.checked.add(String(c['鍵'])); });
    if (type) setType(db, type);
    return db;
  }
  function setType(db, type) {
    const shapes = {}, blocks = {}, columns = {}, fixed = {}, styles = {}, pictures = {};
    rowsOf(type['表の型']).forEach((s) => { shapes[s['表']] = s; });
    rowsOf(type['表の型_まとまり']).forEach((b) => { (blocks[b['表']] = blocks[b['表']] || []).push(b); });
    rowsOf(type['表の型_列']).forEach((c) => { (columns[c['表']] = columns[c['表']] || []).push(c); });
    rowsOf(type['表の型_決まったセル']).forEach((f) => { (fixed[f['表']] = fixed[f['表']] || []).push(f); });
    rowsOf(type['書式']).forEach((s) => { styles[s['書式キー']] = s; });
    rowsOf(type['表の型_絵']).forEach((p) => { pictures[p['名前']] = p['中身']; });
    db.type = { shapes, blocks, columns, fixed, styles, pictures };
  }
  /** 列のまとまりの中で、役割 → 列 */
  function blockRoles(db, table, blockNo) {
    const out = {}; (db.type.columns[table] || []).forEach((c) => { if (String(c['まとまり']) === String(blockNo) && c['役割']) out[c['役割']] = c; });
    return out;
  }
  function blockList(db, table) { return (db.type.blocks[table] || []).slice().sort((a, b) => Number(a['まとまり']) - Number(b['まとまり'])); }

  // ---------------------------------------------------------------- 表ごとの見え方
  /** 北関東の表なら「北関東の違い」を当てた値 */
  function rowValue(row, key, table) { if (table === '北関東' && row._diff && Object.prototype.hasOwnProperty.call(row._diff, key)) return row._diff[key]; return row[key]; }
  function secValue(sec, key, table) { if (table === '北関東' && sec._diff && Object.prototype.hasOwnProperty.call(sec._diff, key)) return sec._diff[key]; return sec[key]; }
  /** 商品の行は 商品 の地域を見る（行には持たない）。空き行・続きの行だけ、行の地域を見る */
  function rowRegion(db, row) { if (row['商品ID']) { const p = db.products.get(String(row['商品ID'])); return (p && p['地域']) || '共通'; } return row['地域'] || '共通'; }
  function rowRegionOk(db, row, table) { const reg = rowRegion(db, row); const t = TABLES[table]; if (!t.region || reg === '共通') return true; return reg === t.region + 'だけ'; }
  function formOf(row, table) { return rowValue(row, '販売形態', table) || ''; }
  // 表に出る価格（PO 2026-11-05）: 特売価格はいつでも入れられる。販売形態は「どちらの価格が表に出るか」を決めるだけ
  /** その表に出る価格の欄。月間特売が入っていて特売価格があれば 特売価格・それ以外は 定番価格（首都圏と北関東で販売形態が違えば、その表の販売形態で） */
  function shownPriceKey(row, table) { return isSaleForm(formOf(row, table)) && !isBlank(rowValue(row, '特売価格', table)) ? '特売価格' : '定番価格'; }
  function shownPrice(p, row, table) { return shownPriceKey(row, table) === '特売価格' ? rowValue(row, '特売価格', table) : p['定番価格']; }
  function rowShown(db, row, table) { return !isOffForm(formOf(row, table)) && rowRegionOk(db, row, table); }
  /** どちらかの表に載る行か（休み・終売・計画でない） */
  function rowOn(row) { return !isOffForm(row['販売形態'] || '') || !isOffForm(rowValue(row, '販売形態', '北関東') || ''); }
  const byOrder = (a, b) => Number(a['順']) - Number(b['順']);
  function sectionsOfKind(db, kind) { return Array.from(db.sections.values()).filter((s) => s['表の種類'] === kind && db.places.has(String(s['区分ID']))); }
  /** 区分の行（全部・順に並べた物） */
  function rowsOfSection(db, secId) { const out = []; db.rows.forEach((r) => { if (String(r['区分ID']) === String(secId)) out.push(r); }); return out.sort(byOrder); }
  /** 区分ID → その区分の行（順に）。全部の区分をまとめて1回で分ける（数百行を区分ごとになめ直さない） */
  function rowsBySection(db) {
    const m = new Map();
    db.rows.forEach((r) => { const k = String(r['区分ID']); if (!m.has(k)) m.set(k, []); m.get(k).push(r); });
    m.forEach((a) => a.sort(byOrder));
    return m;
  }
  function placesOfBlock(db, kind, blockNo) {
    return sectionsOfKind(db, kind).map((s) => db.places.get(String(s['区分ID']))).filter((p) => String(p['まとまり']) === String(blockNo)).sort(byOrder);
  }
  /** 表の形: { table, kind, blocks:[{no, cap, name, sections:[{sec, rows:[row], place:'head'|'body'}], used, over}] } */
  function layoutOf(db, table, memoBySec) {
    const kind = TABLES[table].kind;
    const bySec = memoBySec || rowsBySection(db);
    const blocks = blockList(db, table).map((b) => {
      const no = Number(b['まとまり']);
      const secs = [];
      placesOfBlock(db, kind, no).forEach((p) => {
        const sec = db.sections.get(String(p['区分ID']));
        const rows = (bySec.get(String(p['区分ID'])) || []).filter((r) => rowShown(db, r, table));
        if (!rows.length) return; // 載せる行が無い区分は、表題も出さない
        secs.push({ sec, rows, place: secs.length === 0 ? 'head' : 'body' });
      });
      let used = 0;
      secs.forEach((s) => { used += s.rows.length + (s.place === 'head' ? 0 : 1); });
      const cap = Number(b['上限の行数']);
      return { no, cap, name: b['名前'], row0: Number(b['本文の最初の行']), first: b['最初の列'], last: b['最後の列'], sections: secs, used, over: Math.max(0, used - cap) };
    });
    return { table, kind, blocks };
  }

  // ---------------------------------------------------------------- 動かす（入り口が2つでも、処理はここだけ）
  function orderBetween(a, b) {
    if (a === null && b === null) return ORDER_STEP;
    if (b === null) return a + ORDER_STEP;
    if (a === null) return b - ORDER_STEP;
    return (a + b) / 2;
  }
  /** a と b の間に n 個入れる順の並び */
  function ordersBetween(a, b, n) {
    if (b === null || a === null) { const out = []; let at = a === null ? (b === null ? 0 : b - ORDER_STEP * (n + 1)) : a; for (let i = 0; i < n; i++) { at += ORDER_STEP; out.push(at); } return out; }
    const step = (b - a) / (n + 1), out = [];
    for (let i = 1; i <= n; i++) out.push(a + step * i);
    return out;
  }
  /** 区分をかたまりごと動かす。op={id, toBlock, before}（before＝その区分の前へ・null＝列の末尾）。返す値: 変える欄の一覧 */
  function moveSection(db, op) {
    const place = db.places.get(String(op.id));
    if (!place) throw new Error('区分が見つかりません');
    const sec = db.sections.get(String(op.id));
    const list = placesOfBlock(db, sec['表の種類'], op.toBlock).filter((p) => p !== place);
    let idx = list.length;
    if (op.before) { const at = list.findIndex((p) => String(p['区分ID']) === String(op.before)); if (at >= 0) idx = at; }
    const prev = idx > 0 ? Number(list[idx - 1]['順']) : null;
    const next = idx < list.length ? Number(list[idx]['順']) : null;
    const order = orderBetween(prev, next);
    const changes = [];
    if (String(place['まとまり']) !== String(op.toBlock) || Number(place['順']) !== order) changes.push({ t: '置き場', id: String(op.id), f: { 'まとまり': Number(op.toBlock), '順': order } });
    return changes;
  }
  /** 一緒に動く行（上とつないだ行・続きの行・地域の対の行） */
  function rowGroup(db, row) {
    const list = rowsOfSection(db, row['区分ID']);
    let i = list.indexOf(row);
    while (i > 0 && (list[i]['つなぐ'] === '○' || list[i]['行の種類'] === KIND_CONT)) i--;
    const group = [list[i]]; let j = i + 1;
    while (j < list.length && (list[j]['つなぐ'] === '○' || list[j]['行の種類'] === KIND_CONT)) group.push(list[j++]);
    const head = group[0];
    const p = head['商品ID'] ? db.products.get(String(head['商品ID'])) : null;
    if (p && p['対の商品ID']) { const mate = list.find((r) => String(r['商品ID']) === String(p['対の商品ID'])); if (mate && group.indexOf(mate) < 0) group.push(mate); }
    return group.sort(byOrder);
  }
  /** 商品の行を動かす。op={id（行ID）, toSec, before（行ID・null＝末尾）}。一緒に動く行もついてくる */
  function moveRow(db, op) {
    const row = db.rows.get(String(op.id));
    if (!row) throw new Error('行が見つかりません');
    if (!db.sections.has(String(op.toSec))) throw new Error('区分が見つかりません');
    const group = rowGroup(db, row);
    const list = rowsOfSection(db, op.toSec).filter((r) => group.indexOf(r) < 0);
    let idx = list.length;
    if (op.before) { const at = list.findIndex((r) => String(r['行ID']) === String(op.before)); if (at >= 0) idx = at; else if (group.some((g) => String(g['行ID']) === String(op.before))) return []; }
    const prev = idx > 0 ? Number(list[idx - 1]['順']) : null;
    const next = idx < list.length ? Number(list[idx]['順']) : null;
    const orders = ordersBetween(prev, next, group.length);
    const changes = [];
    group.forEach((g, i) => { const f = {}; if (String(g['区分ID']) !== String(op.toSec)) f['区分ID'] = String(op.toSec); if (Number(g['順']) !== orders[i]) f['順'] = orders[i]; if (Object.keys(f).length) changes.push({ t: '行', id: String(g['行ID']), f }); });
    return changes;
  }
  // ---------------------------------------------------------------- カット商品（≪カット商品≫ の区分へ動く）
  const squashText = (v) => nfkc(v).replace(/\s+/g, '');
  function isCutSection(sec) { return Boolean(sec) && squashText(sec['表題']).indexOf(FORM_CUT) >= 0; }
  /** その月の ≪カット商品≫ の区分（置き場にある物・無ければ null）。9月に寿司弁当の表にあった商品も、10月は焼き物の表の ≪カット商品≫ に載った＝表の種類をまたいでここへ集まる */
  function cutSectionOf(db) { for (const s of db.sections.values()) { if (isCutSection(s) && db.places.has(String(s['区分ID']))) return s; } return null; }
  /** 一緒に動く行（上とつないだ行・続きの行）。地域の対の商品は、販売形態が別なので入れない */
  function cutGroup(db, row) {
    const g = rowGroup(db, row);
    const head = g[0]; const p = head['商品ID'] ? db.products.get(String(head['商品ID'])) : null;
    const mate = p && p['対の商品ID'] ? String(p['対の商品ID']) : '';
    return g.filter((x) => x === row || !mate || String(x['商品ID']) !== mate);
  }
  /** 販売形態を変えるときの、変える欄（入口はここだけ）。カット商品にすると ≪カット商品≫ の区分の末尾へ、カット商品をやめると元の区分（元の区分ID・元の順の近く）へ、つながった行と一緒に動く。北関東だけの販売形態は消す */
  function formChanges(db, row, value) {
    const out = new Map();
    const put = (rid, f) => { const e = out.get(rid) || { t: '行', id: rid, f: {} }; Object.assign(e.f, f); out.set(rid, e); };
    const f = { '販売形態': value };
    if (row._diff && Object.prototype.hasOwnProperty.call(row._diff, '販売形態')) { const d = Object.assign({}, row._diff); delete d['販売形態']; f['北関東の違い'] = Object.keys(d).length ? JSON.stringify(d) : ''; }
    put(String(row['行ID']), f);
    const cut = cutSectionOf(db);
    if (!cut) return Array.from(out.values());
    const cutId = String(cut['区分ID']), was = row['販売形態'] || '';
    const group = cutGroup(db, row);
    if (isCutForm(value) && String(row['区分ID']) !== cutId) {
      const list = rowsOfSection(db, cutId);
      const orders = ordersBetween(list.length ? Number(list[list.length - 1]['順']) : null, null, group.length);
      group.forEach((g, i) => { const gf = { '区分ID': cutId, '順': orders[i] }; if (g !== row && g['商品ID']) gf['販売形態'] = FORM_CUT; put(String(g['行ID']), gf); });
    } else if (isCutForm(was) && !isCutForm(value) && String(row['区分ID']) === cutId) {
      const home = String(row['元の区分ID'] || '');
      if (home && home !== cutId && db.sections.has(home)) {
        const list = rowsOfSection(db, home).filter((r) => group.indexOf(r) < 0);
        const at = Number(row['元の順']);
        let idx = list.findIndex((r) => Number(r['順']) > at); if (idx < 0) idx = list.length;
        const orders = ordersBetween(idx > 0 ? Number(list[idx - 1]['順']) : null, idx < list.length ? Number(list[idx]['順']) : null, group.length);
        group.forEach((g, i) => { const gf = { '区分ID': home, '順': orders[i] }; if (g !== row && g['商品ID'] && isCutForm(g['販売形態'])) gf['販売形態'] = value; put(String(g['行ID']), gf); });
      }
    }
    return Array.from(out.values());
  }

  // ---------------------------------------------------------------- 次の月を作る（アップロードなし・前の月の形から）
  /** 次の月の販売形態: 月間特売・新商品は定番へ、カット商品は終売へ。ほかはそのまま（特売価格は残す） */
  function nextForm(f) { if (isSaleForm(f) || isNewForm(f)) return FORM_REGULAR; if (isCutForm(f)) return FORM_END; return f; }
  /** 前の月の 月の行・月の置き場・月間商品（{c, r}）から、次の月の物を作る。元の形（元の区分ID・元の順・元のまとまり・元の順）は、前の月の最後の形。
   *  カット商品は終売にして、元の区分へ戻す。月間商品は節の見出しと空き行だけ残す（中身は一覧の販売形態から作り直す）。返す値: { rows, places, monthly, counts } */
  function nextMonthTables(src) {
    const counts = { rows: 0, sale: 0, fresh: 0, cut: 0, monthlyDropped: 0 };
    const rc = src.rows.c; const ix = (k) => rc.indexOf(k);
    const iForm = ix('販売形態'), iSec = ix('区分ID'), iOrd = ix('順'), iHome = ix('元の区分ID'), iHomeOrd = ix('元の順'), iDiff = ix('北関東の違い');
    const rows = (src.rows.r || []).map((r0) => {
      const r = r0.slice(); const f = String(r[iForm] || '');
      if (isSaleForm(f)) counts.sale++;
      if (isNewForm(f)) counts.fresh++;
      if (isCutForm(f)) { counts.cut++; if (!isBlank(r[iHome])) { r[iSec] = r[iHome]; r[iOrd] = Number(r[iHomeOrd]) + 0.001; } }
      if (f) r[iForm] = nextForm(f);
      if (!isBlank(r[iDiff])) {
        const d = parseDeco(r[iDiff]);
        if (Object.prototype.hasOwnProperty.call(d, '販売形態')) { d['販売形態'] = nextForm(d['販売形態']); if (d['販売形態'] === r[iForm]) delete d['販売形態']; }
        r[iDiff] = Object.keys(d).length ? JSON.stringify(d) : '';
      }
      r[iHome] = r[iSec]; r[iHomeOrd] = r[iOrd];
      counts.rows++;
      return r;
    });
    const pc = src.places.c; const jB = pc.indexOf('まとまり'), jO = pc.indexOf('順'), jB0 = pc.indexOf('元のまとまり'), jO0 = pc.indexOf('元の順');
    const places = (src.places.r || []).map((r0) => { const r = r0.slice(); r[jB0] = r[jB]; r[jO0] = r[jO]; return r; });
    const mc = src.monthly.c; const kKind = mc.indexOf('行の種類');
    const monthly = (src.monthly.r || []).filter((r) => { const keep = r[kKind] === KIND_SECTION || r[kKind] === KIND_BLANK; if (!keep) counts.monthlyDropped++; return keep; }).map((r) => r.slice());
    return { rows: { c: rc.slice(), r: rows }, places: { c: pc.slice(), r: places }, monthly: { c: mc.slice(), r: monthly }, counts };
  }
  /** 次の月（'YYYY-MM'） */
  function nextMonthOf(m) { const x = /^(\d{4})-(\d{2})$/.exec(String(m)); if (!x) return ''; let y = Number(x[1]), mo = Number(x[2]) + 1; if (mo > 12) { mo = 1; y++; } return y + '-' + String(mo).padStart(2, '0'); }

  // ---------------------------------------------------------------- 月間商品を一覧の販売形態から作り直す
  const MONTHLY_PARTS = [['月間特売', isSaleForm], ['新規商品', isNewForm]];
  const MONTHLY_COLS = ['行ID', '部', '順', '行の種類', '商品名', '呼出番', '規格', '本体価格', '小パックの規格', '小パックの本体価格', '商品ID', '見た目'];
  /** 部ごとに、月間商品に載せたい商品（番号表に載る行の販売形態から・表の順） */
  function monthlyWanted(db) {
    const want = { '月間特売': [], '新規商品': [] }; const seen = { '月間特売': new Set(), '新規商品': new Set() };
    const memo = rowsBySection(db);
    TABLE_KEYS.forEach((t) => { // 表の型が無くても数えられるように、置き場の順（まとまり・順）と行の順で並べる（窓口も同じ部品を使う）
      sectionsOfKind(db, TABLES[t].kind).map((s) => db.places.get(String(s['区分ID']))).sort((a, b) => Number(a['まとまり']) - Number(b['まとまり']) || byOrder(a, b)).forEach((p) => {
        (memo.get(String(p['区分ID'])) || []).forEach((r) => {
          if (r['行の種類'] !== KIND_PRODUCT || !rowShown(db, r, t)) return;
          const pid = String(r['商品ID']); const f = formOf(r, t);
          MONTHLY_PARTS.forEach(([part, test]) => { if (test(f) && !seen[part].has(pid) && db.products.has(pid)) { seen[part].add(pid); want[part].push({ pid, row: r, table: t }); } });
        });
      });
    });
    return want;
  }
  /** 月間商品と一覧の食い違い: missing＝載せたいのに無い商品・extra＝結んだ商品が、もうその販売形態でない行。
   *  月間特売と新商品の両方の商品は、どちらかの部にあれば足りる（10月の手作りは 12 商品のうち 7 が月間特売の部・5 が新規商品の部）。無ければ月間特売の部へ */
  function monthlySync(db) {
    const want = monthlyWanted(db); const missing = [], extra = [];
    const linked = (part) => new Set(db.monthly.filter((g) => g['部'] === part && g['行の種類'] === KIND_PRODUCT && g['商品ID']).map((g) => String(g['商品ID'])));
    const have = { '月間特売': linked('月間特売'), '新規商品': linked('新規商品') };
    const both = new Set(want['月間特売'].map((w) => w.pid).filter((pid) => want['新規商品'].some((w) => w.pid === pid)));
    MONTHLY_PARTS.forEach(([part]) => {
      const ids = new Set(want[part].map((w) => w.pid));
      want[part].forEach((w) => {
        if (have[part].has(w.pid)) return;
        if (both.has(w.pid) && (have['月間特売'].has(w.pid) || have['新規商品'].has(w.pid) || part === '新規商品')) return;
        missing.push({ part, pid: w.pid });
      });
      db.monthly.forEach((g) => { if (g['部'] === part && g['行の種類'] === KIND_PRODUCT && g['商品ID'] && !ids.has(String(g['商品ID']))) extra.push({ part, rowId: String(g['行ID']), pid: String(g['商品ID']) }); });
    });
    return { missing, extra, want };
  }
  /** 小パックの欄「(３個)１９９」→ 規格と本体価格 */
  function splitSmall(v) { const m = /^\s*[(（]([^)）]{1,10})[)）]\s*([0-9０-９]+)\s*$/.exec(String(v === null || v === undefined ? '' : v)); return m ? { spec: m[1], price: num(m[2]) } : { spec: '', price: '' }; }
  /** 月間商品の行を、一覧の販売形態に合わせて作り直す。結んだ行は、まだその販売形態なら残す（続きの行・サマリーから入れた規格も残る）。結んでいない行・節の見出し・空き行も残す。
   *  足りない商品は、表の種類で節を選んで（寿司弁当＝≪米飯≫・焼き物＝≪惣菜≫・無ければその部の最初の節）その節の末尾へ。返す値: { rows（行の物の並び）, added, removed } */
  function monthlyReconcile(db) {
    const sync = monthlySync(db);
    const removeIds = new Set(sync.extra.map((x) => x.rowId));
    const ordered = db.monthly.slice().sort(byOrder);
    let next = 0; ordered.forEach((g) => { const m = /^G(\d+)$/.exec(String(g['行ID'])); if (m) next = Math.max(next, Number(m[1])); });
    const newId = () => { next += 1; return 'G' + String(next).padStart(3, '0'); };
    const out = []; let removed = 0;
    MONTHLY_PARTS.forEach(([part]) => {
      // 部の行（消す行と、その続きの行を外す）
      const rows = []; let dropping = false;
      ordered.forEach((g) => {
        if (g['部'] !== part) return;
        if (g['行の種類'] === KIND_PRODUCT) dropping = removeIds.has(String(g['行ID']));
        else if (g['行の種類'] !== KIND_CONT) dropping = false;
        if (dropping) { if (g['行の種類'] === KIND_PRODUCT) removed++; return; }
        rows.push(Object.assign({}, g));
      });
      // 節に分ける（最初の見出しの前も1つの節）
      const segs = [{ head: null, rows: [] }];
      rows.forEach((g) => { if (g['行の種類'] === KIND_SECTION) segs.push({ head: g, rows: [] }); else segs[segs.length - 1].rows.push(g); });
      const findSeg = (word) => segs.find((s) => s.head && squashText(s.head['商品名']).indexOf(word) >= 0);
      const segFor = (kind) => (kind === '寿司弁当' ? findSeg('米') : (findSeg('惣') || (segs[0].rows.length || !segs[1] ? segs[0] : null))) || segs.find((s) => s.head) || segs[0];
      sync.missing.filter((x) => x.part === part).forEach((x) => {
        const w = sync.want[part].find((y) => y.pid === x.pid); const p = db.products.get(x.pid);
        const q = parseQty(p['名前']); const name = nfkc(q.has ? String(p['名前']).slice(0, q.start) : p['名前']).trim();
        const price = part === '月間特売' ? (isBlank(rowValue(w.row, '特売価格', w.table)) ? p['定番価格'] : rowValue(w.row, '特売価格', w.table)) : p['定番価格'];
        const small = splitSmall(p['小パック']);
        const g = { '行ID': newId(), '部': part, '順': 0, '行の種類': KIND_PRODUCT, '商品名': name, '呼出番': p['呼出番'] || '', '規格': q.has ? toFull(q.count) + q.unit : '', '本体価格': price === undefined ? '' : price, '小パックの規格': small.spec, '小パックの本体価格': small.price, '商品ID': x.pid, '見た目': '' };
        const seg = segFor(p['表の種類']);
        let at = seg.rows.length; while (at > 0 && seg.rows[at - 1]['行の種類'] === KIND_BLANK) at--; // 節の末尾の空き行の前へ
        seg.rows.splice(at, 0, g);
      });
      segs.forEach((s) => { if (s.head) out.push(s.head); s.rows.forEach((g) => out.push(g)); });
    });
    ordered.forEach((g) => { if (MONTHLY_PARTS.every(([p]) => p !== g['部'])) out.push(Object.assign({}, g)); }); // 部の分からない行もそのまま
    out.forEach((g, i) => { g['順'] = (i + 1) * ORDER_STEP; });
    return { rows: out, added: sync.missing.length, removed };
  }

  /** 変える欄を手元に当てる。t: 商品・行・置き場・月間・状態 */
  function applyChanges(db, changes) {
    (changes || []).forEach((c) => {
      let o = null;
      if (c.t === '商品') o = db.products.get(String(c.id));
      else if (c.t === '行') o = db.rows.get(String(c.id));
      else if (c.t === '置き場') o = db.places.get(String(c.id));
      else if (c.t === '月間') o = db.monthly.find((m) => String(m['行ID']) === String(c.id));
      else if (c.t === '区分') o = db.sections.get(String(c.id));
      else if (c.t === '状態') { db.status[c.id] = c.f['形']; return; }
      else if (c.t === '原価') { if (db.costs) { const r = db.costs.get(String(c.id)) || {}; Object.assign(r, c.f); db.costs.set(String(c.id), r); } return; }
      else if (c.t === '新しい商品') { const p = Object.assign({}, c.f, { '商品ID': String(c.id) }); p._deco = parseDeco(p['字の飾り']); db.products.set(String(c.id), p); return; }
      else if (c.t === '消す商品') { const pid = String(c.id); db.products.delete(pid); db.rows.forEach((r, id) => { if (String(r['商品ID']) === pid) db.rows.delete(id); }); if (db.costs) db.costs.delete(pid); return; }
      else if (c.t === '新しい行') { db.rows.set(String(c.id), prepRow(Object.assign({}, c.f, { '行ID': String(c.id) }))); return; }
      else if (c.t === '確かめ済み') { if (c.f && c.f['済み'] === false) db.checked.delete(String(c.id)); else db.checked.add(String(c.id)); return; }
      if (!o) return;
      Object.keys(c.f || {}).forEach((k) => { o[k] = c.f[k]; if (k === '字の飾り') o._deco = parseDeco(c.f[k]); if (k === '北関東の違い') o._diff = parseDeco(c.f[k]); if (k === '続き') o._cont = parseDeco(c.f[k]); });
    });
  }
  /** 手元と違う欄を元に戻すための、前の値（元に戻す・記録に使う） */
  function beforeOf(db, changes) {
    return (changes || []).map((c) => {
      let o = null;
      if (c.t === '商品') o = db.products.get(String(c.id));
      else if (c.t === '行') o = db.rows.get(String(c.id));
      else if (c.t === '置き場') o = db.places.get(String(c.id));
      else if (c.t === '月間') o = db.monthly.find((m) => String(m['行ID']) === String(c.id));
      if (!o) return null;
      const f = {}; Object.keys(c.f || {}).forEach((k) => { f[k] = o[k] === undefined ? '' : o[k]; });
      return { t: c.t, id: c.id, f };
    }).filter(Boolean);
  }

  // ---------------------------------------------------------------- 自動の決め直し
  /** 1つの種類の、列ごとのいちばん多い行数（寿司弁当は首都圏と北関東の多い方） */
  function usedOfKind(db, kind) {
    const memo = rowsBySection(db), out = {};
    KIND_TABLES[kind].forEach((t) => { layoutOf(db, t, memo).blocks.forEach((b) => { out[b.no] = Math.max(out[b.no] || 0, b.used); }); });
    return out;
  }
  function capsOfKind(db, kind) { const out = {}; KIND_TABLES[kind].forEach((t) => { blockList(db, t).forEach((b) => { const no = Number(b['まとまり']); const c = Number(b['上限の行数']); out[no] = out[no] === undefined ? c : Math.min(out[no], c); }); }); return out; }
  /** 小パックのある行が、その列の見える欄に載るか */
  function blockFits(db, kind, secId, blockNo) {
    const t = KIND_TABLES[kind][0]; const roles = blockRoles(db, t, blockNo);
    const small = roles['小パック'];
    if (small && small['隠し'] !== '○') return true;
    return !rowsOfSection(db, secId).some((r) => { const p = r['商品ID'] ? db.products.get(String(r['商品ID'])) : null; return rowOn(r) && p && !isBlank(p['小パック']); });
  }
  /** 自動の決め直し（同じ入力なら必ず同じ結果・動かす数はいちばん少なく）。返す値: { changes, moved:[文], short: {列: 足りない行} } */
  function autoPlace(db, kind) {
    const changes = [], moved = [], work = [];
    const apply = (list) => { work.push(...beforeOf(db, list)); applyChanges(db, list); changes.push(...list); };
    const caps = capsOfKind(db, kind);
    const blocks = Object.keys(caps).map(Number).sort((a, b) => a - b);
    const tick = (blockNo) => usedOfKind(db, kind)[blockNo] - caps[blockNo];
    blocks.forEach((no) => {
      if (tick(no) <= 0) return;
      // まず、空き行を詰める（印刷に出さない）
      placesOfBlock(db, kind, no).forEach((p) => {
        rowsOfSection(db, p['区分ID']).slice().reverse().forEach((r) => {
          if (tick(no) > 0 && r['行の種類'] === KIND_BLANK && !isOffForm(r['販売形態'])) apply([{ t: '行', id: String(r['行ID']), f: { '販売形態': FORM_REST } }]);
        });
      });
      let guard = 0;
      while (tick(no) > 0 && guard++ < 12) {
        const over = tick(no);
        const used = usedOfKind(db, kind);
        const candidates = placesOfBlock(db, kind, no).slice(1).map((p) => {
          const lines = rowsOfSection(db, p['区分ID']).filter((r) => rowOn(r)).length;
          return { id: String(p['区分ID']), lines, order: Number(p['順']) };
        }).filter((c) => c.lines > 0).sort((a, b) => a.lines - b.lines || b.order - a.order);
        const targetFor = (c) => {
          const need = c.lines + 1;
          const fits = blocks.filter((t) => t !== no && caps[t] - used[t] >= need && blockFits(db, kind, c.id, t));
          fits.sort((a, b) => (caps[b] - used[b]) - (caps[a] - used[a]) || Math.abs(a - no) - Math.abs(b - no) || a - b);
          return fits.length ? fits[0] : null;
        };
        let pick = candidates.find((c) => c.lines + 1 >= over && targetFor(c) !== null);
        if (!pick) { const rest = candidates.filter((c) => targetFor(c) !== null); pick = rest.length ? rest[rest.length - 1] : null; }
        if (!pick) break;
        const to = targetFor(pick);
        apply(moveSection(db, { id: pick.id, toBlock: to, before: null }));
        moved.push({ id: pick.id, from: no, to });
      }
    });
    const short = {};
    const usedFinal = usedOfKind(db, kind);
    blocks.forEach((no) => { if (usedFinal[no] > caps[no]) short[no] = usedFinal[no] - caps[no]; });
    return { changes, moved, short };
  }
  /** 自動に戻す: 手直しを捨てて、元の形（前の月の形）から決め直す */
  function resetToAuto(db, kind) {
    const list = [];
    sectionsOfKind(db, kind).forEach((s) => {
      const p = db.places.get(String(s['区分ID']));
      if (Number(p['まとまり']) !== Number(p['元のまとまり']) || Number(p['順']) !== Number(p['元の順'])) list.push({ t: '置き場', id: String(s['区分ID']), f: { 'まとまり': Number(p['元のまとまり']), '順': Number(p['元の順']) } });
    });
    db.rows.forEach((r) => {
      const sec = db.sections.get(String(r['元の区分ID']));
      if (!sec || sec['表の種類'] !== kind) return;
      if (isCutForm(r['販売形態'])) return; // カット商品は ≪カット商品≫ の区分のまま（販売形態で決まる置き場）
      const f = {};
      if (String(r['区分ID']) !== String(r['元の区分ID'])) f['区分ID'] = String(r['元の区分ID']);
      if (Number(r['順']) !== Number(r['元の順'])) f['順'] = Number(r['元の順']);
      if (r['行の種類'] === KIND_BLANK && r['販売形態'] === FORM_REST) f['販売形態'] = '';
      if (Object.keys(f).length) list.push({ t: '行', id: String(r['行ID']), f });
    });
    applyChanges(db, list);
    const auto = autoPlace(db, kind);
    return { changes: list.concat(auto.changes), moved: auto.moved, short: auto.short };
  }

  // ---------------------------------------------------------------- アラート（入れる画面で色とお知らせ・直したらすぐ消える。価格と原価の空は色だけ）
  function pastOf(db) { if (db._pastBy && db._pastSrc === db.past) return db._pastBy; const m = new Map(); db.past.forEach((h) => { if (h['商品ID']) m.set(String(h['商品ID']), h); }); db._pastBy = m; db._pastSrc = db.past; return m; }
  function hintOf(db, pid) { const h = pastOf(db).get(String(pid)); if (!h) return ''; return '8月の番号 ' + (h['8月の番号'] || '載っていない') + '／9月の番号 ' + (h['9月の番号'] || '載っていない'); }
  function alertsOf(db) {
    const out = [];
    const A = (level, kind, key, what, hint, target) => out.push({ level, kind, key, what, hint: hint || '', target: target || {} });
    const memo = rowsBySection(db);
    const placed = new Map(); // 商品ID → その商品が載っている表と行
    const layouts = {};
    TABLE_KEYS.forEach((t) => {
      const lay = layoutOf(db, t, memo); layouts[t] = lay;
      lay.blocks.forEach((b) => {
        b.sections.forEach((s) => s.rows.forEach((r) => {
          if (r['行の種類'] !== KIND_PRODUCT) return;
          const pid = String(r['商品ID'] || '');
          if (!placed.has(pid)) placed.set(pid, { tables: new Set(), row: r, sec: s.sec, block: b.no });
          placed.get(pid).tables.add(t);
        }));
        if (b.over > 0) A(LEVEL_FIX, '上限の行数を超えた', 'over:' + t + ':' + b.no, t + ' の ' + b.no + '列目が ' + b.over + ' 行オーバー（' + b.used + '／' + b.cap + '）', '空きのある列へ動かすか、休みにする商品を選ぶ', { table: t, block: b.no });
      });
    });
    db.rows.forEach((r) => {
      if (!rowOn(r)) return;
      if (!db.sections.has(String(r['区分ID']))) A(LEVEL_FIX, '区分の無い行', 'nosec:' + r['行ID'], '行の区分が見つからない', '', { rowId: r['行ID'] });
      if (r['行の種類'] === KIND_PRODUCT && !db.products.has(String(r['商品ID']))) A(LEVEL_FIX, '商品の無い行', 'noprod:' + r['行ID'], '行の商品が見つからない', '', { rowId: r['行ID'] });
    });
    // 同じ表で同じ呼出番
    const dup = new Map();
    TABLE_KEYS.forEach((t) => {
      const by = new Map();
      placed.forEach((v, pid) => { if (!v.tables.has(t)) return; const p = db.products.get(pid); if (!p) return; const c = String(p['呼出番'] === null || p['呼出番'] === undefined ? '' : p['呼出番']).trim(); if (!c) return; if (!by.has(c)) by.set(c, new Set()); by.get(c).add(pid); });
      by.forEach((ids) => { if (ids.size > 1) ids.forEach((pid) => { if (!dup.has(pid)) dup.set(pid, new Set()); ids.forEach((q) => { if (q !== pid) dup.get(pid).add(q); }); }); });
    });
    dup.forEach((others, pid) => {
      const p = db.products.get(pid);
      A(LEVEL_FIX, '同じ表で同じ呼出番', 'dup:' + pid, '同じ番号が ' + Array.from(others).map((q) => nfkc(db.products.get(q)['名前']).trim()).join('・') + ' にもあります', hintOf(db, pid), { pid, others: Array.from(others) });
    });
    // 休み・終売・計画の番号も数に入れる（番号の管理のために残している）。載っている商品とぶつかったら「確かめる」
    const formOfP = productForms(db);
    const callOf = (x) => String(x['呼出番'] === null || x['呼出番'] === undefined ? '' : x['呼出番']).trim();
    const offByCall = new Map(); // 番号 → 表に載っていない商品（商品の数だけ1回なめる）
    db.products.forEach((q, qid) => { if (placed.has(qid)) return; const c = callOf(q); if (!c) return; if (!offByCall.has(c)) offByCall.set(c, []); offByCall.get(c).push([qid, q]); });
    if (offByCall.size) placed.forEach((v, pid) => {
      const p = db.products.get(pid); if (!p) return;
      const c = callOf(p); if (!c || !offByCall.has(c)) return;
      const key = 'dupoff:' + pid; if (db.checked.has(key)) return;
      const hits = [];
      offByCall.get(c).forEach(([qid, q]) => {
        if (q['表の種類'] !== p['表の種類']) return;
        const a = p['地域'] || '共通', b = q['地域'] || '共通'; if (a !== '共通' && b !== '共通' && a !== b) return;
        hits.push(nfkc(q['名前']).trim() + '（' + (formOfP.get(qid) || FORM_END) + '）');
      });
      if (hits.length) A(LEVEL_CHECK, '休み・終売・計画と同じ番号', key, '番号 ' + c + ' は ' + hits.join('・') + ' と同じです', '番号を変えるか、このままでよいか', { pid });
    });
    // 呼出番・名前の空（価格・原価の空はアラートにしない）
    placed.forEach((v, pid) => {
      const p = db.products.get(pid); if (!p) return;
      if (isBlank(p['呼出番'])) A(LEVEL_FIX, '呼出番が空', 'nocall:' + pid, '呼出番が空です', hintOf(db, pid), { pid });
      if (isBlank(p['名前'])) A(LEVEL_FIX, '名前が空', 'noname:' + pid, '名前が空です', '', { pid });
      TABLE_KEYS.forEach((t) => {
        if (!v.tables.has(t) || isBlank(p['小パック'])) return;
        const blockOf = layouts[t].blocks.find((b) => b.sections.some((s) => s.rows.some((r) => String(r['商品ID']) === pid)));
        if (!blockOf) return;
        const role = blockRoles(db, t, blockOf.no)['小パック'];
        if (!role || role['隠し'] === '○') A(LEVEL_FIX, '欄の無い列へ移した', 'nosmall:' + t + ':' + pid, t + ' の ' + blockOf.no + '列目には小パックの欄が無く、印刷されません', '小パックの欄のある列へ動かす', { pid, table: t, block: blockOf.no });
      });
    });
    // 月間商品
    const callsAll = new Set(); db.products.forEach((p) => { const c = halfCall(p['呼出番']); if (c) callsAll.add(c); });
    db.monthly.forEach((g) => {
      if (g['行の種類'] !== KIND_PRODUCT) return;
      const c = halfCall(g['呼出番']), pid = g['商品ID'] ? String(g['商品ID']) : '', name = g['商品名'];
      const key = String(g['行ID']);
      if (isBlank(g['本体価格'])) A(LEVEL_FIX, '月間商品の本体価格が空', 'monprice:' + key, nfkc(name).trim() + ' の本体価格が空です', '', { monthly: key });
      if (pid && db.products.has(pid)) {
        const pc = String(db.products.get(pid)['呼出番'] || '');
        if (c && pc && c !== pc) A(LEVEL_FIX, '月間商品と番号表で番号が違う', 'mon:' + key, nfkc(name).trim() + ': 月間商品では ' + c + '、番号表では ' + pc, hintOf(db, pid), { monthly: key, pid });
      } else if (c && !pid) {
        if (callsAll.has(c)) { if (!db.checked.has('monlink:' + key)) A(LEVEL_CHECK, '月間商品の行が商品に結ばれていない', 'monlink:' + key, nfkc(name).trim() + ': 番号 ' + c + ' の商品は番号表にありますが、名前が違います', '同じ商品なら結ぶ', { monthly: key }); }
        else if (g['部'] === '新規商品' && !db.checked.has('newmiss:' + key)) A(LEVEL_CHECK, '新規商品が番号表に無い', 'newmiss:' + key, '新規商品 ' + nfkc(name).trim() + '（' + c + '）が番号表にありません', '番号表に載せるか、このままでよいか', { monthly: key });
      }
    });
    // 月間商品と一覧の販売形態の食い違い（番号表の「月間商品」で作り直すと消える）
    if (db.monthly.length) {
      const sync = monthlySync(db);
      if (sync.missing.length) A(LEVEL_FIX, '月間商品に無い', 'monsync:missing', '月間特売・新商品にした ' + sync.missing.length + ' 商品が、月間商品にありません', '番号表の「月間商品」で「一覧から作り直す」を押す', { monthlySync: true });
      if (sync.extra.length && !db.checked.has('monsync:extra')) A(LEVEL_CHECK, '月間商品に余分', 'monsync:extra', '月間商品に、もう月間特売・新商品でない ' + sync.extra.length + ' 商品があります', '番号表の「月間商品」で「一覧から作り直す」を押すか、このままでよいか', { monthlySync: true });
    }
    // 続きの行に価格
    db.rows.forEach((r) => {
      if (r['行の種類'] !== KIND_CONT || !rowOn(r)) return;
      const cont = r._cont || {}; const key = 'cont:' + r['行ID'];
      if (!isBlank(cont['見える価格']) && !db.checked.has(key)) A(LEVEL_CHECK, '続きの行に価格がある', key, '名前の無い行に「' + nfkc(cont['呼出番の欄']).trim() + '」があります（上の商品の別の売り方らしい）', '1つの売り方として商品にするか、このままにするか', { rowId: r['行ID'] });
    });
    // 粗利（原価は窓口が見て、値を返さずに印だけ返す。表示中は手元の値で数える）
    const grossOut = new Set(db.costInfo.grossOut || []);
    if (db.costs) { grossOut.clear(); db.costs.forEach((c, id) => { const g = grossAfter(c['原価'], c['定番価格（包材変更後）']); if (g !== null && (g < GROSS_LOW || g > GROSS_HIGH)) grossOut.add(id); }); }
    grossOut.forEach((pid) => { if (!placed.has(pid)) return; const key = 'gross:' + pid; if (db.checked.has(key)) return; const p = db.products.get(pid); A(LEVEL_CHECK, '粗利がマイナス・極端', key, nfkc(p['名前']).trim() + ' の粗利（包材変更後）が目安 ' + (GROSS_LOW * 100) + '〜' + (GROSS_HIGH * 100) + '% の外です', '原価を表示して確かめる', { pid }); });
    (db.costInfo.badForm || []).forEach((pid) => { const p = db.products.get(pid); if (!p || !placed.has(pid)) return; A(LEVEL_CHECK, '原価の式が読めない', 'badform:' + pid, nfkc(p['名前']).trim() + ' の原価が数でも数字の式でもありません', '原価を表示して確かめる', { pid }); });
    // 元の書類の疑問
    db.doubts.forEach((d) => { const key = 'doubt:' + d['番号']; if (String(d['済み'] || '').trim() === '○' || db.checked.has(key)) return; A(LEVEL_CHECK, '元の書類の疑問', key, nfkc(d['何が']).trim(), nfkc(d['手がかり']).trim().slice(0, 200), {}); });
    const rank = { [LEVEL_FIX]: 0, [LEVEL_CHECK]: 1 };
    out.sort((a, b) => rank[a.level] - rank[b.level] || (a.kind < b.kind ? -1 : a.kind > b.kind ? 1 : 0) || (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
    return out;
  }
  /** 商品ID → 販売形態（その月の行から。地域で違うときは載る方を先に） */
  function productForms(db) { const m = new Map(); db.rows.forEach((r) => { if (!r['商品ID']) return; const k = String(r['商品ID']); const f = r['販売形態'] || ''; if (!m.has(k) || isOffForm(m.get(k))) m.set(k, f); }); return m; }
  /** 空いている番号の候補（around の近くから n 個）。どの商品にも（休み・終売・計画も含む）、8月・9月の番号にも無い番号だけ */
  function freeCalls(db, around, n) {
    const used = new Set();
    db.products.forEach((p) => { const c = halfCall(p['呼出番']); if (c) used.add(c); });
    db.past.forEach((h) => { ['8月の番号', '9月の番号'].forEach((k) => { const c = halfCall(h[k]); if (c) used.add(c); }); });
    const start = Number(halfCall(around) || '0') || (used.size ? Number(Array.from(used).sort()[0]) : 1000);
    const out = [];
    for (let d = 1; d < 9000 && out.length < (n || 4); d++) {
      [start + d, start - d].forEach((x) => { if (out.length < (n || 4) && x >= 1000 && x <= 9999 && !used.has(String(x))) out.push(String(x)); });
    }
    return out;
  }
  /** 抜け（セルの色だけ）: 表に出る価格が空・原価が空 */
  function gapsOf(db) {
    const price = new Set(), cost = new Set();
    const memo = rowsBySection(db);
    TABLE_KEYS.forEach((t) => layoutOf(db, t, memo).blocks.forEach((b) => b.sections.forEach((s) => s.rows.forEach((r) => {
      if (r['行の種類'] !== KIND_PRODUCT) return;
      const pid = String(r['商品ID']); const p = db.products.get(pid); if (!p) return;
      if (isBlank(shownPrice(p, r, t))) price.add(pid);
    }))));
    const emptySet = db.costs ? null : new Set(db.costInfo.empty || []);
    const onProducts = new Set(); db.rows.forEach((r) => { if (r['商品ID'] && rowOn(r)) onProducts.add(String(r['商品ID'])); });
    db.products.forEach((p, pid) => {
      if (!onProducts.has(pid)) return;
      if (db.costs) { const c = db.costs.get(pid); if (!c || isBlank(c['原価'])) cost.add(pid); } else if (emptySet.has(pid)) cost.add(pid);
    });
    return { price, cost };
  }

  // ---------------------------------------------------------------- 値の形（画面と窓口で同じ決まり）
  const LIMITS = { '呼出番': 20, '名前': 200, '指定トレー': 100, '小パック': 100, 'メモ': 500, '定番価格': 30, '特売価格': 30 };
  /** 欄の形を確かめる。問題なければ ''、あれば人が読む文 */
  function fieldProblem(table, key, value) {
    const v = value === null || value === undefined ? '' : value;
    if (typeof v === 'object') return key === '字の飾り' ? '' : '入れられない形です';
    const s = String(v);
    if (s.length > (LIMITS[key] || 2000)) return key + 'が長すぎます';
    if (table === '原価' && (key === '原価' || key.indexOf('定番価格') === 0) && !costFormOk(v)) return key + 'は、数か、数字だけの式（=100+20）で入れてください';
    if (table === '行' && key === '販売形態' && s !== '' && FORMS_ALL.indexOf(s) < 0) return '販売形態は、一覧の中から選んでください';
    if (table === '商品' && key === '地域' && ['共通', '首都圏だけ', '北関東だけ'].indexOf(s) < 0) return '地域は、共通・首都圏だけ・北関東だけのどれかです';
    return '';
  }
  const EDITABLE = {
    '商品': ['呼出番', '名前', '字の飾り', '地域', '指定トレー', '定番価格', '小パック', 'メモ'],
    '原価': COST_KEYS,
    '行': ['販売形態', '特売価格', '区分ID', '順', '北関東の違い'],
    '月間': ['商品名', '呼出番', '規格', '本体価格', '小パックの規格', '小パックの本体価格', '商品ID'],
    '置き場': ['まとまり', '順']
  };

  root.SouzaiCore = { TABLES, TABLE_KEYS, KIND_TABLES, KINDS, KIND_PRODUCT, KIND_BLANK, KIND_CONT, KIND_SECTION, FORMS, FORMS_ALL, FORM_REGULAR, FORM_SALE, FORM_NEW, FORM_SALE_NEW, FORM_CUT, FORM_REST, FORM_END, FORM_PLAN, ADDED_MARK, isAddedProduct, normTray, showCell, shownPriceKey, shownPrice, isOffForm, formHas, toggleForm, productForms, freeCalls, isSaleForm, isNewForm, isCutForm, formOf, rowOn, rowRegion, COST_KEYS, GROSS_LOW, GROSS_HIGH, LEVEL_FIX, LEVEL_CHECK, EDITABLE,
    isCutSection, cutSectionOf, formChanges, nextForm, nextMonthTables, nextMonthOf, MONTHLY_COLS, monthlyWanted, monthlySync, monthlyReconcile, splitSmall,
    isBlank, nfkc, toFull, halfCall, fullDigits, num, evalArith, costFormOk, grossAfter, parseDeco, editRuns, paintRuns, parseQty, setQty, rowsOf, makeDb, setType, blockRoles, blockList, rowValue, secValue, rowRegionOk, rowShown, rowsOfSection, rowsBySection, placesOfBlock, sectionsOfKind, layoutOf, moveSection, moveRow, rowGroup, applyChanges, beforeOf, usedOfKind, capsOfKind, autoPlace, resetToAuto, alertsOf, gapsOf, fieldProblem, hintOf };
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
