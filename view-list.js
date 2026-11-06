(() => {
  'use strict';
  // 一覧: 探す → その場で直す → アラートがその場で消える。1商品1行。見出しと左の2列（番号・名前）は動かない（同じ枠の中で縦横に送る）。
  //   行の色は販売形態から（入れる画面だけ。Excel と見本は見本どおり）。色だけに頼らないよう、販売形態の字も行に出す。
  const App = window.SouzaiApp;
  const { state, $, el, esc, C, SN, toast } = App;
  const FORM_CLASS = App.FORM_CLASS;
  const FORM_NOTE = { '定番': '番号表に載る', '月間特売': '番号表に載る・ピンクの字', '新商品': '番号表に載る・水色のセル', 'カット商品': '≪カット商品≫ に載る・次の月は終売', '休み': '番号表に載らない', '終売': '番号表に載らない・番号の管理に残す', '計画': '番号表に載らない・番号の管理に残す' };
  const CHIPS = ['全部', 'アラート', '抜け', '定番', '月間特売', '新商品', 'カット商品', '休み', '終売', '計画'];
  const view = { built: false, rows: [], rowEls: new Map(), byPid: new Map(), chip: '全部', q: '', terms: null, sec: '', open: '', keys: new Map(), root: null, colCount: 0, menuFor: null, fx: { cell: null, mouse: false } };
  const db = () => state.db;

  // ------------------------------------------------------------ 値の形
  const showVal = (v) => (v === null || v === undefined ? '' : String(v));
  const toStored = (s) => { const t = String(s).trim(); const n = C.nfkc(t); return /^-?\d+(\.\d+)?$/.test(n) ? Number(n) : t; };
  function formOfRow(r) { return r['販売形態'] || ''; }
  function northForm(r) { return r._diff && Object.prototype.hasOwnProperty.call(r._diff, '販売形態') ? r._diff['販売形態'] : null; }
  function matchesForm(r, chip) {
    const has = (f) => f === chip || (f === '月間特売・新商品' && (chip === '月間特売' || chip === '新商品'));
    return has(formOfRow(r)) || (northForm(r) !== null && has(northForm(r)));
  }
  function keyOf(pid, r) {
    if (view.keys.has(pid)) return view.keys.get(pid);
    const p = db().products.get(pid); const sec = db().sections.get(String(r['区分ID']));
    const t = [p['名前'], p['呼出番'], p['小パック'], sec ? sec['表題'] : ''].filter((x) => x !== '' && x !== null && x !== undefined).join(' ');
    const k = SN ? SN.key(t) : C.nfkc(t).toLowerCase();
    view.keys.set(pid, k);
    return k;
  }
  const hitKey = (k, terms) => (SN ? SN.hit(k, terms) : terms.every((x) => k.includes(x)));
  const makeTerms = (q) => (SN ? SN.terms(q) : C.nfkc(q).toLowerCase().split(/\s+/).filter(Boolean));

  // ------------------------------------------------------------ 作る
  function secOrder() {
    const out = [];
    C.KINDS.forEach((kind) => {
      const t = C.KIND_TABLES[kind][0];
      C.blockList(db(), t).forEach((b) => { C.placesOfBlock(db(), kind, b['まとまり']).forEach((p) => out.push(String(p['区分ID']))); });
    });
    return out;
  }
  function collect() {
    const memo = C.rowsBySection(db()); const out = [];
    secOrder().forEach((secId) => { (memo.get(secId) || []).forEach((r) => { if (r['商品ID'] && db().products.has(String(r['商品ID']))) out.push(r); }); });
    return out;
  }
  function headHtml() { return '<tr><th class="c-call" scope="col">番号</th><th class="c-name" scope="col">名前</th><th class="c-form" scope="col">販売形態</th><th class="c-price" scope="col">定番価格</th><th class="c-sale" scope="col">特売価格</th>' + (state.showCosts ? '<th class="c-cost" scope="col">原価</th><th class="c-gross" scope="col">粗利</th><th class="c-cost c-after" scope="col">定番（包材後）</th>' : '') + '<th class="c-small" scope="col">小パック</th><th class="c-tray" scope="col">包材</th><th class="c-sec" scope="col">区分</th></tr>'; }
  function rowHtml(r) {
    const pid = String(r['商品ID']); const p = db().products.get(pid); const sec = db().sections.get(String(r['区分ID']));
    const name = C.nfkc(p['名前']).trim() === '' ? '（名前なし）' : String(p['名前']);
    const costCells = state.showCosts ? '<td class="c-cost"><input data-t="原価" data-f="原価" aria-label="原価"></td><td class="c-gross"><span class="gross"></span></td><td class="c-cost c-after"><input data-t="原価" data-f="定番価格（包材変更後）" aria-label="定番価格（包材変更後）"></td>' : '';
    return '<tr class="lr" data-rid="' + esc(r['行ID']) + '" data-pid="' + esc(pid) + '">'
      + '<td class="c-call"><input data-t="商品" data-f="呼出番" inputmode="numeric" maxlength="20" aria-label="呼出番" value="' + esc(C.showCell(p['呼出番'])) + '"></td>'
      + '<td class="c-name"><button type="button" class="name-btn" data-act="open" aria-expanded="false"><span class="tw" aria-hidden="true">▶</span><span class="nm">' + esc(name) + '</span></button></td>'
      + '<td class="c-form"><button type="button" class="form-btn" data-act="form" aria-haspopup="true"></button></td>'
      + '<td class="c-price"><input data-t="商品" data-f="定番価格" aria-label="定番価格" value="' + esc(C.showCell(p['定番価格'])) + '"></td>'
      + '<td class="c-sale"><input data-t="行" data-f="特売価格" aria-label="特売価格" value="' + esc(C.showCell(r['特売価格'])) + '"></td>'
      + costCells
      + '<td class="c-small"><input data-t="商品" data-f="小パック" aria-label="小パック" value="' + esc(C.showCell(p['小パック'])) + '"></td>'
      + '<td class="c-tray"><input data-t="商品" data-f="指定トレー" role="combobox" aria-autocomplete="list" aria-label="包材" value="' + esc(C.showCell(p['指定トレー'])) + '"></td>'
      + '<td class="c-sec"><span class="c-sec-text" title="' + esc(sec ? sec['表題'] : '') + '">' + esc(sec ? sec['表題'] : '') + '</span></td>'
      + '</tr>';
  }
  function build() {
    const t0 = performance.now();
    const root = $('#view-list'); view.root = root; root.replaceChildren();
    const bar = el('div', 'list-bar');
    const search = el('div', 'search-box'); const input = el('input'); input.id = 'list-search'; input.type = 'search'; input.placeholder = '名前・番号で探す（ひらがな・ローマ字も）'; input.setAttribute('aria-label', '商品を探す'); input.value = view.q; search.append(input);
    const chips = el('div', 'chips'); chips.id = 'chips'; chips.setAttribute('role', 'group'); chips.setAttribute('aria-label', '絞り込み（凡例も兼ねる）');
    CHIPS.forEach((c) => { const b = el('button', 'chip'); b.type = 'button'; b.dataset.chip = c; chips.append(b); });
    const sel = el('select'); sel.id = 'list-section'; sel.setAttribute('aria-label', '区分で絞る');
    const o0 = el('option', '', '区分: 全部'); o0.value = ''; sel.append(o0);
    secOrder().forEach((id) => { const s = db().sections.get(id); if (!s) return; const o = el('option', '', C.nfkc(s['表題'])); o.value = id; sel.append(o); });
    sel.value = view.sec;
    const costBtn = el('button', 'btn', state.showCosts ? '原価を隠す' : '原価を表示'); costBtn.type = 'button'; costBtn.id = 'cost-button';
    bar.append(search, chips, sel, costBtn);
    const hint = el('p', 'list-hint', '番号・価格はその場で打てます（自動で保存）。名前を押すと1商品の入力が開きます。薄い字の価格は番号表に出ません。薄い黄と点線のセルは抜け（空）です。月間特売と新商品の両方の行は、行が薄い赤・名前が薄い水色です。');
    // 数式バー（Excel と同じ）: 選んだセルの中身の全文。縦に送っても動かない
    const fxBar = el('div', 'fx-bar'); fxBar.id = 'fx-bar';
    const fxLabel = el('span', 'fx-label', '選んだセル'); fxLabel.id = 'fx-label';
    const fxInput = el('input'); fxInput.id = 'fx-input'; fxInput.type = 'text'; fxInput.disabled = true; fxInput.placeholder = 'セルを選ぶと、中身の全文がここに出ます（ここでも直せます）'; fxInput.setAttribute('aria-labelledby', 'fx-label');
    fxBar.append(fxLabel, fxInput);
    const others = el('div', 'alert-others'); others.id = 'alert-others'; others.hidden = true;
    const scroll = el('div', 'list-scroll'); scroll.id = 'list-scroll'; scroll.tabIndex = -1;
    view.rows = collect();
    const rowsHtml = view.rows.map(rowHtml).join('');
    const span = state.showCosts ? 11 : 8; view.colCount = span;
    view.fx = { cell: null, mouse: false };
    scroll.innerHTML = '<table class="list-table' + (state.showCosts ? ' has-costs' : '') + '" id="list-table"><thead>' + headHtml() + '</thead><tbody id="list-body"><tr class="new-row"><td colspan="' + span + '"><div class="new-wrap"><input class="new-name" id="new-name" type="text" placeholder="＋ 新しい商品の名前を打つ（候補から写せます）" aria-label="新しい商品の名前"></div></td></tr>' + rowsHtml + '</tbody></table>';
    root.append(bar, hint, fxBar, others, scroll);
    view.rowEls = new Map(); view.byPid = new Map();
    scroll.querySelectorAll('tr.lr').forEach((tr) => { view.rowEls.set(tr.dataset.rid, tr); if (!view.byPid.has(tr.dataset.pid)) view.byPid.set(tr.dataset.pid, []); view.byPid.get(tr.dataset.pid).push(tr); });
    view.rows.forEach((r) => paintRow(r, true));
    view.built = true;
    bind(root);
    applyFilter();
    renderOthers();
    if (view.open) { const tr = view.rowEls.get(view.open); if (tr) openDetail(tr); else view.open = ''; }
    App.perfNote('一覧を作る', t0);
    // 探す鍵は、手が空いたときに作っておく（打ち始めで待たせない）
    window.setTimeout(() => { const rows = view.rows.slice(); let i = 0; const step = () => { const t = performance.now(); while (i < rows.length && performance.now() - t < 8) { keyOf(String(rows[i]['商品ID']), rows[i]); i++; } if (i < rows.length) window.setTimeout(step, 0); }; step(); }, 50);
  }

  // ------------------------------------------------------------ 1行を描き直す（打っている欄は触らない）
  function paintName(tr, p) {
    const nm = tr.querySelector('.nm'); if (!nm) return;
    const text = C.nfkc(p['名前']).trim() === '' ? '（名前なし）' : String(p['名前']);
    const runs = (p._deco && p._deco['名前']) || [];
    if (!runs.length) { if (nm.textContent !== text || nm.childElementCount) nm.textContent = text; return; }
    nm.replaceChildren(); let at = 0;
    runs.slice().sort((a, b) => a[0] - b[0]).forEach((run) => {
      const from = Math.min(run[0], text.length), to = Math.min(run[1], text.length);
      if (from > at) nm.append(document.createTextNode(text.slice(at, from)));
      const s = el('span', '', text.slice(from, to)); if (run[2] && run[2].c) s.style.color = '#' + run[2].c; nm.append(s); at = to;
    });
    if (at < text.length) nm.append(document.createTextNode(text.slice(at)));
  }
  const isCell = (n) => Boolean(n && n.matches && n.matches('input[data-f]'));
  /** 打っている最中のセルか（セルにカーソルがある・数式バーでこのセルを直している） */
  const isEditing = (input) => view.fx.busy === input || document.activeElement === input || (view.fx.cell === input && document.activeElement === $('#fx-input'));
  /** 選んでいないセルは結果を見せる（式は計算した数）。読めない式は赤い印のまま */
  function setInput(input, value) {
    if (!input || isEditing(input)) return;
    const v = C.showCell(value); if (input.value !== v) input.value = v;
    const unread = typeof value === 'string' && value.trim()[0] === '=' && C.evalArith(value) === null;
    input.classList.toggle('is-bad', unread); if (unread) input.title = '数字の式として読めません'; else input.removeAttribute('title');
  }
  /** その行で番号表に出る価格の欄（首都圏と北関東で違えば両方） */
  function shownKeys(r) {
    const p = db().products.get(String(r['商品ID'])); const keys = new Set(); if (!p) return keys;
    (C.KIND_TABLES[p['表の種類']] || []).forEach((t) => { if (C.rowRegionOk(db(), r, t)) keys.add(C.shownPriceKey(r, t)); });
    return keys;
  }
  function paintRow(r, first) {
    const tr = view.rowEls.get(String(r['行ID'])); if (!tr) return;
    const pid = String(r['商品ID']); const p = db().products.get(pid); if (!p) return;
    const form = formOfRow(r); const alerts = state.alertsByPid.get(pid) || [];
    tr.className = 'lr ' + (FORM_CLASS[form] || 'f-rest') + (alerts.length ? ' has-alert' : '');
    const btn = tr.querySelector('.form-btn');
    const north = northForm(r);
    btn.replaceChildren(document.createTextNode((form || '—')), el('small', '', north !== null && north !== form ? '北関東は' + north : '▾'));
    paintName(tr, p);
    const reg = p['地域'] || '共通';
    const nameCell = tr.querySelector('.name-btn');
    nameCell.querySelectorAll('.tag').forEach((t) => t.remove());
    if (reg !== '共通') nameCell.append(el('span', 'tag', reg));
    if (state.gaps.cost.has(pid) && !state.showCosts) nameCell.append(el('span', 'tag tag-gap', '原価の抜け'));
    if (!first) {
      setInput(tr.querySelector('[data-f="呼出番"]'), p['呼出番']); setInput(tr.querySelector('[data-f="定番価格"]'), p['定番価格']); setInput(tr.querySelector('[data-f="小パック"]'), p['小パック']); setInput(tr.querySelector('[data-f="特売価格"]'), r['特売価格']); setInput(tr.querySelector('[data-f="指定トレー"]'), p['指定トレー']);
      const sec = db().sections.get(String(r['区分ID'])); const st = tr.querySelector('.c-sec-text'); if (st && sec) { st.textContent = sec['表題']; st.title = sec['表題']; }
    }
    const keys = shownKeys(r); // 表に出ないほうの価格は薄い字（印刷の色は変えない）
    const cellPrice = tr.querySelector('.c-price'), cellSale = tr.querySelector('.c-sale');
    cellPrice.classList.toggle('is-unshown', !keys.has('定番価格')); cellSale.classList.toggle('is-unshown', !keys.has('特売価格'));
    cellPrice.title = keys.has('定番価格') ? '' : 'いまは番号表に出ません'; cellSale.title = keys.has('特売価格') ? '' : 'いまは番号表に出ません';
    cellPrice.classList.toggle('is-gap', state.gaps.price.has(pid));
    if (state.showCosts) {
      const c = db().costs.get(pid) || {};
      setInput(tr.querySelector('[data-f="原価"]'), c['原価']); setInput(tr.querySelector('[data-f="定番価格（包材変更後）"]'), c['定番価格（包材変更後）']);
      tr.querySelectorAll('.c-cost')[0].classList.toggle('is-gap', state.gaps.cost.has(pid));
      paintGross(tr, c);
    }
    paintNotice(tr, r, alerts);
  }
  function paintGross(tr, c) { const g = C.grossAfter(c['原価'], c['定番価格（包材変更後）']); const s = tr.querySelector('.gross'); if (!s) return; s.textContent = g === null ? '' : (Math.round(g * 1000) / 10).toFixed(1) + '%'; s.classList.toggle('is-out', g !== null && (g < C.GROSS_LOW || g > C.GROSS_HIGH)); }
  // ------------------------------------------------------------ お知らせ（押した行のすぐ下に）
  function alertLine(a) {
    const line = el('div', 'alert-line');
    const lv = el('span', 'level' + (a.level === C.LEVEL_CHECK ? ' is-check' : ''), a.level);
    line.append(lv, el('span', '', a.what));
    if (a.hint) line.append(el('span', 'alert-hint', a.hint));
    alertActions(a).forEach((b) => line.append(b));
    return line;
  }
  function alertActions(a) {
    const out = [];
    if (a.level === C.LEVEL_CHECK) { const b = el('button', 'btn btn-small', 'このままでよい'); b.type = 'button'; b.addEventListener('click', () => markChecked(a, b)); out.push(b); }
    if (a.kind === '月間商品と番号表で番号が違う' && a.target.pid && a.target.monthly) {
      const p = db().products.get(String(a.target.pid));
      const b = el('button', 'btn btn-small', '月間商品を ' + p['呼出番'] + ' に直す'); b.type = 'button';
      b.addEventListener('click', () => { const m = App.change([{ t: '月間', id: String(a.target.monthly), f: { '呼出番': String(p['呼出番']) } }]); if (m) toast(m, true); });
      out.push(b);
    }
    return out;
  }
  async function markChecked(a, button) {
    const done = App.busyButton(button, '記録しています…');
    db().checked.add(a.key); App.refresh([]); renderOthers();
    try { await App.serial(() => App.api('setChecked', { month: state.month, key: a.key, requestId: App.rid() })); }
    catch (error) { db().checked.delete(a.key); App.refresh([]); renderOthers(); toast('記録できませんでした: ' + error.message, true); }
    finally { done(); }
  }
  function paintNotice(tr, r, alerts) {
    const next = tr.nextElementSibling;
    const has = next && next.classList.contains('notice-row');
    if (!alerts.length) { if (has) next.remove(); return; }
    const row = has ? next : el('tr', 'notice-row');
    if (!has) { const td = el('td'); td.colSpan = view.colCount; const box = el('div', 'row-notice'); td.append(box); row.append(td); tr.after(row); }
    const box = row.querySelector('.row-notice'); box.replaceChildren(...alerts.map(alertLine));
    row.hidden = tr.hidden;
  }
  /** 行につかないアラート（元の書類の疑問・月間商品だけの物・上限） */
  function renderOthers() {
    const box = $('#alert-others'); if (!box) return;
    const items = state.alerts.filter((a) => { const pid = a.target && (a.target.pid || (a.target.rowId && App.headPidOfRow(a.target.rowId))); return !pid; });
    if (view.chip !== 'アラート' || !items.length) { box.hidden = true; return; }
    box.hidden = false; box.replaceChildren(el('h2', '', '行につかないアラート ' + items.length + ' 件'));
    items.forEach((a) => {
      const line = alertLine(a);
      if (a.target && a.target.table) { const b = el('button', 'btn btn-small', '番号表で見る'); b.type = 'button'; b.addEventListener('click', () => { App.views.sheet.openTable(a.target.table); App.setView('sheet'); }); line.append(b); }
      box.append(line);
    });
  }

  // ------------------------------------------------------------ 絞り込み
  function counts() {
    const out = { '全部': view.rows.length, 'アラート': state.alerts.length, '抜け': 0 }; // アラートは上の帯の札と同じ数（行につかない物も含む）
    C.FORMS.forEach((f) => { out[f] = 0; });
    view.rows.forEach((r) => {
      const pid = String(r['商品ID']);
      if (state.gaps.price.has(pid) || state.gaps.cost.has(pid)) out['抜け']++;
      C.FORMS.forEach((f) => { if (matchesForm(r, f)) out[f]++; });
    });
    return out;
  }
  function paintChips() {
    const n = counts();
    document.querySelectorAll('#chips .chip').forEach((b) => { const c = b.dataset.chip; b.classList.toggle('is-active', c === view.chip); b.setAttribute('aria-pressed', String(c === view.chip)); b.replaceChildren(document.createTextNode(c + ' '), el('b', '', n[c])); });
  }
  function matches(r) {
    const pid = String(r['商品ID']);
    if (view.sec && String(r['区分ID']) !== view.sec) return false;
    if (view.chip === 'アラート' && !state.alertsByPid.has(pid)) return false;
    if (view.chip === '抜け' && !(state.gaps.price.has(pid) || state.gaps.cost.has(pid))) return false;
    if (C.FORMS.indexOf(view.chip) >= 0 && !matchesForm(r, view.chip)) return false;
    if (view.terms && view.terms.length && !hitKey(keyOf(pid, r), view.terms)) return false;
    return true;
  }
  function applyFilter() {
    view.rows.forEach((r) => {
      const tr = view.rowEls.get(String(r['行ID'])); if (!tr) return;
      const show = matches(r);
      if (tr.hidden === show) tr.hidden = !show;
      const next = tr.nextElementSibling;
      if (next && next.classList.contains('notice-row')) next.hidden = !show;
      const det = next && next.classList.contains('notice-row') ? next.nextElementSibling : next;
      if (det && det.classList.contains('detail-row')) det.hidden = !show;
    });
    paintChips();
  }
  function setChip(c) { view.chip = c; if (view.built) { applyFilter(); renderOthers(); const s = $('#list-scroll'); if (s) s.scrollTop = 0; } }

  // ------------------------------------------------------------ 足す・取り消す: その1行だけ（全部を作り直すと、行が多いぶん待たせる）
  function insertRowOf(rowId) {
    const r = db().rows.get(rowId); if (!r) return null;
    view.rows = collect();
    const i = view.rows.findIndex((x) => String(x['行ID']) === rowId);
    const holder = document.createElement('tbody'); holder.innerHTML = rowHtml(r); const tr = holder.firstElementChild;
    let before = null;
    for (let j = i + 1; j < view.rows.length && !before; j++) before = view.rowEls.get(String(view.rows[j]['行ID'])) || null;
    $('#list-body').insertBefore(tr, before);
    view.rowEls.set(rowId, tr);
    if (!view.byPid.has(tr.dataset.pid)) view.byPid.set(tr.dataset.pid, []);
    view.byPid.get(tr.dataset.pid).push(tr);
    paintRow(r, true);
    return tr;
  }
  function removeRowsOf(pid, rowIds) {
    if (!view.built) return;
    if (view.open && rowIds.indexOf(view.open) >= 0) closeDetail();
    rowIds.forEach((id) => {
      const tr = view.rowEls.get(id); if (!tr) return;
      let next = tr.nextElementSibling;
      while (next && !next.classList.contains('lr')) { const after = next.nextElementSibling; next.remove(); next = after; } // その行のお知らせと入力
      tr.remove(); view.rowEls.delete(id);
    });
    view.byPid.delete(pid); view.keys.delete(pid);
    view.rows = collect();
  }

  // ------------------------------------------------------------ 打つ（その場で直す）
  function cellRaw(input) {
    const tr = input.closest('tr.lr'); if (!tr) return '';
    const t = input.dataset.t, f = input.dataset.f;
    if (t === '商品') { const p = db().products.get(tr.dataset.pid); return p ? p[f] : ''; }
    if (t === '行') { const r = db().rows.get(tr.dataset.rid); return r ? r[f] : ''; }
    return ((db().costs && db().costs.get(tr.dataset.pid)) || {})[f];
  }
  /** quiet＝打っている途中（読めない形でも知らせを出さない。確定のときに知らせる） */
  function commitInput(input, quiet) {
    const tr = input.closest('tr.lr'); if (!tr) return;
    const t = input.dataset.t, f = input.dataset.f; const pid = tr.dataset.pid, rowId = tr.dataset.rid;
    if (f === '指定トレー') return; // 包材は選ぶ・足すときだけ保存する（打ち間違いで増えない）
    let v = input.value;
    if (f === '呼出番') { const c = C.halfCall(v); if (c === null) { input.classList.add('is-bad'); input.title = '呼出番は4桁の数字で入れてください'; if (!quiet) toast('呼出番は4桁の数字で入れてください（' + v + '）', true); return; } v = c; }
    else if (f === '小パック') v = String(v).trim();
    else v = toStored(v);
    const id = t === '行' ? rowId : pid;
    if (showVal(cellRaw(input)) === showVal(v)) { input._dirty = false; input.classList.remove('is-bad'); input.removeAttribute('title'); return; }
    view.fx.busy = input; // 保存の途中で描き直さない（離れる途中は「打っている欄」に見えないので、式が結果の数に置き換わり、それを保存してしまう）
    const problem = App.change([{ t, id, f: { [f]: v } }]);
    view.fx.busy = null;
    if (problem) { input.classList.add('is-bad'); input.title = problem; if (!quiet) toast(problem, true); return; }
    input._dirty = false; input.classList.remove('is-bad'); input.removeAttribute('title');
    if (f === '呼出番' || f === '小パック' || f === '名前') view.keys.delete(pid);
    if (t === '原価') paintGross(tr, db().costs.get(pid) || {});
  }
  const pendingInputs = new Map();
  function schedule(input) { window.clearTimeout(pendingInputs.get(input)); pendingInputs.set(input, window.setTimeout(() => commitInput(input, true), 350)); }
  function unschedule(input) { window.clearTimeout(pendingInputs.get(input)); pendingInputs.delete(input); }
  // ---- 数式バー（セルと同じ保存の道を通る）
  const FX_NAMES = { '呼出番': '番号', '指定トレー': '包材', '定番価格（包材変更後）': '定番（包材後）' };
  function fxShow(input) {
    const fx = $('#fx-input'), label = $('#fx-label'); if (!fx) return;
    const tr = input.closest('tr.lr'); const p = db().products.get(tr.dataset.pid);
    label.textContent = (FX_NAMES[input.dataset.f] || input.dataset.f) + '｜' + (showVal(p && p['呼出番']) || '番号なし');
    fx.disabled = false; fx.value = input.value;
  }
  function selectCell(input) {
    if (view.fx.cell && view.fx.cell !== input) { const old = view.fx.cell.closest('td'); if (old) old.classList.remove('is-fx'); }
    view.fx.cell = input; const td = input.closest('td'); if (td) td.classList.add('is-fx');
  }
  /** セルから離れた: 打った物を保存して、結果の見え方へ戻す */
  function settleCell(input) {
    unschedule(input);
    if (input.dataset.f === '指定トレー') { trayLeave(input); return; }
    if (input._dirty) commitInput(input); // 人が打ったときだけ保存する（結果の見え方の数を保存しない）
    if (!input.classList.contains('is-bad')) { const v = C.showCell(cellRaw(input)); if (input.value !== v) input.value = v; }
  }
  function restoreCell(input) {
    unschedule(input); input.value = input._v0 === undefined ? showVal(cellRaw(input)) : input._v0;
    input.classList.remove('is-bad'); input.removeAttribute('title');
    if (input.dataset.f === '指定トレー') trayClose(); else if (showVal(cellRaw(input)) !== input.value) commitInput(input);
    input._dirty = false;
    if (view.fx.cell === input) $('#fx-input').value = input.value;
  }
  function onFocusIn(e) {
    const i = e.target; if (!isCell(i)) return;
    if (e.relatedTarget !== $('#fx-input') || view.fx.cell !== i) {
      const raw = showVal(cellRaw(i)); i._v0 = raw;
      if (!i.classList.contains('is-bad')) i.value = raw;
      if (!view.fx.mouse) i.select();
    }
    selectCell(i); fxShow(i);
  }
  function onFocusOut(e) {
    const i = e.target; if (!isCell(i) || e.relatedTarget === $('#fx-input')) return;
    settleCell(i);
  }
  function onInput(e) {
    const input = e.target; if (!isCell(input)) return;
    input._dirty = true;
    if (view.fx.cell === input) $('#fx-input').value = input.value;
    if (input.dataset.f === '指定トレー') { trayOpen(input, input.value); return; }
    if (input.dataset.f === '呼出番' && C.halfCall(input.value) === null) { unschedule(input); return; } // 打ちかけの間は確かめない
    schedule(input);
  }
  function onChange(e) { const input = e.target; if (!isCell(input) || !input._dirty) return; unschedule(input); commitInput(input); }
  function nextRowCell(input, step) {
    let tr = input.closest('tr.lr'); const f = input.dataset.f;
    for (;;) {
      tr = step < 0 ? tr.previousElementSibling : tr.nextElementSibling; if (!tr) return null;
      if (tr.classList.contains('lr') && !tr.hidden) { const next = tr.querySelector('input[data-f="' + f + '"]'); if (next) return next; }
    }
  }
  /** Tab の行き先（数式バーから）: 同じ行の右のセル。無ければ次の行の先頭 */
  function nextSideCell(input, step) {
    let tr = input.closest('tr.lr'); const cells = [...tr.querySelectorAll('input[data-f]')]; const i = cells.indexOf(input) + step;
    if (cells[i]) return cells[i];
    for (;;) {
      tr = step < 0 ? tr.previousElementSibling : tr.nextElementSibling; if (!tr) return null;
      if (tr.classList.contains('lr') && !tr.hidden) { const list = [...tr.querySelectorAll('input[data-f]')]; return list[step < 0 ? list.length - 1 : 0]; }
    }
  }
  function goTo(next) { if (next) { next.focus(); next.select(); } }
  /** Enter で確定して下へ・上下の矢印で同じ列の上下へ（スプシと同じ）・Esc で元に戻す・Tab はふつうに右へ */
  function onKey(e) {
    const input = e.target; if (!isCell(input)) return;
    if (input.dataset.f === '指定トレー' && trayKey(e, input)) return;
    if (e.key === 'Escape') { e.preventDefault(); restoreCell(input); input.select(); return; }
    const down = e.key === 'Enter' || e.key === 'ArrowDown', up = e.key === 'ArrowUp';
    if (!down && !up) return;
    e.preventDefault();
    const next = nextRowCell(input, up ? -1 : 1);
    if (next) goTo(next); else input.blur();
  }
  function bindFx() {
    const fx = $('#fx-input');
    fx.addEventListener('input', () => {
      const c = view.fx.cell; if (!c) return;
      c.value = fx.value; c._dirty = true;
      if (c.dataset.f === '指定トレー') { trayOpen(c, fx.value); return; }
      if (c.dataset.f === '呼出番' && C.halfCall(c.value) === null) { unschedule(c); return; }
      schedule(c);
    });
    fx.addEventListener('keydown', (e) => {
      const c = view.fx.cell; if (!c) return;
      if (c.dataset.f === '指定トレー' && trayKey(e, c)) return;
      if (e.key === 'Escape') { e.preventDefault(); restoreCell(c); c.focus(); return; }
      if (e.key === 'Enter') { e.preventDefault(); goTo(nextRowCell(c, 1) || c); return; }
      if (e.key === 'Tab') { e.preventDefault(); goTo(nextSideCell(c, e.shiftKey ? -1 : 1) || c); }
    });
    fx.addEventListener('blur', (e) => { const c = view.fx.cell; if (c && e.relatedTarget !== c) settleCell(c); });
  }

  // ------------------------------------------------------------ 包材を選ぶ欄（選ぶ＋打って絞る。1つの部品を使い回す＝行ごとに作らない）
  const tray = { input: null, items: [], active: -1, memo: new Map() };
  const trayList = () => (db().options['指定トレー'] || []).map((o) => String(o['値']));
  function trayHit(v, q) {
    const nq = C.normTray(q).toLowerCase(); if (C.normTray(v).toLowerCase().includes(nq)) return true;
    if (!SN) return false;
    let k = tray.memo.get(v); if (k === undefined) { k = SN.key(v); tray.memo.set(v, k); }
    const terms = makeTerms(q); return terms.length > 0 && SN.hit(k, terms);
  }
  /** 一覧のセルと1商品の入力で同じ部品。_tray があればその欄の決まり */
  function trayCtx(input) {
    if (input._tray) return input._tray;
    const pid = input.closest('tr.lr').dataset.pid;
    return { raw: () => showVal(db().products.get(pid)['指定トレー']), save: (v) => App.change([{ t: '商品', id: pid, f: { '指定トレー': v } }]) };
  }
  function trayClose() { const m = $('#tray-menu'); if (m) m.hidden = true; tray.input = null; tray.active = -1; }
  const trayIsOpen = (input) => tray.input === input && !$('#tray-menu').hidden;
  function trayOpen(input, q) {
    const t0 = performance.now();
    const m = $('#tray-menu'); tray.input = input;
    const all = trayList(); const cur = trayCtx(input).raw();
    const text = q === undefined ? '' : String(q);
    const filter = text.trim() !== '' && C.normTray(text) !== cur;
    tray.items = filter ? all.filter((v) => trayHit(v, text)) : all;
    tray.active = filter ? (tray.items.length ? 0 : -1) : tray.items.indexOf(cur);
    const list = el('div', 'tray-list'); list.setAttribute('role', 'listbox'); list.setAttribute('aria-label', '包材の一覧');
    tray.items.forEach((v, i) => { const b = el('button', 'tray-item' + (i === tray.active ? ' is-active' : '') + (v === cur ? ' is-current' : '')); b.type = 'button'; b.setAttribute('role', 'option'); b.dataset.i = String(i); b.append(el('span', 'chk', v === cur ? '✓' : ''), document.createTextNode(v)); list.append(b); });
    if (!tray.items.length) list.append(el('p', 'menu-note', '一覧に当たる包材がありません'));
    m.replaceChildren(list);
    const n = C.normTray(text);
    if (n && all.indexOf(n) < 0) { const add = el('button', 'btn btn-small tray-add', '「' + n + '」を新しい包材として足す'); add.type = 'button'; add.id = 'tray-add'; add.dataset.value = n; m.append(add); }
    m.hidden = false;
    trayPlace(input);
    const on = m.querySelector('.tray-item.is-active'); if (on) on.scrollIntoView({ block: 'nearest' });
    App.perfNote('包材を絞る', t0);
  }
  function trayPlace(input) {
    const m = $('#tray-menu'); const rect = input.getBoundingClientRect(); const h = m.offsetHeight, w = m.offsetWidth;
    m.style.left = Math.max(4, Math.min(rect.left, window.innerWidth - w - 4)) + 'px';
    m.style.top = (rect.bottom + h + 4 > window.innerHeight ? Math.max(4, rect.top - h - 4) : rect.bottom + 2) + 'px';
  }
  /** 送られたら欄についていく（カーソルを置いた瞬間の送りで閉じない）。欄が見えなくなったら閉じる */
  function trayFollow(scroll) {
    const r = tray.input.getBoundingClientRect(), s = scroll.getBoundingClientRect();
    if (r.bottom < s.top || r.top > s.bottom) trayClose(); else trayPlace(tray.input);
  }
  function trayMove(step) {
    const n = tray.items.length; if (!n) return;
    tray.active = tray.active < 0 ? (step > 0 ? 0 : n - 1) : (tray.active + step + n) % n;
    $('#tray-menu').querySelectorAll('.tray-item').forEach((b, i) => { b.classList.toggle('is-active', i === tray.active); if (i === tray.active) b.scrollIntoView({ block: 'nearest' }); });
  }
  function trayPick(v) {
    const input = tray.input; if (!input) return;
    const ctx = trayCtx(input); trayClose();
    input.value = v; input._v0 = v; if (view.fx.cell === input) $('#fx-input').value = v;
    if (v !== ctx.raw()) { const m = ctx.save(v); if (m) toast(m, true); }
  }
  /** 欄から離れるとき: 一覧にある値（そろえて同じ）か空なら保存。無い値は足さずに元へ戻す */
  function trayLeave(input) {
    const ctx = trayCtx(input); const n = C.normTray(input.value); const cur = ctx.raw();
    if (tray.input === input) trayClose();
    if (n === cur) { input.value = cur; return; }
    if (!n || trayList().indexOf(n) >= 0) { input.value = n; input._v0 = n; const m = ctx.save(n); if (m) toast(m, true); return; }
    input.value = cur; input._v0 = cur;
    toast('一覧に無い包材です（' + n + '）。足すときは「新しい包材として足す」を押してください。', true);
  }
  /** 返す値: このキーを包材の欄で使ったか */
  function trayKey(e, input) {
    if (e.altKey && e.key === 'ArrowDown') { e.preventDefault(); trayOpen(input); return true; }
    if (!trayIsOpen(input)) return false;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); trayMove(e.key === 'ArrowDown' ? 1 : -1); return true; }
    if (e.key === 'Enter') {
      e.preventDefault();
      const n = C.normTray(input.value);
      if (tray.active >= 0) trayPick(tray.items[tray.active]);
      else if (trayList().indexOf(n) >= 0) trayPick(n);
      else if (n) toast('一覧に無い包材です。足すときは「新しい包材として足す」を押してください。');
      return true;
    }
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); restoreCell(input); return true; }
    if (e.key === 'Tab') trayClose();
    return false;
  }
  function trayAdd(value) {
    const list = db().options['指定トレー'] || (db().options['指定トレー'] = []);
    list.push({ '項目': '指定トレー', '値': value });
    App.serial(() => App.api('addOption', { item: '指定トレー', value, requestId: App.rid() })).catch((error) => {
      const i = list.findIndex((o) => String(o['値']) === value); if (i >= 0) list.splice(i, 1);
      toast('包材を足せませんでした: ' + error.message, true);
    });
    trayPick(value);
  }
  { // 押してもカーソルを欄から動かさない（1つだけ作る）
    const m = $('#tray-menu');
    m.addEventListener('mousedown', (e) => e.preventDefault());
    m.addEventListener('click', (e) => {
      const add = e.target.closest('#tray-add'); if (add) { trayAdd(add.dataset.value); return; }
      const b = e.target.closest('.tray-item'); if (b) trayPick(tray.items[Number(b.dataset.i)]);
    });
  }

  // ------------------------------------------------------------ 販売形態を選ぶ（凡例と同じ色）
  function closeFormMenu() { const m = $('#form-menu'); m.hidden = true; view.menuFor = null; }
  function openFormMenu(btn, r) {
    const m = $('#form-menu'); m.replaceChildren();
    const paint = () => m.querySelectorAll('button').forEach((b) => { const on = C.formHas(formOfRow(r), b.dataset.form); b.classList.toggle('is-current', on); b.setAttribute('aria-checked', String(on)); b.querySelector('.chk').textContent = on ? '✓' : ''; });
    C.FORMS.forEach((f) => {
      const multi = f === C.FORM_SALE || f === C.FORM_NEW; // 月間特売と新商品だけ両方選べる（押しても閉じない）
      const b = el('button'); b.type = 'button'; b.dataset.form = f; b.setAttribute('role', multi ? 'menuitemcheckbox' : 'menuitemradio');
      const text = el('span'); text.append(document.createTextNode(f), el('small', '', FORM_NOTE[f]));
      b.append(el('span', 'chk'), el('span', 'swatch ' + FORM_CLASS[f]), text);
      b.addEventListener('click', () => {
        const next = C.toggleForm(formOfRow(r), f);
        if (next !== formOfRow(r)) setForm(r, next);
        if (multi) { paint(); b.focus(); } else { closeFormMenu(); btn.focus(); }
      });
      m.append(b);
    });
    m.append(el('p', 'menu-note', '月間特売と新商品は両方選べます（ほかの所を押すと閉じます）'));
    paint();
    m.hidden = false; view.menuFor = btn;
    const rect = btn.getBoundingClientRect(); const h = m.offsetHeight, w = m.offsetWidth;
    m.style.left = Math.max(4, Math.min(rect.left, window.innerWidth - w - 4)) + 'px';
    m.style.top = (rect.bottom + h + 4 > window.innerHeight ? Math.max(4, rect.top - h - 4) : rect.bottom + 2) + 'px';
    (m.querySelector('button.is-current') || m.querySelector('button')).focus();
  }
  function setForm(r, value, region) {
    const f = {};
    if (region === '北関東') {
      const d = Object.assign({}, r._diff || {});
      if (value === '' || value === formOfRow(r)) delete d['販売形態']; else d['販売形態'] = value;
      f['北関東の違い'] = Object.keys(d).length ? JSON.stringify(d) : '';
    } else {
      // カット商品にする・やめると、行は ≪カット商品≫ の区分と元の区分の間を動く（決まりは core.js の formChanges の1か所）
      const fromKind = App.kindOfSection(r['区分ID']);
      const changes = C.formChanges(App.state.db, r, value);
      const moved = changes.some((c) => c.f['区分ID'] !== undefined);
      if (moved) App.flush(); // 動きのある直しは、ほかの直しとまとめずに送る（窓口が置き場の手直しに数えないように）
      const p = App.change(changes, moved ? { label: 'カット商品' } : undefined);
      if (p) { toast(p, true); return; }
      if (moved) { App.flush(); const n = changes.filter((c) => c.f['区分ID'] !== undefined).length; toast(C.isCutForm(value) ? '≪カット商品≫ の区分へ移しました' + (n > 1 ? '（つながった ' + n + ' 行）' : '') + '。' : '元の区分へ戻しました' + (n > 1 ? '（つながった ' + n + ' 行）' : '') + '。'); }
      paintChips();
      const toKind = App.kindOfSection(r['区分ID']);
      App.afterFormChange(fromKind);
      if (toKind !== fromKind) App.afterFormChange(toKind);
      return;
    }
    const p = App.change([{ t: '行', id: String(r['行ID']), f }]);
    if (p) { toast(p, true); return; }
    paintChips();
    App.afterFormChange(App.kindOfSection(r['区分ID']));
  }
  document.addEventListener('click', (e) => { if (!$('#form-menu').hidden && !e.target.closest('#form-menu') && !e.target.closest('.form-btn')) closeFormMenu(); });
  document.addEventListener('keydown', (e) => {
    const m = $('#form-menu'); if (m.hidden) return;
    if (e.key === 'Escape') { const b = view.menuFor; closeFormMenu(); if (b) b.focus(); return; }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); const list = [...m.querySelectorAll('button')]; const i = list.indexOf(document.activeElement); list[(i + (e.key === 'ArrowDown' ? 1 : list.length - 1)) % list.length].focus(); }
  });

  // ------------------------------------------------------------ 1商品の入力（すぐ下に開く）
  function closeDetail() {
    const old = document.querySelector('tr.detail-row'); if (old) old.remove();
    if (view.open) { const tr = view.rowEls.get(view.open); if (tr) { const b = tr.querySelector('.name-btn'); b.setAttribute('aria-expanded', 'false'); b.querySelector('.tw').textContent = '▶'; } }
    view.open = '';
  }
  /** 販売形態の選ぶ所（1商品の入力・新しい商品）。一覧の行と同じ決まり */
  function formPicker(getValue, onPick, ariaLabel, forms) {
    const box = el('div', 'form-pick'); box.setAttribute('role', 'group'); box.setAttribute('aria-label', ariaLabel);
    const paint = () => box.querySelectorAll('button').forEach((b) => { const on = C.formHas(getValue(), b.dataset.form); b.classList.toggle('is-on', on); b.setAttribute('aria-pressed', String(on)); b.querySelector('.chk').textContent = on ? '✓' : ''; });
    (forms || C.FORMS).forEach((f) => { const b = el('button'); b.type = 'button'; b.dataset.form = f; b.append(el('span', 'chk'), el('span', 'swatch ' + FORM_CLASS[f]), el('span', '', f)); b.addEventListener('click', () => { const next = C.toggleForm(getValue(), f); if (next !== getValue()) onPick(next); paint(); }); box.append(b); });
    paint(); box.repaint = paint;
    return box;
  }
  function field(label, node, cls) { const w = el('label', cls || ''); w.append(document.createTextNode(label), node); return w; }
  function textInput(value, onCommit, attrs) {
    const i = el('input'); i.value = showVal(value); Object.assign(i, attrs || {});
    let timer = 0;
    i.addEventListener('input', () => { window.clearTimeout(timer); timer = window.setTimeout(() => onCommit(i.value, i), 500); });
    i.addEventListener('change', () => { window.clearTimeout(timer); onCommit(i.value, i); });
    return i;
  }
  function openDetail(tr) {
    closeDetail();
    const rowId = tr.dataset.rid, pid = tr.dataset.pid;
    const r = db().rows.get(rowId), p = db().products.get(pid);
    view.open = rowId;
    const b = tr.querySelector('.name-btn'); b.setAttribute('aria-expanded', 'true'); b.querySelector('.tw').textContent = '▼';
    const row = el('tr', 'detail-row'); const td = el('td'); td.colSpan = view.colCount;
    const box = el('div', 'detail'); box.style.setProperty('--detail-w', Math.min($('#list-scroll').clientWidth, window.innerWidth) + 'px');
    td.append(box); row.append(td);
    const after = tr.nextElementSibling && tr.nextElementSibling.classList.contains('notice-row') ? tr.nextElementSibling : tr;
    after.after(row);
    renderDetail(box, r, p);
    const s = $('#list-scroll'); const rr = row.getBoundingClientRect(), sr = s.getBoundingClientRect();
    if (rr.bottom > sr.bottom) s.scrollTop += Math.min(rr.bottom - sr.bottom + 8, rr.top - sr.top - 60);
  }
  function renderDetail(box, r, p) {
    const pid = String(p['商品ID']); const kind = p['表の種類'];
    box.replaceChildren();
    const title = el('div', 'detail-title'); title.append(el('h3', '', '1商品の入力'), el('span', 'hint-line', '打つと自動で保存します'));
    const close = el('button', 'btn btn-small', '閉じる'); close.type = 'button'; close.addEventListener('click', closeDetail);
    const copy = el('button', 'btn btn-small', '写して新しい商品'); copy.type = 'button'; copy.addEventListener('click', () => startCopy(p, r));
    title.append(copy);
    if (C.isAddedProduct(p)) { // 足した商品だけ取り消せる（消えるので2回押す）
      const undo = el('button', 'btn btn-small btn-danger', 'この商品を取り消す'); undo.type = 'button'; undo.id = 'detail-cancel';
      undo.addEventListener('click', () => {
        if (undo.dataset.sure !== '1') { undo.dataset.sure = '1'; undo.textContent = 'もう一度押すと、この商品と行を消します'; window.setTimeout(() => { if (undo.isConnected) { undo.dataset.sure = ''; undo.textContent = 'この商品を取り消す'; } }, 5000); return; }
        App.cancelProduct(pid, undo);
      });
      title.append(undo);
    }
    title.append(close); box.append(title);
    const grid = el('div', 'detail-grid');
    const save1 = (f, v) => App.change([{ t: '商品', id: pid, f: { [f]: v } }]);
    // 名前（字を選んで赤くする・戻す。入数と連動）
    const nameInput = textInput(p['名前'], (v, i) => { if (String(v) === String(p['名前'])) return; const old = String(p['名前']); const deco = Object.assign({}, p._deco || {}); if (deco['名前']) deco['名前'] = C.editRuns(deco['名前'], old, v); const f = { '名前': v }; if (deco['名前']) f['字の飾り'] = Object.keys(deco).length ? JSON.stringify(deco) : ''; const m = App.change([{ t: '商品', id: pid, f }]); if (m) { toast(m, true); return; } view.keys.delete(pid); paintPreview(); syncQty(); });
    const nameWrap = field('名前', nameInput, 'wide');
    const tools = el('div', 'name-tools');
    const preview = el('span', 'name-preview');
    function paintPreview() { preview.replaceChildren(); const text = String(p['名前'] || ''); let at = 0; ((p._deco && p._deco['名前']) || []).slice().sort((a, b) => a[0] - b[0]).forEach((run) => { const from = Math.min(run[0], text.length), to = Math.min(run[1], text.length); if (from > at) preview.append(document.createTextNode(text.slice(at, from))); const s = el('span', '', text.slice(from, to)); if (run[2] && run[2].c) s.style.color = '#' + run[2].c; preview.append(s); at = to; }); if (at < text.length) preview.append(document.createTextNode(text.slice(at))); }
    const paint = (props) => { const a = nameInput.selectionStart, z = nameInput.selectionEnd; if (a === null || a === z) { toast('赤くしたい字を、名前の欄で選んでください。'); return; } const deco = Object.assign({}, p._deco || {}); deco['名前'] = C.paintRuns(deco['名前'] || [], a, z, props); if (!deco['名前'].length) delete deco['名前']; const m = App.change([{ t: '商品', id: pid, f: { '字の飾り': Object.keys(deco).length ? JSON.stringify(deco) : '' } }]); if (m) toast(m, true); paintPreview(); };
    const red = el('button', 'btn btn-small', '選んだ字を赤くする'); red.type = 'button'; red.addEventListener('mousedown', (e) => e.preventDefault()); red.addEventListener('click', () => paint({ c: 'FF0000' }));
    const unred = el('button', 'btn btn-small', '赤を戻す'); unred.type = 'button'; unred.addEventListener('mousedown', (e) => e.preventDefault()); unred.addEventListener('click', () => paint(null));
    tools.append(red, unred, preview); nameWrap.append(tools); grid.append(nameWrap); paintPreview();
    // 入数（名前の中のカッコと連動）
    const units = (db().options['入数の単位'] || []).map((o) => String(o['値']));
    const qtyCount = el('input'); qtyCount.inputMode = 'numeric'; qtyCount.setAttribute('aria-label', '入数の数');
    const qtyUnit = el('select'); qtyUnit.setAttribute('aria-label', '入数の単位'); [''].concat(units).forEach((u) => { const o = el('option', '', u || '単位'); o.value = u; qtyUnit.append(o); });
    function syncQty() { const q = C.parseQty(p['名前']); if (document.activeElement !== qtyCount) qtyCount.value = q.count || ''; if (q.unit && units.indexOf(q.unit) < 0) { const o = el('option', '', q.unit); o.value = q.unit; qtyUnit.append(o); } qtyUnit.value = q.unit || ''; }
    const applyQty = () => { const v = C.setQty(p['名前'], qtyCount.value, qtyUnit.value); if (v === p['名前']) return; const old = String(p['名前']); const deco = Object.assign({}, p._deco || {}); const f = { '名前': v }; if (deco['名前']) { deco['名前'] = C.editRuns(deco['名前'], old, v); f['字の飾り'] = JSON.stringify(deco); } const m = App.change([{ t: '商品', id: pid, f }]); if (m) { toast(m, true); return; } view.keys.delete(pid); if (document.activeElement !== nameInput) nameInput.value = v; paintPreview(); };
    qtyCount.addEventListener('change', applyQty); qtyUnit.addEventListener('change', applyQty);
    const qty = el('div', 'qty'); qty.append(qtyCount, qtyUnit); grid.append(field('入数（名前のカッコと連動）', qty)); syncQty();
    // 呼出番
    const hint = el('div', 'hint-line', C.hintOf(db(), pid));
    const callInput = textInput(p['呼出番'], (v, i) => { const c = C.halfCall(v); if (c === null) { if (document.activeElement !== i) toast('呼出番は4桁の数字で入れてください', true); i.classList.add('is-bad'); return; } i.classList.remove('is-bad'); if (String(p['呼出番'] ?? '') === c) return; save1('呼出番', c); view.keys.delete(pid); paintFree(); }, { inputMode: 'numeric', maxLength: 20 });
    const free = el('div', 'free-calls');
    function paintFree() { free.replaceChildren(document.createTextNode('空いている番号: ')); C.freeCalls(db(), p['呼出番'], 4).forEach((n) => { const b = el('button', 'btn btn-small', n); b.type = 'button'; b.addEventListener('click', () => { callInput.value = n; save1('呼出番', n); view.keys.delete(pid); paintFree(); }); free.append(b); }); }
    const callWrap = field('呼出番', callInput); callWrap.append(hint, free); grid.append(callWrap); paintFree();
    // 販売形態（北関東だけ違うときは別に選べる）
    const formWrap = el('div', 'field wide'); formWrap.append(document.createTextNode('販売形態（月間特売と新商品は両方選べます）'));
    const formBox = formPicker(() => formOfRow(r), (next) => { setForm(r, next); paintShown(); paintNorth(); }, '販売形態');
    formWrap.append(formBox); grid.append(formWrap);
    // 北関東だけ違うとき（首都圏と同じなら開かない）
    const northWrap = el('div', 'field wide');
    const northCheck = el('input'); northCheck.type = 'checkbox'; northCheck.className = 'north-check';
    const northLabel = el('label', 'north-label'); northLabel.append(northCheck, document.createTextNode('北関東だけ違う'));
    const northBox = formPicker(() => (northForm(r) === null ? formOfRow(r) : northForm(r)), (next) => { setForm(r, next, '北関東'); paintShown(); }, '北関東の販売形態', C.FORMS.filter((f) => !C.isCutForm(f))); // カット商品は置き場が変わるので、首都圏と北関東で分けない
    function paintNorth() { northCheck.checked = northForm(r) !== null; northBox.hidden = !northCheck.checked; northBox.repaint(); }
    northCheck.addEventListener('change', () => { northBox.hidden = !northCheck.checked; if (!northCheck.checked) { setForm(r, '', '北関東'); paintShown(); } northBox.repaint(); });
    northWrap.append(northLabel, northBox); paintNorth();
    if (kind === '寿司弁当' && (p['地域'] || '共通') === '共通') grid.append(northWrap);
    // 価格は隣に
    const priceInput = textInput(p['定番価格'], (v) => { const n = toStored(v); if (showVal(n) !== showVal(p['定番価格'])) { save1('定番価格', n); paintShown(); } });
    const saleInput = textInput(r['特売価格'], (v) => { const n = toStored(v); if (showVal(n) !== showVal(r['特売価格'])) { App.change([{ t: '行', id: String(r['行ID']), f: { '特売価格': n } }]); paintShown(); } });
    const shownLine = el('div', 'hint-line wide');
    function paintShown() { const parts = []; (C.KIND_TABLES[kind] || []).forEach((t) => { if (C.rowRegionOk(db(), r, t)) parts.push(t + ' ' + (C.showCell(C.shownPrice(p, r, t)) || '（空）') + '（' + C.shownPriceKey(r, t) + '）'); }); shownLine.textContent = '番号表に出る価格: ' + parts.join('・'); }
    paintShown();
    grid.append(field('定番価格', priceInput), field('特売価格', saleInput), shownLine);
    grid.append(field('小パック', textInput(p['小パック'], (v) => { const t = String(v).trim(); if (t !== showVal(p['小パック'])) { save1('小パック', t); view.keys.delete(pid); } })));
    // 包材（選ぶ＋打って絞る。一覧と同じ部品）
    const trayInput = el('input'); trayInput.value = showVal(p['指定トレー']); trayInput.setAttribute('role', 'combobox'); trayInput.setAttribute('aria-autocomplete', 'list'); trayInput.className = 'tray-input';
    trayInput._tray = { raw: () => showVal(p['指定トレー']), save: (v) => save1('指定トレー', v) };
    trayInput.addEventListener('click', () => trayOpen(trayInput));
    trayInput.addEventListener('input', () => trayOpen(trayInput, trayInput.value));
    trayInput.addEventListener('keydown', (e) => { if (trayKey(e, trayInput)) return; if (e.key === 'Enter') e.preventDefault(); });
    trayInput.addEventListener('blur', () => trayLeave(trayInput));
    grid.append(field('包材', trayInput));
    // 区分（動かす核を通す）
    const secSel = el('select'); C.sectionsOfKind(db(), kind).forEach((s) => { const o = el('option', '', C.nfkc(s['表題'])); o.value = String(s['区分ID']); secSel.append(o); }); secSel.value = String(r['区分ID']);
    secSel.addEventListener('change', () => { if (App.doMove({ kind: 'row', id: String(r['行ID']), toSec: secSel.value, before: null }, '区分を変える')) { const st = view.rowEls.get(String(r['行ID'])).querySelector('.c-sec-text'); const sec = db().sections.get(secSel.value); st.textContent = sec['表題']; view.keys.delete(pid); } });
    grid.append(field('区分', secSel));
    // 地域
    const regWrap = el('div', 'field'); regWrap.append(document.createTextNode('地域'));
    const tabs = el('div', 'region-tabs'); tabs.setAttribute('role', 'group'); tabs.setAttribute('aria-label', '地域');
    ['共通', '首都圏だけ', '北関東だけ'].forEach((reg) => { const b = el('button', (p['地域'] || '共通') === reg ? 'is-active' : '', reg); b.type = 'button'; b.addEventListener('click', () => { const m = App.change([{ t: '商品', id: pid, f: { '地域': reg } }]); if (m) { toast(m, true); return; } tabs.querySelectorAll('button').forEach((x) => x.classList.toggle('is-active', x === b)); paintRow(r); }); tabs.append(b); });
    regWrap.append(tabs); if (kind === '寿司弁当') grid.append(regWrap);
    grid.append(field('メモ', textInput(p['メモ'], (v) => { if (String(v) !== showVal(p['メモ'])) save1('メモ', v); }), 'wide'));
    // 隠しの計算
    if (state.showCosts) {
      const c = db().costs.get(pid) || {};
      C.COST_KEYS.forEach((k) => grid.append(field(k, textInput(c[k], (v, i) => { const n = String(v).trim().startsWith('=') ? String(v).trim() : toStored(v); if (showVal(n) === showVal(c[k])) return; const m = App.change([{ t: '原価', id: pid, f: { [k]: n } }]); if (m) { i.classList.add('is-bad'); toast(m, true); return; } i.classList.remove('is-bad'); paintRow(r); }))));
    } else {
      grid.append(field('原価', Object.assign(el('input'), { disabled: true, value: '「原価を表示」を押すと入れられます' })));
    }
    box.append(grid);
    // この商品の記録
    const logBox = el('div', 'detail-log'); const logBtn = el('button', 'btn btn-small', 'この商品の記録を見る'); logBtn.type = 'button';
    logBtn.addEventListener('click', async () => {
      const done = App.busyButton(logBtn, '読んでいます…');
      try { const data = await App.api('getHistory', { target: pid, limit: 20, withCosts: state.showCosts }); const rows = data.rows || []; const t = el('table'); t.append(el('tbody')); if (!rows.length) { logBox.append(el('p', 'hint-line', 'まだ記録はありません（行の記録は 記録 のタブにあります）。')); } rows.forEach((x) => { const tr2 = el('tr'); [String(x['時刻']).replace('T', ' ').slice(0, 16), x['メール'], x['欄'] || x['操作'], x['前の値'], x['後の値']].forEach((v) => tr2.append(el('td', '', v))); t.tBodies[0].append(tr2); }); if (rows.length) logBox.append(t); logBtn.remove(); }
      catch (error) { toast(error.message, true); } finally { done(); }
    });
    logBox.append(logBtn); box.append(logBox);
  }

  // ------------------------------------------------------------ 新しい商品（空きマスに打つ → 候補 → 写す／そのまま）
  const fresh = { box: null, from: null, fromLabel: '', sec: '', form: C.FORM_NEW };
  function suggestions(q) {
    const terms = makeTerms(q); if (!terms.length) return [];
    const out = [];
    view.rows.forEach((r) => { const pid = String(r['商品ID']); if (hitKey(keyOf(pid, r), terms)) { const p = db().products.get(pid); out.push({ kind: 'マスター', name: String(p['名前']), from: pid, call: p['呼出番'], sec: String(r['区分ID']), rowKind: p['表の種類'] }); } });
    db().past.forEach((h) => { if (h['商品ID'] || !h['名前']) return; const k = SN ? SN.key(String(h['名前'])) : C.nfkc(h['名前']); if (hitKey(k, terms)) out.push({ kind: '過去の名前', name: String(h['名前']), from: '', call: h['9月の番号'] || h['8月の番号'] || '', rowKind: h['表の種類'] }); });
    return out.slice(0, 12);
  }
  function closeFresh() { if (fresh.box) { fresh.box.closest('tr')?.remove(); fresh.box = null; fresh.from = null; fresh.fromLabel = ''; fresh.sec = ''; fresh.form = C.FORM_NEW; } }
  function renderFresh(q, preset) {
    const anchor = document.querySelector('tr.new-row');
    if (!fresh.box) { const row = el('tr', 'detail-row'); const td = el('td'); td.colSpan = view.colCount; fresh.box = el('div', 'suggest'); td.append(fresh.box); row.append(td); anchor.after(row); }
    const box = fresh.box; box.replaceChildren();
    const list = el('div', 'suggest-list');
    const sugg = suggestions(q);
    if (sugg.length) list.append(el('span', 'hint-line', '似た商品を選ぶと、中身を写した新しい行を1つ作ります。元の商品は変わりません。番号は空で始まります。'));
    sugg.forEach((s) => { const picked = fresh.from && fresh.from === s.from; const b = el('button', picked ? 'is-picked' : ''); b.type = 'button'; b.textContent = (picked ? '✓ ' : '') + '［' + s.kind + '］' + s.name + (s.call ? '（' + s.call + '）' : ''); b.setAttribute('aria-pressed', String(Boolean(picked))); b.addEventListener('click', () => { fresh.from = s.from || null; fresh.fromLabel = s.from ? s.name + (s.call ? '（' + s.call + '）' : '') : ''; if (s.kind === '過去の名前') $('#new-name').value = s.name; if (s.sec) fresh.sec = s.sec; renderFresh($('#new-name').value); }); list.append(b); });
    box.append(list);
    const form = el('div', 'suggest-form');
    const sec = el('select'); sec.id = 'new-sec'; secOrder().forEach((id) => { const s = db().sections.get(id); if (!s) return; const o = el('option', '', '［' + s['表の種類'] + '］' + C.nfkc(s['表題'])); o.value = id; sec.append(o); });
    if (preset && preset.sec) fresh.sec = preset.sec;
    if (fresh.sec) sec.value = fresh.sec;
    const effect = el('p', 'new-effect'); effect.id = 'new-effect';
    const paintEffect = () => { const s = db().sections.get(sec.value); const name = $('#new-name').value.trim(); effect.textContent = '押すと: 「' + name + '」を［' + (s ? C.nfkc(s['表題']) : '') + '］へ足す・番号は空' + (fresh.from ? '・元の「' + fresh.fromLabel + '」はそのまま' : '') + '（足したあとで取り消せます）'; };
    sec.addEventListener('change', () => { fresh.sec = sec.value; paintEffect(); });
    const fl = el('label'); fl.append(document.createTextNode('区分'), sec);
    const f2 = el('div', 'field'); f2.append(document.createTextNode('販売形態'), formPicker(() => fresh.form, (next) => { fresh.form = next; }, '新しい商品の販売形態'));
    const go = el('button', 'btn btn-primary', fresh.from ? '写して足す' : 'この名前で足す'); go.type = 'button'; go.id = 'new-go'; go.addEventListener('click', () => addProduct(go));
    const no = el('button', 'btn', 'やめる'); no.type = 'button'; no.addEventListener('click', () => { closeFresh(); $('#new-name').value = ''; });
    form.append(fl, f2); box.append(form);
    const actions = el('div', 'new-actions'); actions.append(effect, go, no); box.append(actions); paintEffect();
  }
  async function addProduct(button) {
    const name = $('#new-name').value.trim(); if (!name) { toast('名前を入れてください。', true); return; }
    const toSec = $('#new-sec').value, form = fresh.form;
    const pid = 'P' + App.rid(), rowId = 'R' + App.rid(); const from = fresh.from, fromLabel = fresh.fromLabel;
    const done = App.busyButton(button, '足しています…');
    App.flush();
    const t0 = performance.now();
    try {
      const data = await App.serial(() => App.api('addProduct', { month: state.month, id: pid, rowId, toSec, from: from || '', form, fields: { '名前': name }, requestId: App.rid() }));
      const product = C.rowsOf(data.product)[0], row = C.rowsOf(data.row)[0];
      C.applyChanges(db(), [{ t: '新しい商品', id: pid, f: product }, { t: '新しい行', id: rowId, f: row }]);
      if (from && db().costs) { const c = db().costs.get(from); if (c) db().costs.set(pid, Object.assign({}, c, { '商品ID': pid })); }
      closeFresh(); $('#new-name').value = '';
      const tr = insertRowOf(rowId); // 足した1行だけ差し込む（全部を作り直さない）
      App.refresh([{ t: '新しい商品', id: pid }]);
      App.afterFormChange(db().products.get(pid)['表の種類']);
      applyFilter();
      if (tr) { openDetail(tr); tr.scrollIntoView({ block: 'center' }); const c = tr.querySelector('[data-f="呼出番"]'); if (c) c.focus(); }
      App.perfNote('商品を足す', t0);
      toast('「' + name + '」を足しました。' + (from ? '元の「' + fromLabel + '」は変わっていません。' : ''), false, { label: '取り消す', run: () => App.cancelProduct(pid) });
    } catch (error) { if (error.code === 'SESSION') App.showLogin(error.message); else toast('足せませんでした: ' + error.message, true); }
    finally { done(); }
  }
  function startCopy(p, r) {
    closeDetail(); const s = $('#list-scroll'); s.scrollTop = 0;
    const input = $('#new-name'); input.value = String(p['名前']); fresh.from = String(p['商品ID']); fresh.fromLabel = String(p['名前']) + (p['呼出番'] ? '（' + p['呼出番'] + '）' : '');
    renderFresh(input.value, { sec: String(r['区分ID']), kind: p['表の種類'] });
    input.focus(); input.select();
    toast('名前を直して「写して足す」を押してください。元の商品は変わりません。番号は空で始まります。');
  }

  // ------------------------------------------------------------ つなぐ
  function bind(root) {
    const scroll = $('#list-scroll');
    scroll.addEventListener('input', onInput);
    scroll.addEventListener('change', onChange);
    scroll.addEventListener('keydown', onKey);
    scroll.addEventListener('focusin', onFocusIn);
    scroll.addEventListener('focusout', onFocusOut);
    scroll.addEventListener('mousedown', (e) => { view.fx.mouse = isCell(e.target); });
    document.addEventListener('mouseup', () => { view.fx.mouse = false; });
    bindFx();
    scroll.addEventListener('click', (e) => {
      if (isCell(e.target) && e.target.dataset.f === '指定トレー') { if (!trayIsOpen(e.target)) trayOpen(e.target); return; }
      const act = e.target.closest('[data-act]'); if (!act) return;
      const tr = act.closest('tr.lr'); if (!tr) return;
      if (act.dataset.act === 'open') { if (view.open === tr.dataset.rid) closeDetail(); else openDetail(tr); }
      if (act.dataset.act === 'form') { e.stopPropagation(); if (view.menuFor === act) closeFormMenu(); else openFormMenu(act, db().rows.get(tr.dataset.rid)); }
    });
    scroll.addEventListener('scroll', () => { if (view.menuFor) closeFormMenu(); if (tray.input) trayFollow(scroll); }, { passive: true });
    root.querySelector('#chips').addEventListener('click', (e) => { const b = e.target.closest('.chip'); if (b) setChip(b.dataset.chip); });
    root.querySelector('#list-section').addEventListener('change', (e) => { view.sec = e.target.value; applyFilter(); scroll.scrollTop = 0; });
    let timer = 0;
    root.querySelector('#list-search').addEventListener('input', (e) => { window.clearTimeout(timer); const v = e.target.value; timer = window.setTimeout(() => { const t0 = performance.now(); view.q = v; view.terms = v.trim() ? makeTerms(v) : null; applyFilter(); scroll.scrollTop = 0; App.perfNote('探す', t0); }, 120); });
    root.querySelector('#cost-button').addEventListener('click', async (e) => {
      const btn = e.currentTarget;
      if (state.showCosts) { App.hideCosts(); build(); return; }
      const done = App.busyButton(btn, '原価を読んでいます…');
      const shown = await App.showCosts();
      if (shown) { done.stage('列を作っています…'); await new Promise((r) => window.setTimeout(r, 20)); build(); }
      done();
    });
    let nt = 0;
    root.querySelector('#new-name').addEventListener('input', (e) => { window.clearTimeout(nt); const v = e.target.value; nt = window.setTimeout(() => { if (v.trim()) renderFresh(v); else closeFresh(); }, 150); });
    root.querySelector('#new-name').addEventListener('keydown', (e) => { if (e.key === 'Escape') { closeFresh(); e.target.value = ''; } });
  }

  App.views.list = {
    show() { if (!state.db) return; if (state.readOnly) { App.views.reader.show(); return; } if (!view.built) build(); else { applyFilter(); renderOthers(); } },
    reset() { view.built = false; view.keys = new Map(); view.open = ''; tray.memo = new Map(); closeFresh(); closeFormMenu(); trayClose(); if (App.views.reader) App.views.reader.reset(); },
    update(touched) {
      if (!view.built || state.readOnly) return;
      const t0 = performance.now();
      touched.forEach((pid) => { const trs = view.byPid.get(String(pid)); if (trs) trs.forEach((tr) => { const r = db().rows.get(tr.dataset.rid); if (r) paintRow(r); }); });
      paintChips();
      if (view.chip !== '全部' || view.q || view.sec) { /* 直したとたん行が消えると迷うので、絞り込みは押し直すまで保つ */ }
      renderOthers();
      App.perfNote('一覧の部分更新', t0);
    },
    setChip,
    skeleton() { const root = $('#view-list'); root.hidden = false; root.replaceChildren(); const s = el('div', 'skeleton'); s.setAttribute('aria-hidden', 'true'); for (let i = 0; i < 12; i++) s.append(el('i')); root.append(s); },
    rebuild() { view.built = false; if (state.view === 'list') build(); },
    removeRows: removeRowsOf,
    closeOpen() { view.open = ''; }
  };
})();
