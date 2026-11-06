(() => {
  'use strict';
  // 番号表の見本: セルの格子（sheet-model.js）から描く。Excel と同じ元なので、見本と Excel は分かれない。
  //   色は見本どおり（ピンクの字・水色のセル）。見出しの行と左の行番号は動かない。動かす: 区分の見出しを掴むとかたまりごと・商品の行を掴むと1行だけ（処理は App.doMove の1つ）。
  const App = window.SouzaiApp;
  const { state, $, el, C, S, SN, toast } = App;
  const TABS = ['月間商品', '首都圏', '北関東', '焼き物'];
  const view = { built: false, table: '首都圏', zoom: 'fit', sheet: null, cells: new Map(), dirty: true, merged: '', movePanel: false, drag: null, pick: null, lastRows: 0 };
  const db = () => state.db;
  const colPx = (w) => Math.round((w === null || w === undefined ? 8.43 : Number(w)) * 7 + 5);
  const BORDER = { thin: '1px solid #000', hair: '1px dotted #666', medium: '2px solid #000', thick: '3px solid #000', double: '3px double #000', dotted: '1px dotted #000', dashed: '1px dashed #000' };
  // 書式は組み合わせごとに1つのクラスにして共有する（セルごとに書かない＝描くのが速い）。CSP に触れない形（画面の中で作る表）
  const cssMemo = new Map();
  const sheetCss = new CSSStyleSheet();
  document.adoptedStyleSheets = [...document.adoptedStyleSheets, sheetCss];
  function cssOf(styles, key, fc, fl) {
    const k = key + '|' + fc + '|' + fl; if (cssMemo.has(k)) return cssMemo.get(k);
    const st = styles[key] || {};
    const color = fc !== undefined ? fc : (st['字の色'] || ''), fill = fl !== undefined ? fl : (st['塗り'] || '');
    const name = 'sx' + cssMemo.size;
    const safe = (v) => String(v).replace(/[^A-Za-z0-9#.\- ]/g, '');
    const rule = [];
    if (st['大きさ']) rule.push('font-size:' + (Number(st['大きさ']) * 1.333).toFixed(1) + 'px');
    if (st['太字'] === '○') rule.push('font-weight:700');
    if (st['斜体'] === '○') rule.push('font-style:italic');
    if (color) rule.push('color:#' + safe(color));
    rule.push('background-color:' + (fill ? '#' + safe(fill) : '#fff'));
    [['罫線（左）', 'left'], ['罫線（右）', 'right'], ['罫線（上）', 'top'], ['罫線（下）', 'bottom']].forEach(([cn, side]) => { if (BORDER[st[cn]]) rule.push('border-' + side + ':' + BORDER[st[cn]]); });
    if (st['横の位置']) rule.push('text-align:' + safe(st['横の位置']));
    rule.push('vertical-align:' + (st['縦の位置'] === 'center' ? 'middle' : (st['縦の位置'] ? safe(st['縦の位置']) : 'bottom')));
    if (st['折り返し'] === '○') rule.push('white-space:normal');
    sheetCss.insertRule('.sheet-grid td.' + name + '{' + rule.join(';') + '}', sheetCss.cssRules.length);
    const out = { name, shrink: st['縮小して全体を表示'] === '○', fontPx: st['大きさ'] ? Number(st['大きさ']) * 1.333 : 14.7, fmt: st['表示形式'] || 'General' };
    cssMemo.set(k, out); return out;
  }
  function textOf(cell, fmt) {
    let v = cell.v;
    if (cell.f !== undefined) v = typeof cell.cached === 'number' ? cell.cached : (cell.cachedText !== undefined ? cell.cachedText : '');
    if (v === undefined || v === null) return '';
    if (typeof v === 'number') { if (fmt && fmt.indexOf('#,##0') === 0) return v.toLocaleString('ja-JP'); if (fmt && fmt.indexOf('%') >= 0) return (v * 100).toFixed(2) + '%'; return String(v); }
    return String(v);
  }
  const visualLen = (s) => { let n = 0; for (const ch of s) n += ch.charCodeAt(0) < 0x2E80 && !/[！-～]/.test(ch) ? 0.55 : 1; return n; };
  function sigOf(cell, text) { return [text, cell.s, cell.fc, cell.fl, cell.runs ? JSON.stringify(cell.runs) : ''].join('|'); }
  function paintCell(box, cell, styles, w, text) {
    const td = box.td; const c = cssOf(styles, cell.s, cell.fc, cell.fl);
    if (box.sx !== c.name) { if (box.sx) td.classList.remove(box.sx); td.classList.add(c.name); box.sx = c.name; }
    let shrunk = '';
    if (c.shrink && text) { const need = visualLen(text) * c.fontPx; if (need > w - 4) shrunk = Math.max(7, c.fontPx * (w - 4) / need).toFixed(1) + 'px'; }
    if (td.style.fontSize !== shrunk) td.style.fontSize = shrunk;
    td.replaceChildren();
    if (cell.runs && cell.runs.length && typeof cell.v === 'string') {
      let at = 0; const t = cell.v;
      cell.runs.slice().sort((a, b) => a[0] - b[0]).forEach((r) => {
        const from = Math.max(at, Math.min(r[0], t.length)), to = Math.max(from, Math.min(r[1], t.length));
        if (from > at) td.append(document.createTextNode(t.slice(at, from)));
        const span = el('span', '', t.slice(from, to)); const p = r[2] || {}; if (p.c) span.style.color = '#' + p.c; if (p.b) span.style.fontWeight = '700'; td.append(span); at = to;
      });
      if (at < t.length) td.append(document.createTextNode(t.slice(at)));
    } else if (text !== '') td.textContent = text;
  }

  // ------------------------------------------------------------ 帯
  function buildFrame() {
    const root = $('#view-sheet'); root.replaceChildren();
    const bar = el('div', 'sheet-bar');
    const tabs = el('div', 'sheet-tabs'); tabs.setAttribute('role', 'tablist'); tabs.setAttribute('aria-label', 'シート');
    TABS.forEach((t) => { const b = el('button', t === view.table ? 'is-active' : '', t); b.type = 'button'; b.dataset.tab = t; b.setAttribute('role', 'tab'); b.setAttribute('aria-selected', String(t === view.table)); b.addEventListener('click', () => openTable(t)); tabs.append(b); });
    const zoom = el('div', 'zoom-group'); zoom.setAttribute('role', 'group'); zoom.setAttribute('aria-label', '大きさ');
    [['fit', '全体'], ['full', '大きく']].forEach(([k, label]) => { const b = el('button', view.zoom === k ? 'is-active' : '', label); b.type = 'button'; b.dataset.zoom = k; b.addEventListener('click', () => { view.zoom = k; zoom.querySelectorAll('button').forEach((x) => x.classList.toggle('is-active', x === b)); applyZoom(); }); zoom.append(b); });
    const mode = el('span', 'mode-label'); mode.id = 'mode-label';
    const move = el('button', 'btn', '名前や番号で移す'); move.type = 'button'; move.id = 'move-open'; move.addEventListener('click', () => { view.movePanel = !view.movePanel; renderMovePanel(); });
    const undo = el('button', 'btn', '元に戻す'); undo.type = 'button'; undo.id = 'move-undo'; undo.addEventListener('click', () => { App.undoLast(); paintBar(); });
    const auto = el('button', 'btn', '自動に戻す'); auto.type = 'button'; auto.id = 'move-auto'; auto.addEventListener('click', () => askAuto(auto));
    if (state.readOnly) bar.append(tabs, zoom); else bar.append(tabs, zoom, mode, move, undo, auto);
    const usage = el('div', 'usage'); usage.id = 'usage'; usage.setAttribute('aria-label', '列ごとの使い方');
    const note = el('p', 'sheet-note', '区分の見出しを掴むとかたまりごと、商品の行を掴むとその1行だけ動きます。Esc で取り消せます。首都圏と北関東はいつも一緒に動きます。');
    const panel = el('div', 'move-panel'); panel.id = 'move-panel'; panel.hidden = true;
    const scroll = el('div', 'sheet-scroll'); scroll.id = 'sheet-scroll';
    if (state.readOnly) root.append(bar, scroll); else root.append(bar, usage, note, panel, scroll);
    scroll.addEventListener('pointerdown', onPointerDown);
    view.built = true;
  }
  function kindNow() { return view.table === '月間商品' ? '' : C.TABLES[view.table].kind; }
  function paintBar() {
    const kind = kindNow(); const mode = $('#mode-label'); if (!mode) return;
    const manual = kind && db().status[kind] === '手直し';
    mode.textContent = !kind ? '月間商品' : (manual ? '手直し（自動では動きません）' : '自動（いつも連動）');
    mode.classList.toggle('is-manual', Boolean(manual));
    $('#move-auto').disabled = !manual; $('#move-undo').disabled = !state.undo.length;
    $('#move-open').disabled = !kind;
  }
  function paintUsage(preview) {
    const box = $('#usage'); if (!box) return;
    box.replaceChildren();
    if (view.table === '月間商品') { box.append(monthlySyncBar()); return; }
    const lay = C.layoutOf(db(), view.table);
    lay.blocks.forEach((b) => {
      const used = preview && preview.used && preview.used[b.no] !== undefined ? preview.used[b.no] : b.used;
      const item = el('div', 'usage-item' + (used > b.cap ? ' is-over' : '') + (preview && preview.target === b.no ? ' is-target' : ''));
      const bar = el('span', 'usage-bar'); const fill = el('i'); fill.style.width = Math.min(100, Math.round(used / b.cap * 100)) + '%'; bar.append(fill);
      item.append(el('span', '', b.no + '列目'), bar, el('span', '', used + '／' + b.cap + (used > b.cap ? '（超えています）' : '')));
      box.append(item);
    });
  }

  /** 月間商品と一覧の販売形態のずれ（数）と「一覧から作り直す」。決まりは core.js の monthlyReconcile の1か所 */
  function monthlySyncBar() {
    const bar = el('div', 'mon-sync');
    if (state.readOnly) { bar.append(el('span', 'hint-line', '月間商品は、その月の月間特売と新商品の表です。')); return bar; }
    const sync = C.monthlySync(db());
    if (!sync.missing.length && !sync.extra.length) { bar.append(el('span', 'hint-line', '月間商品は、一覧の月間特売・新商品と合っています。')); return bar; }
    const msg = el('span'); msg.append(document.createTextNode('一覧の販売形態とずれています: 足りない '), el('b', '', String(sync.missing.length)), document.createTextNode('・余分 '), el('b', '', String(sync.extra.length)));
    const b = el('button', 'btn btn-small', '一覧から作り直す'); b.type = 'button'; b.id = 'monthly-rebuild';
    b.title = '月間特売・新商品にした商品を足し、もうそうでない商品の行を外します。結んでいない行・見出し・サマリーから入れた規格は残します。';
    b.addEventListener('click', () => rebuildMonthly(b));
    bar.append(msg, b);
    return bar;
  }
  async function rebuildMonthly(b) {
    const done = App.busyButton(b, '作り直しています…');
    const t0 = performance.now();
    App.flush();
    const res = C.monthlyReconcile(db());
    try {
      await App.serial(() => App.api('replaceMonthly', { month: state.month, rows: { c: C.MONTHLY_COLS, r: res.rows.map((g) => C.MONTHLY_COLS.map((h) => (g[h] === undefined || g[h] === null ? '' : g[h]))) }, requestId: App.rid() }));
      db().monthly = res.rows; App.refresh([]); render(true); paintUsage();
      App.perfNote('月間商品を作り直す', t0);
      App.toast('月間商品を作り直しました（足した ' + res.added + '・外した ' + res.removed + '）。');
    } catch (error) { App.toast('作り直せませんでした: ' + error.message, true); }
    finally { done(); }
  }

  // ------------------------------------------------------------ 格子を描く（変わったセルだけ描き直す）
  function render(force) {
    const t0 = performance.now();
    const sheet = S.buildOne(db(), view.table, C.rowsBySection(db()));
    view.sheet = sheet;
    const scroll = $('#sheet-scroll');
    const cols = []; const byC = new Map(sheet.cols.map((c) => [c.c, c]));
    for (let c = 1; c <= sheet.lastCol; c++) { const d = byC.get(c); if (!d || !d.hidden) cols.push({ c, w: colPx(d ? d.w : null) }); }
    const hiddenRows = new Set(sheet.hiddenRows || []);
    const defaultH = (sheet.page.defaultRowHeight || 15) * 1.333;
    const rowsList = []; for (let r = 1; r <= sheet.lastRow; r++) if (!hiddenRows.has(r)) rowsList.push(r);
    const colIndex = new Map(cols.map((c, i) => [c.c, i]));
    // つないだセル
    const span = new Map(), covered = new Set();
    sheet.merges.forEach((x) => {
      const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(x); if (!m) return;
      const c1 = S.letterToCol(m[1]), r1 = Number(m[2]), c2 = S.letterToCol(m[3]), r2 = Number(m[4]);
      const vis = cols.filter((c) => c.c >= c1 && c.c <= c2); if (!vis.length) return;
      const rs = rowsList.filter((r) => r >= r1 && r <= r2); if (!rs.length) return;
      span.set(r1 * 2000 + vis[0].c, { cs: vis.length, rs: rs.length });
      rs.forEach((r) => vis.forEach((c) => { if (!(r === r1 && c.c === vis[0].c)) covered.add(r * 2000 + c.c); }));
    });
    const structure = cols.map((c) => c.c + ':' + c.w).join(',') + '|' + rowsList.map((r) => (sheet.rowHeights[r] || '') + (r <= (sheet.meta.hr || 0) ? 'h' : '')).join(',') + '|' + sheet.merges.join(',');
    const headRow = sheet.meta.hr || 0;
    const grabbable = view.table !== '月間商品' && !state.readOnly;
    let table = scroll.querySelector('table');
    let updated = 0;
    if (force || !table || view.merged !== structure) {
      view.merged = structure; view.cells = new Map();
      table = el('table', 'sheet-grid'); table.id = 'sheet-grid';
      const cg = el('colgroup'); const gw = el('col'); gw.style.width = '30px'; cg.append(gw);
      let total = 30; cols.forEach((c) => { const x = el('col'); x.style.width = c.w + 'px'; cg.append(x); total += c.w; }); table.append(cg);
      table.style.width = total + 'px'; view.naturalWidth = total;
      const body = el('tbody');
      let top = 0;
      rowsList.forEach((r) => {
        const tr = el('tr'); const h = sheet.rowHeights[r] !== undefined ? sheet.rowHeights[r] * 1.333 : defaultH; tr.style.height = h.toFixed(1) + 'px';
        const isHead = headRow && r <= headRow;
        const g = el('th', 'gutter', r); g.scope = 'row'; if (isHead) { g.classList.add('is-head'); g.style.top = top + 'px'; } tr.append(g);
        cols.forEach((c) => {
          const key = r * 2000 + c.c; if (covered.has(key)) return;
          const td = el('td'); td.dataset.r = r; td.dataset.c = c.c;
          const sp = span.get(key); if (sp) { if (sp.cs > 1) td.colSpan = sp.cs; if (sp.rs > 1) td.rowSpan = sp.rs; }
          if (isHead) { td.classList.add('is-head'); td.style.top = top + 'px'; }
          tr.append(td);
          view.cells.set(key, { td, sig: null, w: sp ? cols.slice(colIndex.get(c.c), colIndex.get(c.c) + sp.cs).reduce((a, x) => a + x.w, 0) : c.w });
        });
        if (isHead) top += h;
        body.append(tr);
      });
      table.append(body);
      scroll.replaceChildren(table);
    }
    view.cells.forEach((box, key) => {
      const cell = sheet.grid.cells.get(key) || { r: Math.floor(key / 2000), c: key % 2000 };
      const c0 = cssOf(sheet._db.type.styles, cell.s, cell.fc, cell.fl);
      const text = textOf(cell, c0.fmt);
      const sig = sigOf(cell, text);
      if (sig !== box.sig) { paintCell(box, cell, sheet._db.type.styles, box.w, text); box.sig = sig; updated++; }
    });
    markGrab(grabbable);
    applyZoom();
    paintBar(); paintUsage();
    view.dirty = false;
    App.perfNote('見本を描く（' + view.table + '・' + updated + 'セル）', t0);
  }
  function metaOf(r, c) {
    const sheet = view.sheet; if (!sheet || !sheet.meta || !sheet.meta.blocks) return null;
    const b = sheet.meta.blocks.find((x) => c >= x.first && c <= x.last); if (!b) return null;
    const m = b.rows.find((x) => x.r === r);
    return m ? { block: b, row: m } : { block: b, row: null };
  }
  function markGrab(on) {
    const sheet = view.sheet; const blocks = (sheet.meta && sheet.meta.blocks) || [];
    view.cells.forEach((box) => box.td.classList.remove('is-grab'));
    if (!on) return;
    blocks.forEach((b) => b.rows.forEach((m) => { for (let c = b.first; c <= b.last; c++) { const x = view.cells.get(m.r * 2000 + c); if (x) x.td.classList.add('is-grab'); } }));
  }
  function applyZoom() {
    const scroll = $('#sheet-scroll'), table = scroll && scroll.querySelector('table'); if (!table) return;
    const z = view.zoom === 'fit' ? Math.min(1, (scroll.clientWidth - 2) / (view.naturalWidth || 1)) : 1;
    table.style.zoom = z.toFixed(3);
  }
  window.addEventListener('resize', () => { if (state.view === 'sheet') applyZoom(); });

  function openTable(t) {
    view.table = t;
    document.querySelectorAll('.sheet-tabs button').forEach((b) => { const on = b.dataset.tab === t; b.classList.toggle('is-active', on); b.setAttribute('aria-selected', String(on)); });
    if (view.built && state.view === 'sheet') { render(); $('#sheet-scroll').scrollTop = 0; renderMovePanel(); }
    else view.dirty = true;
  }

  // ------------------------------------------------------------ ドラッグ（行き先が光る・数字が先に変わる・超えると赤・Esc で取り消し・端で流れる）
  function hitCell(x, y) { const n = document.elementFromPoint(x, y); const td = n && n.closest && n.closest('td[data-r]'); if (!td) return null; return { r: Number(td.dataset.r), c: Number(td.dataset.c), td }; }
  function onPointerDown(e) {
    if (state.readOnly || e.button !== 0 || e.pointerType === 'touch' || view.table === '月間商品') return;
    const hit = hitCell(e.clientX, e.clientY); if (!hit) return;
    const m = metaOf(hit.r, hit.c); if (!m || !m.row) return;
    const item = m.row.kind === 'section' ? { kind: 'sec', id: m.row.secId, from: m.block.no } : { kind: 'row', id: m.row.rowId, secId: m.row.secId, from: m.block.no };
    view.drag = { item, x0: e.clientX, y0: e.clientY, active: false, target: null, label: labelOf(item) };
    document.addEventListener('pointermove', onPointerMove);
    document.addEventListener('pointerup', onPointerUp);
    document.addEventListener('keydown', onDragKey);
    e.preventDefault();
  }
  function labelOf(item) {
    if (item.kind === 'sec') { const s = db().sections.get(String(item.id)); return (s ? C.nfkc(s['表題']) : '') + ' のかたまり'; }
    const r = db().rows.get(String(item.id)); const p = r && r['商品ID'] ? db().products.get(String(r['商品ID'])) : null;
    return p ? C.nfkc(p['名前']).trim() : (r && r['行の種類'] === C.KIND_BLANK ? '空き行' : '続きの行');
  }
  function endDrag() {
    document.removeEventListener('pointermove', onPointerMove); document.removeEventListener('pointerup', onPointerUp); document.removeEventListener('keydown', onDragKey);
    window.clearInterval(view.autoScroll); view.autoScroll = 0;
    $('#drag-ghost').hidden = true;
    const line = document.querySelector('.drop-line'); if (line) line.remove();
    document.querySelectorAll('.sheet-grid td.is-source').forEach((td) => td.classList.remove('is-source'));
    view.drag = null;
    if (view.built) paintUsage();
  }
  function onDragKey(e) { if (e.key === 'Escape') { endDrag(); toast('動かすのをやめました。'); } }
  function markSource(item) {
    const blocks = view.sheet.meta.blocks;
    blocks.forEach((b) => b.rows.forEach((m) => {
      const hit = item.kind === 'sec' ? m.secId === String(item.id) : m.rowId === String(item.id);
      if (hit) for (let c = b.first; c <= b.last; c++) { const x = view.cells.get(m.r * 2000 + c); if (x) x.td.classList.add('is-source'); }
    }));
  }
  /** 指している所 → 行き先（どの列のどこの前か）。線を出す場所も返す */
  function targetAt(x, y) {
    const d = view.drag; const hit = hitCell(x, y); if (!hit) return null;
    const b = view.sheet.meta.blocks.find((q) => hit.c >= q.first && hit.c <= q.last); if (!b) return null;
    const rect = hit.td.getBoundingClientRect(); const upper = y < rect.top + rect.height / 2;
    const m = b.rows.find((q) => q.r === hit.r);
    if (d.item.kind === 'sec') {
      const order = []; b.rows.forEach((q) => { if (q.kind === 'section' && order.indexOf(q.secId) < 0) order.push(q.secId); });
      if (!m) return { toBlock: b.no, before: null, lineRow: lastRowOf(b) + 1, block: b };
      const i = order.indexOf(m.secId);
      const first = firstRowOf(b, m.secId);
      const onFirst = m.r === first || m.r === Math.min(...b.rows.filter((q) => q.secId === m.secId).map((q) => q.r));
      if (onFirst && upper) return { toBlock: b.no, before: i === 0 && m.r === b.rows[0].r ? order[0] : m.secId, lineRow: first, block: b };
      const next = order[i + 1] || null;
      return { toBlock: b.no, before: next, lineRow: next ? firstRowOf(b, next) : lastRowOf(b) + 1, block: b };
    }
    // 行
    if (!m) { const last = b.rows[b.rows.length - 1]; if (!last) return null; return { toSec: last.secId, before: null, lineRow: lastRowOf(b) + 1, block: b }; }
    if (m.kind === 'section') { const rows = b.rows.filter((q) => q.secId === m.secId && q.kind !== 'section'); return { toSec: m.secId, before: rows[0] ? rows[0].rowId : null, lineRow: rows[0] ? rows[0].r : m.r + 1, block: b }; }
    if (upper) return { toSec: m.secId, before: m.rowId, lineRow: m.r, block: b };
    const rows = b.rows.filter((q) => q.secId === m.secId && q.kind !== 'section');
    const i = rows.findIndex((q) => q.rowId === m.rowId); const next = rows[i + 1];
    return { toSec: m.secId, before: next ? next.rowId : null, lineRow: next ? next.r : m.r + 1, block: b };
  }
  const firstRowOf = (b, secId) => Math.min(...b.rows.filter((q) => q.secId === secId).map((q) => q.r));
  const lastRowOf = (b) => (b.rows.length ? Math.max(...b.rows.map((q) => q.r)) : b.start - 1);
  function drawLine(t) {
    const scroll = $('#sheet-scroll'); let line = scroll.querySelector('.drop-line');
    if (!line) { line = el('div', 'drop-line'); scroll.append(line); }
    const b = t.block; const sr = scroll.getBoundingClientRect();
    const refCell = (r) => view.cells.get(r * 2000 + b.first) || null;
    let top; const cell = refCell(t.lineRow);
    if (cell) top = cell.td.getBoundingClientRect().top; else { const last = refCell(t.lineRow - 1) || refCell(lastRowOf(b)); top = last ? last.td.getBoundingClientRect().bottom : sr.top; }
    const c1 = view.cells.get((cell ? t.lineRow : Math.max(lastRowOf(b), b.start)) * 2000 + b.first), c2 = view.cells.get((cell ? t.lineRow : Math.max(lastRowOf(b), b.start)) * 2000 + b.last);
    const left = c1 ? c1.td.getBoundingClientRect().left : sr.left; let right = c2 ? c2.td.getBoundingClientRect().right : sr.right;
    // 隠し列は描かないので、見えている一番右を探す
    if (!c2) { for (let c = b.last; c >= b.first; c--) { const x = view.cells.get((cell ? t.lineRow : Math.max(lastRowOf(b), b.start)) * 2000 + c); if (x) { right = x.td.getBoundingClientRect().right; break; } } }
    line.style.left = (left - sr.left + scroll.scrollLeft) + 'px'; line.style.width = Math.max(30, right - left) + 'px'; line.style.top = (top - sr.top + scroll.scrollTop - 2) + 'px';
  }
  /** 動かしたときの列ごとの行数（手元で当てて、数えて、戻す） */
  function simulate(t) {
    const d = view.drag; const database = db(); const kind = kindNow();
    let changes = [];
    try { changes = d.item.kind === 'sec' ? C.moveSection(database, { id: d.item.id, toBlock: t.toBlock, before: t.before }) : C.moveRow(database, { id: d.item.id, toSec: t.toSec, before: t.before }); } catch { return null; }
    const before = C.beforeOf(database, changes);
    C.applyChanges(database, changes);
    const lay = C.layoutOf(database, view.table);
    const used = {}; lay.blocks.forEach((b) => { used[b.no] = b.used; });
    C.applyChanges(database, before);
    return { used, changed: changes.length > 0 };
  }
  function onPointerMove(e) {
    const d = view.drag; if (!d) return;
    if (!d.active) { if (Math.abs(e.clientX - d.x0) + Math.abs(e.clientY - d.y0) < 6) return; d.active = true; markSource(d.item); $('#drag-ghost').hidden = false; $('#drag-ghost').textContent = d.label; startAutoScroll(); }
    d.px = e.clientX; d.py = e.clientY;
    const g = $('#drag-ghost'); g.style.left = (e.clientX + 14) + 'px'; g.style.top = (e.clientY + 12) + 'px';
    const t = targetAt(e.clientX, e.clientY);
    const key = t ? JSON.stringify([t.toBlock, t.toSec, t.before, t.lineRow, t.block.no]) : '';
    if (key === d.key) return; d.key = key; d.target = t;
    if (!t) { const line = document.querySelector('.drop-line'); if (line) line.remove(); paintUsage(); g.classList.remove('is-over'); return; }
    drawLine(t);
    const sim = simulate(t);
    if (sim) { paintUsage({ used: sim.used, target: t.block.no }); const over = sim.used[t.block.no] > t.block.cap; g.classList.toggle('is-over', over); g.textContent = d.label + (over ? '（上限を超えます。動かせます）' : ''); }
  }
  function startAutoScroll() {
    window.clearInterval(view.autoScroll);
    view.autoScroll = window.setInterval(() => {
      const d = view.drag; if (!d || d.px === undefined) return;
      const s = $('#sheet-scroll'), r = s.getBoundingClientRect(); const edge = 36, step = 18;
      if (d.py < r.top + edge) s.scrollTop -= step; else if (d.py > r.bottom - edge) s.scrollTop += step;
      if (d.px < r.left + edge) s.scrollLeft -= step; else if (d.px > r.right - edge) s.scrollLeft += step;
    }, 40);
  }
  function onPointerUp(e) {
    const d = view.drag; if (!d) return;
    const active = d.active; const t = active ? targetAt(e.clientX, e.clientY) : null; const item = d.item;
    endDrag();
    if (!active) { selectForMove(item); return; }
    if (!t) { toast('行き先が表の外だったので、動かしませんでした。'); return; }
    const t0 = performance.now();
    const op = item.kind === 'sec' ? { kind: 'sec', id: String(item.id), toBlock: t.toBlock, before: t.before || null } : { kind: 'row', id: String(item.id), toSec: String(t.toSec), before: t.before || null };
    if (App.doMove(op, item.kind === 'sec' ? '区分を動かす' : '行を動かす')) { render(); App.perfNote('ドラッグ→見本が変わる', t0); flashMoved(op); }
  }
  function flashMoved(op) {
    const blocks = view.sheet.meta.blocks;
    blocks.forEach((b) => b.rows.forEach((m) => {
      const hit = op.kind === 'sec' ? m.secId === op.id : m.rowId === op.id;
      if (hit) for (let c = b.first; c <= b.last; c++) { const x = view.cells.get(m.r * 2000 + c); if (x) { x.td.classList.add('is-found'); window.setTimeout(() => x.td.classList.remove('is-found'), 1400); } }
    }));
  }
  // 押しただけ（動かさなかった）: 「ここに移す」の何をへ入れる
  function selectForMove(item) { view.pick = item; if (view.movePanel) renderMovePanel(); }

  // ------------------------------------------------------------ 名前や番号で移す（同じ処理の別の入り口）
  function candidates(q) {
    const kind = kindNow(); if (!kind) return [];
    const out = [];
    const terms = q.trim() ? (SN ? SN.terms(q) : C.nfkc(q).toLowerCase().split(/\s+/).filter(Boolean)) : null;
    const hit = (t) => { if (!terms) return true; const k = SN ? SN.key(t) : C.nfkc(t).toLowerCase(); return SN ? SN.hit(k, terms) : terms.every((x) => k.includes(x)); };
    C.sectionsOfKind(db(), kind).forEach((s) => { if (hit(String(s['表題']))) out.push({ kind: 'sec', id: String(s['区分ID']), label: '［区分］' + C.nfkc(s['表題']) }); });
    db().rows.forEach((r) => { const p = r['商品ID'] ? db().products.get(String(r['商品ID'])) : null; if (!p || p['表の種類'] !== kind) return; if (hit(String(p['名前']) + ' ' + String(p['呼出番'] ?? ''))) out.push({ kind: 'row', id: String(r['行ID']), label: (p['呼出番'] ? p['呼出番'] + ' ' : '') + C.nfkc(p['名前']).trim() }); });
    return out.slice(0, terms ? 24 : 12);
  }
  function renderMovePanel() {
    const panel = $('#move-panel'); if (!panel) return;
    if (!view.movePanel || view.table === '月間商品') { panel.hidden = true; return; }
    panel.hidden = false; panel.replaceChildren(el('h2', '', '名前や番号で選んで、ここに移す'));
    const row = el('div', 'move-row');
    const input = el('input'); input.type = 'search'; input.id = 'move-search'; input.placeholder = '何を: 名前・番号で探す'; input.value = view.moveQ || ''; input.setAttribute('aria-label', '動かす物を探す');
    row.append(input); panel.append(row);
    const list = el('div', 'move-list'); list.id = 'move-list'; panel.append(list);
    const dest = el('div', 'move-row'); dest.id = 'move-dest'; panel.append(dest);
    const paintList = () => {
      list.replaceChildren();
      candidates(input.value).forEach((c) => { const b = el('button', view.pick && view.pick.kind === c.kind && String(view.pick.id) === c.id ? 'is-picked' : '', c.label); b.type = 'button'; b.title = c.label; b.addEventListener('click', () => { view.pick = c; paintList(); paintDest(); }); list.append(b); });
    };
    const paintDest = () => {
      dest.replaceChildren();
      const pick = view.pick; const kind = kindNow(); if (!pick) { dest.append(el('span', 'hint-line', '上から動かす物を選んでください（見本の行を押しても選べます）。')); return; }
      if (pick.kind === 'sec') {
        const blockSel = el('select'); blockSel.id = 'move-to-block'; C.blockList(db(), view.table).forEach((b) => { const o = el('option', '', b['まとまり'] + '列目'); o.value = b['まとまり']; blockSel.append(o); });
        const place = db().places.get(String(pick.id)); if (place) blockSel.value = place['まとまり'];
        const beforeSel = el('select'); beforeSel.id = 'move-to-before';
        const fill = () => { beforeSel.replaceChildren(); const o0 = el('option', '', 'この列の末尾へ'); o0.value = ''; beforeSel.append(o0); C.placesOfBlock(db(), kind, blockSel.value).forEach((p) => { if (String(p['区分ID']) === String(pick.id)) return; const s = db().sections.get(String(p['区分ID'])); const o = el('option', '', C.nfkc(s['表題']) + ' の前へ'); o.value = String(p['区分ID']); beforeSel.append(o); }); };
        blockSel.addEventListener('change', fill); fill();
        const l1 = el('label'); l1.append(document.createTextNode('どの列へ'), blockSel); const l2 = el('label'); l2.append(document.createTextNode('どこへ'), beforeSel);
        const go = el('button', 'btn btn-primary', 'ここに移す'); go.type = 'button'; go.id = 'move-go'; go.addEventListener('click', () => runMove({ kind: 'sec', id: String(pick.id), toBlock: Number(blockSel.value), before: beforeSel.value || null }));
        dest.append(l1, l2, go);
      } else {
        const secSel = el('select'); secSel.id = 'move-to-sec'; C.sectionsOfKind(db(), kind).forEach((s) => { const o = el('option', '', C.nfkc(s['表題'])); o.value = String(s['区分ID']); secSel.append(o); });
        const r = db().rows.get(String(pick.id)); if (r) secSel.value = String(r['区分ID']);
        const beforeSel = el('select'); beforeSel.id = 'move-to-before';
        const fill = () => { beforeSel.replaceChildren(); const o0 = el('option', '', 'この区分の末尾へ'); o0.value = ''; beforeSel.append(o0); C.rowsOfSection(db(), secSel.value).forEach((x) => { if (String(x['行ID']) === String(pick.id)) return; const p = x['商品ID'] ? db().products.get(String(x['商品ID'])) : null; if (!p) return; const o = el('option', '', C.nfkc(p['名前']).trim() + ' の前へ'); o.value = String(x['行ID']); beforeSel.append(o); }); };
        secSel.addEventListener('change', fill); fill();
        const l1 = el('label'); l1.append(document.createTextNode('どの区分へ'), secSel); const l2 = el('label'); l2.append(document.createTextNode('どこへ'), beforeSel);
        const go = el('button', 'btn btn-primary', 'ここに移す'); go.type = 'button'; go.id = 'move-go'; go.addEventListener('click', () => runMove({ kind: 'row', id: String(pick.id), toSec: secSel.value, before: beforeSel.value || null }));
        dest.append(l1, l2, go);
      }
    };
    let timer = 0;
    input.addEventListener('input', () => { window.clearTimeout(timer); view.moveQ = input.value; timer = window.setTimeout(paintList, 120); });
    paintList(); paintDest();
  }
  function runMove(op) {
    const t0 = performance.now();
    if (App.doMove(op, op.kind === 'sec' ? '区分を移す' : '行を移す')) { render(); flashMoved(op); App.perfNote('ここに移す→見本が変わる', t0); toast('移しました。'); renderMovePanel(); }
    else toast('もうその場所にあります。');
  }
  async function askAuto(button) {
    const kind = kindNow(); if (!kind) return;
    if (button.dataset.sure !== '1') { button.dataset.sure = '1'; button.dataset.idleLabel = '自動に戻す'; button.textContent = 'もう一度押すと、手直しを捨てて決め直します'; window.setTimeout(() => { if (button.dataset.sure === '1') { button.dataset.sure = ''; button.textContent = '自動に戻す'; delete button.dataset.idleLabel; } }, 5000); return; }
    button.dataset.sure = ''; delete button.dataset.idleLabel; button.textContent = '自動に戻す';
    const done = App.busyButton(button, '決め直しています…');
    App.flush();
    const t0 = performance.now();
    try {
      const data = await App.serial(() => App.api('setPlacementMode', { month: state.month, kind, requestId: App.rid() }));
      C.resetToAuto(db(), kind); // 手元でも同じ部品で（窓口の返事と同じ結果）
      if (data.changes && data.changes.length) C.applyChanges(db(), data.changes);
      db().status[kind] = '自動'; state.undo = [];
      App.refresh([]); App.perfNote('自動に戻す', t0);
      if (state.view === 'sheet') render();
      toast('元の形から決め直しました。');
    } catch (error) { toast('決め直せませんでした: ' + error.message, true); }
    finally { done(); paintBar(); }
  }

  App.views.sheet = {
    show() {
      if (!state.db) return;
      if (!view.built) { buildFrame(); view.dirty = true; }
      renderMovePanel();
      if (!view.dirty) { applyZoom(); paintBar(); paintUsage(); return; }
      const scroll = $('#sheet-scroll');
      if (scroll.querySelector('table')) { render(); return; }
      // はじめて描くときは、先に印を出してから描く（待ちをフリーズに見せない）
      const wait = el('p', 'sheet-note'); wait.append(App.spinner(), document.createTextNode(' 番号表を描いています…')); scroll.replaceChildren(wait);
      paintBar(); paintUsage();
      window.requestAnimationFrame(() => window.setTimeout(() => { if (state.db && state.view === 'sheet') render(); else view.dirty = true; }, 0));
    },
    reset() { view.built = false; view.dirty = true; view.merged = ''; view.cells = new Map(); view.pick = null; view.movePanel = false; cssMemo.clear(); sheetCss.replaceSync(''); if (view.drag) endDrag(); },
    markDirty() { view.dirty = true; if (state.view === 'sheet' && view.built && !view.drag) render(); },
    openTable
  };
})();
