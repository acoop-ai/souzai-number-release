// セルの格子を作る部品（画面の部品を持たない）。見本の絵と Excel は、この格子から作る（2つに分かれない）。
//   段1 の tools/generate_tables.py の決まりを移した物。列のまとまりの詰め方は「月の置き場」から決める。
//   原価（db.costs）が無いときは、隠し列を空で作る（見本の絵は隠し列を出さないので、そのままでよい）。Excel は窓口から原価を受け取ってから作る。
(function (root) {
  'use strict';
  const C = root.SouzaiCore;
  const isBlank = C.isBlank;
  function letterToCol(s) { let n = 0; for (let i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
  function colToLetter(c) { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - m - 1) / 26); } return s; }
  function parseRef(ref) { const m = /^([A-Z]+)(\d+)$/.exec(String(ref)); return { c: letterToCol(m[1]), r: Number(m[2]) }; }
  function parseOverrides(text) { const out = {}; String(text || '').replace(/\//g, '／').split('／').forEach((part) => { const i = part.indexOf(':'); if (i > 0) out[part.slice(0, i).trim()] = part.slice(i + 1).trim(); }); return out; }
  const isArith = (v) => typeof v === 'string' && v[0] === '=' && /^[0-9.+\-*/() ]+$/.test(v.slice(1));
  const isPinkFont = (rgb) => rgb === 'FF00FF' || rgb === 'FF3399' || rgb === 'FF33CC';
  function isNewFill(rgb) { if (!rgb || rgb === 'FFFFFF') return false; const r = parseInt(rgb.slice(0, 2), 16), g = parseInt(rgb.slice(2, 4), 16), b = parseInt(rgb.slice(4, 6), 16); return !(r === g && g === b); }

  class Grid {
    constructor() { this.cells = new Map(); }
    at(r, c) { const k = r * 2000 + c; let o = this.cells.get(k); if (!o) { o = { r, c }; this.cells.set(k, o); } return o; }
    get(r, c) { return this.cells.get(r * 2000 + c); }
  }
  /** 値を入れる（空は入れない）。式は決まった形だけ式にする（ほかは文字。= で始まる見出しが式に化けない） */
  function put(cell, v, formulaOk) {
    if (v === null || v === undefined || v === '') return;
    if (typeof v === 'string' && v[0] === '=' && formulaOk !== false && (v.indexOf('=+((') === 0 || isArith(v) || v.indexOf('=IF(') === 0)) { cell.f = v.slice(1); return; }
    cell.v = v;
  }
  function style(cell, key, fc, fl) { cell.s = key || null; if (fc !== undefined) cell.fc = fc; if (fl !== undefined) cell.fl = fl; }

  function pageOf(sh, lastRow) {
    return { printArea: 'A1:' + sh['印刷の右の列'] + lastRow, orientation: sh['向き'] || 'landscape', paper: sh['用紙'] ? Number(sh['用紙']) : null, scale: sh['縮小（%）'] ? Number(sh['縮小（%）']) : null,
      margins: { left: sh['余白（左）'], right: sh['余白（右）'], top: sh['余白（上）'], bottom: sh['余白（下）'], header: sh['余白（ヘッダー）'], footer: sh['余白（フッター）'] },
      hCenter: sh['横の中央'] === '○', vCenter: sh['縦の中央'] === '○', zoom: sh['表示倍率'] ? Number(sh['表示倍率']) : null, defaultRowHeight: sh['既定の行の高さ'] ? Number(sh['既定の行の高さ']) : null,
      header: { l: sh['ヘッダー（左）'], c: sh['ヘッダー（中）'], r: sh['ヘッダー（右）'] }, footer: { l: sh['フッター（左）'], c: sh['フッター（中）'], r: sh['フッター（右）'] } };
  }
  function parseAnchor(text) {
    const one = (s) => { const p = s.split('+'); const ref = parseRef(p[0]); return { col: ref.c - 1, colOff: Number(p[1]), row: ref.r - 1, rowOff: Number(p[2]) }; };
    const a = String(text).split(':');
    return { from: one(a[0]), to: one(a[1]) };
  }
  function applyFixed(sheet, rows, shift) {
    const type = sheet._db.type;
    rows.forEach((f) => {
      const kind = f['種類'];
      if (kind === '絵') { const data = type.pictures[f['値']]; sheet.pictures.push({ name: String(f['値']), data: data || '', anchor: parseAnchor(f['セル']) }); return; }
      if (kind === '行の高さ') { if (!isBlank(f['値'])) sheet.rowHeights[Number(f['セル']) + shift] = Number(f['値']); return; }
      if (kind === 'セル') {
        const ref = parseRef(f['セル']);
        const cell = sheet.grid.at(ref.r + shift, ref.c);
        let v = f['値'];
        if (typeof v === 'string' && v.indexOf('{月}') >= 0) v = v.replace(/\{月\}/g, String(sheet._monthNum));
        put(cell, v);
        style(cell, f['書式キー']);
        if (f['字の飾り']) cell.runs = C.parseDeco(f['字の飾り']);
        return;
      }
      if (kind === 'つなぐ') {
        const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(String(f['セル']));
        if (m) sheet.merges.push(m[1] + (Number(m[2]) + shift) + ':' + m[3] + (Number(m[4]) + shift));
      }
    });
  }

  function newSheet(db, table, monthNum) {
    const sh = db.type.shapes[table];
    const cols = (db.type.columns[table] || []).map((c) => ({ c: letterToCol(c['列']), w: isBlank(c['幅']) ? null : Number(c['幅']), hidden: c['隠し'] === '○', raw: c }));
    return { key: table, name: sh['シート名'], tabColor: sh['タブの色'] || '', cols, rowHeights: {}, hiddenRows: [], grid: new Grid(), merges: [], pictures: [], page: null, lastRow: 0, lastCol: letterToCol(sh['印刷の右の列']), meta: { blocks: [], hr: Number(sh['見出しの行']) }, _db: db, _monthNum: monthNum };
  }

  function costOf(db, id) { return db.costs ? (db.costs.get(String(id)) || null) : null; }
  function writeProduct(sheet, r, roles, p, row, table) {
    const db = sheet._db;
    const cost = costOf(db, p['商品ID']);
    const col = (k) => (roles[k] ? letterToCol(roles[k]['列']) : 0);
    const deco = (table === '北関東' && p._deco && p._deco['北関東']) ? p._deco['北関東'] : (p._deco || {});
    const vals = { '名前': p['名前'], '指定トレー': p['指定トレー'], '呼出番': C.fullDigits(p['呼出番']), '見える価格': C.shownPrice(p, row, table), '小パック': p['小パック'] };
    if (cost) { vals['原価'] = cost['原価']; vals['粗利（包材値上前）'] = cost['粗利（包材値上前）']; vals['定番価格（包材値上前）'] = cost['定番価格（包材値上前）']; vals['定番価格（包材変更後）'] = cost['定番価格（包材変更後）']; }
    Object.keys(vals).forEach((k) => {
      if (!col(k)) return;
      const cell = sheet.grid.at(r, col(k));
      put(cell, vals[k], k === '原価');
      if (deco[k] && typeof cell.v === 'string') cell.runs = deco[k];
    });
    if (col('粗利（包材変更後）') && cost) {
      const L = roles['定番価格（包材変更後）']['列'], D = roles['原価']['列'];
      sheet.grid.at(r, col('粗利（包材変更後）')).f = '+((' + L + r + '-' + D + r + ')/' + L + r + ')';
      const v = C.grossAfter(cost['原価'], cost['定番価格（包材変更後）']);
      if (v !== null) sheet.grid.at(r, col('粗利（包材変更後）')).cached = v;
    }
  }
  function writeCont(sheet, r, roles, row) {
    const db = sheet._db, cont = row._cont || {}, cost = costOf(db, row['行ID']);
    const col = (k) => (roles[k] ? letterToCol(roles[k]['列']) : 0);
    const vals = { '呼出番': cont['呼出番の欄'], '見える価格': cont['見える価格'], '小パック': cont['小パック'] };
    if (cost) { vals['原価'] = cost['原価']; vals['粗利（包材値上前）'] = cost['粗利（包材値上前）']; vals['定番価格（包材値上前）'] = cost['定番価格（包材値上前）']; vals['定番価格（包材変更後）'] = cost['定番価格（包材変更後）']; }
    Object.keys(vals).forEach((k) => { if (col(k)) { const cell = sheet.grid.at(r, col(k)); put(cell, vals[k], k === '原価'); const deco = row._deco || {}; if (deco[k] && typeof cell.v === 'string') cell.runs = deco[k]; } });
    if (col('粗利（包材変更後）') && cost && !isBlank(cost['定番価格（包材変更後）'])) {
      const L = roles['定番価格（包材変更後）']['列'], D = roles['原価']['列'];
      sheet.grid.at(r, col('粗利（包材変更後）')).f = '+((' + L + r + '-' + D + r + ')/' + L + r + ')';
      const v = C.grossAfter(cost['原価'], cost['定番価格（包材変更後）']);
      if (v !== null) sheet.grid.at(r, col('粗利（包材変更後）')).cached = v;
    }
  }

  // ------------------------------------------------------------ 番号表（首都圏・北関東・焼き物）
  function numberSheet(db, table, monthNum, memo) {
    const sheet = newSheet(db, table, monthNum);
    const sh = db.type.shapes[table];
    const styles = db.type.styles;
    const hr = Number(sh['見出しの行']);
    const fixedRows = db.type.fixed[table] || [];
    applyFixed(sheet, fixedRows, 0);
    String(sh['隠す行'] || '').split(',').forEach((r) => { if (String(r).trim()) sheet.hiddenRows.push(Number(r)); });
    const colsByNo = {};
    sheet.cols.forEach((c) => { colsByNo[c.c] = c.raw; });
    const lay = C.layoutOf(db, table, memo);
    const saleFont = sh['月間特売の字の色'] || 'FF00FF';
    const newFill = sh['新商品の塗り'] || '';
    let lastRow = Number(sh['本文の最後の行']);
    lay.blocks.forEach((b) => {
      const roles = C.blockRoles(db, table, b.no);
      const nameCol = letterToCol(roles['名前']['列']);
      const first = letterToCol(b.first), last = letterToCol(b.last);
      const start = b.row0;
      const capEnd = start + b.cap - 1;
      let r = start;
      const kinds = {};
      const metaRows = [];
      const typeBlock = (db.type.blocks[table] || []).find((x) => Number(x['まとまり']) === b.no) || {};
      let headDone = false;
      b.sections.forEach((s) => {
        const sec = s.sec; const secId = String(sec['区分ID']);
        const title = C.secValue(sec, '表題', table);
        const deco = (table === '北関東' && sec._diff && sec._diff['字の飾り']) || sec._deco || {};
        const ownPlace = sec['表題の置き場（元）'];
        if (s.place === 'head') {
          headDone = true;
          const cell = sheet.grid.at(hr, nameCol);
          put(cell, title);
          if (deco['名前'] && typeof cell.v === 'string') cell.runs = deco['名前'];
          const sk = ownPlace === '見出しの行' ? C.secValue(sec, '表題の書式', table) : (typeBlock['表題の書式（見出しの行）'] || C.secValue(sec, '表題の書式', table));
          style(cell, sk);
          metaRows.push({ r: hr, kind: 'section', secId, head: true });
        } else {
          const cell = sheet.grid.at(r, nameCol);
          put(cell, title);
          if (deco['名前'] && typeof cell.v === 'string') cell.runs = deco['名前'];
          const callCell = sheet.grid.at(r, letterToCol(roles['呼出番']['列']));
          put(callCell, C.secValue(sec, '表題の続き', table));
          if (deco['呼出番'] && typeof callCell.v === 'string') callCell.runs = deco['呼出番'];
          kinds[r] = { k: 'section', sec, ownPlace, typeBlock };
          metaRows.push({ r, kind: 'section', secId, head: false });
          r += 1;
        }
        s.rows.forEach((row) => {
          const kind = row['行の種類'];
          if (kind === C.KIND_PRODUCT) {
            const p = db.products.get(String(row['商品ID']));
            if (p) writeProduct(sheet, r, roles, p, row, table);
            kinds[r] = { k: 'product', row };
          } else if (kind === C.KIND_CONT) {
            writeCont(sheet, r, roles, row);
            kinds[r] = { k: 'product', row };
          } else {
            kinds[r] = { k: 'blank', row };
          }
          if (row['つなぐ'] === '○' && r > start) sheet.merges.push(roles['名前']['列'] + (r - 1) + ':' + roles['名前']['列'] + r);
          metaRows.push({ r, kind: kind === C.KIND_PRODUCT ? 'product' : (kind === C.KIND_CONT ? 'cont' : 'blank'), secId, rowId: String(row['行ID']), pid: row['商品ID'] ? String(row['商品ID']) : '' });
          r += 1;
        });
      });
      if (!headDone) { const cell = sheet.grid.at(hr, nameCol); style(cell, typeBlock['表題の書式（見出しの行）'] || null); }
      sheet.meta.blocks.push({ no: b.no, first, last, nameCol, roles, used: b.used, cap: b.cap, over: b.over, rows: metaRows, start });
      lastRow = Math.max(lastRow, r - 1);
      // 書式（まとまりの全部の列・本文の行）
      const colRole = {}; Object.keys(roles).forEach((k) => { colRole[letterToCol(roles[k]['列'])] = k; });
      for (let rr = start; rr <= Math.max(capEnd, r - 1); rr++) {
        const info = kinds[rr] || { k: 'blank' };
        let over = {};
        if (info.k === 'section') over = parseOverrides(C.secValue(info.sec, '見た目', table));
        else if (info.row) over = parseOverrides(C.rowValue(info.row, '見た目', table));
        for (let c = first; c <= last; c++) {
          const col = colsByNo[c]; if (!col) continue;
          const key = { product: '書式（商品）', section: '書式（区分の表題）', blank: '書式（空き）' }[info.k];
          let sk = col[key];
          const role = colRole[c] || '';
          let fc, fl;
          const isName = c === nameCol;
          if (isName && info.k === 'section') {
            sk = info.ownPlace === '表の中' ? (C.secValue(info.sec, '表題の書式', table) || sk) : (info.typeBlock['表題の書式（表の中）'] || sk);
          } else {
            if (Object.prototype.hasOwnProperty.call(over, role)) sk = over[role];
            if (isName && info.k === 'product' && info.row && info.row['行の種類'] === C.KIND_PRODUCT) {
              const st = styles[sk] || {};
              const form = C.formOf(info.row, table);
              const sale = C.isSaleForm(form);
              const isNew = C.isNewForm(form);
              if (sale && !isPinkFont(st['字の色'])) fc = saleFont;
              if (!sale && isPinkFont(st['字の色'])) fc = '';
              if (isNew && newFill && !isNewFill(st['塗り'])) fl = newFill;
              if (!isNew && isNewFill(st['塗り'])) fl = '';
            }
          }
          style(sheet.grid.at(rr, c), sk, fc, fl);
        }
      }
    });
    // まとまりの外の列（印刷範囲の右の端など）
    const inBlock = new Set();
    lay.blocks.forEach((b) => { for (let c = letterToCol(b.first); c <= letterToCol(b.last); c++) inBlock.add(c); });
    sheet.cols.forEach((col) => {
      if (inBlock.has(col.c)) return;
      for (let rr = hr + 1; rr <= lastRow; rr++) style(sheet.grid.at(rr, col.c), col.raw['書式（空き）']);
    });
    sheet.lastRow = lastRow;
    sheet.page = pageOf(sh, lastRow);
    sheet.layout = lay;
    return sheet;
  }

  // ------------------------------------------------------------ 月間商品
  function monthlySheet(db, monthNum) {
    const table = '月間商品';
    const sheet = newSheet(db, table, monthNum);
    const sh = db.type.shapes[table];
    const roleCol = {}; sheet.cols.forEach((c) => { if (c.raw['役割']) roleCol[c.raw['役割']] = c.c; });
    const parts = (db.type.blocks[table] || []).slice().sort((a, b) => Number(a['題の行']) - Number(b['題の行']));
    let shift = 0;
    let lastRow = Number(sh['本文の最後の行']);
    const metaRows = [];
    parts.forEach((p, i) => {
      applyFixed(sheet, (db.type.fixed[table] || []).filter((f) => f['まとまり'] === p['まとまり']), shift);
      const start = Number(p['本文の最初の行']) + shift;
      const cap = Number(p['上限の行数']);
      const rows = db.monthly.filter((g) => g['部'] === p['まとまり']).sort((a, b) => Number(a['順']) - Number(b['順']));
      const end = start + Math.max(cap, rows.length) - 1;
      const kinds = {};
      rows.forEach((g, k) => {
        const r = start + k;
        const kind = g['行の種類'];
        put(sheet.grid.at(r, roleCol['商品名']), g['商品名']);
        if (kind !== C.KIND_SECTION) {
          put(sheet.grid.at(r, roleCol['呼出番']), C.fullDigits(g['呼出番']));
          put(sheet.grid.at(r, roleCol['規格']), g['規格']);
          put(sheet.grid.at(r, roleCol['本体価格']), g['本体価格']);
          put(sheet.grid.at(r, roleCol['小パックの規格']), g['小パックの規格']);
          put(sheet.grid.at(r, roleCol['小パックの本体価格']), g['小パックの本体価格']);
        }
        kinds[r] = kind === C.KIND_SECTION ? 'section' : (kind === C.KIND_BLANK ? 'blank' : 'product');
        metaRows.push({ r, kind: kinds[r], monthly: String(g['行ID']) });
      });
      const L = colToLetter;
      for (let r = start; r <= end; r++) {
        [['本体価格', '税込価格'], ['小パックの本体価格', '小パックの税込価格']].forEach(([price, tax]) => {
          const pc = sheet.grid.get(r, roleCol[price]);
          const pv = pc ? pc.v : undefined;
          if (typeof pv === 'string' && pv.trim() && !/^[0-9.]+$/.test(pv.trim())) return; // 本体価格の欄が文字（よりどり3パック など）なら税込は空
          const pl = L(roleCol[price]);
          const cell = sheet.grid.at(r, roleCol[tax]);
          cell.f = 'IF(' + pl + r + '="","",ROUND(' + pl + r + '*108/100,0))';
          const n = C.num(pv);
          if (n !== null) cell.cached = Math.sign(n) * Math.round(Math.abs(n) * 108 / 100 + 1e-9);
          else cell.cachedText = '';
        });
        const k = kinds[r] || 'blank';
        const g = (r - start >= 0 && r - start < rows.length) ? rows[r - start] : null;
        const over = g ? parseOverrides(g['見た目']) : {};
        sheet.cols.forEach((col) => {
          const key = { product: '書式（商品）', section: '書式（区分の表題）', blank: '書式（空き）' }[k];
          const role = col.raw['役割'] || '';
          const sk = Object.prototype.hasOwnProperty.call(over, role) ? over[role] : col.raw[key];
          style(sheet.grid.at(r, col.c), sk);
        });
      }
      if (i + 1 < parts.length) shift = Math.max(shift, end + 1 - Number(parts[i + 1]['題の行']));
      lastRow = Math.max(lastRow, end);
    });
    sheet.lastRow = lastRow;
    sheet.page = pageOf(sh, lastRow);
    sheet.meta.monthly = metaRows;
    return sheet;
  }

  /** ブック全部（Excel の順: 月間商品・首都圏・北関東・焼き物） */
  function buildBook(db) {
    const monthNum = Number(String(db.month).split('-')[1]);
    const memo = C.rowsBySection(db);
    const sheets = [monthlySheet(db, monthNum)];
    C.TABLE_KEYS.forEach((t) => sheets.push(numberSheet(db, t, monthNum, memo)));
    return { month: db.month, sheets, styles: db.type.styles, shape: db.type.shapes };
  }
  function buildOne(db, table, memo) {
    const monthNum = Number(String(db.month).split('-')[1]);
    return table === '月間商品' ? monthlySheet(db, monthNum) : numberSheet(db, table, monthNum, memo);
  }

  root.SouzaiSheet = { buildBook, buildOne, letterToCol, colToLetter, parseOverrides };
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
