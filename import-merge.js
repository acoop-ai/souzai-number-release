// 手作りの番号表から読んだ「月の束」（import-table.js）を、いまのマスター（商品・区分・原価）に結ぶ部品（画面の部品を持たない）。
//   結んだ結果は、窓口の importMonth にそのまま送れる形と、人が確かめる「変わる所」の一覧の2つ。
//   決まり: 区分は 表の種類＋表題（同じ表題が2つあれば出てくる順）で結ぶ。商品は ①番号と名前が同じ ②名前が同じ ③番号が同じで名前が似ている の順に結び、どれでもなければ新しい商品。
//   マスター（商品・原価）を直すのは、いちばん新しい月より新しい月のときだけ（前の月を入れても、いまのマスターを古い値に戻さない）。
(function (root) {
  'use strict';
  const C = root.SouzaiCore;
  const PRODUCT_FIELDS = ['呼出番', '名前', '字の飾り', '指定トレー', '定番価格', '小パック']; // 番号表から直す欄（地域・メモは直さない）
  const IMPORT_MARK = '取り込み'; // 新しい商品の 元の商品ID（足した商品と同じく、ほかの月で使っていなければ取り消せる）
  const keyOf = (v) => C.nfkc(v).replace(/\s+/g, '');
  const same = (a, b) => String(a === null || a === undefined ? '' : a) === String(b === null || b === undefined ? '' : b);
  function bigrams(s) { const out = new Map(); for (let i = 0; i < s.length - 1; i++) { const g = s.slice(i, i + 2); out.set(g, (out.get(g) || 0) + 1); } return out; }
  /** 名前の似ている度合い（0〜1・2字ずつの重なり） */
  function similar(a, b) {
    const x = bigrams(keyOf(a)), y = bigrams(keyOf(b)); let both = 0, nx = 0, ny = 0;
    x.forEach((n) => { nx += n; }); y.forEach((n) => { ny += n; });
    x.forEach((n, g) => { if (y.has(g)) both += Math.min(n, y.get(g)); });
    return nx + ny ? (2 * both) / (nx + ny) : 0;
  }
  const numOf = (id, re) => { const m = re.exec(String(id)); return m ? Number(m[1]) : 0; };

  /** 区分を結ぶ。返す値: { map: 束の区分ID → マスターの区分ID, added: [新しい区分の行の物] } */
  function mapSections(bundleSecs, db) {
    const map = new Map(), added = [];
    const groupOf = (list) => { const g = new Map(); list.forEach((s) => { const k = s['表の種類'] + '|' + keyOf(s['表題']); if (!g.has(k)) g.set(k, []); g.get(k).push(s); }); return g; };
    const byNo = (a, b) => numOf(a['区分ID'], /-(\d+)$/) - numOf(b['区分ID'], /-(\d+)$/);
    const master = groupOf(Array.from(db.sections.values()).slice().sort(byNo));
    const next = {}; db.sections.forEach((s) => { const k = s['表の種類']; next[k] = Math.max(next[k] || 0, numOf(s['区分ID'], /-(\d+)$/)); });
    const bundle = groupOf(bundleSecs.slice().sort(byNo));
    bundle.forEach((list, k) => {
      const have = master.get(k) || [];
      list.forEach((s, i) => {
        if (have[i]) { map.set(String(s['区分ID']), String(have[i]['区分ID'])); return; }
        const kind = s['表の種類']; next[kind] = (next[kind] || 0) + 1;
        const id = kind + '-' + String(next[kind]).padStart(2, '0');
        added.push(Object.assign({}, s, { '区分ID': id }));
        map.set(String(s['区分ID']), id);
      });
    });
    return { map, added };
  }

  /** 商品を結ぶ。返す値: { map: 束の商品ID → マスターの商品ID, how: 束の商品ID → 結び方, added: [束の商品] } */
  function mapProducts(bundleProducts, db) {
    const map = new Map(), how = new Map(), added = [];
    const free = new Set(db.products.keys());
    const all = Array.from(db.products.values());
    const take = (bp, p, why) => { map.set(String(bp['商品ID']), String(p['商品ID'])); how.set(String(bp['商品ID']), why); free.delete(String(p['商品ID'])); };
    const prefer = (bp, list) => list.find((p) => p['表の種類'] === bp['表の種類'] && (p['地域'] || '共通') === (bp['地域'] || '共通')) || list.find((p) => p['表の種類'] === bp['表の種類']) || list[0];
    const pending = bundleProducts.slice();
    const pass = (test, why) => {
      for (let i = pending.length - 1; i >= 0; i--) {
        const bp = pending[i];
        const hits = all.filter((p) => free.has(String(p['商品ID'])) && test(bp, p));
        if (!hits.length) continue;
        if (why !== '番号と名前' && hits.length > 1 && !hits.some((p) => p['表の種類'] === bp['表の種類'])) continue; // 名前だけ・番号だけで2つ以上に当たり、決められない
        take(bp, prefer(bp, hits), why); pending.splice(i, 1);
      }
    };
    const callOf = (p) => C.halfCall(p['呼出番']) || '';
    pass((bp, p) => callOf(bp) !== '' && callOf(bp) === callOf(p) && keyOf(bp['名前']) === keyOf(p['名前']), '番号と名前');
    pass((bp, p) => keyOf(bp['名前']) !== '' && keyOf(bp['名前']) === keyOf(p['名前']), '名前');
    pass((bp, p) => callOf(bp) !== '' && callOf(bp) === callOf(p) && similar(bp['名前'], p['名前']) >= 0.5, '番号と似た名前');
    pending.forEach((bp) => added.push(bp));
    return { map, how, added };
  }

  /** 月の束をマスターに結ぶ。opts: { latest: いちばん新しい月, costs: 原価（Map・無ければ原価は比べない）, source: ファイル名 }。
   *  返す値: { ok, payload（importMonth に送る物）, preview（人が確かめる物） } */
  function mergeBundle(bundle, db, opts) {
    opts = opts || {};
    const month = bundle.month; const T = bundle.tables || {};
    const stop = (bundle.problems || []).filter((p) => p.level === '止める');
    const isNewest = !opts.latest || month > opts.latest;
    const secs = mapSections(C.rowsOf(T['区分']), db);
    const bundleProducts = C.rowsOf(T['商品']);
    const prods = mapProducts(bundleProducts, db);
    // 新しい商品の番号（いまの SZ の続き）
    let next = 0; db.products.forEach((p, id) => { next = Math.max(next, numOf(id, /^SZ(\d+)$/)); });
    prods.added.forEach((bp) => { next += 1; prods.map.set(String(bp['商品ID']), 'SZ' + String(next).padStart(4, '0')); });
    const pid = (v) => (C.isBlank(v) ? '' : (prods.map.get(String(v)) || ''));
    const sid = (v) => (C.isBlank(v) ? '' : (secs.map.get(String(v)) || ''));
    const P_COLS = ['商品ID', '表の種類', '地域', '対の商品ID', '呼出番', '名前', '字の飾り', '指定トレー', '定番価格', '小パック', 'メモ', '元の商品ID'];
    const S_COLS = ['区分ID', '表の種類', '表題', '表題の続き', '表題の書式', '表題の置き場（元）', '見た目', '字の飾り', '北関東の違い'];
    const COST_COLS = ['商品ID'].concat(C.COST_KEYS);
    const pick = (o, cols, over) => cols.map((h) => { const v = over && Object.prototype.hasOwnProperty.call(over, h) ? over[h] : o[h]; return v === undefined || v === null ? '' : v; });
    const newProducts = { c: P_COLS, r: prods.added.map((bp) => pick(bp, P_COLS, { '商品ID': pid(bp['商品ID']), '対の商品ID': pid(bp['対の商品ID']), '元の商品ID': IMPORT_MARK })) };
    const newSections = { c: S_COLS, r: secs.added.map((s) => pick(s, S_COLS)) };
    const costRows = C.rowsOf(T['原価']).filter((r) => !/^R/.test(String(r['商品ID']))); // 続きの行（行ID で持つ原価）は入れない
    const addedIds = new Set(prods.added.map((bp) => String(bp['商品ID'])));
    const newCosts = { c: COST_COLS, r: costRows.filter((r) => addedIds.has(String(r['商品ID']))).map((r) => pick(r, COST_COLS, { '商品ID': pid(r['商品ID']) })) };
    // 直す欄（いちばん新しい月より新しいときだけ）
    const updates = [], changed = [];
    if (isNewest) {
      bundleProducts.forEach((bp) => {
        if (addedIds.has(String(bp['商品ID']))) return;
        const p = db.products.get(pid(bp['商品ID'])); if (!p) return;
        const f = {};
        PRODUCT_FIELDS.forEach((k) => { if (!same(bp[k], p[k])) { f[k] = bp[k] === undefined || bp[k] === null ? '' : bp[k]; changed.push({ id: p['商品ID'], name: C.nfkc(p['名前']).trim(), field: k, before: p[k], after: f[k] }); } });
        if (Object.keys(f).length) updates.push({ t: '商品', id: String(p['商品ID']), f });
      });
      if (opts.costs) costRows.forEach((r) => {
        if (addedIds.has(String(r['商品ID']))) return;
        const id = pid(r['商品ID']); if (!id) return; const now = opts.costs.get(id) || {};
        const f = {}; C.COST_KEYS.forEach((k) => { if (!same(r[k], now[k])) f[k] = r[k] === undefined || r[k] === null ? '' : r[k]; });
        if (Object.keys(f).length) updates.push({ t: '原価', id, f });
      });
    }
    // その月の表（区分・商品の番号を付け替える）
    const rowsT = T['月の行_' + month] || { c: [], r: [] }; const rc = rowsT.c;
    const iSec = rc.indexOf('区分ID'), iHome = rc.indexOf('元の区分ID'), iPid = rc.indexOf('商品ID');
    let lost = 0; // 結べなかった番号（あれば送らない）
    const mapOrLose = (fn, v) => { const out = fn(v); if (!out) lost++; return out; };
    const rows = { c: rc.slice(), r: (rowsT.r || []).map((r0) => { const r = r0.slice(); r[iSec] = mapOrLose(sid, r[iSec]); if (!C.isBlank(r[iHome])) r[iHome] = mapOrLose(sid, r[iHome]); if (!C.isBlank(r[iPid])) r[iPid] = mapOrLose(pid, r[iPid]); return r; }) };
    const placeT = T['月の置き場_' + month] || { c: [], r: [] };
    const places = { c: placeT.c.slice(), r: (placeT.r || []).map((r0) => { const r = r0.slice(); r[0] = mapOrLose(sid, r[0]); return r; }) };
    const monT = T['月間商品_' + month] || { c: C.MONTHLY_COLS, r: [] }; const jPid = monT.c.indexOf('商品ID');
    const monthly = { c: monT.c.slice(), r: (monT.r || []).map((r0) => { const r = r0.slice(); if (!C.isBlank(r[jPid])) r[jPid] = mapOrLose(pid, r[jPid]); return r; }) };
    const payload = { month, tables: { '月の行': rows, '月の置き場': places, '月間商品': monthly }, newProducts, newCosts, newSections, updates, source: String(opts.source || '').slice(0, 120) };
    const byField = {}; changed.forEach((x) => { byField[x.field] = (byField[x.field] || 0) + 1; });
    const preview = {
      month, isNewest, rows: rows.r.length, productRows: rows.r.filter((r) => !C.isBlank(r[iPid])).length,
      matched: prods.map.size - prods.added.length, how: countBy(Array.from(prods.how.values())),
      added: prods.added.map((bp) => ({ name: C.nfkc(bp['名前']).trim(), call: bp['呼出番'] || '', kind: bp['表の種類'] })),
      newSections: secs.added.map((s) => C.nfkc(s['表題']).trim()),
      changed, byField, costChanged: updates.filter((u) => u.t === '原価').length,
      forms: countBy(rows.r.map((r) => r[rc.indexOf('販売形態')]).filter((f) => f)),
      problems: bundle.problems || [], stop, lost
    };
    const size = JSON.stringify(payload).length;
    const ok = !stop.length && !lost && size < 190000;
    if (size >= 190000) preview.problems = preview.problems.concat([{ level: '止める', what: '送る中身が大きすぎます（' + Math.round(size / 1000) + ' KB）', where: '' }]);
    return { ok, payload, preview };
  }
  function countBy(list) { const o = {}; list.forEach((k) => { o[k] = (o[k] || 0) + 1; }); return o; }

  root.SouzaiImportMerge = { mergeBundle, mapSections, mapProducts, similar, IMPORT_MARK, PRODUCT_FIELDS };
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
