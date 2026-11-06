(() => {
  'use strict';
  // 月を足す・取り込む（本部と PO・守りは窓口）。左＝前の月から作る（アップロードなし）・右＝Excel から作る・入れる（手作りの番号表・サマリー。.xls でも .xlsx でも）。
  //   ファイルは画面の外の作業場所（import-worker.js）で読む＝読んでいる間も画面は止まらない。読むだけで、窓口へ送るのは「まとめて取り込む」を押したときだけ。
  //   取り込みの型（全システム共通）: 選ぶたびに積み上がる・何の物か何月かは中身から決める・同じ物は外して知らせる・決められない物だけ人に聞く。
  //   最初の取り込み（PO だけ・1回だけ）もここ（種のファイルを、この PC から窓口へ送る。公開の画面には置かない）。
  const App = window.SouzaiApp;
  const { $, el, C, state, toast, busyButton, api, rid, monthText } = App;
  const M = window.SouzaiImportMerge;
  const V = '0.2.1';
  const panel = $('#month-panel'), button = $('#month-add');
  const queue = []; // { key, file, name, status: 'reading'|'ready'|'bad'|'dup'|'done', kind, month, result, merge, summary, note, rid }
  let costsMemo = null; // 比べるための原価（この画面の中だけ・端末に残さない）
  const latest = () => state.months[state.months.length - 1] || '';

  // ------------------------------------------------------------ 画面の外で読む
  const W = { worker: null, seq: 0, wait: new Map() };
  function readInWorker(file) {
    if (!W.worker) {
      W.worker = new Worker('import-worker.js?v=' + V);
      W.worker.onmessage = (e) => { const job = W.wait.get(e.data.id); if (!job) return; W.wait.delete(e.data.id); job(e.data); };
      W.worker.onerror = () => { W.wait.forEach((job) => job({ ok: false, message: '読む部品を動かせませんでした。画面を読み込み直してください。' })); W.wait.clear(); W.worker = null; };
    }
    const t = state.typeTables || {}; const typeTables = { '書式': t['書式'], '表の型_列': t['表の型_列'], '表の型': t['表の型'], '表の型_まとまり': t['表の型_まとまり'] };
    return file.arrayBuffer().then((buf) => new Promise((resolve) => { const id = ++W.seq; W.wait.set(id, resolve); W.worker.postMessage({ id, name: file.name, bytes: buf, typeTables }, [buf]); }));
  }

  // ------------------------------------------------------------ 開く・閉じる（押した帯のすぐ下に開く）
  function open() { render(); panel.hidden = false; button.setAttribute('aria-expanded', 'true'); const f = panel.querySelector('button'); if (f) f.focus(); }
  function close() { panel.hidden = true; button.setAttribute('aria-expanded', 'false'); }
  button.addEventListener('click', () => { if (panel.hidden) open(); else close(); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !panel.hidden) { close(); button.focus(); } });

  function render() {
    const from = latest(), to = C.nextMonthOf(from);
    panel.replaceChildren();
    const head = el('div', 'mp-head'); const h = el('h2', '', '月を足す・取り込む'); h.id = 'month-panel-title';
    const x = el('button', 'btn btn-small', '× 閉じる'); x.type = 'button'; x.addEventListener('click', close);
    head.append(h, x);
    const cols = el('div', 'mp-cols');
    // 左: 前の月から作る
    const a = el('div', 'mp-card');
    a.append(el('h3', '', '前の月から作る（アップロードなし）'), el('p', '', monthText(to) + ' を、' + monthText(from) + ' の形から作ります。'));
    const ul = el('ul'); ['月間特売・新商品 → 定番（特売価格は残ります）', 'カット商品 → 休み（元の区分へ戻ります）', '並び・置き場・休み・計画は、そのまま', '終売の商品 → 次の月の一覧に載りません（番号の確かめには残ります）', '月間商品は見出しだけ。月間特売・新商品を決めたら、番号表の「月間特売」か「新規商品」のタブで「一覧から作り直す」'].forEach((t) => ul.append(el('li', '', t)));
    const make = el('button', 'btn btn-primary', '＋ ' + monthText(to) + ' を作る'); make.type = 'button'; make.id = 'month-make';
    make.addEventListener('click', () => createNext(make, from, to));
    a.append(ul, make);
    // 右: Excel から作る・入れる
    const b = el('div', 'mp-card');
    b.append(el('h3', '', 'Excel から作る・入れる'), el('p', 'hint-line', '手作りの番号表（その月を作る）と、サマリー（月間特売を入れる）を読めます。.xls でも .xlsx でもかまいません。何の物か・何月かは、中身から決めます。'));
    const input = el('input'); input.type = 'file'; input.multiple = true; input.accept = '.xls,.xlsx,.xlsm,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'; input.hidden = true; input.id = 'month-files';
    input.addEventListener('change', () => { addFiles(input.files); input.value = ''; });
    const pickB = el('button', 'btn', 'ファイルを選ぶ'); pickB.type = 'button'; pickB.addEventListener('click', () => input.click());
    const drop = el('div', 'mp-drop', 'ここへファイルを落としても入ります（何度でも足せます）');
    ['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('is-over'); }));
    ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('is-over'); }));
    drop.addEventListener('drop', (e) => addFiles(e.dataTransfer.files));
    const list = el('ul', 'mp-queue'); list.id = 'month-queue';
    const foot = el('div', 'mp-foot'); const go = el('button', 'btn btn-primary', 'まとめて取り込む'); go.type = 'button'; go.id = 'month-commit';
    go.addEventListener('click', () => commitAll(go));
    foot.append(go, el('span', 'hint-line', '番号表は月ごとに作り、サマリーはその月の月間特売に入れます（月がまだ無ければ、前の月から作ってから入れます）。'));
    b.append(pickB, input, drop, list, foot);
    cols.append(a, b);
    panel.append(head, cols);
    paintQueue();
  }

  // ------------------------------------------------------------ 前の月から作る（1往復）
  async function createNext(btn, from, to) {
    const done = busyButton(btn, monthText(to) + ' を作っています…');
    App.flush();
    const t0 = performance.now();
    try {
      const data = await App.serial(() => api('createMonth', { from, to, requestId: rid() }));
      App.perfNote('次の月を作る', t0);
      close();
      await App.loadMonth(to);
      const c = data.counts || {};
      toast(monthText(to) + ' を作りました（月間特売 ' + (c.sale || 0) + '・新商品 ' + (c.fresh || 0) + ' を定番に、カット商品 ' + (c.cut || 0) + ' を休みに' + (c.ended ? '・終売 ' + c.ended + ' は持ち越さず' : '') + '）。店にはまだ非公開です。');
    } catch (error) {
      if (error.code === 'SESSION') App.showLogin(error.message); else toast('作れませんでした: ' + error.message, true);
    } finally { done(); }
  }

  // ------------------------------------------------------------ 積み上げ（選ぶたびに足す・同じ物は外す）
  function addFiles(files) {
    Array.from(files || []).forEach((file) => {
      const key = file.name + '|' + file.size;
      if (queue.some((q) => q.key === key && q.status !== 'done')) { toast('同じファイル（' + file.name + '）は、もう入っています。'); return; }
      const item = { key, file, name: file.name, status: 'reading', note: '読んでいます…', rid: rid() };
      queue.push(item); paintQueue();
      const t0 = performance.now();
      readInWorker(file).then(async (res) => {
        item.ms = Math.round(performance.now() - t0);
        if (!res.ok) { item.status = 'bad'; item.note = res.message; paintQueue(); return; }
        item.kind = res.kind; item.result = res.result; item.month = res.result && res.result.month;
        if (res.kind === 'table') await planTable(item);
        else if (res.kind === 'summary') planSummary(item);
        else { item.status = 'bad'; item.note = '商品呼出番号表でもサマリーでもないようです（見えるシートの名前で見分けています）。'; }
        markDups(); paintQueue();
        App.perfNote('取り込む物を読む（' + (item.kind || '?') + '）', t0);
      });
    });
  }
  /** 同じ月の同じ物は、先に入れた物だけ残す（外したことを行に出す） */
  function markDups() {
    const seen = new Map();
    queue.forEach((q) => {
      if (q.status === 'dup') q.status = 'ready';
      if (q.status !== 'ready') return;
      const k = q.kind + '|' + q.month;
      if (seen.has(k)) { q.status = 'dup'; q.note = '同じ月の' + (q.kind === 'table' ? '番号表' : 'サマリー') + 'が先に入っているので、外しました。'; } else seen.set(k, q);
    });
  }
  async function costsForCompare() {
    if (state.db && state.db.costs) return state.db.costs;
    if (costsMemo) return costsMemo;
    try { const d = await api('getCosts', { month: state.month, requestId: rid() }); costsMemo = new Map(C.rowsOf(d.costs).map((r) => [String(r['商品ID']), r])); } catch { costsMemo = null; }
    return costsMemo;
  }
  async function planTable(item) {
    const b = item.result;
    if (b && !b.tables) { item.status = 'bad'; item.note = '取り込めません（下の「止めた理由」）。'; item.merge = { preview: { problems: b.problems || [], lost: 0 } }; return; }
    if (!b || !b.month) { item.status = 'bad'; item.note = '何月の番号表か、決められませんでした。'; return; }
    if (state.months.includes(b.month)) { item.status = 'bad'; item.note = monthText(b.month) + ' は、もうあります（その月を直すときは、その月の一覧で直します）。'; return; }
    const costs = await costsForCompare();
    item.merge = M.mergeBundle(b, state.db, { latest: latest(), costs, source: item.name });
    item.status = item.merge.ok ? 'ready' : 'bad';
    item.note = item.merge.ok ? '' : '取り込めません（下の「止めた理由」）。';
  }

  // ------------------------------------------------------------ 一覧を描く
  function paintQueue() {
    const list = $('#month-queue'); if (!list) return;
    list.replaceChildren();
    queue.filter((q) => q.status !== 'done').forEach((q) => {
      const li = el('li', 'mp-item' + (q.status === 'bad' ? ' is-bad' : '') + (q.status === 'dup' ? ' is-dup' : ''));
      const line = el('div', 'mp-line');
      line.append(el('span', 'mp-name', q.name));
      if (q.status === 'reading') line.append(App.spinner(), el('span', '', '読んでいます…'));
      else if (q.kind === 'table' && q.month) line.append(el('span', 'mp-what', '→ ' + monthText(q.month) + 'の番号表'));
      else if (q.kind === 'summary' && q.month) line.append(el('span', 'mp-what', '→ ' + monthText(q.month) + 'のサマリー'));
      const xb = el('button', 'btn btn-small mp-x', '✕ 外す'); xb.type = 'button'; xb.setAttribute('aria-label', q.name + ' を外す');
      xb.addEventListener('click', () => { const i = queue.indexOf(q); if (i >= 0) queue.splice(i, 1); markDups(); paintQueue(); });
      line.append(xb);
      li.append(line);
      if (q.note) li.append(el('div', q.status === 'bad' ? 'mp-stop' : 'hint-line', q.note));
      if (q.status === 'ready' && q.kind === 'table') li.append(tableSummary(q));
      if ((q.status === 'ready' || q.status === 'dup') && q.kind === 'summary') li.append(summarySummary(q));
      if (q.status === 'bad' && q.merge) li.append(problemsBox(q.merge.preview));
      list.append(li);
    });
    const go = $('#month-commit'); if (go) go.disabled = !queue.some((q) => q.status === 'ready');
  }
  function tableSummary(q) {
    const p = q.merge.preview; const box = el('div');
    const fields = Object.entries(p.byField).map(([k, n]) => k + ' ' + n).join('・');
    box.append(el('div', '', '行 ' + p.rows + '（商品 ' + p.productRows + '）・いまの商品に結べた ' + p.matched + '・新しい商品 ' + p.added.length + '・新しい区分 ' + p.newSections.length + '・直す欄 ' + p.changed.length + (fields ? '（' + fields + '）' : '') + (p.costChanged ? '・原価を直す商品 ' + p.costChanged : '')));
    if (!p.isNewest) box.append(el('div', 'hint-line', 'いちばん新しい月より前の月なので、マスター（商品・原価）は直しません。その月の表だけを作ります。'));
    const d = el('details'); d.append(el('summary', '', 'くわしく（変わる所）'));
    if (p.added.length) d.append(miniTable('新しい商品', ['名前', '番号', '表'], p.added.map((a) => [a.name, a.call, a.kind])));
    if (p.newSections.length) d.append(miniTable('新しい区分', ['表題'], p.newSections.map((s) => [s])));
    if (p.changed.length) d.append(miniTable('直す欄', ['商品', '欄', '前', '後'], p.changed.slice(0, 200).map((c) => [c.name, c.field, C.showCell(c.before), C.showCell(c.after)])));
    if (p.changed.length > 200) d.append(el('div', 'hint-line', 'ほか ' + (p.changed.length - 200) + ' 欄'));
    const checks = (p.problems || []).filter((x) => x.level !== '止める');
    if (checks.length) d.append(miniTable('確かめる所（取り込むと、アラートにも出ます）', ['何が', 'どこ'], checks.map((x) => [x.what, x.where || ''])));
    box.append(d);
    return box;
  }
  function problemsBox(p) {
    const box = el('div');
    const stop = (p.problems || []).filter((x) => x.level === '止める');
    if (stop.length) box.append(miniTable('止めた理由', ['何が', 'どこ'], stop.map((x) => [x.what, x.where || ''])));
    if (p.lost) box.append(el('div', 'mp-stop', '結べない番号が ' + p.lost + ' か所あります。'));
    return box;
  }
  function miniTable(title, heads, rows) {
    const wrap = el('div'); wrap.append(el('strong', '', title + '（' + rows.length + '）'));
    const t = el('table'); const hr = el('tr'); heads.forEach((h) => hr.append(el('th', '', h))); t.append(hr);
    rows.forEach((r) => { const tr = el('tr'); r.forEach((v) => tr.append(el('td', '', v === undefined || v === null ? '' : String(v)))); t.append(tr); });
    wrap.append(t); return wrap;
  }

  // ------------------------------------------------------------ まとめて取り込む（番号表は古い月から・そのあとサマリー）
  async function commitAll(btn) {
    const tables = queue.filter((q) => q.status === 'ready' && q.kind === 'table').sort((a, b) => (a.month < b.month ? -1 : 1));
    const sums = queue.filter((q) => q.status === 'ready' && q.kind === 'summary').sort((a, b) => (a.month < b.month ? -1 : 1));
    if (!tables.length && !sums.length) return;
    const done = busyButton(btn, '取り込んでいます…');
    // 途中で月を開き直すとパネルが閉じる＝画面全体の待ちの印に段階の文字を出し続ける（月を開く部品が印を消すので、開いた後に出し直す）
    let stageText = '取り込んでいます…';
    const say = (t) => { if (t) stageText = t; done.stage(stageText); App.showProgress(stageText); };
    const loadKeep = async (m) => { await App.loadMonth(m); say(); };
    say();
    App.flush();
    const report = []; let lastMonth = '';
    try {
      for (const q of tables) {
        say(monthText(q.month) + ' を作っています…');
        const t0 = performance.now();
        const data = await App.serial(() => api('importMonth', Object.assign({}, q.merge.payload, { requestId: q.rid })));
        App.perfNote('Excel から月を作る', t0);
        q.status = 'done'; lastMonth = q.month;
        if (!state.months.includes(q.month)) { state.months.push(q.month); state.months.sort(); }
        const c = data.counts || {}; report.push(monthText(q.month) + 'を作りました（行 ' + c.rows + '・新しい商品 ' + c.newProducts + '・直した欄 ' + c.updated + '）');
      }
      if (tables.length) await loadKeep(lastMonth || latest()); // マスターが変わったので読み直す（サマリーの突き合わせも新しいマスターで）
      for (const q of sums) {
        say(monthText(q.month) + ' の月間特売を入れています…');
        const msg = await applySummary(q, loadKeep);
        q.status = 'done'; lastMonth = q.month; report.push(msg);
      }
      if (lastMonth && lastMonth !== state.month) await loadKeep(lastMonth);
      close();
      toast(report.join('／'));
    } catch (error) {
      if (error.code === 'SESSION') App.showLogin(error.message);
      else { toast('途中で止まりました: ' + error.message + (report.length ? '（済んだ物: ' + report.join('／') + '）' : ''), true); paintQueue(); }
    } finally { App.hideProgress(); done(); }
  }

  // ------------------------------------------------------------ サマリー（月間厳選特の文）→ その月の月間特売
  /** サマリーを読んだ結果に、いまのマスターの商品の候補を付ける（人が選び直せる） */
  function planSummary(item) {
    const S = window.SouzaiImportSummary || null; const r = item.result;
    if (!S || !r || !r.month) { item.status = 'bad'; item.note = '何月のサマリーか、決められませんでした。'; return; }
    const stop = (r.problems || []).filter((p) => p.level === '止める');
    if (stop.length || !(r.items || []).length) { item.status = 'bad'; item.note = stop.length ? stop.map((p) => p.what).join('・') : '「月間厳選特」の行が見つかりませんでした。'; return; }
    const products = Array.from(state.db.products.values());
    item.summary = { matches: S.matchItems(r.items, products, { monthlyRows: state.db.monthly, words: state.words || undefined }) };
    item.summary.pick = item.summary.matches.map((m) => (m && m.kind === 'one' && m.sure && m.best ? String(m.best['商品ID']) : '')); // 人が選び直した商品ID（'' は結ばない）。種類ごとの物（family）は部品の答えのまま
    item.status = 'ready'; item.note = '';
    const can = state.months.includes(r.month) || r.month === C.nextMonthOf(latest()) || queue.some((q) => q.kind === 'table' && q.month === r.month && q.status === 'ready');
    if (!can) { item.status = 'bad'; item.note = monthText(r.month) + ' がまだありません。先にその月を作ってください（いちばん新しい月の次の月なら、ここで作ってから入れます）。'; }
  }
  function summarySummary(q) {
    const r = q.result; const box = el('div');
    const m = q.summary.matches; const isSure = (x, i) => (x && x.kind === 'family' ? x.sure : Boolean(q.summary.pick[i]));
    const sure = m.filter(isSure).length;
    box.append(el('div', '', '月間特売 ' + r.items.length + ' 件（商品に結べた ' + sure + '・確かめる ' + (r.items.length - sure) + '）' + (r.period ? '・' + r.period : '') + ((r.leftovers || []).length ? '・読めなかった文 ' + r.leftovers.length : '')));
    const d = el('details'); d.open = sure < r.items.length; d.append(el('summary', '', 'くわしく（結ぶ商品を選び直せます）'));
    const t = el('table'); const hr = el('tr'); ['サマリーの文', '規格と価格（本体）', '結ぶ商品'].forEach((h) => hr.append(el('th', '', h))); t.append(hr);
    r.items.forEach((it, i) => {
      const tr = el('tr'); tr.append(el('td', '', it.name || it.text), el('td', '', (it.specs || []).map((s) => (s['規格'] ? s['規格'] + ' ' : '') + s['本体価格'] + '円').join('／')));
      const td = el('td'); const x = m[i];
      if (x && x.kind === 'family') { // 種類ごとの物は、種類ごとに結ぶ（確かでない種類は、入れたあと一覧で選ぶ）
        td.append(document.createTextNode('種類ごと ' + (x.parts || []).length + '（確か ' + (x.parts || []).filter((y) => y.sure).length + '）'));
      } else {
        const sel = el('select'); sel.setAttribute('aria-label', (it.name || '') + ' に結ぶ商品');
        const o0 = el('option', '', '結ばない（月間商品に文だけ入れる）'); o0.value = ''; sel.append(o0);
        ((x && x.candidates) || []).forEach((c) => { const p = state.db.products.get(String(c['商品ID'])); if (!p) return; const o = el('option', '', C.nfkc(p['名前']).trim() + '（' + (p['呼出番'] || '番号なし') + '・' + c['点数'] + '点）'); o.value = String(c['商品ID']); sel.append(o); });
        sel.value = q.summary.pick[i] || '';
        sel.addEventListener('change', () => { q.summary.pick[i] = sel.value; });
        td.append(sel);
        if (x && !x.sure && x.why) td.append(el('div', 'hint-line', x.why));
      }
      tr.append(td); t.append(tr);
    });
    d.append(t);
    if ((r.leftovers || []).length) d.append(miniTable('読めなかった文', ['文'], r.leftovers.map((s) => [s])));
    box.append(d);
    return box;
  }
  /** その月を開き（無ければ前の月から作り）、結んだ商品を月間特売にして特売価格を入れ、月間商品の月間特売の部をサマリーの行で作り直す */
  async function applySummary(q, loadKeep) {
    const S = window.SouzaiImportSummary; const r = q.result;
    let made = '';
    if (!state.months.includes(r.month)) {
      const from = latest();
      if (C.nextMonthOf(from) !== r.month) throw new Error(monthText(r.month) + ' がまだありません。');
      await App.serial(() => api('createMonth', { from, to: r.month, requestId: q.rid + '-m' }));
      state.months.push(r.month); made = monthText(r.month) + 'を前の月から作り、';
    }
    if (state.month !== r.month) await loadKeep(r.month);
    const db = state.db;
    // 人が選び直した商品を、確かな候補として渡す（family は部品の答えのまま）
    const matches = q.summary.matches.map((m, i) => {
      if (!m || m.kind === 'family') return m;
      const id = q.summary.pick[i]; if (!id) return Object.assign({}, m, { sure: false });
      const p = db.products.get(id); const c = (m.candidates || []).find((x) => String(x['商品ID']) === id) || { '商品ID': id, '名前': p ? p['名前'] : '', '呼出番': p ? p['呼出番'] : '', '点数': 100, '理由': '人が選んだ' };
      return Object.assign({}, m, { best: c, sure: true });
    });
    let maxG = 0; db.monthly.forEach((g) => { const x = /^G(\d+)$/.exec(String(g['行ID'])); if (x) maxG = Math.max(maxG, Number(x[1])); });
    const plan = S.toMonthlyRows(r.items, matches, { products: Array.from(db.products.values()), monthlyRows: db.monthly, usedIds: db.monthly.map((g) => String(g['行ID'])), firstNo: maxG + 1 });
    // 番号表の行: 月間特売を足す（新商品はそのまま）・特売価格
    const changes = [];
    (plan.changes || []).forEach((c) => {
      db.rows.forEach((row) => {
        if (String(row['商品ID']) !== String(c['商品ID'])) return;
        const f = {}; const next = C.isSaleForm(row['販売形態']) ? row['販売形態'] : C.toggleForm(row['販売形態'] || C.FORM_REGULAR, C.FORM_SALE);
        if (next !== row['販売形態'] && !C.isOffForm(row['販売形態'])) f['販売形態'] = next;
        if (c['特売価格'] !== undefined && c['特売価格'] !== '' && String(c['特売価格']) !== String(row['特売価格'])) f['特売価格'] = c['特売価格'];
        if (Object.keys(f).length) changes.push({ t: '行', id: String(row['行ID']), f });
      });
    });
    if (changes.length) { const p = App.change(changes, { label: 'サマリーから' }); if (p) throw new Error(p); await App.flush(); }
    // 月間商品: 月間特売の部を、サマリーの行（節の見出しつき）で置き換える。新規商品の部はそのまま。そのあと一覧の販売形態に合わせる（サマリーに無い月間特売の商品を足す）
    const keep = db.monthly.filter((g) => g['部'] !== '月間特売').sort((a, b) => Number(a['順']) - Number(b['順']));
    db.monthly = (plan.rows || []).concat(keep); const res = C.monthlyReconcile(db);
    await App.serial(() => api('replaceMonthly', { month: r.month, rows: { c: C.MONTHLY_COLS, r: res.rows.map((g) => C.MONTHLY_COLS.map((h) => (g[h] === undefined || g[h] === null ? '' : g[h]))) }, label: 'サマリーから', requestId: q.rid }));
    db.monthly = res.rows; App.refresh([]);
    const asked = (plan.toPick || []).length;
    return made + monthText(r.month) + 'の月間特売に ' + (plan.rows || []).filter((g) => g['行の種類'] === C.KIND_PRODUCT).length + ' 行を入れました（番号表の行 ' + changes.length + ' か所' + (asked ? '・商品に結んでいない行 ' + asked + '＝番号表の「月間特売」か「新規商品」のタブで確かめる' : '') + '）';
  }

  // ------------------------------------------------------------ 最初の取り込み（PO だけ・1回だけ）
  function showSeed(principal) {
    const p = $('#seed-panel'); p.replaceChildren(); p.hidden = false;
    if (principal.role !== 'PO') {
      const h = el('h1', '', '商品呼出番号表は、準備中です'); h.id = 'seed-title';
      p.append(h, el('p', '', '本部で最初のデータを入れています。入ったあとに、この画面を開き直すと使えます。'));
      return;
    }
    const h = el('h1', '', '最初の取り込み（管理者だけ・1回だけ）'); h.id = 'seed-title';
    p.append(h, el('p', '', 'この PC の中の「種」のファイル（種_2026-10.json）を選ぶと、表ごとに窓口へ送ります。ファイルは公開の画面には置きません。終わると、ふつうの画面が開きます。'));
    const input = el('input'); input.type = 'file'; input.accept = '.json,application/json'; input.hidden = true;
    const pick = el('button', 'btn btn-primary', '種のファイルを選ぶ'); pick.type = 'button'; pick.id = 'seed-pick';
    pick.addEventListener('click', () => input.click());
    const steps = el('ol', 'seed-steps'); steps.id = 'seed-steps';
    input.addEventListener('change', () => { const f = input.files && input.files[0]; input.value = ''; if (f) runSeed(f, pick, steps); });
    p.append(pick, input, steps);
  }
  async function runSeed(file, pick, steps) {
    let seed;
    try { seed = JSON.parse(await file.text()); } catch { toast('種のファイルを読めませんでした（JSON ではありません）。', true); return; }
    if (!seed || !/^20\d{2}-\d{2}$/.test(String(seed.month)) || !seed.tables || typeof seed.tables !== 'object') { toast('種のファイルの形が違います（month と tables がありません）。', true); return; }
    const names = Object.keys(seed.tables);
    steps.replaceChildren(); const items = names.map((n) => { const li = el('li', '', n + '（' + (seed.tables[n].r || []).length + ' 行）'); steps.append(li); return li; });
    const fin = el('li', '', '仕上げ（取り込み済みの印）'); steps.append(fin);
    const done = busyButton(pick, '送っています…');
    const t0 = performance.now();
    try {
      for (let i = 0; i < names.length; i++) {
        const t = seed.tables[names[i]]; items[i].className = 'is-now';
        const chunks = []; let cur = [], size = 0; // 1回に送るのは 15 万字まで（窓口の上限 20 万字）
        (t.r || []).forEach((r) => { const n = JSON.stringify(r).length; if (cur.length && size + n > 150000) { chunks.push(cur); cur = []; size = 0; } cur.push(r); size += n; });
        chunks.push(cur);
        for (let k = 0; k < chunks.length; k++) {
          done.stage(names[i] + ' を送っています…（' + (i + 1) + '／' + names.length + '）');
          // 操作番号は中身から決める＝止まってからもう一度押しても、送り済みの分は窓口が前の結果を返す（二重に書かない）
          await api('seedImport', { table: names[i], first: k === 0, c: t.c, r: chunks[k], requestId: await seedId(seed.month, names[i], k, chunks[k]) });
        }
        items[i].className = 'is-done';
      }
      fin.className = 'is-now';
      await api('seedImport', { final: true, requestId: await seedId(seed.month, '仕上げ', 0, names) });
      fin.className = 'is-done';
      App.perfNote('最初の取り込み', t0);
      toast('最初の取り込みが済みました。画面を開き直します。');
      window.setTimeout(() => location.reload(), 800);
    } catch (error) {
      toast('途中で止まりました: ' + error.message + '（同じファイルをもう一度選ぶと、送り済みの分は二重にせず、続きから送ります）', true);
    } finally { done(); }
  }
  async function seedId(month, table, k, rows) {
    const bytes = new TextEncoder().encode(month + '|' + table + '|' + k + '|' + JSON.stringify(rows));
    const hash = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
    return 'seed-' + Array.from(hash.slice(0, 16), (b) => b.toString(16).padStart(2, '0')).join('');
  }

  App.views.intake = { showSeed, open, close };
})();
