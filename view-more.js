(() => {
  'use strict';
  // 記録のタブ（だれが・いつ・何を。前の値に戻す）と、設定（本部と PO: パックの値段／PO だけ: 上限の行数の3段・A4 の見立て・月を消す）。
  const App = window.SouzaiApp;
  const { state, $, el, C, toast } = App;
  const log = { rows: [], q: '', loaded: false };
  const when = (v) => String(v ?? '').replace('T', ' ').slice(0, 16);
  const canRevert = (x) => x['欄'] && !x.masked && ['商品', '原価', '行', '月間', '置き場'].includes(x['種類']) && C.EDITABLE[x['種類']] && C.EDITABLE[x['種類']].includes(x['欄']);
  function renderLog() {
    const root = $('#view-log'); root.replaceChildren();
    const bar = el('div', 'log-bar');
    const input = el('input'); input.type = 'search'; input.id = 'log-search'; input.placeholder = '名前・人・欄で絞る'; input.value = log.q; input.setAttribute('aria-label', '記録を絞る');
    const again = el('button', 'btn', '読み直す'); again.type = 'button'; again.id = 'log-reload'; again.addEventListener('click', () => loadLog(again));
    bar.append(input, again, el('span', 'hint-line', 'だれが・いつ・何を変えたか。原価の値は、原価を表示している人にだけ出ます。'));
    root.append(bar);
    const wrap = el('div'); wrap.id = 'log-table-wrap'; root.append(wrap);
    input.addEventListener('input', () => { log.q = input.value; paintLogTable(); });
    paintLogTable();
    root.append(perfDetails());
  }
  function paintLogTable() {
    const wrap = $('#log-table-wrap'); if (!wrap) return; wrap.replaceChildren();
    const q = C.nfkc(log.q).toLowerCase().trim();
    const rows = log.rows.filter((x) => !q || C.nfkc([x['対象の名前'], x['メール'], x['欄'], x['操作']].join(' ')).toLowerCase().includes(q));
    if (!rows.length) { wrap.append(el('p', 'hint-line', log.loaded ? '記録はまだありません。' : '読んでいます…')); return; }
    const t = el('table', 'data-table'); t.id = 'log-table';
    const head = el('thead'); const hr = el('tr'); ['いつ', 'だれ', '操作', '何を', '欄', '前の値', '後の値', ''].forEach((h) => hr.append(el('th', '', h))); head.append(hr); t.append(head);
    const body = el('tbody');
    rows.forEach((x) => {
      const tr = el('tr');
      [when(x['時刻']), x['メール'], x['操作'], x['対象の名前'] || x['対象'], x['欄'], x['前の値'], x['後の値']].forEach((v) => tr.append(el('td', '', v)));
      const td = el('td');
      if (canRevert(x)) { const b = el('button', 'btn btn-small', '前の値に戻す'); b.type = 'button'; b.addEventListener('click', () => revert(x, b)); td.append(b); }
      if ((x['操作'] === '新しい商品' || x['操作'] === '写して新しい商品') && state.db.products.has(String(x['対象']))) { const b = el('button', 'btn btn-small btn-danger', '取り消す'); b.type = 'button'; b.addEventListener('click', async () => { if (b.dataset.sure !== '1') { b.dataset.sure = '1'; b.textContent = 'もう一度押すと消します'; return; } if (await App.cancelProduct(String(x['対象']), b)) await loadLog(); }); td.append(b); }
      tr.append(td); body.append(tr);
    });
    t.append(body); wrap.append(t);
  }
  async function revert(x, button) {
    const before = x['前の値'];
    const m = App.change([{ t: x['種類'], id: String(x['対象']), f: { [x['欄']]: before } }], { label: '前の値に戻す' });
    if (m) { toast(m, true); return; }
    const done = App.busyButton(button, '戻しています…');
    App.flush();
    try { await App.serial(() => Promise.resolve()); toast('前の値に戻しました。戻したことも記録に残ります。'); await loadLog(); } finally { done(); }
  }
  async function loadLog(button) {
    const done = button ? App.busyButton(button, '読んでいます…') : () => {};
    const t0 = performance.now();
    try { const data = await App.api('getHistory', { limit: 200, withCosts: state.showCosts }); log.rows = data.rows || []; log.loaded = true; paintLogTable(); App.perfNote('記録を開く', t0); }
    catch (error) { if (error.code === 'SESSION') App.showLogin(error.message); else toast('記録を読めませんでした: ' + error.message, true); log.loaded = true; }
    finally { done(); }
  }
  /** この画面の待ち時間（中身は持たない。操作の名前とミリ秒だけ・閉じると消える） */
  function perfDetails() {
    const d = el('details', 'perf'); d.append(el('summary', '', 'この画面の待ち時間（新しい順）'));
    d.addEventListener('toggle', () => {
      if (!d.open) return;
      d.querySelector('table')?.remove();
      const t = el('table', 'data-table'); const hr = el('tr'); ['何を', 'ミリ秒', '窓口の中', '大きさ KB'].forEach((h) => hr.append(el('th', '', h))); t.append(hr);
      App.perfLog.slice().reverse().slice(0, 60).forEach((p) => { const tr = el('tr'); [p.what, p.ms, p.server ?? '', p.kb ?? ''].forEach((v) => tr.append(el('td', '', v))); t.append(tr); });
      d.append(t);
    });
    return d;
  }
  App.views.log = {
    show() { renderLog(); loadLog(); },
    reset() { log.rows = []; log.loaded = false; }
  };

  // ------------------------------------------------------------ 設定（パックの値段は本部と PO・ほかは PO だけ）。画面で隠すだけでなく、窓口が役割で通す
  // 上限の行数はシートごとに3段から選ぶ（PO 2026-11-06）。首都圏と北関東はいつも一緒に動くので1つ。値は窓口も同じ一覧で確かめる（Code.gs の CAP_CHOICES）
  const CAP_ROWS = [{ label: '首都圏・北関東', note: 'いつも一緒に動きます', table: '首都圏', tables: ['首都圏', '北関東'], choices: [54, 57, 60] }, { label: '焼き物', note: '', table: '焼き物', tables: ['焼き物'], choices: [48, 51, 54] }];
  const estimate = (table, cap) => { try { return App.S && App.S.estimateScale ? App.S.estimateScale(state.db, table, cap) : null; } catch { return null; } };
  const estText = (e) => (e && e.scale ? '印刷 約 ' + Math.round(e.scale * (e.scale <= 1 ? 100 : 1)) + '%' + (e.namePt ? '・商品名 約 ' + Number(e.namePt).toFixed(1) + 'pt' : '') : '');
  function renderSettings() {
    const root = $('#view-settings'); root.replaceChildren();
    const isPo = state.principal && state.principal.role === 'PO';
    root.append(el('h2', '', '設定'), packSection());
    if (!isPo) return; // ここから下は管理者だけ（窓口も PO のアカウントだけ通す）
    root.append(el('h2', 'set-h-po', '管理者だけの設定'));
    // 1) A3 の番号表の上限の行数
    const s1 = el('section', 'set-sec'); s1.append(el('h3', '', '番号表（A3）の上限の行数'), el('p', 'hint-line', 'シートごとに3つから選びます。印刷は、選んだ行数に合わせて A3 1枚に自動で縮めます。変えたことは記録に残ります。'));
    CAP_ROWS.forEach((row) => {
      const now = Number((C.blockList(state.db, row.table)[0] || {})['上限の行数']);
      const line = el('div', 'set-row'); const name = el('div', 'set-name'); name.append(el('strong', '', row.label)); if (row.note) name.append(el('small', '', row.note));
      const group = el('div', 'seg'); group.setAttribute('role', 'radiogroup'); group.setAttribute('aria-label', row.label + 'の上限の行数');
      row.choices.forEach((cap) => {
        const b = el('button', 'seg-btn' + (cap === now ? ' is-on' : '')); b.type = 'button'; b.setAttribute('role', 'radio'); b.setAttribute('aria-checked', String(cap === now)); b.dataset.cap = String(cap);
        b.append(el('strong', '', cap + ' 行')); const e = estText(estimate(row.table, cap)); if (e) b.append(el('small', '', e));
        b.addEventListener('click', () => saveCap(row, cap, b));
        group.append(b);
      });
      line.append(name, group); s1.append(line);
    });
    // 2) A4 の月間商品（自動）
    const s2 = el('section', 'set-sec'); s2.append(el('h3', '', 'A4 の月間商品（月間特売・新規商品）'), el('p', 'hint-line', '上限は決めず、それぞれ A4 1枚に自動で合わせます（幅いっぱいが基本・行が多いときだけ縮める）。商品名が 8pt より小さくなるときは、空き行を詰め、それでも入らなければアラートで知らせます。'));
    ['月間特売', '新規商品'].forEach((t) => { const e = estText(estimate(t)); if (e) s2.append(el('p', 'set-est', App.monthText(state.month) + 'の' + t + ': ' + e)); });
    // 3) 月を消す（いちばん新しい・非公開の月だけ。窓口も同じ決まりで確かめる）
    const s3 = el('section', 'set-sec'); s3.append(el('h3', '', '月を消す'), el('p', 'hint-line', '試しに作った月などを、まとめて消します。消せるのは、いちばん新しい月で、店に非公開の月だけです。消した月は、データのブックに「消した月_…」のシートとして残ります（元に戻せます）。'));
    const last = state.months[state.months.length - 1];
    const why = state.months.length < 2 ? '月が1つだけなので消せません。' : (state.published && state.published.has(last) ? App.monthText(last) + ' は店に公開中です。先に、その月を開いて「店に公開中」を押し、非公開にしてください。' : '');
    const del = el('button', 'btn btn-danger', last ? App.monthText(last) + ' を消す…' : '消す…'); del.type = 'button'; del.id = 'month-delete'; del.disabled = !last || Boolean(why);
    const ask = el('div', 'set-ask'); ask.hidden = true;
    del.addEventListener('click', () => {
      ask.replaceChildren(el('span', '', App.monthText(last) + ' を消します。よろしいですか？'));
      const yes = el('button', 'btn btn-danger', '消す'); yes.type = 'button'; yes.id = 'month-delete-yes';
      const no = el('button', 'btn', 'やめる'); no.type = 'button'; no.addEventListener('click', () => { ask.hidden = true; del.focus(); });
      yes.addEventListener('click', () => deleteMonth(last, yes));
      ask.append(yes, no); ask.hidden = false; no.focus();
    });
    s3.append(del, ask); if (why) s3.append(el('p', 'hint-line', why));
    root.append(s1, s2, s3);
  }
  // パックの値段（焼き物の表の ≪パック販売用値段特≫。PO 2026-11-07「BY 設定画面で数値を入れる」）。打つと自動で保存・絵はその場で描き直す（Excel と番号表と同じ絵）
  const pack = { rows: null, month: '', timer: 0, saved: '', pending: null };
  const digits = (v) => C.nfkc(String(v ?? '')).replace(/[,，円\s]/g, '');
  function packSection() {
    const sec = el('section', 'set-sec'); sec.id = 'pack-sec';
    if (pack.month !== state.month || !pack.rows) { pack.month = state.month; pack.rows = C.packOf(state.db).map((r) => r.map(String)); pack.saved = C.packText(C.packOf(state.db)); }
    sec.append(el('h3', '', 'パックの値段（焼き物の表の「' + C.PACK_TITLE + '」）'),
      el('p', 'hint-line', App.monthText(state.month) + 'の値段です。打つと自動で保存し、右の絵（番号表と Excel に入る絵）がその場で変わります。次の月を作ると、この値段を引き継ぎます。'));
    const wrap = el('div', 'pack-wrap');
    const left = el('div', 'pack-left');
    const table = el('table', 'pack-edit'); table.id = 'pack-edit';
    const hr = el('tr'); ['', '1個', '', '2個', '3個', ''].forEach((h) => hr.append(el('th', '', h))); const head = el('thead'); head.append(hr); table.append(head);
    const body = el('tbody'); table.append(body);
    const state1 = el('p', 'pack-state'); state1.id = 'pack-state'; state1.setAttribute('aria-live', 'polite');
    const add = el('button', 'btn', '＋ 段を足す'); add.type = 'button'; add.id = 'pack-add';
    const tools = el('div', 'pack-tools'); tools.append(add, state1);
    const fig = el('figure', 'pack-preview'); const img = el('img'); img.id = 'pack-img'; img.alt = ''; fig.append(img, el('figcaption', '', '番号表と Excel に入る絵'));
    const paint = () => {
      body.replaceChildren();
      pack.rows.forEach((r, i) => {
        const tr = el('tr'); tr.append(el('th', '', (i + 1) + '段目'));
        [0, null, 1, 2].forEach((k) => {
          const td = el('td');
          if (k === null) { td.textContent = '⇒'; td.className = 'pack-arrow'; tr.append(td); return; }
          const input = el('input'); input.type = 'text'; input.inputMode = 'numeric'; input.value = r[k]; input.dataset.i = String(i); input.dataset.k = String(k);
          input.setAttribute('aria-label', (i + 1) + '段目の' + (k + 1) + '個の値段（円）');
          input.addEventListener('input', () => { pack.rows[i][k] = input.value; changed(false); });
          input.addEventListener('change', () => changed(true));
          td.append(input, el('span', 'pack-yen', '円')); tr.append(td);
        });
        const tdx = el('td'); const del = el('button', 'btn btn-small', '消す'); del.type = 'button'; del.disabled = pack.rows.length <= 1; del.setAttribute('aria-label', (i + 1) + '段目を消す');
        del.addEventListener('click', () => { pack.rows.splice(i, 1); paint(); changed(true); }); tdx.append(del); tr.append(tdx);
        body.append(tr);
      });
      add.disabled = pack.rows.length >= C.PACK_MAX_ROWS;
      add.title = add.disabled ? C.PACK_MAX_ROWS + ' 段までです' : '';
      check();
    };
    // 確かめて、絵を描き直す。よければ保存の予約（now＝欄を離れた・段を消したときは、すぐ）
    const check = () => {
      const nums = pack.rows.map((r) => r.map((v) => { const d = digits(v); return /^[0-9]+$/.test(d) ? Number(d) : NaN; }));
      body.querySelectorAll('input').forEach((x) => { const v = nums[Number(x.dataset.i)][Number(x.dataset.k)]; x.classList.toggle('is-bad', !(Number.isInteger(v) && v >= 1 && v <= C.PACK_MAX_PRICE)); });
      const why = C.packProblem(nums);
      const drawn = !why && window.SouzaiPack ? window.SouzaiPack.draw(nums) : null;
      if (drawn) img.src = drawn.url;
      return why ? { why } : { nums, text: C.packText(nums) };
    };
    const changed = (now) => {
      const r = check();
      window.clearTimeout(pack.timer);
      if (r.why) { state1.className = 'pack-state is-bad'; state1.textContent = r.why + '（入れ終わると保存します）'; return; }
      if (r.text === pack.saved) { state1.className = 'pack-state is-ok'; state1.textContent = '✓ 保存済み（この値段で Excel を作ります）'; return; }
      state1.className = 'pack-state'; state1.textContent = '打ち終わると保存します…';
      pack.pending = { month: pack.month, nums: r.nums, text: r.text, note: state1 };
      pack.timer = window.setTimeout(flushPack, now ? 0 : 700);
    };
    add.addEventListener('click', () => { pack.rows.push(['', '', '']); paint(); const last = body.querySelector('tr:last-child input'); if (last) last.focus(); changed(true); });
    left.append(table, tools); wrap.append(left, fig); sec.append(wrap);
    paint(); changed(false);
    return sec;
  }
  /** 元に戻す・やり直す: 値段を記録の値にして保存し、設定の画面を開いていれば描き直す */
  async function applyPack(month, text, what) {
    if (month !== state.month) throw new Error('開いている月が替わったので、パックの値段は戻せません');
    const nums = C.parsePack(text) || C.packOf(state.db);
    pack.month = month; pack.rows = nums.map((r) => r.map(String)); window.clearTimeout(pack.timer); pack.pending = null;
    pack.saved = C.packText(nums); // 描き直しが、同じ値の保存をもう一度予約しないように（記録は積まない）
    if (state.view === 'settings') renderSettings();
    await savePack(month, nums, C.packText(nums), $('#pack-state') || el('p'), true);
    toast(what + ': パックの値段（' + App.monthText(month) + '）');
  }
  function flushPack() { window.clearTimeout(pack.timer); pack.timer = 0; const p = pack.pending; pack.pending = null; if (p) savePack(p.month, p.nums, p.text, p.note); }
  async function savePack(month, nums, text, note, noHistory) {
    const t0 = performance.now();
    const beforeText = month === pack.month ? pack.saved : '';
    note.className = 'pack-state is-busy'; note.replaceChildren(App.spinner ? App.spinner() : '', document.createTextNode(' 保存しています…'));
    try {
      await App.serial(() => App.api('savePack', { month, rows: nums, requestId: App.rid() }));
      if (month === state.month) { state.db.status[C.PACK_KIND] = text; App.refresh([]); }
      if (month === pack.month) pack.saved = text; App.perfNote('パックの値段を保存する', t0);
      // 元に戻す・やり直すの記録（続けて打った分は1回）
      if (!noHistory && beforeText && beforeText !== text) App.remember({ label: 'パックの値段（' + App.monthText(month) + '）', mergeKey: 'pack|' + month, undo: () => applyPack(month, beforeText, '元に戻しました'), redo: () => applyPack(month, text, 'やり直しました') });
      if (C.packText(pack.rows.map((r) => r.map((v) => Number(digits(v))))) === text) { note.className = 'pack-state is-ok'; note.textContent = '✓ 保存しました（番号表と Excel の絵が変わりました）'; }
    } catch (error) {
      note.className = 'pack-state is-bad'; note.textContent = '保存できませんでした: ' + error.message;
      if (error.code === 'SESSION') App.showLogin(error.message);
    }
  }
  async function saveCap(row, cap, b) {
    if (b.classList.contains('is-on')) return;
    const done = App.busyButton(b, '保存しています…'); const t0 = performance.now();
    try {
      await App.serial(() => App.api('saveCaps', { table: row.table, cap, requestId: App.rid() }));
      row.tables.forEach((t) => C.blockList(state.db, t).forEach((x) => { x['上限の行数'] = cap; }));
      App.refresh([]); App.perfNote('上限の行数を変える', t0);
      toast(row.label + ' の上限の行数を ' + cap + ' 行にしました。');
    } catch (error) { toast('保存できませんでした: ' + error.message, true); }
    finally { done(); renderSettings(); }
  }
  async function deleteMonth(month, b) {
    const done = App.busyButton(b, '消しています…'); const t0 = performance.now();
    try {
      await App.serial(() => App.api('deleteMonth', { month, requestId: App.rid() }));
      App.perfNote('月を消す', t0);
      const back = state.months.filter((m) => m !== month);
      toast(App.monthText(month) + ' を消しました（データのブックに控えのシートが残っています）。');
      await App.loadMonth(back[back.length - 1]);
    } catch (error) { toast('消せませんでした: ' + error.message, true); done(); }
  }
  App.views.settings = { show() { renderSettings(); }, reset() { flushPack(); pack.rows = null; pack.month = ''; } }; // 月を替える前に、打ちかけの値段を元の月へ保存する
})();
