(() => {
  'use strict';
  // 記録のタブ（だれが・いつ・何を。前の値に戻す）と、設定（PO だけ・上限の行数の3段・A4 の見立て・月を消す）。
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

  // ------------------------------------------------------------ 設定（PO だけ）。画面で隠すだけでなく、窓口が PO のアカウントだけ通す
  // 上限の行数はシートごとに3段から選ぶ（PO 2026-11-06）。首都圏と北関東はいつも一緒に動くので1つ。値は窓口も同じ一覧で確かめる（Code.gs の CAP_CHOICES）
  const CAP_ROWS = [{ label: '首都圏・北関東', note: 'いつも一緒に動きます', table: '首都圏', tables: ['首都圏', '北関東'], choices: [54, 57, 60] }, { label: '焼き物', note: '', table: '焼き物', tables: ['焼き物'], choices: [48, 51, 54] }];
  const estimate = (table, cap) => { try { return App.S && App.S.estimateScale ? App.S.estimateScale(state.db, table, cap) : null; } catch { return null; } };
  const estText = (e) => (e && e.scale ? '印刷 約 ' + Math.round(e.scale * (e.scale <= 1 ? 100 : 1)) + '%' + (e.namePt ? '・商品名 約 ' + Number(e.namePt).toFixed(1) + 'pt' : '') : '');
  function renderSettings() {
    const root = $('#view-settings'); root.replaceChildren();
    root.append(el('h2', '', '設定（管理者だけ）'));
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
  App.views.settings = { show() { renderSettings(); } };
})();
