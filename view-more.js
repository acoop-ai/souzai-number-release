(() => {
  'use strict';
  // 記録のタブ（だれが・いつ・何を。前の値に戻す）と、設定（PO だけ・上限の行数）。
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
  function renderSettings() {
    const root = $('#view-settings'); root.replaceChildren();
    root.append(el('h2', '', '設定（管理者だけ）'));
    root.append(el('p', 'hint-line', '上限の行数を変えられます。変えたことは記録に残ります。'));
    const t = el('table', 'data-table settings-table'); const hr = el('tr'); ['表', '列', '上限の行数'].forEach((h) => hr.append(el('th', '', h))); t.append(hr);
    C.TABLE_KEYS.forEach((table) => {
      C.blockList(state.db, table).forEach((b) => {
        const tr = el('tr'); tr.append(el('td', '', table), el('td', '', b['まとまり'] + '列目'));
        const td = el('td'); const i = el('input'); i.type = 'number'; i.min = '1'; i.max = '200'; i.value = b['上限の行数']; i.setAttribute('aria-label', table + ' ' + b['まとまり'] + '列目の上限');
        i.addEventListener('change', async () => {
          const cap = Number(i.value); if (!Number.isInteger(cap) || cap < 1 || cap > 200) { toast('上限の行数は 1 から 200 です。', true); i.value = b['上限の行数']; return; }
          try { await App.serial(() => App.api('saveSettings', { table, block: String(b['まとまり']), cap, requestId: App.rid() })); b['上限の行数'] = cap; App.refresh([]); toast('上限の行数を保存しました。'); }
          catch (error) { i.value = b['上限の行数']; toast('保存できませんでした: ' + error.message, true); }
        });
        td.append(i); tr.append(td); t.append(tr);
      });
    });
    root.append(t);
    root.append(el('p', 'hint-line', '最初の取り込み（実データをスプレッドシートへ入れる）は、本物にする段で行います。'));
  }
  App.views.settings = { show() { renderSettings(); } };
})();
