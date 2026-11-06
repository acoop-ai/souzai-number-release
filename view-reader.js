(() => {
  'use strict';
  // 見るだけの人の一覧: 入力欄でなく文字の表。窓口が返した物（その月の表に載る行・表に出る価格）だけを描く。守りは窓口（画面で隠すだけにしない）。
  //   文字は全部 esc を通す。見出しと左の2列は動かない（一覧と同じ枠の形）。
  const App = window.SouzaiApp;
  const { state, $, el, esc, C, SN } = App;
  const CHIPS = ['全部', '定番', '月間特売', '新商品', 'カット商品'];
  const view = { built: false, rows: [], els: new Map(), chip: '全部', q: '', terms: null, sec: '', keys: new Map() };
  const db = () => state.db;
  const has = (form, chip) => form === chip || (form === C.FORM_SALE_NEW && (chip === C.FORM_SALE || chip === C.FORM_NEW));
  function secOrder() {
    const out = [];
    C.KINDS.forEach((kind) => { const t = C.KIND_TABLES[kind][0]; C.blockList(db(), t).forEach((b) => { C.placesOfBlock(db(), kind, b['まとまり']).forEach((p) => out.push(String(p['区分ID']))); }); });
    return out;
  }
  /** 1行の見え方: 載っている表ごとの販売形態と価格。首都圏と北関東で違えば添える */
  function lineOf(r) {
    const p = db().products.get(String(r['商品ID'])); if (!p) return null;
    const tables = (C.KIND_TABLES[p['表の種類']] || []).filter((t) => C.rowShown(db(), r, t));
    if (!tables.length) return null;
    const main = tables[0]; const form = C.formOf(r, main); const price = C.showCell(C.shownPrice(p, r, main));
    const other = tables[1];
    const otherForm = other && C.formOf(r, other) !== form ? C.formOf(r, other) : '';
    const otherPrice = other && C.showCell(C.shownPrice(p, r, other)) !== price ? C.showCell(C.shownPrice(p, r, other)) : '';
    return { r, p, form, price, otherForm, otherPrice, forms: other ? [form, C.formOf(r, other)] : [form] };
  }
  function keyOf(line) {
    const pid = String(line.p['商品ID']); if (view.keys.has(pid)) return view.keys.get(pid);
    const sec = db().sections.get(String(line.r['区分ID']));
    const t = [line.p['名前'], line.p['呼出番'], line.p['小パック'], line.p['指定トレー'], sec ? sec['表題'] : ''].filter((x) => !C.isBlank(x)).join(' ');
    const k = SN ? SN.key(t) : C.nfkc(t).toLowerCase(); view.keys.set(pid, k); return k;
  }
  function matches(line) {
    if (view.sec && String(line.r['区分ID']) !== view.sec) return false;
    if (view.chip !== '全部' && !line.forms.some((f) => has(f, view.chip))) return false;
    if (view.terms && view.terms.length) { const k = keyOf(line); if (!(SN ? SN.hit(k, view.terms) : view.terms.every((x) => k.includes(x)))) return false; }
    return true;
  }
  function paintChips() {
    document.querySelectorAll('#chips .chip').forEach((b) => {
      const c = b.dataset.chip; const n = c === '全部' ? view.rows.length : view.rows.filter((l) => l.forms.some((f) => has(f, c))).length;
      b.classList.toggle('is-active', c === view.chip); b.setAttribute('aria-pressed', String(c === view.chip)); b.replaceChildren(document.createTextNode(c + ' '), el('b', '', n));
    });
  }
  function applyFilter() {
    view.rows.forEach((line) => { const tr = view.els.get(String(line.r['行ID'])); if (tr) tr.hidden = !matches(line); });
    paintChips();
  }
  function rowHtml(line) {
    const { r, p } = line; const sec = db().sections.get(String(r['区分ID'])); const title = sec ? sec['表題'] : '';
    const reg = p['地域'] && p['地域'] !== '共通' ? '<span class="tag">' + esc(p['地域']) + '</span>' : '';
    return '<tr class="lr ' + (App.FORM_CLASS[line.form] || 'f-rest') + '" data-rid="' + esc(r['行ID']) + '" data-pid="' + esc(p['商品ID']) + '">'
      + '<td class="c-call">' + esc(C.showCell(p['呼出番'])) + '</td>'
      + '<td class="c-name"><span class="nm"></span>' + reg + '</td>'
      + '<td class="c-form">' + esc(line.form) + (line.otherForm ? '<small>北関東は' + esc(line.otherForm) + '</small>' : '') + '</td>'
      + '<td class="c-rprice num">' + esc(line.price) + (line.otherPrice ? '<small>北関東 ' + esc(line.otherPrice) + '</small>' : '') + '</td>'
      + '<td class="c-tray-text">' + esc(C.showCell(p['指定トレー'])) + '</td>'
      + '<td class="c-small">' + esc(C.showCell(p['小パック'])) + '</td>'
      + '<td class="c-sec"><span class="c-sec-text" title="' + esc(title) + '">' + esc(title) + '</span></td>'
      + '</tr>';
  }
  function build() {
    const t0 = performance.now();
    const root = $('#view-list'); root.replaceChildren();
    const bar = el('div', 'list-bar');
    const search = el('div', 'search-box'); const input = el('input'); input.id = 'list-search'; input.type = 'search'; input.placeholder = '名前・番号で探す（ひらがな・ローマ字も）'; input.setAttribute('aria-label', '商品を探す'); input.value = view.q; search.append(input);
    const chips = el('div', 'chips'); chips.id = 'chips'; chips.setAttribute('role', 'group'); chips.setAttribute('aria-label', '絞り込み');
    CHIPS.forEach((c) => { const b = el('button', 'chip'); b.type = 'button'; b.dataset.chip = c; chips.append(b); });
    const sel = el('select'); sel.id = 'list-section'; sel.setAttribute('aria-label', '区分で絞る');
    const o0 = el('option', '', '区分: 全部'); o0.value = ''; sel.append(o0);
    secOrder().forEach((id) => { const s = db().sections.get(id); if (!s) return; const o = el('option', '', C.nfkc(s['表題'])); o.value = id; sel.append(o); });
    sel.value = view.sec;
    bar.append(search, chips, sel);
    const hint = el('p', 'list-hint', 'いまの月の、番号表に載る商品です。価格は番号表に出る価格です。');
    const scroll = el('div', 'list-scroll'); scroll.id = 'list-scroll'; scroll.tabIndex = -1;
    const memo = C.rowsBySection(db()); view.rows = [];
    secOrder().forEach((secId) => { (memo.get(secId) || []).forEach((r) => { if (!r['商品ID']) return; const line = lineOf(r); if (line) view.rows.push(line); }); });
    scroll.innerHTML = '<table class="list-table reader-table" id="list-table"><thead><tr><th class="c-call" scope="col">番号</th><th class="c-name" scope="col">名前</th><th class="c-form" scope="col">販売形態</th><th class="c-rprice" scope="col">価格</th><th class="c-tray-text" scope="col">包材</th><th class="c-small" scope="col">小パック</th><th class="c-sec" scope="col">区分</th></tr></thead><tbody id="list-body">' + view.rows.map(rowHtml).join('') + '</tbody></table>';
    root.append(bar, hint, scroll);
    view.els = new Map();
    scroll.querySelectorAll('tr.lr').forEach((tr) => view.els.set(tr.dataset.rid, tr));
    view.rows.forEach((line) => {
      const tr = view.els.get(String(line.r['行ID'])); const text = C.nfkc(line.p['名前']).trim() === '' ? '（名前なし）' : String(line.p['名前']);
      App.paintRunsInto(tr.querySelector('.nm'), text, (line.p._deco && line.p._deco['名前']) || []);
    });
    view.built = true;
    chips.addEventListener('click', (e) => { const b = e.target.closest('.chip'); if (b) { view.chip = b.dataset.chip; applyFilter(); scroll.scrollTop = 0; } });
    sel.addEventListener('change', () => { view.sec = sel.value; applyFilter(); scroll.scrollTop = 0; });
    let timer = 0;
    input.addEventListener('input', () => { window.clearTimeout(timer); const v = input.value; timer = window.setTimeout(() => { view.q = v; view.terms = v.trim() ? (SN ? SN.terms(v) : C.nfkc(v).toLowerCase().split(/\s+/).filter(Boolean)) : null; applyFilter(); scroll.scrollTop = 0; }, 120); });
    applyFilter();
    App.perfNote('見るだけの一覧を作る', t0);
  }
  App.views.reader = {
    show() { if (!state.db) return; if (!view.built) build(); },
    reset() { view.built = false; view.keys = new Map(); view.q = ''; view.terms = null; view.sec = ''; view.chip = '全部'; }
  };
})();
