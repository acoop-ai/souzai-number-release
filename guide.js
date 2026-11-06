'use strict';
// ご利用ガイド。言葉（本文）はログインの後に窓口から受け取る（公開の画面に本文を置かない）。図は画面の側に置き、窓口は節ごとに図の名前だけを返す。
//   節 = { id, title, fig: 図の名前, steps: [{ h: 見出し, t: 補い }]（図の番号の丸 ①②… の説明・順番が番号）, items: [短い文] }
//   図の名前は、下の FIGS（図解）と guide-img/figs.js（写真と丸の位置・tests/guide_shots.js が作る）にある物だけを出す。無い名前は図なしで言葉だけ出す。
//   図解は styles.css の部品（札・行の色・保存の印・ボタン）をそのまま使う＝色と名前は画面と同じ。窓口の文字は全部 textContent（HTML として入れない）。
//   型: 左に目次・右に本文・双方向に連動。目次を押した見出しは、自分でスクロールするまで優先する。
(function () {
  const CONFIG = window.APP_CONFIG || {};
  const IS_DEV = CONFIG.dev === true && /^(127\.0\.0\.1|localhost)$/.test(location.hostname);
  const GAS_URL = String(CONFIG.gasUrl || '');
  const OK = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(GAS_URL) || (IS_DEV && GAS_URL === '/__gas');
  const SHOTS = window.GUIDE_SHOTS || {};
  const own = (o, k) => typeof k === 'string' && Object.prototype.hasOwnProperty.call(o, k);
  const el = (t, c, x) => { const n = document.createElement(t); if (c) n.className = c; if (x !== undefined && x !== null) n.textContent = String(x); return n; };
  const status = document.getElementById('guide-status');
  const scroller = document.getElementById('guide-scroll');

  // ------------------------------------------------------------ 図解（写真でない図）。見本の字は画面の部品の名前と同じにする
  // 販売形態ごとの色: [販売形態, 一覧の行の印（styles.css の .lr.f-*）, 一覧の色の名前, 番号表・Excel の見本（ink＝ピンクの字・fill＝水色のセル・cut・off）, 番号表・Excel の説明]
  const FORM_ROWS = [
    ['定番', 'f-reg', '薄い緑', '', 'そのまま'],
    ['月間特売', 'f-sale', '薄い赤', 'ink', 'ピンクの字'],
    ['新商品', 'f-new', '薄い水色', 'fill', '水色のセル'],
    ['月間特売・新商品', 'f-both', '行は薄い赤・名前は薄い水色', 'ink fill', 'ピンクの字・水色のセル'],
    ['カット商品', 'f-cut', '薄い紫', 'cut', '焼き物の表に載る'],
    ['休み', 'f-rest', '白', 'off', '載らない'],
    ['終売', 'f-off', '灰色', 'off', '載らない（番号の管理に残る）'],
    ['計画', 'f-off', '灰色', 'off', '載らない（番号の管理に残る）']
  ];
  // flow＝矢印でつなぐ箱・cards＝並べるだけの箱。ui＝箱の上に置く画面の部品の見本 [種類, 字]
  const FIGS = {
    'flow-month': { kind: 'flow', ui: [['quiet', '＋ 月を足す'], ['tab', '一覧'], ['sheettab', '月間特売'], ['state', '店に公開'], ['btn', 'Excel']] },
    'flow-cut': { kind: 'flow', ui: [['chip', 'カット商品'], ['sec', '≪カット商品≫'], ['chip', '定番'], ['chip', '休み']] },
    'flow-excel': { kind: 'flow', ui: [['btn', 'Excel'], ['file', '商品呼出番号表_2026年10月.xlsx'], ['sheets', ''], ['paper', '']] },
    'cards-save': { kind: 'cards', ui: [['save', '✓ 自動保存'], ['saving', '保存中…'], ['saveerr', '⚠ 保存できていません（自動で送り直します）']] },
    'cards-ask': { kind: 'flow', ui: [['cells', '9003|【見本】ちらし寿司'], ['cells', '定番価格|100 → 120'], ['who', '本部']] },
    'forms': { kind: 'forms', rows: FORM_ROWS },
    'forms-view': { kind: 'forms', rows: FORM_ROWS.slice(0, 5) }
  };
  function figOf(name) {
    if (own(FIGS, name)) return FIGS[name];
    if (own(SHOTS, name) && /^guide-img\/[a-z0-9-]+\.png(\?v=[0-9a-f]+)?$/.test(String(SHOTS[name].src))) return Object.assign({ kind: 'shot' }, SHOTS[name]);
    return null;
  }

  // 画面の部品の見本（styles.css の部品をそのまま使う。押せない見本なので button にしない）
  function ui(kind, text) {
    const bar = (child) => { const b = el('span', 'gui-bar'); b.append(child); return b; };
    switch (kind) {
      case 'quiet': return bar(el('span', 'btn btn-quiet btn-small', text));
      case 'btn': return bar(el('span', 'btn', text));
      case 'tab': { const b = el('span', 'gui-bar gui-tabs'); b.append(el('span', 'is-active', text), el('span', '', '番号表'), el('span', '', '記録')); return b; }
      case 'alert': { const c = el('span', 'alert-chip is-zero'); c.textContent = text; return bar(c); }
      case 'sheettab': case 'sheets': { // sheettab＝その1枚だけ（選んだ形）・sheets＝4枚
        const b = el('span', 'gui-sheettabs');
        [['月間特売', 't-mon'], ['新規商品', 't-mon'], ['首都圏', 't-shu'], ['北関東', 't-kita'], ['焼き物', 't-yaki']].filter(([t]) => kind === 'sheets' || t === text).forEach(([t, c]) => b.append(el('span', c + (kind === 'sheettab' ? ' is-active' : ''), t)));
        return b;
      }
      case 'chip': { const c = el('span', 'chip', text); c.dataset.chip = text; return c; }
      case 'sec': return el('span', 'gui-sec', text);
      case 'file': { const b = el('span', 'gui-file'); b.append(el('i', '', 'X'), document.createTextNode(text)); return b; }
      case 'paper': { const b = el('span', 'gui-paper'); for (let i = 0; i < 5; i++) b.append(el('i')); return b; }
      case 'save': return bar(el('span', 'save-state', text));
      case 'saving': { const s = el('span', 'save-state'); s.append(el('span', 'spinner'), document.createTextNode(' ' + text)); return bar(s); }
      case 'saveerr': return bar(el('span', 'save-state is-error', text));
      case 'cells': { const b = el('span', 'gui-cells'); String(text).split('|').forEach((t) => b.append(el('span', '', t))); return b; }
      case 'who': return el('span', 'gui-who', text);
      case 'state': case 'state-off': { const b = el('span', 'gui-bar'); b.append(el('strong', '', '2026年11月'), el('span', 'gui-state' + (kind === 'state-off' ? ' is-off' : ''), text)); return b; } // 月と、店に見えるかどうか
      default: return el('span', '', text);
    }
  }
  function num(n) { const c = el('span', 'gnum', n); c.setAttribute('aria-hidden', 'true'); return c; }

  // 流れの箱（番号・部品の見本・見出し・補い）
  function flowFig(f, steps) {
    const box = el('div', f.kind === 'flow' ? 'gflow' : 'gflow gcards');
    const list = el('ol', 'gflow-list');
    steps.forEach((s, i) => {
      const li = el('li', 'gflow-step'); li.append(num(i + 1));
      const u = f.ui && f.ui[i]; if (u) { const w = el('span', 'gflow-ui'); w.setAttribute('aria-hidden', 'true'); w.append(ui(u[0], u[1])); li.append(w); }
      li.append(el('strong', '', s.h)); if (s.t) li.append(el('span', 'gflow-t', s.t));
      list.append(li);
    });
    box.append(list);
    return box;
  }
  // 販売形態と色の見本（一覧の行の色は styles.css の .lr.f-* をそのまま使う）
  function formsFig(f) {
    const wrap = el('div', 'gforms-wrap');
    const t = el('table', 'gforms');
    const head = el('tr'); ['販売形態', '一覧（行の色）', '番号表・Excel'].forEach((h) => { const th = el('th', '', h); th.scope = 'col'; head.append(th); });
    const thead = el('thead'); thead.append(head); t.append(thead);
    const body = el('tbody');
    f.rows.forEach(([form, cls, colorName, sheet, sheetText]) => {
      const tr = el('tr');
      const c1 = el('td'); form.split('・').forEach((x) => { const c = el('span', 'chip', x); c.dataset.chip = x; c1.append(c); });
      const c2 = el('td'); const row = el('span', 'gm lr ' + cls); row.setAttribute('aria-hidden', 'true'); row.append(el('span', 'gm-call', '9001'), el('span', 'gm-name', '【見本】商品')); c2.append(row, el('span', 'gm-text', colorName));
      const c3 = el('td');
      if (sheet === 'off') c3.append(el('span', 'goff', sheetText));
      else if (sheet === 'cut') { c3.append(el('span', 'gui-sec', '≪カット商品≫'), el('span', 'gm-text', sheetText)); }
      else { const s = el('span', 'gsheet' + sheet.split(' ').filter(Boolean).map((k) => ' is-' + k).join(''), '【見本】商品'); s.setAttribute('aria-hidden', 'true'); c3.append(s, el('span', 'gm-text', sheetText)); }
      tr.append(c1, c2, c3); body.append(tr);
    });
    t.append(body); wrap.append(t);
    return wrap;
  }
  // 写真＋番号の丸（丸の位置は％なので、どの大きさでもずれない）。丸か手順にさわると、同じ番号の枠が光る
  function shotFig(f) {
    const fig = el('figure', 'gshot'); fig.style.setProperty('--gw', f.w + 'px');
    const scroll = el('div', 'gshot-scroll');
    const frame = el('div', 'gshot-frame');
    const img = el('img'); img.alt = String(f.alt || ''); img.width = f.w; img.height = f.h; img.loading = 'lazy'; img.decoding = 'async'; img.src = f.src;
    frame.append(img);
    (f.marks || []).forEach((m, i) => {
      const n = String(i + 1);
      const b = el('span', 'gbox'); b.dataset.n = n; b.style.left = m[2] + '%'; b.style.top = m[3] + '%'; b.style.width = m[4] + '%'; b.style.height = m[5] + '%';
      const c = num(n); c.classList.add('gmark'); c.dataset.n = n; c.style.left = m[0] + '%'; c.style.top = m[1] + '%';
      frame.append(b, c);
    });
    scroll.append(frame); fig.append(scroll);
    const cap = el('figcaption', 'gcap'); cap.append(el('span', '', '見本の写真（作り物のデータ）'));
    if (f.w > 640) {
      const z = el('button', 'btn btn-small gzoom', '大きく見る'); z.type = 'button'; z.setAttribute('aria-pressed', 'false');
      z.addEventListener('click', () => { const on = fig.classList.toggle('is-zoom'); z.textContent = on ? '画面の幅に戻す' : '大きく見る'; z.setAttribute('aria-pressed', String(on)); });
      cap.append(z);
    }
    fig.append(cap);
    return fig;
  }
  function stepList(steps) {
    const ol = el('ol', 'gsteps');
    steps.forEach((s, i) => { const li = el('li'); li.dataset.n = String(i + 1); li.append(num(i + 1), el('strong', '', s.h)); if (s.t) li.append(el('span', 'gstep-t', s.t)); ol.append(li); });
    return ol;
  }
  function linkHot(sec) {
    const set = (n) => sec.querySelectorAll('[data-n]').forEach((x) => x.classList.toggle('is-hot', x.dataset.n === n));
    sec.addEventListener('mouseover', (e) => { const t = e.target.closest('[data-n]'); set(t ? t.dataset.n : ''); });
    sec.addEventListener('mouseleave', () => set(''));
    sec.addEventListener('click', (e) => { const t = e.target.closest('.gmark, .gsteps li'); if (t) set(t.dataset.n); });
  }

  // ------------------------------------------------------------ 目次と本文の連動
  let pinned = '';
  const toc = document.getElementById('guide-toc');
  const narrow = window.matchMedia('(max-width: 760px)');
  const topGap = () => (narrow.matches ? toc.offsetHeight : 0); // スマホでは目次の帯が上に残る分
  function setCurrent(id) {
    document.querySelectorAll('#guide-toc a').forEach((a) => {
      const on = a.dataset.id === id; a.classList.toggle('is-current', on);
      if (on) { a.setAttribute('aria-current', 'true'); if (toc.scrollWidth > toc.clientWidth + 1) toc.scrollLeft = a.offsetLeft - (toc.clientWidth - a.offsetWidth) / 2; } else a.removeAttribute('aria-current');
    });
  }
  function spy() {
    if (pinned) { setCurrent(pinned); return; }
    const secs = [...document.querySelectorAll('.guide-section')]; if (!secs.length) return;
    const top = scroller.getBoundingClientRect().top + topGap() + 40; let cur = secs[0];
    secs.forEach((s) => { if (s.getBoundingClientRect().top <= top) cur = s; });
    if (scroller.scrollTop + scroller.clientHeight >= scroller.scrollHeight - 2) cur = secs[secs.length - 1];
    setCurrent(cur.id.replace('g-', ''));
  }
  // 画面の外の節は描いていない（大きさが仮）ので、着いてから描いた分のずれを、少し後に直して着地させる（自分で送ったらやめる）
  function jump(id) {
    const s = document.getElementById('g-' + id); if (!s) return;
    pinned = id; setCurrent(id);
    const land = () => { const d = s.getBoundingClientRect().top - scroller.getBoundingClientRect().top - topGap() - 8; if (Math.abs(d) > 1) scroller.scrollTop += d; };
    land(); [60, 180, 400].forEach((ms) => window.setTimeout(() => { if (pinned === id) land(); }, ms));
    s.classList.remove('is-landing'); void s.offsetWidth; s.classList.add('is-landing');
  }
  ['wheel', 'touchstart', 'keydown', 'mousedown'].forEach((ev) => scroller.addEventListener(ev, () => { pinned = ''; }, { passive: true }));
  scroller.addEventListener('scroll', spy, { passive: true });
  // 印刷のときは、まだ読んでいない写真も読む
  window.addEventListener('beforeprint', () => document.querySelectorAll('.gshot img').forEach((i) => { i.loading = 'eager'; }));

  const textList = (a) => (Array.isArray(a) ? a.filter((x) => typeof x === 'string' && x) : []);
  const stepsOf = (a) => (Array.isArray(a) ? a.filter((s) => s && typeof s.h === 'string').map((s) => ({ h: s.h, t: typeof s.t === 'string' ? s.t : '' })) : []);
  function render(g) {
    document.title = String(g.title || document.title);
    const sections = Array.isArray(g.sections) ? g.sections.slice() : [];
    if (IS_DEV) sections.unshift({ id: 'proto', title: 'この試作について', items: ['この試作は手元だけで動いていて、本物のシートには書きません。', '右上のメニューの「見本をはじめの形に戻す」で、はじめの形に戻ります。'] });
    const box = document.getElementById('guide-content');
    sections.forEach((s) => {
      const id = String(s.id || '').replace(/[^a-z0-9-]/gi, ''); if (!id) return;
      const a = el('a', '', s.title); a.href = '#g-' + id; a.dataset.id = id; a.addEventListener('click', (e) => { e.preventDefault(); jump(id); }); toc.append(a);
      const sec = el('section', 'guide-section'); sec.id = 'g-' + id; sec.append(el('h2', '', s.title));
      const steps = stepsOf(s.steps), fig = figOf(s.fig);
      if (fig && (fig.kind === 'flow' || fig.kind === 'cards')) sec.append(flowFig(fig, steps));
      else if (fig && fig.kind === 'forms') { sec.append(formsFig(fig)); if (steps.length) sec.append(stepList(steps)); }
      else if (fig && fig.kind === 'shot') { const pair = el('div', 'gpair' + (fig.w <= 640 || (fig.w <= 1000 && fig.w / fig.h < 1.6) ? ' is-side' : '')); pair.append(shotFig(fig)); if (steps.length) pair.append(stepList(steps)); sec.append(pair); linkHot(sec); }
      else if (steps.length) sec.append(stepList(steps));
      const items = textList(s.items);
      if (items.length) { const ul = el('ul', 'gnotes'); items.forEach((t) => ul.append(el('li', '', t))); sec.append(ul); }
      box.append(sec);
    });
    status.hidden = true; document.getElementById('guide-root').hidden = false; spy();
  }
  function session() { try { const o = JSON.parse(localStorage.getItem('souzai-number-session-v1') || 'null'); return o && o.t && Date.now() < o.e ? o.t : ''; } catch { return ''; } }
  function askLogin(msg) { status.replaceChildren(document.createTextNode(msg + ' '), Object.assign(el('a', '', '画面でログインする'), { href: 'index.html' })); }
  async function start() {
    if (!OK) { status.textContent = '窓口の設定が終わると、ガイドを表示します。'; return; }
    const s = session(); if (!s) { askLogin('ログインすると、ガイドを表示します。'); return; }
    try {
      const res = await fetch(GAS_URL, { method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'getGuide', data: {}, session: s }) });
      const out = await res.json();
      if (!out || out.ok !== true) { askLogin(out && out.message ? out.message : '読めませんでした。'); return; }
      render(out.data);
    } catch { status.textContent = '読み込めませんでした。通信を確かめて、開き直してください。'; }
  }
  window.addEventListener('DOMContentLoaded', start);
})();
