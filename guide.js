'use strict';
// ご利用ガイド。本文はログインの後に窓口から受け取る（公開の画面に本文を置かない）。文字は全部 textContent（窓口の文字を HTML として入れない）。
//   型: 左に目次・右に本文・双方向に連動。目次を押した見出しは、自分でスクロールするまで優先する。
(function () {
  const CONFIG = window.APP_CONFIG || {};
  const IS_DEV = CONFIG.dev === true && /^(127\.0\.0\.1|localhost)$/.test(location.hostname);
  const GAS_URL = String(CONFIG.gasUrl || '');
  const OK = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(GAS_URL) || (IS_DEV && GAS_URL === '/__gas');
  const el = (t, c, x) => { const n = document.createElement(t); if (c) n.className = c; if (x !== undefined && x !== null) n.textContent = String(x); return n; };
  const status = document.getElementById('guide-status');
  const scroller = document.getElementById('guide-scroll');
  function session() { try { const o = JSON.parse(localStorage.getItem('souzai-number-session-v1') || 'null'); return o && o.t && Date.now() < o.e ? o.t : ''; } catch { return ''; } }
  function askLogin(msg) { status.replaceChildren(document.createTextNode(msg + ' '), Object.assign(el('a', '', '画面でログインする'), { href: 'index.html' })); }
  let pinned = '';
  function setCurrent(id) { document.querySelectorAll('#guide-toc a').forEach((a) => { const on = a.dataset.id === id; a.classList.toggle('is-current', on); if (on) a.setAttribute('aria-current', 'true'); else a.removeAttribute('aria-current'); }); }
  function spy() {
    if (pinned) { setCurrent(pinned); return; }
    const secs = [...document.querySelectorAll('.guide-section')]; if (!secs.length) return;
    const top = scroller.getBoundingClientRect().top + 40; let cur = secs[0];
    secs.forEach((s) => { if (s.getBoundingClientRect().top <= top) cur = s; });
    if (scroller.scrollTop + scroller.clientHeight >= scroller.scrollHeight - 2) cur = secs[secs.length - 1];
    setCurrent(cur.id.replace('g-', ''));
  }
  function jump(id) {
    const s = document.getElementById('g-' + id); if (!s) return;
    pinned = id; setCurrent(id);
    scroller.scrollTop += s.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 8;
    s.classList.remove('is-landing'); void s.offsetWidth; s.classList.add('is-landing');
  }
  ['wheel', 'touchstart', 'keydown', 'mousedown'].forEach((ev) => scroller.addEventListener(ev, () => { pinned = ''; }, { passive: true }));
  scroller.addEventListener('scroll', spy, { passive: true });
  function render(g) {
    document.title = g.title;
    if (IS_DEV) g.sections = [{ id: 'proto', title: 'この試作について', items: ['この試作は手元だけで動いていて、本物のシートには書きません。', '右上のメニューの「見本をはじめの形に戻す」で、はじめの形に戻ります。'] }].concat(g.sections);
    const toc = document.getElementById('guide-toc'), box = document.getElementById('guide-content');
    if (Array.isArray(g.flow) && g.flow.length) {
      const list = el('ol', 'guide-flow'); list.setAttribute('aria-label', '使う順');
      g.flow.forEach((step, i) => { const li = el('li'); li.append(el('strong', '', step.what), el('span', '', step.where)); list.append(li); if (i < g.flow.length - 1) { const a = el('li', 'arrow', '→'); a.setAttribute('aria-hidden', 'true'); list.append(a); } });
      box.append(list);
    }
    g.sections.forEach((s) => {
      const a = el('a', '', s.title); a.href = '#g-' + s.id; a.dataset.id = s.id; a.addEventListener('click', (e) => { e.preventDefault(); jump(s.id); }); toc.append(a);
      const sec = el('section', 'guide-section'); sec.id = 'g-' + s.id; sec.append(el('h2', '', s.title));
      const ul = el('ul'); s.items.forEach((t) => ul.append(el('li', '', t))); sec.append(ul); box.append(sec);
    });
    status.hidden = true; document.getElementById('guide-root').hidden = false; spy();
  }
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
