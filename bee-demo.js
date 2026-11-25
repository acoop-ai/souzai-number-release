'use strict';
// ハチの見本（0.2.8 で作った・0.2.9 で自分のペースで見る形に）。右上のアカウントのメニュー（▾）の「ハチの見本を見る」から。
//   この画面で出るハチの場面（短い待ち・長い待ち・つながらない・30 秒・空・利用ガイドを開くとき）を、本物と同じ部品・同じ置き場・同じ字で、もふ・むぎわら の 12 場面。
//   0.2.9（2026-11-25 PO「ハチの見本見ました　でも短すぎて見られません」）:
//     既定は「自動で進まない」＝1つの場面を出したまま止まる。「次へ ▶」「◀ 前へ」（キーの → ←）で進める。左の「場面の一覧」から見たい場面へ直に飛べる。
//     いま何番目か（3 / 12）・場面の名前・どんな時に出るかの1行を大きめに。「自動で流す」を選んだときだけ 1 場面 10 秒（長い待ちは字が替わってから 6 秒）・一時停止できる。
//     待ちのハチ（8 の字）は場面を出している間ずっと飛ぶ（短い待ちも出したまま）。長い待ちは本物と同じ 8 秒で字が替わる（それまでの秒を出す）。
//   見るだけ: 窓口を呼ばない（api を使わない）・データ・ログインの札・この PC に覚えた絵の選び（mitsubachi_style）を変えない（絵の案はカードごとに渡す）。
//   カードの釦（もう一度・閉じる など）を押しても、次の場面へ進むだけ。
//   上の帯と同じ所に見本の帯（← 元の画面へ戻る）。Esc でも戻る。戻ると元の画面のまま（開いていた月・表・送った位置）。
//   字は全部 textContent（部品も同じ）。外から読む物は無い。
(() => {
  const App = window.SouzaiApp, MB = window.Mitsubachi;
  if (!App || !MB || !App.fillEmpty) return; // 部品が読めない端末では、見本を出さない（メニューにも出さない）
  const { $, el, monthText, state } = App;
  const NAME = { B: 'もふ', C: 'むぎわら' };
  const NO_MONTH_TEXT = 'まだ店に公開している月がありません。本部からのお知らせをお待ちください。'; // 窓口 Code.gs の viewData_ の決まった文（店に公開の月がまだ無いとき）
  const GUIDE_TEXT = '利用ガイドを読み込んでいます…'; // guide.js の待ちと同じ字
  const AUTO_MS = 10000, AFTER_SLOW_MS = 6000, SLOW_AT_MS = 8300; // 自動で流すときの 1 場面・長い待ちで字が替わってから・字が替わる時刻（カードが出る 0.3 秒＋8 秒＝本物と同じ）
  const NARROW = 760; // これより狭い幅では、場面の一覧は「☰ 場面の一覧」で開け閉め
  // 字に出す月（いま開いている月の隣。無ければいまの月）
  const months = () => { const i = state.months.indexOf(state.month); const m = state.months[i + 1] || state.months[i - 1] || state.month; return { m: monthText(m), now: monthText(state.month) }; };

  // ---------------------------------------------------------------- 場面（play(絵の案, 次へ) → 止める関数）
  const SCENES = [
    { key: 'short', name: '短い待ち', when: '開く・月の切り替えなどが 0.3 秒より長くかかったとき、真ん中に出ます。終わるとすぐ消えます。',
      note: '見本では出したままにしています（ハチはずっと 8 の字に飛びます）。「↺ この場面をもう一度」で、0.3 秒たって出る所から見られます。',
      play(st) { App.showProgress(months().m + ' を読み込んでいます…', { style: st }); return () => App.hideProgress(); } },
    { key: 'long', name: '長い待ち（8 秒を超えた）', when: '8 秒を超えても終わらないとき、下の字が「時間がかかっています」に替わります。', slow: true,
      note: '', // 字が替わるまでの秒を出す（tick）
      play(st) { App.showProgress(months().m + ' を読み込んでいます…', { style: st }); return () => App.hideProgress(); } },
    { key: 'offline', name: 'つながらない', when: '通信が切れたとき、困り顔と理由と「もう一度」が出ます。「もう一度」で同じ操作をやり直せます。',
      note: '見本なので、カードの釦を押すと次の場面へ進みます。',
      play(st, next) { const x = months(); App.showRetry({ title: x.m + ' を開けませんでした', text: App.NET_MSG.offline, retry: next, closeLabel: '閉じる（' + x.now + ' のまま）', close: next, style: st }); return () => App.hideProgress(); } },
    { key: 'timeout', name: '30 秒たっても返事がない', when: '30 秒たっても返事が無いとき、待つのをやめて、理由と「もう一度」を出します。',
      note: '見本なので、カードの釦を押すと次の場面へ進みます。',
      play(st, next) { const x = months(); App.showRetry({ title: x.m + ' を開けませんでした', text: App.NET_MSG.timeout, retry: next, closeLabel: '閉じる（' + x.now + ' のまま）', close: next, style: st }); return () => App.hideProgress(); } },
    { key: 'empty', name: '空（見られる月がまだ無い）', when: '店のアカウントで、店に公開の月がまだ無いとき。止まったハチと、次の一手の釦です。',
      note: '見本なので、カードの釦を押すと次の場面へ進みます。',
      play(st, next) { const sg = stage(''); const p = el('section', 'empty-panel'); sg.append(p); App.fillEmpty(p, NO_MONTH_TEXT, next, next, st); return () => sg.remove(); } },
    { key: 'guide', name: '利用ガイドを開くとき', when: '右上の「？ 利用ガイド」を開いて 0.3 秒より長くかかったとき、本文の場所の真ん中に出ます。',
      note: '',
      play(st) { const sg = stage('is-guide'); const place = el('div', 'wait-place'); place.setAttribute('role', 'status'); sg.append(place); const h = MB.wait(place, { text: GUIDE_TEXT, style: st }); return () => { h.end(); sg.remove(); }; } }
  ];
  const STEPS = [];
  ['B', 'C'].forEach((st) => SCENES.forEach((sc) => STEPS.push({ st, sc })));

  // ---------------------------------------------------------------- 見本の帯（上の帯と同じ所）・左の場面の一覧・場面の上の進める釦と説明・下の画面を押せなくする透明の幕
  let run = null;
  function stage(cls) { const sg = el('div', 'bee-demo-stage' + (cls ? ' ' + cls : '')); sg.style.top = run.top + 'px'; sg.style.left = run.left + 'px'; document.body.append(sg); return sg; }
  function btn(cls, text, title, fn) { const b = el('button', cls, text); b.type = 'button'; if (title) b.title = title; b.addEventListener('click', fn); return b; }
  function build() {
    // 上の帯（← 元の画面へ戻る は、上の帯の ← 戻る と同じ所）
    const bar = el('div', 'bee-demo-bar'); bar.id = 'bee-demo'; bar.setAttribute('role', 'region'); bar.setAttribute('aria-label', 'ハチの見本');
    const nav = el('nav', 'nav-fixed'); nav.setAttribute('aria-label', '戻る');
    const back = btn('nav-btn bd-back', '← 元の画面へ戻る', '見本を終えて、元の画面へ戻ります（Esc でも戻ります）', () => end(''));
    back.setAttribute('aria-label', '見本を終えて、元の画面へ戻る');
    const top = btn('nav-btn bd-top', '↑ TOP', '場面の一覧と場面を、いちばん上へ', () => { run.list.scrollTop = 0; document.querySelectorAll('.bee-demo-stage').forEach((sg) => { sg.scrollTop = 0; }); });
    top.setAttribute('aria-label', 'TOP（いちばん上へ）');
    nav.append(back, top);
    bar.append(nav, el('span', 'bd-badge', 'ハチの見本'), el('span', 'bd-lead', '待ち・つながらない・空のときのハチを、1場面ずつ見られます（見るだけ）'), el('span', 'bd-keys', 'キー: → 次へ・← 前へ・Esc 戻る'));
    // 左の場面の一覧（もふ 1〜6・むぎわら 7〜12）
    const list = el('nav', 'bee-demo-list'); list.id = 'bee-demo-list'; list.setAttribute('aria-label', '場面の一覧');
    list.append(el('p', 'bdl-title', '場面の一覧（12）'));
    const items = [];
    ['B', 'C'].forEach((st) => {
      list.append(el('p', 'bdl-head', '絵: ' + NAME[st]));
      STEPS.forEach((s, i) => {
        if (s.st !== st) return;
        const b = btn('bdl-item', '', s.sc.when, () => { closeList(); show(i); });
        b.append(el('span', 'bdl-n', String(i + 1)), el('span', 'bdl-name', s.sc.name));
        b.dataset.step = String(i + 1);
        list.append(b); items.push(b);
      });
    });
    // 場面の上: 進める釦（前へ・何番目・次へ・もう一度・自動で流す）・場面の名前・どんな時に出るか
    const cap = el('div', 'bee-demo-cap'); cap.id = 'bee-demo-cap';
    const row = el('div', 'bd-row');
    const prev = btn('btn bd-prev', '◀ 前へ', '前の場面（キーの ←）', () => go(-1));
    const count = el('strong', 'bd-count');
    const next = btn('btn btn-primary bd-next', '次へ ▶', '次の場面（キーの →）', () => go(+1));
    const replay = btn('btn bd-replay', '↺ ', 'いまの場面を、はじめから出し直します', () => show(run.i)); replay.append(el('span', 'bd-long', 'この場面を'), document.createTextNode('もう一度')); // 狭い幅では「↺ もう一度」
    const autoBox = el('span', 'bd-autobox');
    const autoBtn = btn('btn bd-auto', '自動で流す', '1 場面 10 秒で次へ進みます（もう一度押すとやめる）', () => setAuto(!run.auto)); autoBtn.setAttribute('aria-pressed', 'false');
    const pauseBtn = btn('btn bd-pause', '‖ 一時停止', 'いまの場面のまま止める（もう一度押すと続ける）', () => setPaused(!run.paused)); pauseBtn.setAttribute('aria-pressed', 'false'); pauseBtn.hidden = true;
    const leftTxt = el('span', 'bd-left');
    autoBox.append(autoBtn, pauseBtn, leftTxt);
    const listBtn = btn('btn bd-listbtn', '☰ 場面の一覧', '見たい場面を選ぶ', () => toggleList()); listBtn.setAttribute('aria-expanded', 'false'); listBtn.setAttribute('aria-controls', 'bee-demo-list');
    row.append(prev, count, next, listBtn, replay, autoBox);
    const title = el('div', 'bd-title'); title.setAttribute('aria-live', 'polite');
    const style = el('span', 'bd-style'), name = el('strong', 'bd-name'); title.append(style, name);
    const when = el('p', 'bd-when'), note = el('p', 'bd-note'), still = el('p', 'bd-still', 'この PC は『動きを減らす』の設定なので、ハチは止まって見えます。');
    still.hidden = !MB.deviceReduced();
    cap.append(row, title, when, note, still);
    const shield = el('div', 'bee-demo-shield'); shield.setAttribute('aria-hidden', 'true');
    document.body.append(shield, list, cap, bar);
    Object.assign(run, { bar, list, cap, shield, items, count, style, name, when, note, prev, next, replay, autoBtn, pauseBtn, leftTxt, listBtn, back });
  }
  const narrow = () => window.innerWidth <= NARROW;
  // 帯・一覧・説明の大きさに合わせて、場面の置き場（上と左）を決める（場面の説明が替わったとき・幅が替わったときも。狭い幅では帯と説明が何行にもなる）
  function place() {
    if (!run) return;
    const tb = $('#topbar'), nw = narrow(); // PC の幅では上の帯と同じ高さ（← 元の画面へ戻る が ← 戻る と同じ高さに来る）
    run.bar.style.minHeight = tb && tb.offsetHeight && !nw ? tb.offsetHeight + 'px' : '';
    run.list.classList.toggle('is-drop', nw);
    run.listBtn.hidden = !nw;
    if (!nw) { run.list.hidden = false; run.listBtn.setAttribute('aria-expanded', 'false'); } else if (!run.listOpen) run.list.hidden = true;
    const barH = run.bar.getBoundingClientRect().height; // 端数のまま（説明と場面の間に 1px の隙間も重なりも作らない）
    run.left = nw ? 0 : run.list.offsetWidth;
    run.cap.style.top = barH + 'px'; run.cap.style.left = run.left + 'px';
    run.capH = run.cap.getBoundingClientRect().height;
    run.top = barH + run.capH;
    run.list.style.top = (nw ? run.top : barH) + 'px';
    document.querySelectorAll('.bee-demo-stage').forEach((sg) => { sg.style.top = run.top + 'px'; sg.style.left = run.left + 'px'; });
    const pg = $('#progress'); pg.style.top = run.top + 'px'; pg.style.left = run.left + 'px'; // 本物の幕を、場面の置き場の中に（真ん中に寄せる作りはそのまま）
  }
  function toggleList(on) {
    if (!run || !narrow()) return;
    run.listOpen = on === undefined ? !run.listOpen : on;
    run.list.hidden = !run.listOpen; run.listBtn.setAttribute('aria-expanded', String(run.listOpen));
    if (run.listOpen && run.items[run.i]) run.items[run.i].focus();
  }
  function closeList() { if (run && run.listOpen) { toggleList(false); run.listBtn.focus(); } }
  function paint(i) {
    const { st, sc } = STEPS[i];
    run.count.textContent = (i + 1) + ' / ' + STEPS.length;
    run.style.textContent = '絵: ' + NAME[st];
    run.name.textContent = sc.name;
    run.when.textContent = sc.when;
    run.note.textContent = sc.slow ? '本物と同じ 8 秒で、カードの下の字が替わります（あと 9 秒）。' : sc.note; run.note.hidden = !sc.note && !sc.slow; // 長い待ちは tick が秒を数える（先に同じ長さの字を入れて置き場を決める）
    run.prev.disabled = i === 0;
    run.next.textContent = i === STEPS.length - 1 ? '終える ▶' : '次へ ▶';
    run.next.title = i === STEPS.length - 1 ? '見本を終えて、元の画面へ戻ります' : '次の場面（キーの →）';
    run.items.forEach((b, k) => { if (k === i) b.setAttribute('aria-current', 'step'); else b.removeAttribute('aria-current'); });
    place();
  }
  function stopScene() { const s = run.stop; run.stop = null; if (s) { try { s(); } catch { /* 片付けの失敗で止めない */ } } }
  function show(i) {
    stopScene();
    run.i = i; paint(i);
    const { st, sc } = STEPS[i];
    const me = ++run.seq;
    const keep = document.activeElement; // 本物のエラーのカードは「もう一度」に手を移すので、見本の釦に手を戻す（キーで続けて進められる）
    run.stop = sc.play(st, () => { if (run && run.seq === me) go(+1); }); // カードの釦＝次の場面へ（二度押しで2つ進まない）
    const mine = keep && keep !== document.body && keep.isConnected && (run.bar.contains(keep) || run.cap.contains(keep) || run.list.contains(keep));
    (mine ? keep : run.next).focus({ preventScroll: true });
    run.t0 = Date.now();
    run.ms = sc.slow ? SLOW_AT_MS + AFTER_SLOW_MS : AUTO_MS; run.deadline = run.t0 + run.ms; run.rest = run.ms;
    tick();
  }
  function go(d) { if (!run) return; const j = run.i + d; if (j < 0) return; if (j >= STEPS.length) { end('ハチの見本は、ここまでです（元の画面のままです）'); return; } show(j); }
  // 0.25 秒ごと: 長い待ちの「字が替わるまで あと◯秒」と、自動で流すときの「あと◯秒で次へ」
  function tick() {
    if (!run) return;
    const { sc } = STEPS[run.i];
    if (sc.slow) {
      const changed = !!document.querySelector('#progress .mb-wait.mb-slow');
      const sec = Math.max(1, Math.ceil((SLOW_AT_MS - (Date.now() - run.t0)) / 1000));
      const t = changed ? '✓ カードの下の字が替わりました（本物と同じ、8 秒を超えたとき）。' : '本物と同じ 8 秒で、カードの下の字が替わります（あと ' + sec + ' 秒）。';
      if (run.note.textContent !== t) { run.note.textContent = t; if (run.cap.getBoundingClientRect().height !== run.capH) place(); } // 説明の行数が替われば、場面の置き場も置き直す
    }
    if (!run.auto) return;
    if (run.paused) { run.leftTxt.textContent = '止めています'; return; }
    if (run.i === STEPS.length - 1) { setAuto(false); run.leftTxt.textContent = '最後の場面です'; return; }
    const rest = run.deadline - Date.now();
    if (rest <= 0) { go(+1); return; }
    const t = 'あと ' + Math.ceil(rest / 1000) + ' 秒で次へ';
    if (run.leftTxt.textContent !== t) run.leftTxt.textContent = t;
  }
  function setAuto(on) {
    run.auto = on; run.paused = false;
    run.autoBtn.setAttribute('aria-pressed', String(on)); run.autoBtn.textContent = on ? '✓ 自動で流す' : '自動で流す';
    run.pauseBtn.hidden = !on; run.pauseBtn.textContent = '‖ 一時停止'; run.pauseBtn.setAttribute('aria-pressed', 'false');
    run.leftTxt.textContent = '';
    if (on) { run.deadline = Date.now() + run.ms; tick(); }
    place();
  }
  function setPaused(on) {
    if (!run.auto) return;
    if (on) run.rest = Math.max(0, run.deadline - Date.now()); else run.deadline = Date.now() + run.rest;
    run.paused = on; run.pauseBtn.textContent = on ? '▶ 続ける' : '‖ 一時停止'; run.pauseBtn.setAttribute('aria-pressed', String(on));
    tick();
  }
  function start() {
    if (run) return;
    if (!$('#progress').hidden) return; // 本物の待ちの最中は始めない
    $('#account-menu').hidden = true; $('#account-button').setAttribute('aria-expanded', 'false');
    run = { i: 0, seq: 0, stop: null, auto: false, paused: false, listOpen: false, top: 0, left: 0, t0: 0, ms: AUTO_MS, deadline: 0, rest: 0, ticker: 0 };
    build();
    run.ticker = window.setInterval(tick, 250);
    window.addEventListener('resize', place, { passive: true });
    show(0);
  }
  function end(message) {
    if (!run) return;
    window.clearInterval(run.ticker);
    stopScene(); App.hideProgress();
    run.bar.remove(); run.cap.remove(); run.list.remove(); run.shield.remove();
    const pg = $('#progress'); pg.style.top = ''; pg.style.left = '';
    window.removeEventListener('resize', place);
    run = null;
    $('#account-button').focus();
    if (message) App.toast(message);
  }
  // キー: Esc＝見本を終える（本物の幕の Esc〈閉じる〉より先に受ける）・→ 次へ・← 前へ
  window.addEventListener('keydown', (e) => {
    if (!run || e.altKey || e.ctrlKey || e.metaKey) return;
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); end(''); return; }
    if (e.key === 'ArrowRight') { e.preventDefault(); e.stopPropagation(); go(+1); return; }
    if (e.key === 'ArrowLeft') { e.preventDefault(); e.stopPropagation(); go(-1); }
  }, true);

  // ---------------------------------------------------------------- メニューの入口（ハチの絵を選ぶ所の下）
  const host = $('#bee-pick'); if (!host) return;
  const open = el('button', 'bee-demo-open'); open.type = 'button'; open.id = 'bee-demo-open';
  open.append(el('strong', '', 'ハチの見本を見る'), el('small', '', '待ち・つながらない・空のときのハチを、1場面ずつ自分のペースで見ます（見るだけ）'));
  open.addEventListener('click', start);
  host.append(open);
})();
