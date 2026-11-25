'use strict';
// ハチの見本（0.2.8・2026-11-25 PO「もっとわかりやすいので作らないとわかりません」）。右上のアカウントのメニュー（▾）の「ハチの見本を見る」から。
//   この画面で出るハチの場面（短い待ち・長い待ち・つながらない・30 秒・空・利用ガイドを開くとき）を、本物と同じ部品・同じ置き場・同じ字で、
//   もふ → むぎわら の順に 3〜5 秒ずつ出す。
//   見るだけ: 窓口を呼ばない（api を使わない）・データ・ログインの札・この PC に覚えた絵の選び（mitsubachi_style）を変えない（絵の案はカードごとに渡す）。
//   カードの釦（もう一度・閉じる など）を押しても、次の場面へ進むだけ。
//   上の帯と同じ所に見本の帯（← 元の画面へ戻る・いまの場面・前／止める／次）。Esc でも戻る。戻ると元の画面のまま（開いていた月・表・送った位置）。
//   字は全部 textContent（部品も同じ）。外から読む物は無い。
(() => {
  const App = window.SouzaiApp, MB = window.Mitsubachi;
  if (!App || !MB || !App.fillEmpty) return; // 部品が読めない端末では、見本を出さない（メニューにも出さない）
  const { $, el, monthText, state } = App;
  const NAME = { B: 'もふ', C: 'むぎわら' };
  const NO_MONTH_TEXT = 'まだ店に公開している月がありません。本部からのお知らせをお待ちください。'; // 窓口 Code.gs の viewData_ の決まった文（店に公開の月がまだ無いとき）
  const GUIDE_TEXT = '利用ガイドを読み込んでいます…'; // guide.js の待ちと同じ字
  const SLOW_DEMO_MS = 2500; // 長い待ち: 本物は 8 秒で字が替わる。見本は 2.5 秒に縮める（帯の説明にそう書く）
  // 字に出す月（いま開いている月の隣。無ければいまの月）
  const months = () => { const i = state.months.indexOf(state.month); const m = state.months[i + 1] || state.months[i - 1] || state.month; return { m: monthText(m), now: monthText(state.month) }; };

  // ---------------------------------------------------------------- 場面（play(絵の案, 次へ) → 止める関数）
  const SCENES = [
    { key: 'short', name: '短い待ち（月の切り替えなど）', ms: 3400,
      note: '0.3 秒を超えたときだけ、真ん中のカードにハチが出て、終わるとすぐ消えます（0.3 秒までに終われば何も出ません）。',
      play(st) { App.showProgress(months().m + ' を読み込んでいます…', { style: st }); const t = window.setTimeout(() => App.hideProgress(), 2300); return () => { window.clearTimeout(t); App.hideProgress(); }; } },
    { key: 'long', name: '長い待ち（8 秒を超えた）', ms: 5600,
      note: '8 秒を超えると、下の字が「時間がかかっています」に替わります（見本では 2.5 秒に縮めています）。',
      play(st) { App.showProgress(months().m + ' を読み込んでいます…', { style: st, slowMs: SLOW_DEMO_MS }); return () => App.hideProgress(); } },
    { key: 'offline', name: 'つながらない（通信が切れた）', ms: 5000,
      note: 'やり直せば直るときだけ、困り顔と「もう一度」が出ます（見本なので、押しても次の場面へ進むだけです）。',
      play(st, next) { const x = months(); App.showRetry({ title: x.m + ' を開けませんでした', text: App.NET_MSG.offline, retry: next, closeLabel: '閉じる（' + x.now + ' のまま）', close: next, style: st }); return () => App.hideProgress(); } },
    { key: 'timeout', name: '30 秒たっても返事がない', ms: 5000,
      note: '30 秒で待つのをやめて、理由と「もう一度」を出します（見本なので、押しても次の場面へ進むだけです）。',
      play(st, next) { const x = months(); App.showRetry({ title: x.m + ' を開けませんでした', text: App.NET_MSG.timeout, retry: next, closeLabel: '閉じる（' + x.now + ' のまま）', close: next, style: st }); return () => App.hideProgress(); } },
    { key: 'empty', name: '空（店に公開の月がまだ無い・店のアカウント）', ms: 4600,
      note: '止まったハチと、ひとことと、次の一手の釦です（動きません）。',
      play(st, next) { const sg = stage(''); const p = el('section', 'empty-panel'); sg.append(p); App.fillEmpty(p, NO_MONTH_TEXT, next, next, st); return () => sg.remove(); } },
    { key: 'guide', name: '利用ガイドを開くとき', ms: 4200,
      note: '利用ガイドの本文の場所の真ん中に、ハチのカードが出ます（0.2.8 で、左上に出ていたのを直しました）。',
      play(st) { const sg = stage('is-guide'); const place = el('div', 'wait-place'); place.setAttribute('role', 'status'); sg.append(place); const h = MB.wait(place, { text: GUIDE_TEXT, style: st }); return () => { h.end(); sg.remove(); }; } }
  ];
  const STEPS = [];
  ['B', 'C'].forEach((st) => SCENES.forEach((sc) => STEPS.push({ st, sc })));

  // ---------------------------------------------------------------- 見本の帯（上の帯と同じ所）・場面の説明・下の画面を押せなくする透明の幕
  let run = null;
  function stage(cls) { const sg = el('div', 'bee-demo-stage' + (cls ? ' ' + cls : '')); sg.style.top = run.top + 'px'; document.body.append(sg); return sg; }
  function btn(cls, text, title, fn) { const b = el('button', 'nav-btn ' + cls, text); b.type = 'button'; if (title) b.title = title; b.addEventListener('click', fn); return b; }
  function build() {
    const bar = el('div', 'bee-demo-bar'); bar.id = 'bee-demo'; bar.setAttribute('role', 'region'); bar.setAttribute('aria-label', 'ハチの見本');
    const nav = el('nav', 'nav-fixed'); nav.setAttribute('aria-label', '戻る');
    const back = btn('bd-back', '← 元の画面へ戻る', '見本を終えて、元の画面へ戻ります（Esc でも戻ります）', () => end(''));
    back.setAttribute('aria-label', '見本を終えて、元の画面へ戻る'); nav.append(back);
    const badge = el('span', 'bd-badge', 'ハチの見本'), count = el('span', 'bd-count'), style = el('span', 'bd-style'), name = el('strong', 'bd-name');
    name.setAttribute('aria-live', 'polite');
    const ctl = el('span', 'bd-ctl');
    const prev = btn('bd-prev', '◀ 前', '前の場面', () => go(-1));
    const pause = btn('bd-pause', '‖ 止める', 'いまの場面のまま止める（もう一度押すと続ける）', () => setPaused(!run.paused));
    pause.setAttribute('aria-pressed', 'false');
    const next = btn('bd-next', '次 ▶', '次の場面', () => go(+1));
    ctl.append(prev, pause, next);
    bar.append(nav, badge, count, style, name, ctl);
    const cap = el('div', 'bee-demo-cap'); cap.id = 'bee-demo-cap';
    const note = el('span', 'bd-note'), still = el('span', 'bd-still', 'この PC は『動きを減らす』の設定なので、ハチは止まって見えます。');
    still.hidden = !MB.deviceReduced();
    cap.append(note, still);
    const shield = el('div', 'bee-demo-shield'); shield.setAttribute('aria-hidden', 'true');
    document.body.append(shield, bar, cap);
    Object.assign(run, { bar, cap, shield, count, style, name, note, prev, pause, next, back });
  }
  // 帯と説明の高さに合わせて、説明と場面の上端を決める（場面の説明が替わったとき・幅が替わったときも。狭い幅では帯が2〜3行になる）
  function place() {
    if (!run) return;
    const tb = $('#topbar'); // PC の幅では上の帯と同じ高さ（← 元の画面へ戻る が ← 戻る と同じ高さに来る）。狭い幅（上の帯が何行にもなる）では、見本の帯は中身の高さ
    run.bar.style.minHeight = tb && tb.offsetHeight && window.innerWidth > 760 ? tb.offsetHeight + 'px' : '';
    run.cap.style.top = run.bar.offsetHeight + 'px';
    run.top = run.bar.offsetHeight + run.cap.offsetHeight;
    document.querySelectorAll('.bee-demo-stage').forEach((sg) => { sg.style.top = run.top + 'px'; });
  }
  function paint(i) {
    const { st, sc } = STEPS[i];
    run.count.textContent = (i + 1) + ' / ' + STEPS.length;
    run.style.textContent = '絵: ' + NAME[st];
    run.name.textContent = sc.name;
    run.note.textContent = sc.note;
    run.prev.disabled = i === 0;
    run.next.textContent = i === STEPS.length - 1 ? '終える ▶' : '次 ▶';
    place();
  }
  function stopScene() { window.clearTimeout(run.timer); run.timer = 0; const s = run.stop; run.stop = null; if (s) { try { s(); } catch { /* 片付けの失敗で止めない */ } } }
  function show(i) {
    stopScene();
    run.i = i; paint(i);
    const { st, sc } = STEPS[i];
    const me = ++run.seq;
    run.stop = sc.play(st, () => { if (run && run.seq === me) go(+1); }); // カードの釦＝次の場面へ（二度押しで2つ進まない）
    if (!run.paused) run.timer = window.setTimeout(() => go(+1), sc.ms);
    const a = document.activeElement; if (!a || a === document.body) run.next.focus();
  }
  function go(d) { if (!run) return; const j = run.i + d; if (j < 0) return; if (j >= STEPS.length) { end('ハチの見本は、ここまでです（元の画面のままです）'); return; } show(j); }
  function setPaused(on) {
    run.paused = on; run.pause.textContent = on ? '▶ 続ける' : '‖ 止める'; run.pause.setAttribute('aria-pressed', String(on));
    window.clearTimeout(run.timer); run.timer = 0;
    if (!on) run.timer = window.setTimeout(() => go(+1), STEPS[run.i].sc.ms);
  }
  function start() {
    if (run) return;
    if (!$('#progress').hidden) return; // 本物の待ちの最中は始めない
    $('#account-menu').hidden = true; $('#account-button').setAttribute('aria-expanded', 'false');
    run = { i: 0, seq: 0, timer: 0, stop: null, paused: false, top: 0 };
    build();
    window.addEventListener('resize', place, { passive: true });
    show(0);
  }
  function end(message) {
    if (!run) return;
    stopScene(); App.hideProgress();
    run.bar.remove(); run.cap.remove(); run.shield.remove();
    window.removeEventListener('resize', place);
    run = null;
    $('#account-button').focus();
    if (message) App.toast(message);
  }
  // Esc＝見本を終える（本物の幕の Esc〈閉じる〉より先に受ける）
  window.addEventListener('keydown', (e) => { if (!run || e.key !== 'Escape') return; e.preventDefault(); e.stopPropagation(); end(''); }, true);

  // ---------------------------------------------------------------- メニューの入口（ハチの絵を選ぶ所の下）
  const host = $('#bee-pick'); if (!host) return;
  const open = el('button', 'bee-demo-open'); open.type = 'button'; open.id = 'bee-demo-open';
  open.append(el('strong', '', 'ハチの見本を見る'), el('small', '', '待ち・つながらない・空の出方を、もふ → むぎわら の順に見ます（見るだけ）'));
  open.addEventListener('click', start);
  host.append(open);
})();
