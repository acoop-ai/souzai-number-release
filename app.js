(() => {
  'use strict';
  // 商品呼出番号表の画面の土台（版 souzai-number-api-0.2.4）。通信・ログイン・保存の列・待ちの印・上の帯・Excel。一覧は view-list.js・見本は view-sheet.js・記録と設定は view-more.js。
  // 業務データ・アカウント表・鍵は置かない。だれか・何ができるかは、毎回 窓口が決める。原価は本部の人が開くとき（たたんでいなければ）・「原価を表示」・Excel のときだけ受け取り、端末に残さない。見るだけの人には窓口が返さない。
  // 信頼境界: 窓口から届く文字（商品名など）は、人が打った文字。画面に出すときは textContent か esc() を通す。
  const APP_VERSION = 'souzai-number-api-0.2.4';
  const CONFIG = window.APP_CONFIG || {};
  const IS_DEV = CONFIG.dev === true && /^(127\.0\.0\.1|localhost)$/.test(location.hostname);
  const GAS_URL = String(CONFIG.gasUrl || '');
  const GAS_URL_OK = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(GAS_URL) || (IS_DEV && GAS_URL === '/__gas');
  const REDIRECT_URI = String(CONFIG.oauthRedirectUri || '');
  const SESSION_KEY = 'souzai-number-session-v1';
  const VIEW_KEY = 'souzai-number-view-v1';
  const PREF_KEY = 'souzai-number-pref-v1'; // アカウントごとの好み（原価をたたむ・見るだけか）。読めないときは表示
  const SESSION_MSG = 'ログインの期限が切れました。もう一度ログインしてください。';
  const GIVEUP_MS = 30000, SLOW_MS = 8000, HINT_MS = 300, SAVE_WAIT_MS = 700;
  const SLOW_TEXT = '時間がかかっています（通信を確かめています）…';
  const L = window.SouzaiLogin, SN = window.SearchNorm || null, C = window.SouzaiCore, S = window.SouzaiSheet, X = window.SouzaiXlsx;
  const $ = (s) => document.querySelector(s);
  const el = (tag, cls, txt) => { const n = document.createElement(tag); if (cls) n.className = cls; if (txt !== undefined && txt !== null) n.textContent = String(txt); return n; };
  const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const store = () => { try { return window.localStorage; } catch { return null; } };
  const appError = (message, code) => { const e = new Error(message); e.code = code || 'ERROR'; return e; };
  const rid = () => 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 12);
  const monthText = (m) => { const x = /^(\d{4})-(\d{2})$/.exec(String(m)); return x ? x[1] + '年' + Number(x[2]) + '月' : String(m); };

  const state = { clientId: '', session: null, principal: null, db: null, months: [], month: '', view: 'list', alerts: [], alertsByPid: new Map(), gaps: { price: new Set(), cost: new Set() }, showCosts: false, readOnly: false, viewPictures: null, logRow: 0, undo: [], redo: [], tokenClient: null, tokenJob: null, gisLoading: null, loggingIn: false, leaving: false, words: null };
  const views = {}; // 画面ごとの部品（list・reader・sheet・log・settings）
  const FORM_CLASS = { '定番': 'f-reg', '月間特売': 'f-sale', '新商品': 'f-new', '月間特売・新商品': 'f-both', 'カット商品': 'f-cut', '休み': 'f-rest', '終売': 'f-off', '計画': 'f-off' };
  function readPref() { try { const o = JSON.parse(store()?.getItem(PREF_KEY) || 'null'); return o && typeof o === 'object' && o.accounts ? o : { last: '', accounts: {} }; } catch { return { last: '', accounts: {} }; } }
  function writePref(pref) { try { store()?.setItem(PREF_KEY, JSON.stringify(pref)); } catch { /* 覚えられなくても動く */ } }
  /** この端末で最後に入った人が、原価を一緒に受け取るか（たたんだ人と見るだけの人には頼まない） */
  function wantsCosts(email) { const pref = readPref(); const who = pref.accounts[email || pref.last] || {}; return who.view !== true && who.hide !== true; }
  function noteAccount(email, isView) { const pref = readPref(); pref.last = email; const who = pref.accounts[email] || (pref.accounts[email] = {}); if (isView) who.view = true; else delete who.view; writePref(pref); }
  function setCostFold(email, hide) { const pref = readPref(); const who = pref.accounts[email] || (pref.accounts[email] = {}); if (hide) who.hide = true; else delete who.hide; writePref(pref); }
  /** 字の飾りを付けて書く（一覧と見るだけの一覧で同じ） */
  function paintRunsInto(node, text, runs) {
    if (!runs || !runs.length) { if (node.textContent !== text || node.childElementCount) node.textContent = text; return; }
    node.replaceChildren(); let at = 0;
    runs.slice().sort((a, b) => a[0] - b[0]).forEach((run) => { const from = Math.min(run[0], text.length), to = Math.min(run[1], text.length); if (from > at) node.append(document.createTextNode(text.slice(at, from))); const s = el('span', '', text.slice(from, to)); if (run[2] && run[2].c) s.style.color = '#' + run[2].c; node.append(s); at = to; });
    if (at < text.length) node.append(document.createTextNode(text.slice(at)));
  }

  // ------------------------------------------------------------ 待ちの印・お知らせ（フリーズに見せない: 0.3 秒から印・8 秒で「時間がかかっています」・30 秒で止める）
  const perfLog = [];
  function perfNote(what, t0, note, extra) { const row = { what, ms: Math.round(performance.now() - t0), note: note || '', at: Date.now() }; if (extra && extra.server) row.server = Number(extra.server.t) || 0; if (extra && Number.isFinite(extra.kb)) row.kb = extra.kb; perfLog.push(row); if (perfLog.length > 200) perfLog.shift(); }
  function spinner() { const s = el('span', 'spinner'); s.setAttribute('aria-hidden', 'true'); return s; }
  /** ボタンを待ちの形にする。返す値: 元に戻す関数（stage(文字) で段階の文字を替える）。0.3 秒より短い待ちでは印を出さない */
  function busyButton(button, label) {
    const noop = () => {}; noop.stage = () => {}; if (!button) return noop;
    const original = button.dataset.idleLabel || button.textContent; button.dataset.idleLabel = original;
    let textNow = label; let shown = false;
    const show = () => { shown = true; button.replaceChildren(spinner(), document.createTextNode(textNow)); };
    button.disabled = true; button.setAttribute('aria-busy', 'true');
    const hint = window.setTimeout(show, HINT_MS);
    let slow = window.setTimeout(() => { textNow = SLOW_TEXT; show(); }, SLOW_MS);
    const done = () => { window.clearTimeout(hint); window.clearTimeout(slow); button.disabled = false; button.removeAttribute('aria-busy'); button.textContent = original; delete button.dataset.idleLabel; };
    done.stage = (t) => { textNow = t; if (shown) show(); window.clearTimeout(slow); slow = window.setTimeout(() => { textNow = SLOW_TEXT; show(); }, SLOW_MS); };
    return done;
  }
  let progressSlow = 0;
  function showProgress(message) { const n = $('#progress'); n.querySelector('span').textContent = message; n.hidden = false; window.clearTimeout(progressSlow); progressSlow = window.setTimeout(() => { n.querySelector('span').textContent = SLOW_TEXT; }, SLOW_MS); }
  function hideProgress() { window.clearTimeout(progressSlow); $('#progress').hidden = true; }
  function toast(message, error, action) {
    const n = $('#toast'); n.textContent = message; n.classList.toggle('is-error', Boolean(error)); n.hidden = false;
    if (action) { const b = el('button', 'btn btn-small', action.label); b.type = 'button'; b.addEventListener('click', () => { n.hidden = true; action.run(); }); n.append(b); }
    window.clearTimeout(toast.timer); toast.timer = window.setTimeout(() => { n.hidden = true; }, action ? 12000 : (error ? 8000 : 4200));
  }
  function showNotice(message, title) { const n = $('#notice'); if (!message) { n.hidden = true; return; } n.hidden = false; n.querySelector('strong').textContent = title || '確かめてください'; n.querySelector('p').textContent = message; }

  // ------------------------------------------------------------ 通信
  async function api(action, data = {}, authenticated = true) {
    if (!GAS_URL_OK) throw appError('窓口の接続先が、まだ設定されていません。', 'SETUP');
    const body = { action, data };
    if (data && data.requestId) body.requestId = data.requestId;
    if (authenticated) {
      if (!state.session || Date.now() > state.session.e) { clearSession(); throw appError(SESSION_MSG, 'SESSION'); }
      body.session = state.session.t;
    }
    const sent = body.session;
    const controller = typeof AbortController === 'function' ? new AbortController() : null;
    let gaveUp = false;
    const timer = controller ? window.setTimeout(() => { gaveUp = true; controller.abort(); }, GIVEUP_MS) : 0;
    const t0 = performance.now();
    const noReply = () => appError(gaveUp ? '30 秒たっても返事がありません。通信を確かめて、もう一度お試しください（入れた物は消えません）。' : '通信できませんでした。通信を確かめて、もう一度お試しください（入れた物は消えません）。', 'NETWORK');
    let response;
    try { response = await fetch(GAS_URL, { method: 'POST', mode: 'cors', redirect: 'follow', cache: 'no-store', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body), signal: controller?.signal }); }
    catch { if (timer) window.clearTimeout(timer); perfNote('api:' + action, t0, 'error'); throw noReply(); }
    if (!response.ok) { if (timer) window.clearTimeout(timer); throw appError('通信に失敗しました（' + response.status + '）。時間をおいて、もう一度お試しください。', 'NETWORK'); }
    let result, size = 0;
    try { const raw = await response.text(); size = raw.length; result = JSON.parse(raw); } catch { throw gaveUp ? noReply() : appError('窓口の返事を読めませんでした。少し待ってから、もう一度お試しください。', 'NETWORK'); }
    finally { if (timer) window.clearTimeout(timer); }
    perfNote('api:' + action, t0, result && result.ok === false ? String(result.code || 'ERROR') : '', { server: result && result.ms, kb: Math.round(size / 102.4) / 10 });
    if (!result || result.ok !== true) {
      const error = appError(result?.message || '処理を完了できませんでした。', result?.code || 'ERROR');
      if (error.code === 'SESSION' && authenticated && state.session && state.session.t === sent) clearSession();
      throw error;
    }
    return result.data;
  }
  function readSession() { try { const o = JSON.parse(store()?.getItem(SESSION_KEY) || 'null'); if (!o || typeof o.t !== 'string' || !o.e || Date.now() > o.e - 60000) return null; return { t: o.t, e: Number(o.e) }; } catch { return null; } }
  function saveSession(v) { state.session = v; try { store()?.setItem(SESSION_KEY, JSON.stringify(v)); } catch {} }
  function clearSession() { state.session = null; try { store()?.removeItem(SESSION_KEY); } catch {} }

  // ------------------------------------------------------------ 保存の列（0.7 秒のまとめ＋1本ずつ＋同じ操作番号は1回＋楽観更新）
  const save = { pending: new Map(), before: new Map(), label: '直す', timer: 0, running: 0, chain: Promise.resolve(), failed: false, lastAt: '' };
  const stampNow = () => { const now = new Date(); save.lastAt = String(now.getHours()).padStart(2, '0') + ':' + String(now.getMinutes()).padStart(2, '0'); };
  function setSaveState(kind) {
    const n = $('#save-state'); n.classList.toggle('is-error', kind === 'error');
    if (kind === 'wait' || kind === 'saving') { n.replaceChildren(spinner(), document.createTextNode(' 保存中…')); return; }
    if (kind === 'error') { n.textContent = '⚠ 保存できていません（自動で送り直します）'; return; }
    n.textContent = '✓ 自動保存' + (save.lastAt ? ' ' + save.lastAt : '');
  }
  const hasUnsaved = () => save.pending.size > 0 || save.running > 0;
  /** 順番に1本ずつ送る（保存・動かす・自動の決め直しが追い越さない） */
  function serial(fn) { save.running++; const job = save.chain.then(fn, fn); const after = job.catch(() => {}).then(() => { save.running--; if (!hasUnsaved() && !save.failed) setSaveState('ok'); }); save.chain = after; return job; }
  function queueChanges(changes, label) {
    changes.forEach((c) => { const k = c.t + '|' + c.id; const e = save.pending.get(k) || { t: c.t, id: c.id, f: {} }; Object.assign(e.f, c.f); save.pending.set(k, e); });
    if (label) save.label = label;
    setSaveState('wait');
    window.clearTimeout(save.timer); save.timer = window.setTimeout(flush, SAVE_WAIT_MS);
  }
  function flush() {
    window.clearTimeout(save.timer);
    if (!save.pending.size) return save.chain;
    const batch = Array.from(save.pending.values()); save.pending.clear();
    const label = save.label; save.label = '直す';
    const before = new Map(save.before); save.before.clear();
    const requestId = rid();
    setSaveState('saving');
    serial(() => sendBatch(batch, label, requestId, before));
    return save.chain;
  }
  async function sendBatch(batch, label, requestId, before) {
    const t0 = performance.now();
    for (let attempt = 0; ; attempt++) {
      try {
        const data = await api('saveRows', { month: state.month, changes: batch, label, known: state.logRow, withCosts: state.showCosts, requestId });
        save.failed = false; stampNow();
        applyOthers(data);
        perfNote('保存', t0);
        return;
      } catch (error) {
        if (error.code === 'SESSION') { save.failed = true; setSaveState('error'); showLogin(error.message); return; }
        if (error.code === 'NETWORK' || error.code === 'BUSY') { // 通信の失敗・混雑だけ送り直す（同じ操作番号なので二重にならない）
          save.failed = true; setSaveState('error');
          await new Promise((r) => window.setTimeout(r, [1000, 3000, 8000, 15000][Math.min(attempt, 3)]));
          continue;
        }
        // 断られた（やり直しても変わらない）: 手元を元へ戻して理由を出す
        save.failed = false;
        const back = [];
        batch.forEach((c) => { const f = {}; Object.keys(c.f).forEach((k) => { const key = c.t + '|' + c.id + '|' + k; if (before.has(key)) f[k] = before.get(key); }); if (Object.keys(f).length) back.push({ t: c.t, id: c.id, f }); });
        C.applyChanges(state.db, back); refresh(back);
        toast('保存できませんでした: ' + error.message, true);
        return;
      }
    }
  }
  function applyOthers(data) {
    if (data && Number.isFinite(data.logRow)) state.logRow = data.logRow;
    const others = (data && data.others) || [];
    if (others.length) { C.applyChanges(state.db, others); refresh(others); toast('ほかの人の直しを ' + others.length + ' 欄 取り込みました。'); }
    if (data && data.reload) toast('ほかの人が商品を足すか、取り消しました。', false, { label: '読み直す', run: () => loadMonth(state.month) });
  }
  /** 直す入口はここだけ: 形を確かめ → 手元に当てる（楽観更新）→ 保存の列へ → 画面を部分更新。返す値: 問題の文（無ければ ''） */
  function change(changes, opts) {
    opts = opts || {};
    if (state.readOnly) return '';
    const list = (changes || []).filter((c) => Object.keys(c.f || {}).length);
    if (!list.length) return '';
    for (const c of list) for (const k of Object.keys(c.f)) { const p = C.fieldProblem(c.t, k, c.f[k]); if (p) return p; }
    C.beforeOf(state.db, list).forEach((b) => Object.keys(b.f).forEach((k) => { const key = b.t + '|' + b.id + '|' + k; if (!save.before.has(key)) save.before.set(key, b.f[k]); }));
    const before = opts.noHistory ? null : beforeAll(state.db, list);
    C.applyChanges(state.db, list);
    queueChanges(list, opts.label);
    if (before) remember({ steps: [{ before, after: cloneChanges(list), saveLabel: opts.label || '直す' }] });
    refresh(list);
    return '';
  }
  /** 手元が変わったあとの計算（アラート・抜け・札） */
  function recompute() {
    const db = state.db; if (!db) return;
    if (state.readOnly) { state.alerts = []; state.alertsByPid = new Map(); return; }
    state.alerts = C.alertsOf(db);
    const byPid = new Map();
    state.alerts.forEach((a) => {
      let pid = a.target && a.target.pid;
      if (!pid && a.target && a.target.rowId) pid = headPidOfRow(a.target.rowId);
      if (!pid) return;
      pid = String(pid);
      if (!byPid.has(pid)) byPid.set(pid, []);
      byPid.get(pid).push(a);
    });
    state.alertsByPid = byPid;
    state.gaps = C.gapsOf(db);
    const n = state.alerts.length;
    $('#alert-count').textContent = String(n);
    $('#alert-chip').classList.toggle('is-zero', n === 0);
  }
  /** 続きの行のアラートは、すぐ上の商品の行へ */
  function headPidOfRow(rowId) {
    const r = state.db.rows.get(String(rowId)); if (!r) return '';
    if (r['商品ID']) return String(r['商品ID']);
    const list = C.rowsOfSection(state.db, r['区分ID']); let i = list.indexOf(r);
    while (i > 0 && !list[i]['商品ID']) i--;
    return list[i] && list[i]['商品ID'] ? String(list[i]['商品ID']) : '';
  }
  /** 変えた所と、アラートが変わった所だけ描き直す */
  function refresh(changes) {
    const prev = state.alertsByPid;
    recompute();
    const touched = new Set();
    (changes || []).forEach((c) => { if (c.t === '商品' || c.t === '原価' || c.t === '新しい商品') touched.add(String(c.id)); else if (c.t === '行') { const r = state.db.rows.get(String(c.id)); if (r && r['商品ID']) touched.add(String(r['商品ID'])); } });
    const sig = (m, pid) => (m.get(pid) || []).map((a) => a.key + a.what).join('|');
    new Set([...prev.keys(), ...state.alertsByPid.keys()]).forEach((pid) => { if (sig(prev, pid) !== sig(state.alertsByPid, pid)) touched.add(pid); });
    views.list && views.list.update(touched);
    views.sheet && views.sheet.markDirty();
  }

  // ------------------------------------------------------------ 動かす（ドラッグも「ここに移す」も同じ）・元に戻す・自動
  function kindOfSection(secId) { const s = state.db.sections.get(String(secId)); return s ? s['表の種類'] : ''; }
  /** op: { kind:'sec', id, toBlock, before } か { kind:'row', id, toSec, before } */
  function doMove(op, label) {
    const db = state.db;
    if (state.readOnly) return false;
    let changes, kind;
    try {
      if (op.kind === 'sec') { kind = kindOfSection(op.id); changes = C.moveSection(db, { id: op.id, toBlock: op.toBlock, before: op.before || null }); }
      else { const r = db.rows.get(String(op.id)); kind = kindOfSection(r['区分ID']); changes = C.moveRow(db, { id: op.id, toSec: op.toSec, before: op.before || null }); }
    } catch (e) { toast(e.message, true); return false; }
    if (!changes.length) return false;
    const before = C.beforeOf(db, changes);
    const statusBefore = db.status[kind];
    C.applyChanges(db, changes);
    db.status[kind] = '手直し';
    const entry = remember({ label: label || '動かす', steps: [{ before, after: cloneChanges(changes), saveLabel: '元に戻す' }], kind, statusBefore, noMerge: true });
    refresh(changes);
    const requestId = rid();
    const t0 = performance.now();
    flush();
    setSaveState('saving');
    serial(async () => {
      try {
        const data = await api('moveRows', { month: state.month, op, requestId });
        if (data.changes && data.changes.length) C.applyChanges(db, data.changes); // 窓口の計算が正
        perfNote('動かす', t0); stampNow();
        views.sheet && views.sheet.markDirty();
      } catch (error) {
        C.applyChanges(db, before); db.status[kind] = statusBefore; forget(entry);
        refresh(before);
        toast('動かせませんでした: ' + error.message, true);
      }
    });
    return true;
  }
  // ------------------------------------------------------------ 元に戻す・やり直す（PO 2026-11-07「どの動きでも一つ前に戻るで、いくつも修正前に戻れる」）
  //   1回の操作＝1つの記録（steps: [{ before, after, saveLabel }]）。直す入口（change・動かす・自動の置き直し・自動に戻す・新しい商品・パックの値段）が積む。
  //   開いている月の、この画面の間だけ覚える（開き直したら、記録のタブの「前の値に戻す」）。同じ欄を続けて打った分は1回にまとめる。
  //   ほかの人があとから同じ欄を直していたら、その欄は戻さない（人の直しを消さない）。月を作る・取り込む・消す・店に公開は入れない（それぞれに戻す道がある）
  const HISTORY_MAX = 200, MERGE_MS = 4000;
  const cloneChanges = (list) => list.map((c) => ({ t: c.t, id: c.id, f: Object.assign({}, c.f) }));
  /** 直す前の値（原価も。core の beforeOf は原価の表を見ないので、ここで足す） */
  function beforeAll(db, list) {
    const out = C.beforeOf(db, list);
    list.forEach((c) => { if (c.t !== '原価' || !db.costs) return; const o = db.costs.get(String(c.id)) || {}; const f = {}; Object.keys(c.f).forEach((k) => { f[k] = o[k] === undefined ? '' : o[k]; }); out.push({ t: '原価', id: c.id, f }); });
    return out;
  }
  const fieldKey = (st) => (st.after.length === 1 && Object.keys(st.after[0].f).length === 1 ? st.after[0].t + '|' + st.after[0].id + '|' + Object.keys(st.after[0].f)[0] : '');
  /** 記録に積む（やり直すは捨てる）。同じ1欄を続けて直したときは、前の記録の「後」だけを替える */
  function remember(entry) {
    if (state.readOnly) return null;
    entry.at = Date.now();
    const last = state.undo[state.undo.length - 1];
    const key = !entry.noMerge && !entry.undo && entry.steps && entry.steps.length === 1 ? fieldKey(entry.steps[0]) : '';
    if (key && last && !last.noMerge && !last.undo && last.steps.length === 1 && fieldKey(last.steps[0]) === key && entry.at - last.at < MERGE_MS && last.steps[0].saveLabel === entry.steps[0].saveLabel) {
      last.steps[0].after = entry.steps[0].after; last.at = entry.at; last.label = '';
      state.redo = []; paintUndo(); return last;
    }
    if (entry.mergeKey && last && last.mergeKey === entry.mergeKey && entry.at - last.at < MERGE_MS) { last.redo = entry.redo; last.at = entry.at; if (entry.label) last.label = entry.label; state.redo = []; paintUndo(); return last; }
    state.undo.push(entry); if (state.undo.length > HISTORY_MAX) state.undo.shift();
    state.redo = []; paintUndo();
    return entry;
  }
  /** 行と区分の置き場の写し（自動の置き直し・自動に戻すの、前と後を比べるため） */
  function snapLayout() { const db = state.db; const cp = (m) => new Map([...m].map(([k, v]) => [k, Object.assign({}, v)])); return { rows: cp(db.rows), places: cp(db.places) }; }
  function diffLayout(snap) {
    const db = state.db; const before = [], after = [];
    [['行', db.rows, snap.rows], ['置き場', db.places, snap.places]].forEach(([t, now, old]) => now.forEach((o, id) => {
      const p = old.get(id); if (!p) return; const fb = {}, fa = {};
      Object.keys(o).forEach((k) => { if (k.charAt(0) === '_' || typeof o[k] === 'object') return; if (String(o[k] ?? '') !== String(p[k] ?? '')) { fb[k] = p[k] === undefined ? '' : p[k]; fa[k] = o[k]; } });
      if (Object.keys(fa).length) { before.push({ t, id, f: fb }); after.push({ t, id, f: fa }); }
    }));
    return { before, after };
  }
  function forget(entry) { const i = state.undo.lastIndexOf(entry); if (i >= 0) state.undo.splice(i, 1); paintUndo(); }
  /** 直前の記録に、同じ操作の続き（自動の置き直しなど）を足す。1秒より前の記録なら別の記録にする */
  function rememberMore(step) {
    const last = state.undo[state.undo.length - 1];
    if (last && !last.undo && Date.now() - last.at < 1500) { last.steps.push(step); last.noMerge = true; paintUndo(); return last; }
    return remember({ steps: [step], noMerge: true });
  }
  // 記録の名前（ボタンとお知らせに出す）: 1欄なら「番号 名前 欄 前→後」
  const short = (v) => { const t = String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim(); return t === '' ? '（空）' : (t.length > 14 ? t.slice(0, 14) + '…' : t); };
  function whoOf(c) {
    const db = state.db;
    if (c.t === '商品' || c.t === '原価') { const p = db.products.get(String(c.id)); return p ? [halfOf(p['呼出番']), short(C.nfkc(p['名前']))].filter(Boolean).join(' ') : String(c.id); }
    if (c.t === '行') { const r = db.rows.get(String(c.id)); const p = r && r['商品ID'] ? db.products.get(String(r['商品ID'])) : null; return p ? [halfOf(p['呼出番']), short(C.nfkc(p['名前']))].filter(Boolean).join(' ') : '空き行'; }
    if (c.t === '置き場') { const sc = db.sections.get(String(c.id)); return sc ? short(C.nfkc(sc['表題'])) : '区分'; }
    if (c.t === '月間') { const m = db.monthly.find((x) => String(x['行ID']) === String(c.id)); return m ? short(m['商品名']) : '月間商品'; }
    return '';
  }
  const halfOf = (v) => (v === null || v === undefined ? '' : C.nfkc(String(v)).trim());
  function labelOf(entry) {
    if (entry.label) return entry.label;
    const st = entry.steps[0]; const a = st.after;
    if (a.length === 1 && Object.keys(a[0].f).length === 1) {
      const k = Object.keys(a[0].f)[0]; const b = (st.before.find((x) => x.t === a[0].t && String(x.id) === String(a[0].id)) || { f: {} }).f[k];
      const fmt = (v) => (k === '粗利（包材値上前）' && v !== '' && v !== null && v !== undefined ? C.grossText(v) : short(v));
      return whoOf(a[0]) + ' ' + k + ' ' + fmt(b) + '→' + fmt(a[0].f[k]);
    }
    const n = entry.steps.reduce((s, x) => s + x.after.length, 0);
    return (whoOf(a[0]) || '') + ' ほか ' + n + ' か所';
  }
  function paintUndo() {
    const u = $('#undo-btn'), r = $('#redo-btn'); if (!u || !r) return;
    const grp = $('#undo-group'); if (grp) grp.hidden = Boolean(state.readOnly);
    const lu = state.undo[state.undo.length - 1], lr = state.redo[state.redo.length - 1];
    u.disabled = !lu; r.disabled = !lr;
    u.title = lu ? '元に戻す（' + state.undo.length + '）: ' + labelOf(lu) + '（Ctrl+Z）' : '元に戻す物はまだありません';
    r.title = lr ? 'やり直す: ' + labelOf(lr) + '（Ctrl+Y）' : 'やり直す物はありません';
    u.setAttribute('aria-label', '元に戻す' + (state.undo.length ? '（' + state.undo.length + '）' : ''));
    if (views.sheet && views.sheet.paintBar) views.sheet.paintBar();
  }
  /** 記録の1つを、後→前（戻す）か 前→後（やり直す）へ当てる。ほかの人があとから直した欄は飛ばす */
  function applyEntry(entry, back) {
    const db = state.db; let skipped = 0; const touched = [];
    flush(); // 打ちかけの直しを先に送る（ラベルを混ぜない）
    const steps = back ? entry.steps.slice().reverse() : entry.steps;
    for (const st of steps) {
      const want = back ? st.after : st.before, to = back ? st.before : st.after; // want＝いまそうなっているはずの値
      const now = beforeAll(db, cloneChanges(want));
      const list = [];
      to.forEach((c) => {
        const cur = now.find((x) => x.t === c.t && String(x.id) === String(c.id)); const exp = want.find((x) => x.t === c.t && String(x.id) === String(c.id));
        const f = {};
        Object.keys(c.f).forEach((k) => { const same = !cur || !exp || exp.f[k] === undefined || String(cur.f[k] ?? '') === String(exp.f[k] ?? ''); if (same) f[k] = c.f[k]; else skipped++; });
        if (Object.keys(f).length) list.push({ t: c.t, id: c.id, f });
      });
      if (!list.length) continue;
      const saveLabel = st.saveLabel === 'カット商品' ? 'カット商品' : (back ? '元に戻す' : (st.saveLabel || '直す'));
      const p = change(list, { label: saveLabel, noHistory: true });
      if (p) { toast(p, true); return false; }
      flush();
      touched.push(...list);
    }
    if (views.sheet) views.sheet.markDirty();
    flashRows(touched);
    return { skipped };
  }
  async function undo() {
    const e = state.undo.pop(); if (!e) return;
    const label = labelOf(e);
    if (e.undo) { paintUndo(); try { await e.undo(); if (e.redo) state.redo.push(e); } catch (error) { toast('元に戻せませんでした: ' + error.message, true); } paintUndo(); return; }
    const r = applyEntry(e, true); if (!r) { state.undo.push(e); paintUndo(); return; }
    state.redo.push(e); paintUndo();
    toast('元に戻しました: ' + label + (r.skipped ? '（ほかの人があとから直した ' + r.skipped + ' 欄は、そのままにしました）' : ''), false, { label: 'やり直す', run: () => redo() });
  }
  async function redo() {
    const e = state.redo.pop(); if (!e) return;
    const label = labelOf(e);
    if (e.redo) { paintUndo(); try { await e.redo(); state.undo.push(e); } catch (error) { toast('やり直せませんでした: ' + error.message, true); } paintUndo(); return; }
    const r = applyEntry(e, false); if (!r) { state.redo.push(e); paintUndo(); return; }
    state.undo.push(e); paintUndo();
    toast('やり直しました: ' + label + (r.skipped ? '（ほかの人があとから直した ' + r.skipped + ' 欄は、そのままにしました）' : ''));
  }
  const undoLast = undo; // 番号表の「元に戻す」も同じ記録を戻す
  function flashRows(list) {
    const pids = new Set();
    list.forEach((c) => { if (c.t === '商品' || c.t === '原価') pids.add(String(c.id)); else if (c.t === '行') { const r = state.db.rows.get(String(c.id)); if (r && r['商品ID']) pids.add(String(r['商品ID'])); } });
    let first = null;
    pids.forEach((pid) => document.querySelectorAll('tr.lr[data-pid="' + CSS.escape(pid) + '"]').forEach((tr) => { tr.classList.remove('is-undone'); void tr.offsetWidth; tr.classList.add('is-undone'); if (!first) first = tr; }));
    if (first && state.view === 'list') first.scrollIntoView({ block: 'nearest' });
  }
  $('#undo-btn').addEventListener('click', () => undo());
  $('#redo-btn').addEventListener('click', () => redo());
  // Ctrl+Z／Ctrl+Y（Ctrl+Shift+Z）。打っている欄の中では、その欄の字の戻し（ブラウザの物）を優先する
  document.addEventListener('keydown', (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey || state.readOnly || !state.db) return;
    const t = e.target; if (t && (t.closest('input, textarea, select, [contenteditable="true"]'))) return;
    const k = e.key.toLowerCase();
    if (k === 'z' && !e.shiftKey) { e.preventDefault(); undo(); }
    else if (k === 'y' || (k === 'z' && e.shiftKey)) { e.preventDefault(); redo(); }
  });
  /** 販売形態を変えたあと: 自動の月は、上限を超えたら置き場を決め直す（置き場はいつも連動）。手直しの月は動かさない */
  function afterFormChange(kind) {
    const db = state.db;
    if (db.status[kind] === '手直し') return;
    const used = C.usedOfKind(db, kind), caps = C.capsOfKind(db, kind);
    if (!Object.keys(caps).some((b) => (used[b] || 0) > caps[b])) return;
    const snap = snapLayout();
    const res = C.autoPlace(db, kind); // 手元でも同じ部品で決める（0 往復で見える）
    const d = diffLayout(snap); if (d.after.length) rememberMore({ before: d.before, after: d.after, saveLabel: 'カット商品' }); // 戻すときは手直しに数えない（販売形態で決まる置き場）
    refresh(res.changes);
    flush();
    const requestId = rid();
    serial(async () => {
      try { const data = await api('runAuto', { month: state.month, kind, requestId }); if (data.changes && data.changes.length) { C.applyChanges(db, data.changes); refresh(data.changes); } }
      catch (error) { toast('置き場の決め直しを保存できませんでした: ' + error.message, true); }
    });
    if (res.moved.length) toast('自動で ' + res.moved.length + ' 区分を動かしました。');
  }

  // ------------------------------------------------------------ 中身を受け取る
  function applyWorkspace(data) {
    $('#month-panel').hidden = true; $('#seed-panel').hidden = true;
    if (data && data.notSeeded) { // 最初の取り込みの前: PO は取り込みの画面・ほかは準備中（中身は窓口が返していない）
      state.principal = data.principal; state.db = null; state.readOnly = true;
      $('#login-panel').hidden = true; showNotice('');
      ['list', 'sheet', 'log', 'settings'].forEach((v) => { $('#view-' + v).hidden = true; });
      $('#month-nav').hidden = true; $('#view-tabs').hidden = true; $('#tb-right').hidden = false;
      ['#alert-chip', '#save-state', '#excel-button', '#guide-link'].forEach((s) => { $(s).hidden = true; });
      $('#account-label').textContent = state.principal.email.split('@')[0];
      $('#account-role').textContent = state.principal.email + '（' + state.principal.roleLabel + '）';
      $('#app-version').textContent = '版 ' + APP_VERSION;
      if (views.intake) views.intake.showSeed(state.principal);
      return;
    }
    ['#excel-button', '#guide-link'].forEach((s) => { $(s).hidden = false; });
    state.months = data.months || [];
    state.month = data.month;
    // 店に公開している月（本部・PO の画面だけ受け取る。店は公開の月だけが months に来る）
    state.published = new Set(C.rowsOf((data.tables && data.tables['月の状態']) || []).filter((r) => r['表の種類'] === PUBLISH_KIND && r['形'] === '公開').map((r) => String(r['月'])));
    state.principal = data.principal;
    state.logRow = Number(data.logRow) || 0;
    state.undo = []; state.redo = []; paintUndo();
    const isView = state.principal.role === 'VIEW';
    state.readOnly = isView;
    noteAccount(state.principal.email, isView);
    const db = C.makeDb(data.tables, data.type, data.month);
    state.typeTables = isView ? null : data.type; // 手作りの番号表を読むとき、書式の鍵をアプリの表の型から選ぶ（取り込みの画面だけが使う）
    db.costInfo = data.costInfo || db.costInfo;
    const wanted = !isView && wantsCosts(state.principal.email);
    db.costs = data.costs && wanted ? new Map(C.rowsOf(data.costs).map((r) => [String(r['商品ID']), r])) : null; // たたんでいる人に、この端末で別の人が入った直後に届いた分は捨てる
    state.showCosts = Boolean(db.costs);
    state.viewPictures = isView ? (data.pictures || null) : null;
    state.db = db;
    if (SN && data.words && !state.words) { state.words = data.words; try { SN.load(data.words); } catch {} }
    recompute();
    $('#login-panel').hidden = true; showNotice('');
    $('#month-nav').hidden = false; $('#view-tabs').hidden = isView; $('#tb-right').hidden = false; // 店は番号表だけ（PO 2026-11-07「店はマスターを見ない・一覧はなし」）＝切り替えのタブを出さない
    $('#month-label').textContent = monthText(state.month);
    document.title = '商品呼出番号表 ' + monthText(state.month);
    const i = state.months.indexOf(state.month);
    $('#month-prev').disabled = i <= 0; $('#month-next').disabled = i < 0 || i >= state.months.length - 1;
    document.body.classList.toggle('is-view', isView); // 見るだけの人の画面（守りは窓口。画面で隠すだけにしない）
    $('#month-prev').hidden = isView && state.months.length < 2; $('#month-next').hidden = isView && state.months.length < 2; $('#alert-chip').hidden = isView; $('#save-state').hidden = isView; // 店は公開の月が2つ以上あれば選べる
    paintPublish();
    $('#month-add').hidden = isView; // 月を足す・取り込むのは本部と PO（守りは窓口）
    document.querySelector('[data-view="log"]').hidden = isView;
    $('#account-label').textContent = state.principal.email.split('@')[0];
    $('#account-role').textContent = state.principal.email + '（' + state.principal.roleLabel + '）';
    document.querySelector('[data-view="settings"]').hidden = isView; // 設定は本部と PO（パックの値段は本部も・上限の行数と月を消すは PO だけ。守りは窓口）
    $('#app-version').textContent = '版 ' + APP_VERSION;
    if (isView && state.view !== 'sheet') state.view = 'sheet';
    Object.values(views).forEach((v) => v.reset && v.reset());
    setView(state.view, true);
    if (wanted && !data.costs) fetchCostsThenShow(); // この端末で先に入った人がたたんでいて、頼まなかったとき
  }
  /** 開くときの頼み（本部の人で、たたんでいなければ原価も一緒に＝往復を増やさない） */
  async function fetchOpen(month) {
    const data = {}; if (month) data.month = month;
    if (wantsCosts()) data.withCosts = true;
    try { return await api('getWorkspace', data); }
    catch (error) { if (error.code === 'DENIED' && data.withCosts) { delete data.withCosts; return api('getWorkspace', data); } throw error; } // 好みを覚えていない端末の、見るだけの人
  }
  // ------------------------------------------------------------ 店に公開（本部・PO。新しく作った月は非公開から。守りは窓口）
  const PUBLISH_KIND = '店に公開';
  function paintPublish() {
    const b = $('#month-publish'); if (!b) return;
    b.hidden = state.readOnly || !state.month;
    const on = state.published && state.published.has(state.month);
    b.setAttribute('aria-pressed', String(Boolean(on)));
    b.textContent = on ? '店に公開中' : '店に非公開';
    b.title = on ? '店のアカウントに、この月を見せています。押すと非公開に戻します' : 'この月は、店のアカウントには見えません。押すと店に公開します';
  }
  async function togglePublish(to, b) {
    const month = state.month; const done = busyButton(b || $('#month-publish'), to ? '公開しています…' : '非公開にしています…');
    const t0 = performance.now();
    try {
      await serial(() => api('setPublished', { month, published: to, requestId: rid() }));
      if (to) state.published.add(month); else state.published.delete(month);
      perfNote('店に公開を切り替える', t0);
      toast(monthText(month) + (to ? ' を店に公開しました。店の画面で選べます。' : ' を店に非公開にしました。'), false, { label: '元に戻す', run: () => togglePublish(!to) });
    } catch (error) { toast('切り替えられませんでした: ' + error.message, true); }
    finally { done(); paintPublish(); }
  }
  $('#month-publish').addEventListener('click', (e) => togglePublish(!(state.published && state.published.has(state.month)), e.currentTarget));

  async function loadMonth(month) {
    if (hasUnsaved()) { flush(); await save.chain; }
    const t0 = performance.now();
    showProgress(monthText(month) + ' を読み込んでいます…');
    try { applyWorkspace(await fetchOpen(month)); perfNote('月の切り替え', t0); }
    catch (error) { if (error.code === 'SESSION') showLogin(error.message); else toast(error.message, true, { label: 'もう一度', run: () => loadMonth(month) }); }
    finally { hideProgress(); }
  }

  // ------------------------------------------------------------ 画面の切り替え
  function setView(name, force) {
    if (state.readOnly && name !== 'sheet') name = 'sheet'; // 店は番号表だけ
    if (state.view === name && !force) return;
    state.view = name;
    try { store()?.setItem(VIEW_KEY, name); } catch {}
    document.querySelectorAll('#view-tabs button').forEach((b) => { const on = b.dataset.view === name; b.classList.toggle('is-active', on); b.setAttribute('aria-selected', String(on)); });
    // 隠している間も作った物を捨てない（content-visibility＝戻すときに描き直さない）
    ['list', 'sheet', 'log', 'settings'].forEach((v) => { const n = $('#view-' + v); n.hidden = false; n.classList.toggle('is-off', v !== name); });
    if (views[name]) views[name].show();
  }
  document.querySelectorAll('#view-tabs button').forEach((b) => b.addEventListener('click', () => setView(b.dataset.view)));
  $('#month-prev').addEventListener('click', () => { const i = state.months.indexOf(state.month); if (i > 0) loadMonth(state.months[i - 1]); });
  $('#month-next').addEventListener('click', () => { const i = state.months.indexOf(state.month); if (i >= 0 && i < state.months.length - 1) loadMonth(state.months[i + 1]); });
  $('#alert-chip').addEventListener('click', () => { setView('list'); views.list.setChip('アラート'); });

  // ------------------------------------------------------------ 開く・閉じるメニュー
  function bindMenu(button, menu) {
    const close = () => { menu.hidden = true; button.setAttribute('aria-expanded', 'false'); };
    button.addEventListener('click', (e) => { e.stopPropagation(); const open = menu.hidden; document.querySelectorAll('.pop-right').forEach((m) => { m.hidden = true; }); menu.hidden = !open; button.setAttribute('aria-expanded', String(open)); });
    menu.addEventListener('click', (e) => e.stopPropagation());
    document.addEventListener('click', close);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !menu.hidden) { close(); button.focus(); } });
    return close;
  }
  bindMenu($('#account-button'), $('#account-menu'));

  // ------------------------------------------------------------ Excel（1種類・非表示の列あり）。作ったら原価を捨てる（原価をたたんでいれば）
  function download(bytes, name) {
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = el('a'); a.href = url; a.download = name; document.body.append(a); a.click(); a.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  async function exportExcel() {
    const button = $('#excel-button'); const done = busyButton(button, '保存を待っています…');
    const t0 = performance.now();
    const db = state.db; const keep = db.costs;
    try {
      if (state.readOnly) { // 見るだけ: 開くときに受け取った物だけで作る（0往復）。非表示の原価・粗利・定番（包材前後）の列は空
        done.stage('ファイルを作っています…');
        await new Promise((r) => window.setTimeout(r, 0));
        db.type.pictures = {}; C.rowsOf(state.viewPictures).forEach((p) => { db.type.pictures[p['名前']] = p['中身']; });
      } else {
        flush(); await save.chain;
        done.stage('原価を読んでいます…');
        const data = await api('getExportData', { month: state.month, requestId: rid() });
        done.stage('ファイルを作っています…');
        await new Promise((r) => window.setTimeout(r, 0));
        db.type.pictures = {}; C.rowsOf(data.pictures).forEach((p) => { db.type.pictures[p['名前']] = p['中身']; });
        db.costs = new Map(C.rowsOf(data.costs).map((r) => [String(r['商品ID']), r]));
      }
      const bytes = X.writeBook(S.buildBook(db), new Date()); const name = X.fileName(state.month);
      download(bytes, name);
      perfNote('Excel', t0);
      if (IS_DEV) window.__lastExcel = { name, base64: bytesToBase64(bytes) }; // 手元の確かめの時だけ（試験が中身を比べる）
      toast('落としました: ' + name);
    } catch (error) {
      if (error.code === 'SESSION') showLogin(error.message); else toast('Excel を作れませんでした: ' + error.message, true, { label: 'もう一度', run: () => exportExcel() });
    } finally { db.costs = keep; db.type.pictures = {}; done(); }
  }
  function bytesToBase64(bytes) { let s = ''; for (let i = 0; i < bytes.length; i += 32768) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 32768)); return btoa(s); }
  $('#excel-button').addEventListener('click', () => exportExcel());

  // ------------------------------------------------------------ 原価（はじめから表示。たたむと次もたたんだまま＝この端末にアカウントごとに覚える）
  async function showCosts() {
    const t0 = performance.now();
    try {
      const data = await api('getCosts', { month: state.month, requestId: rid() });
      state.db.costs = new Map(C.rowsOf(data.costs).map((r) => [String(r['商品ID']), r]));
      state.showCosts = true;
      setCostFold(state.principal.email, false);
      recompute();
      perfNote('原価を表示', t0);
      return true;
    } catch (error) {
      if (error.code === 'SESSION') showLogin(error.message); else toast('原価を読めませんでした: ' + error.message, true);
      return false;
    }
  }
  function hideCosts() { state.db.costs = null; state.showCosts = false; setCostFold(state.principal.email, true); recompute(); }
  async function fetchCostsThenShow() { if (await showCosts() && views.list) views.list.rebuild(); }
  // ------------------------------------------------------------ 足した商品を取り消す（窓口が確かめる: 足した商品だけ・ほかの月で使っていないこと）
  async function cancelProduct(pid, button) {
    const done = busyButton(button, '取り消しています…');
    flush();
    const t0 = performance.now();
    try {
      const data = await serial(() => api('cancelProduct', { month: state.month, id: String(pid), requestId: rid() }));
      const gone = new Set(data.rows || []);
      const touches = (u) => (u.steps || []).some((st) => st.before.concat(st.after).some((b) => (b.t === '行' && gone.has(String(b.id))) || ((b.t === '商品' || b.t === '原価') && String(b.id) === String(pid))));
      state.undo = state.undo.filter((u) => !touches(u)); state.redo = state.redo.filter((u) => !touches(u)); paintUndo();
      C.applyChanges(state.db, [{ t: '消す商品', id: String(pid) }]);
      if (views.list) views.list.removeRows(String(pid), [...gone].map(String)); // その行だけ消す（全部を作り直さない）
      refresh([]);
      perfNote('取り消す', t0);
      toast('足した商品を取り消しました。元の商品は変わっていません。');
      return true;
    } catch (error) {
      if (error.code === 'SESSION') showLogin(error.message); else toast('取り消せませんでした: ' + error.message, true);
      return false;
    } finally { done(); }
  }

  // ------------------------------------------------------------ ログイン（スマレコ v88・畜産の形）
  function setLoginButton(enabled, label) { const b = $('#login-button'); b.disabled = !enabled; b.textContent = label; }
  function setLoginHelp(t) { $('#login-help').textContent = t || ''; }
  const HELP = { redirect_going: 'Google のログイン画面へ移ります…', redirect_slow: '画面が移らないときは、もう一度押してください。', redirect_failed: 'ログインできませんでした。もう一度お試しください。', redirect_denied: 'ログインが取り消されました。', redirect_state: 'ログインをやり直してください。', popup_closed: 'ログインの窓が閉じられました。もう一度押してください。', popup_failed_to_open: 'ログインの窓が開けません。窓を止める設定を確かめてください。', gis_missing: 'ログインの部品を読み込めません。通信を確かめてください。', slow: '時間がかかっています。もう一度押してください。', unknown: 'ログインできませんでした。もう一度お試しください。' };
  function showLogin(message) {
    state.db = null; state.readOnly = false; document.body.classList.remove('is-view');
    ['list', 'sheet', 'log', 'settings'].forEach((v) => { $('#view-' + v).hidden = true; });
    $('#month-nav').hidden = true; $('#view-tabs').hidden = true; $('#tb-right').hidden = true;
    $('#month-panel').hidden = true; $('#seed-panel').hidden = true;
    $('#login-panel').hidden = false;
    document.title = '商品呼出番号表';
    showNotice(message || '');
    renderDevAccounts();
    ensureBootstrap().catch(() => {});
  }
  async function ensureBootstrap() {
    if (IS_DEV) { setLoginButton(false, '手元の確かめでは Google を使いません'); setLoginHelp('下の見本のアカウントで入ってください。'); return true; }
    if (!GAS_URL_OK) { setLoginButton(false, 'ログインの準備中'); setLoginHelp('窓口の設定が終わると、ログインできます。'); return false; }
    if (state.clientId) { setLoginButton(true, 'Googleでログイン'); setLoginHelp(''); return true; }
    try { const d = await api('bootstrap', {}, false); state.clientId = d.clientId || ''; if (!state.clientId) throw new Error('no client'); setLoginButton(true, 'Googleでログイン'); setLoginHelp(''); return true; }
    catch { setLoginButton(false, 'ログインの準備中'); setLoginHelp('接続先を確かめられません。少し待ってから開き直してください。'); return false; }
  }
  function isStandalone() { try { return window.navigator.standalone === true || window.matchMedia('(display-mode: standalone)').matches; } catch { return false; } }
  function beginLogin() {
    if (state.loggingIn || !state.clientId || !L) return;
    const choice = L.chooseMode({ ua: navigator.userAgent, standalone: isStandalone(), redirectUri: REDIRECT_URI, origin: location.origin, hasUnsaved: hasUnsaved() });
    if (choice.mode === 'redirect') {
      let st; try { st = L.randomState(window.crypto); L.rememberState(store(), st, Date.now()); } catch { setLoginHelp(HELP.redirect_failed); return; }
      state.leaving = true; setLoginButton(false, 'Googleへ移動しています…'); setLoginHelp(HELP.redirect_going);
      location.assign(L.buildRedirectUrl(state.clientId, REDIRECT_URI, st));
      window.setTimeout(() => { state.leaving = false; setLoginButton(true, 'Googleでログイン'); setLoginHelp(HELP.redirect_slow); }, 15000);
    } else startPopupLogin();
  }
  function loadGis() {
    if (window.google?.accounts?.oauth2) return Promise.resolve();
    if (state.gisLoading) return state.gisLoading;
    state.gisLoading = new Promise((resolve, reject) => { const s = document.createElement('script'); s.src = 'https://accounts.google.com/gsi/client'; s.async = true; s.onload = () => resolve(); s.onerror = () => { state.gisLoading = null; reject(new Error('gis')); }; document.head.append(s); });
    return state.gisLoading;
  }
  async function startPopupLogin() {
    if (!window.google?.accounts?.oauth2) { setLoginButton(false, 'ログインの窓を開いています…'); try { await loadGis(); } catch { setLoginButton(true, 'Googleでログイン'); setLoginHelp(HELP.gis_missing); return; } }
    if (!state.tokenClient) state.tokenClient = google.accounts.oauth2.initTokenClient({ client_id: state.clientId, scope: 'openid email profile', prompt: 'select_account', callback: (r) => state.tokenJob && state.tokenJob.onResponse(r), error_callback: (e) => state.tokenJob && state.tokenJob.onError(e) });
    const slow = window.setTimeout(() => { setLoginButton(true, 'Googleでログイン'); setLoginHelp(HELP.slow); }, 20000);
    state.tokenJob = {
      onResponse(r) { window.clearTimeout(slow); state.tokenJob = null; if (r && r.access_token) finishLogin(r.access_token); else { setLoginButton(true, 'Googleでログイン'); setLoginHelp(r && r.error === 'access_denied' ? HELP.redirect_denied : HELP.unknown); } },
      onError(e) { const type = (e && e.type) || 'unknown'; setLoginButton(true, 'Googleでログイン'); if (type === 'popup_closed') { setLoginHelp(HELP.popup_closed); return; } window.clearTimeout(slow); state.tokenJob = null; setLoginHelp(HELP[type] || HELP.unknown); }
    };
    setLoginButton(false, 'ログインの窓を開いています…'); setLoginHelp('');
    try { state.tokenClient.requestAccessToken(); } catch { window.clearTimeout(slow); state.tokenJob = null; setLoginButton(true, 'Googleでログイン'); setLoginHelp(HELP.unknown); }
  }
  /** Google の札を窓口に確かめてもらい、アプリの札を受け取る（Google の札はここで捨てる）。札と中身を1往復で */
  async function finishLogin(accessToken) {
    state.loggingIn = true; showProgress('ログインして、読み込んでいます…');
    const t0 = performance.now();
    try {
      const data = await api('startSession', { accessToken, withWorkspace: true, withCosts: wantsCosts() }, false);
      saveSession({ t: data.session, e: Number(data.expiresAt) || Date.now() + 7 * 86400000 });
      applyWorkspace(data.workspace);
      perfNote('ログイン', t0);
    } catch (error) { clearSession(); showLogin(error.message); }
    finally { state.loggingIn = false; hideProgress(); }
  }
  $('#login-button').addEventListener('click', beginLogin);
  async function signOut() { const s = state.session; if (hasUnsaved()) { flush(); await save.chain; } clearSession(); if (s) { try { await fetch(GAS_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'endSession', data: {}, session: s.t }) }); } catch {} } showLogin('ログアウトしました。'); }
  $('#signout-button').addEventListener('click', signOut);

  // ------------------------------------------------------------ 手元の確かめ: 見本のアカウント（公開用の config.js には無い）
  function renderDevAccounts() {
    if (!IS_DEV) return;
    const accounts = Array.isArray(CONFIG.devAccounts) ? CONFIG.devAccounts : [];
    const box = $('#dev-account-list'); box.replaceChildren();
    accounts.forEach((a) => { const b = el('button', 'btn', a.label); b.type = 'button'; b.dataset.dev = a.key; b.addEventListener('click', () => finishLogin(a.token)); box.append(b); });
    $('#dev-accounts').hidden = false;
    const sw = $('#dev-switch-list'); sw.replaceChildren();
    accounts.forEach((a) => { const b = el('button', '', a.label); b.type = 'button'; b.dataset.devSwitch = a.key; b.addEventListener('click', async () => { $('#account-menu').hidden = true; if (hasUnsaved()) { flush(); await save.chain; } clearSession(); await finishLogin(a.token); }); sw.append(b); });
    $('#dev-switch').hidden = false;
  }
  $('#dev-reset').addEventListener('click', async () => { $('#account-menu').hidden = true; if (hasUnsaved()) { flush(); await save.chain; } await fetch('/__dev/reset', { method: 'POST' }); clearSession(); location.reload(); });

  // ------------------------------------------------------------ ほかの人の直しを1分ごとに
  window.setInterval(async () => {
    if (document.visibilityState !== 'visible' || !state.db || state.readOnly || hasUnsaved() || !state.session) return;
    const month = state.month;
    try { const data = await api('getChanges', { since: state.logRow, withCosts: state.showCosts }); if (month === state.month) applyOthers({ logRow: data.logRow, others: data.changes || [], reload: data.reload }); } catch { /* 次の分でまた見る */ }
  }, 60000);
  window.addEventListener('beforeunload', (e) => { if (hasUnsaved() && !state.leaving) { flush(); e.preventDefault(); e.returnValue = ''; } });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') flush(); });

  async function start() {
    const redirect = L ? L.consumeRedirect(window.__oauthRedirectHash, store() || { getItem: () => null, setItem() {}, removeItem() {} }, Date.now()) : null;
    window.__oauthRedirectHash = null;
    try { const v = store()?.getItem(VIEW_KEY); if (['list', 'sheet', 'log', 'settings'].includes(v)) state.view = v; } catch {}
    renderDevAccounts();
    if (redirect && redirect.accessToken) { await finishLogin(redirect.accessToken); return; }
    state.session = readSession();
    if (state.session) {
      const t0 = performance.now();
      $('#login-panel').hidden = true; if (state.view === 'list' && views.list) views.list.skeleton();
      showProgress('前回のログインで開いています…');
      try { applyWorkspace(await fetchOpen()); perfNote('開く', t0); }
      catch (error) { clearSession(); showLogin(error.message); }
      finally { hideProgress(); }
      return;
    }
    showLogin('');
    if (redirect && redirect.error) setLoginHelp(HELP[redirect.error] || HELP.unknown);
  }

  window.SouzaiApp = { state, views, FORM_CLASS, paintRunsInto, cancelProduct, $, el, esc, C, S, SN, api, rid, toast, spinner, busyButton, change, refresh, recompute, flush, serial, doMove, undoLast, undo, redo, remember, snapLayout, diffLayout, paintUndo, afterFormChange, showCosts, hideCosts, setView, monthText, perfNote, perfLog, kindOfSection, headPidOfRow, IS_DEV, showLogin, loadMonth, hasUnsaved, showProgress, hideProgress };
  // 部品が全部読み込まれてから始める（defer の順に読まれ、DOMContentLoaded は全部のあと）
  window.addEventListener('DOMContentLoaded', () => { start().catch(() => {}); });
})();
