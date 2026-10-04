/* EXCEL FORMULA QUEST — app */
(function () {
  'use strict';
  const F = window.EFQFormula, D = window.EFQData;
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = F.fmt;
  const clone = (o) => JSON.parse(JSON.stringify(o));
  const todayKey = () => new Date().toISOString().slice(0, 10);
  const COMPETENCY = ['BUILD', 'READ', 'PREDICT', 'DETECT', 'REPAIR', 'VERIFY', 'AI REVIEW', 'TRANSFER'];
  const MASTERY = ['NEW', 'DISCOVERED', 'PRACTICED', 'INDEPENDENT', 'VERIFIED', 'TRANSFERRED'];
  const MASTERY_JA = { NEW: '未体験', DISCOVERED: '意味を理解', PRACTICED: '支援付きで成功', INDEPENDENT: 'ヒントなしで成功', VERIFIED: '検証まで成功', TRANSFERRED: '実務相当で成功' };
  const LEVELS = [
    { k: 'A', name: 'FULL GUIDE', d: '意味＋構文＋候補' },
    { k: 'B', name: 'PARTIAL', d: '一部だけ空欄' },
    { k: 'C', name: 'FUNCTION ONLY', d: '関数名と構文だけ' },
    { k: 'D', name: 'JOB ONLY', d: '依頼文だけ' },
  ];
  const JOB_NAME = Object.fromEntries(D.jobs.map((j) => [j.id, j.skill]));

  // ===================== STATE =====================
  const KEY = 'efq.v1';
  const fresh = () => ({ skills: {}, mastery: {}, fnsUsed: [], chapters: {}, support: 0, log: [], tomorrow: null, sound: false, narration: false, timer: false, map: null, kpi: {}, realWins: 0 });
  let S;
  try { S = Object.assign(fresh(), JSON.parse(localStorage.getItem(KEY) || '{}')); } catch (e) { S = fresh(); }
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(S)); } catch (e) { /* private mode */ } };
  const sessionStart = Date.now();
  const session = { jobs: new Set(), skills: {}, fns: new Set() };

  function rec(skill, ok, assisted) {
    for (const bag of [S.skills, session.skills]) {
      const s = bag[skill] || (bag[skill] = { ok: 0, ng: 0, as: 0 });
      if (!ok) s.ng++; else if (assisted) s.as++; else s.ok++;
    }
    save();
  }
  function stars(skill, bag = S.skills) {
    const s = bag[skill];
    if (!s) return S.map && S.map[skill] ? S.map[skill] : 1;
    return Math.max(1, Math.min(5, 1 + Math.round(s.ok + s.as * 0.5)));
  }
  function raise(job, level) {
    if (!job) return;
    const cur = MASTERY.indexOf(S.mastery[job] || 'NEW');
    const n = MASTERY.indexOf(level);
    if (level === 'VERIFIED' && cur < 2) return; // 検証は「作れた」後
    if (n > cur) { S.mastery[job] = level; save(); }
  }
  function logWeek(kind) { S.log.push({ d: todayKey(), k: kind }); S.log = S.log.slice(-300); save(); }
  function useFns(list) { for (const f of list) { session.fns.add(f); if (!S.fnsUsed.includes(f)) S.fnsUsed.push(f); } save(); }

  // ===================== SOUND / VOICE =====================
  const SFX = { click: '01_click', correct: '02_correct', discovery: '03_discovery', hint: '04_hint', warning: '05_warning', verified: '06_verified', clear: '07_clear', ai: '08_ai_check' };
  const audioCache = {};
  function sfx(name) {
    if (!S.sound || !SFX[name]) return;
    try { const a = audioCache[name] || (audioCache[name] = new Audio('assets/sfx/' + SFX[name] + '.wav')); a.currentTime = 0; a.volume = 0.6; a.play().catch(() => {}); } catch (e) { /* ignore */ }
  }
  function speak(key) {
    const v = D.voices[key]; if (!v) return;
    showCaption(v.text);
    if (!S.sound || !S.narration) return;
    const fallback = () => {
      if (!('speechSynthesis' in window)) return;
      speechSynthesis.cancel();
      const u = new SpeechSynthesisUtterance(v.text); u.lang = 'ja-JP'; u.rate = 1.05; speechSynthesis.speak(u);
    };
    try { const a = new Audio(v.url); a.play().catch(fallback); a.onerror = fallback; } catch (e) { fallback(); }
  }
  let capTimer;
  function showCaption(text) {
    const el = $('#caption'); if (!el) return;
    el.innerHTML = '<span class="cap-ico">🎙</span>' + esc(text);
    el.classList.add('show'); clearTimeout(capTimer);
    capTimer = setTimeout(() => el.classList.remove('show'), Math.max(4500, text.length * 140));
  }
  function toast(msg, kind) {
    const t = $('#toast'); t.textContent = msg; t.className = 'toast show ' + (kind || '');
    clearTimeout(toast.tm); toast.tm = setTimeout(() => (t.className = 'toast'), 2600);
  }

  // ===================== SHEET RENDERING =====================
  function cellHTML(v, o, colHeader) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'object' && v.err) return '<span class="err">' + esc(v.err) + '</span>';
    if (F.isErr(v)) return '<span class="err">' + esc(v.e) + '</span>';
    if (typeof v === 'number') return esc(Number.isInteger(v) || colHeader === '稼働時間' ? (colHeader === '稼働時間' ? v.toFixed(1) : String(v)) : String(v));
    let s = esc(v);
    if (o.showSpaces) s = s.replace(/^( +)/, (m) => '<span class="sp">' + '␣'.repeat(m.length) + '</span>').replace(/( +)$/, (m) => '<span class="sp">' + '␣'.repeat(m.length) + '</span>');
    if (o.showWide) s = s.replace(/[！-～]/g, (m) => '<span class="wide" title="全角文字">' + m + '</span>');
    return s;
  }
  function sheetHTML(wb, name, o = {}) {
    const sh = wb[name];
    const n = sh.headers.length;
    const cols = Array.from({ length: n }, (_, i) => F.numToCol(i + 1));
    let h = '<div class="sheet-wrap"><table class="sheet" data-sheet="' + esc(name) + '"><thead><tr><th class="corner"></th>';
    cols.forEach((c) => (h += '<th class="colh" data-col="' + c + '" title="クリックで式に ' + c + ':' + c + ' を入れる">' + c + '</th>'));
    h += '</tr></thead><tbody><tr class="hdr" data-row="1"><td class="rn">1</td>';
    sh.headers.forEach((t, i) => (h += '<td data-col="' + cols[i] + '" data-row="1">' + esc(t) + '</td>'));
    h += '</tr>';
    sh.rows.forEach((row, i) => {
      const r = i + 2;
      h += '<tr data-row="' + r + '" class="' + (o.rowClass ? o.rowClass(r) || '' : '') + '"><td class="rn">' + r + '</td>';
      cols.forEach((c, ci) => {
        let v = row[ci];
        let extra = '';
        if (o.computed && o.computed[c]) { const res = F.evaluate(o.computed[c](r), wb, name); v = res.value; extra = ' computed'; }
        const isTextNum = typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v)) && /数/.test(sh.headers[ci] || '');
        const cls = (o.cellClass ? o.cellClass(c, r) || '' : '') + extra + (typeof v === 'number' ? ' num' : '') + (isTextNum ? ' textnum' : '') + ((v && typeof v === 'object') ? ' errc' : '');
        h += '<td data-col="' + c + '" data-row="' + r + '" class="' + cls + '">' + cellHTML(v, o, sh.headers[ci]) + '</td>';
      });
      h += '</tr>';
    });
    h += '</tbody></table></div>';
    return h;
  }
  function tabsHTML(names, active) {
    return '<div class="tabs">' + names.map((n) => '<button class="tab' + (n === active ? ' on' : '') + '" data-tab="' + esc(n) + '">' + esc(n) + '</button>').join('') + '</div>';
  }

  // 式 → 意味ブロック
  const BLOCK_COLORS = 6;
  function formulaBarHTML(formula, wb, sheet, o = {}) {
    const ex = o.plain ? null : F.explain(formula, wb, sheet);
    let code = esc(formula);
    if (ex && ex.fn) {
      // 引数ごとに色付け
      let src = F.normalizeSource(formula.trim()); src = src.replace(/^=+\s*/, '');
      let out = '', last = 0;
      ex.blocks.forEach((b, i) => { const p = b.node.p, e = b.node.end; out += esc(src.slice(last, p)) + '<span class="arg a' + (i % BLOCK_COLORS) + '" data-col="' + (b.col || '') + '" data-sheet="' + esc(b.sheet || '') + '">' + esc(src.slice(p, e)) + '</span>'; last = e; });
      code = '=' + out + esc(src.slice(last));
    }
    let h = '<div class="fbar"><span class="fx">fx</span><code class="fcode">' + code + '</code>' + (o.result !== undefined ? '<span class="fres ' + (F.isErr(o.result) ? 'bad' : '') + '">→ ' + esc(fmt(o.result)) + '</span>' : '') + '</div>';
    if (ex && ex.fn && o.blocks !== false) {
      h += '<div class="blocks">';
      ex.blocks.forEach((b, i) => (h += '<div class="blk a' + (i % BLOCK_COLORS) + '" data-col="' + (b.col || '') + '" data-sheet="' + esc(b.sheet || '') + '"><small>' + esc(b.label) + '</small><b>' + esc(b.meaning) + '</b></div>'));
      h += '</div>';
      if (ex.pairs && ex.pairs.length) h += '<div class="pairs">条件：' + ex.pairs.map((p) => '<span class="pair">' + esc(p) + '</span>').join('<span class="and">かつ</span>') + '</div>';
    }
    return h;
  }
  // ブロックにマウスを乗せると列をハイライト（式→データ）
  function bindBlockHover(root) {
    $$('[data-col]', root).forEach((el) => {
      if (!el.classList.contains('blk') && !el.classList.contains('arg')) return;
      const col = el.dataset.col; if (!col) return;
      const on = () => $$('.sheet td[data-col="' + col + '"]', root).forEach((td) => td.classList.add('colhl'));
      const off = () => $$('.sheet td.colhl', root).forEach((td) => td.classList.remove('colhl'));
      el.addEventListener('mouseenter', on); el.addEventListener('mouseleave', off);
      el.addEventListener('focus', on); el.addEventListener('blur', off);
    });
  }
  function chipMeaning(wb, sheet, text) {
    try {
      const ast = F.parse(text);
      if (ast.k === 'ref') {
        const sh = wb[ast.sheet || sheet]; if (!sh) return '';
        const h = sh.headers[F.colToNum(ast.c1) - 1];
        if (ast.r1 !== null && ast.r1 === ast.r2) return h + '：' + fmt(F.cellValue(sh, ast.c1, ast.r1));
        return (ast.sheet ? ast.sheet + 'の' : '') + h;
      }
      if (ast.k === 'str') return '文字「' + ast.v + '」';
    } catch (e) { /* ignore */ }
    return '';
  }

  // ===================== FORMULA X-RAY =====================
  function mountXray(el, wb, sheet, formula, opts = {}) {
    const res = F.evaluate(formula, wb, sheet);
    const top = res.ast && res.ast.k === 'call' ? res.ast.fn : null;
    let tr = res.trace.find((t) => t.fn === top) || res.trace.find((t) => t.fn === 'XLOOKUP' || t.fn === 'VLOOKUP') || res.trace[0];
    if (!tr) {
      el.innerHTML = '<div class="xray"><div class="xr-head"><span class="xr-badge">FORMULA X-RAY</span></div>' + formulaBarHTML(formula, wb, sheet, { result: res.value }) +
        '<p class="note">この式は配列計算のため、行ごとの分解表示はありません。結果は <b>' + esc(fmt(res.value)) + '</b>。</p></div>';
      return;
    }
    const tsheet = tr.sheet || sheet;
    const stages = [];
    const isLookup = tr.fn === 'XLOOKUP' || tr.fn === 'VLOOKUP';
    if (isLookup) {
      stages.push({ label: '探す値「' + fmt(tr.value) + '」', alive: tr.scanRows, scan: true });
      stages.push({ label: tr.notFound ? '見つからない…' : '同じ値の行を発見', alive: tr.matchRow ? [tr.matchRow] : [], match: true });
      stages.push({ label: tr.notFound ? '→ ' + (F.isErr(tr.result) ? tr.result.e : fmt(tr.result)) : '同じ行の値を持ってくる', alive: tr.matchRow ? [tr.matchRow] : [], glow: true });
    } else {
      stages.push({ label: '全' + tr.scanRows.length + '行', alive: tr.scanRows });
      tr.steps.forEach((s, i) => stages.push({ label: '①②③④⑤'[i] + ' ' + s.label, alive: s.rows, col: s.col }));
      const cnt = /COUNT/.test(tr.fn);
      stages.push({ label: cnt ? '残った行を数える' : (tr.fn === 'AVERAGE' || tr.fn === 'AVERAGEIFS' ? '対象を平均' : '対象だけ合計'), alive: tr.steps.length ? tr.steps[tr.steps.length - 1].rows : tr.scanRows, glow: true });
    }
    let k = opts.startAt != null ? opts.startAt : 0, timer = null;
    const meaning = opts.meaning;
    function eqHTML() {
      if (isLookup) {
        if (tr.notFound) return '<div class="xr-eq bad">「' + esc(fmt(tr.value)) + '」は ' + esc(tsheet) + ' に見つかりません → <b>' + esc(F.isErr(tr.result) ? tr.result.e : fmt(tr.result)) + '</b>' + (tr.hiddenError ? '（IFERRORで ' + esc(tr.hiddenError) + ' が隠されています）' : '') + '</div>';
        return '<div class="xr-eq">' + esc(tr.returnCol + tr.returnRow) + ' → <b>' + esc(fmt(tr.result)) + '</b></div>';
      }
      if (/COUNT/.test(tr.fn)) return '<div class="xr-eq">' + (tr.steps.length ? tr.steps[tr.steps.length - 1].rows.length : tr.rows.length) + ' 行 → <b>' + esc(fmt(tr.result)) + '</b></div>';
      const vals = tr.values.map((v) => fmt(v));
      let s = '<div class="xr-eq">' + (vals.length ? vals.join(' + ') : '対象なし') + ' = <b>' + esc(fmt(tr.result)) + '</b></div>';
      if (tr.skipped && tr.skipped.length) s += '<div class="xr-warn">⚠ 文字として入っている値は合計されません：' + tr.skipped.map((x) => '行' + x.row + '「' + esc(x.v) + '」').join('、') + '</div>';
      return s;
    }
    function render() {
      const st = stages[k];
      const alive = new Set(st.alive);
      const glowRows = new Set(isLookup ? (tr.matchRow ? [tr.returnRow] : []) : tr.rows);
      const valueCol = isLookup ? tr.returnCol : tr.valueCol;
      const off = tr.rowOffset || 0;
      const html = sheetHTML(wb, tsheet, {
        showSpaces: true, showWide: true,
        rowClass: (r) => {
          if (!tr.scanRows.includes(r) && !(st.glow && glowRows.has(r))) return 'out';
          if (st.scan) return 'scan';
          return alive.has(r) || (st.glow && glowRows.has(r)) ? 'alive' : 'dim';
        },
        cellClass: (c, r) => {
          let cls = '';
          if (st.col && c === st.col && alive.has(r)) cls += ' hit';
          if (isLookup && c === tr.lookupCol && (st.scan || (st.match && r === tr.matchRow))) cls += st.scan ? ' scanc' : ' hit';
          if (st.glow && c === valueCol && glowRows.has(r)) cls += ' glow';
          if (!st.glow && !isLookup && k > 0 && c === valueCol && alive.has(r + off)) cls += ' soft';
          return cls;
        },
      });
      el.innerHTML = '<div class="xray">' +
        '<div class="xr-head"><span class="xr-badge">FORMULA X-RAY</span><span class="xr-sub">式 → データ → 結果 を追跡</span></div>' +
        formulaBarHTML(formula, wb, sheet, { result: res.value }) +
        '<ol class="xr-steps">' + stages.map((s, i) => '<li class="' + (i === k ? 'on' : i < k ? 'done' : '') + '" data-k="' + i + '"><span>' + (i + 1) + '</span>' + esc(s.label) + (i > 0 && !isLookup && !s.glow ? '<em>' + s.alive.length + '行</em>' : '') + '</li>').join('<li class="arrow">▶</li>') + '</ol>' +
        html +
        '<div class="xr-reason" aria-live="polite">行をクリックすると「なぜ対象？なぜ対象外？」が分かります。</div>' +
        (st.glow ? eqHTML() + (meaning ? '<div class="xr-meaning">' + esc(meaning) + '</div>' : '') : '') +
        '<div class="xr-ctrl"><button class="btn ghost sm" data-x="prev"' + (k === 0 ? ' disabled' : '') + '>◀ 前</button>' +
        '<button class="btn ghost sm" data-x="play">▶ 自動再生</button>' +
        '<button class="btn sm" data-x="next"' + (k === stages.length - 1 ? ' disabled' : '') + '>次 ▶</button></div></div>';
      bindBlockHover(el);
      $$('[data-x]', el).forEach((b) => (b.onclick = () => { stop(); if (b.dataset.x === 'prev') k--; if (b.dataset.x === 'next') k++; if (b.dataset.x === 'play') { k = 0; play(); } render(); }));
      $$('.xr-steps li[data-k]', el).forEach((li) => (li.onclick = () => { stop(); k = +li.dataset.k; render(); }));
      $$('tbody tr[data-row]', el).forEach((row) => (row.onclick = () => { const r = +row.dataset.row; if (r > 1) $('.xr-reason', el).innerHTML = reason(r); }));
      if (st.glow && opts.onDone) { const f = opts.onDone; opts.onDone = null; f(res); }
    }
    function reason(r) {
      if (isLookup) {
        if (r === tr.matchRow) return '✔ 行' + r + '：探す値「' + esc(fmt(tr.value)) + '」と一致。同じ行の ' + esc(tr.returnCol) + '列 を持ってきます。';
        const v = F.cellValue(wb[tsheet], tr.lookupCol, r);
        return '✖ 行' + r + '：「' + esc(fmt(v)) + '」は探す値「' + esc(fmt(tr.value)) + '」と一致しません。';
      }
      if (!tr.scanRows.includes(r)) return '✖ 行' + r + '：式の範囲に入っていません（範囲の外）。';
      for (const s of tr.steps) if (!s.rows.includes(r)) {
        const v = F.cellValue(wb[s.sheet || tsheet], s.col, r);
        return '✖ 行' + r + '：' + esc(s.col) + '列が「' + esc(fmt(v)).replace(/^ | $/g, '␣') + '」なので「' + esc(s.label) + '」に合わず対象外。';
      }
      const sk = (tr.skipped || []).find((x) => x.row === r + (tr.rowOffset || 0));
      if (sk) return '⚠ 行' + r + '：条件は合うけれど、値「' + esc(sk.v) + '」が文字なので合計されません。';
      return '✔ 行' + r + '：すべての条件に合うので対象。' + (tr.valueCol ? ' ' + esc(tr.valueCol) + (r + (tr.rowOffset || 0)) + ' = ' + esc(fmt(F.cellValue(wb[tsheet], tr.valueCol, r + (tr.rowOffset || 0)))) : '');
    }
    function play() { stop(); timer = setInterval(() => { if (k >= stages.length - 1) return stop(); k++; render(); if (k === stages.length - 1) { sfx('discovery'); stop(); } }, 1200); }
    function stop() { if (timer) clearInterval(timer); timer = null; }
    render();
    if (opts.auto !== false) play();
    return { res, stop };
  }

  // ===================== SCREENS =====================
  const app = () => $('#app');
  let run = null; // 現在のミッション
  let finalTimer = null;

  function setTop(title, extra) {
    $('#topTitle').innerHTML = title ? esc(title) : '';
    $('#topExtra').innerHTML = extra || '';
    $('#exitBtn').hidden = !run;
  }

  // ---------- HOME ----------
  function home() {
    stopRun();
    setTop('');
    const tm = S.tomorrow;
    const week = weekSummary();
    const chaptersDone = D.chapters.filter((c) => S.chapters[c]).length;
    app().innerHTML = `
      <section class="hero">
        <div class="hero-bg"></div>
        <div class="hero-inner">
          <div class="hero-text">
            <p class="eyebrow">Just-in-Time Excel Performance Support</p>
            <h1><span class="h-excel">EXCEL</span> <span class="h-formula">FORMULA</span> <span class="h-quest">QUEST</span></h1>
            <p class="tagline">関数を覚えるな。仕事を解決せよ。</p>
          </div>
          <div class="hero-chars" aria-hidden="true">
            <img src="assets/img/chars/support.jpg" alt="" class="hc hc2">
            <img src="assets/img/chars/navi.jpg" alt="" class="hc hc1">
          </div>
        </div>
      </section>
      ${tm ? `<section class="panel tomorrow">
        <div class="tm-head"><span class="badge yellow">TOMORROW MISSION</span><span class="muted">${esc(tm.date)} に設定</span></div>
        <p class="tm-mission">${esc(tm.mission)}</p>
        <div class="row gap">
          <button class="btn good" id="tmDone">✓ 自分のExcelで使えた！</button>
          <button class="btn ghost" id="tmHelp">⚡ 困った → 10秒RESCUE</button>
          <button class="btn ghost" id="tmLater">まだ使っていない</button>
        </div></section>` : ''}
      <section class="entry">
        <h2 class="entry-q">今日、どのExcel仕事をラクにする？</h2>
        <div class="jobs">
          ${D.jobs.map((j) => `<button class="job" data-job="${j.id}">
            <img src="assets/img/icons/${j.icon}" alt="" class="job-ico">
            <span class="job-t">${esc(j.title)}</span>
            <span class="job-s">${esc(j.sub)}</span>
            <span class="job-m">${masteryDots(j.id)}</span>
          </button>`).join('')}
        </div>
      </section>
      <section class="home-grid">
        <button class="panel action" id="goChallenge"><span class="a-ico">⏱</span><b>60秒 WORK CHALLENGE</b><span>テストではありません。あなたに役立つMISSIONを探します。</span>${S.map ? '<em>前回のMAPあり</em>' : ''}</button>
        <button class="panel action" id="goMap"><span class="a-ico">🗺</span><b>QUEST MAP</b><span>FIRST SUCCESS → … → THE BROKEN REPORT → BRIDGE</span><em>${chaptersDone} / ${D.chapters.length} 章</em></button>
        <button class="panel action" id="goRescue"><span class="a-ico">⚡</span><b>10秒 RESCUE</b><span>仕事中に式を忘れたら。コピーしてすぐ戻れる。</span></button>
        ${S.mastery.sum || Object.keys(S.mastery).length ? `<button class="panel action" id="goCard"><span class="a-ico">🪪</span><b>FORMULA CARD</b><span>できるようになった仕事を見る</span></button>` : ''}
      </section>
      ${week ? weekHTML(week) : ''}
      <footer class="principles">
        <span>JOB FIRST</span><span>LOGIC BEFORE FORMULA</span><span>NO ERROR ≠ CORRECT</span><span>VERIFY BEFORE TRUST</span><span>FAILURE = DISCOVERY</span><span>競争相手は、少し前の自分。</span>
      </footer>`;
    $$('.job').forEach((b) => (b.onclick = () => { sfx('click'); const j = D.jobs.find((x) => x.id === b.dataset.job); startMission(j.mission); }));
    $('#goChallenge').onclick = () => { sfx('click'); challenge(); };
    $('#goMap').onclick = () => { sfx('click'); questMap(); };
    $('#goRescue').onclick = () => openRescue();
    if ($('#goCard')) $('#goCard').onclick = () => { stopRun(); app().innerHTML = '<div class="mission solo"></div>'; renderCard($('.mission'), { standalone: true }); };
    if (tm) {
      $('#tmDone').onclick = () => realWin();
      $('#tmHelp').onclick = () => openRescue(tm.job === 'lookup' ? 'xlookup' : tm.job === 'judge' ? 'if' : tm.job === 'error' ? 'iferror' : tm.job === 'clean' ? 'trim' : 'sumifs');
      $('#tmLater').onclick = () => toast('大丈夫。明日の仕事で1回だけ試してみよう。');
    }
    $$('.week-next').forEach((b) => (b.onclick = () => startMission(b.dataset.m)));
    window.scrollTo(0, 0);
  }
  function masteryDots(job) {
    const m = MASTERY.indexOf(S.mastery[job] || 'NEW');
    return MASTERY.slice(1).map((x, i) => '<i class="' + (i < m ? 'on' : '') + '" title="' + MASTERY_JA[x] + '"></i>').join('') + (m > 0 ? '<small>' + MASTERY_JA[MASTERY[m]] + '</small>' : '');
  }
  function realWin() {
    const tm = S.tomorrow; if (!tm) return;
    S.realWins++; raise(tm.job, 'TRANSFERRED'); logWeek(tm.job); logWeek('real');
    S.tomorrow = null; save(); sfx('clear');
    modal(`<div class="realwin"><div class="big-emoji">🏭</div><h2>REAL-WORK SUCCESS</h2>
      <p>学習の終了をCLEARとしない。<br><b>実務で使えた瞬間を、本当のCLEARとする。</b></p>
      <p class="muted">「${esc(tm.mission)}」達成！</p>
      <button class="btn primary" data-close>次の +1 SKILL へ</button></div>`, home);
  }

  // ---------- WEEK ----------
  function weekSummary() {
    const since = new Date(Date.now() - 6 * 864e5).toISOString().slice(0, 10);
    const recent = S.log.filter((x) => x.d >= since);
    if (!recent.length) return null;
    const groups = { sum: '条件集計', lookup: 'Lookup', judge: '自動判定', count: '件数集計', clean: 'データ整形', error: 'エラー確認', verify: 'Verify', ai: 'AI確認', rescue: 'RESCUE', real: '実務で使用' };
    const c = {};
    recent.forEach((x) => { if (groups[x.k]) c[x.k] = (c[x.k] || 0) + 1; });
    const jobKeys = Object.keys(c).filter((k) => JOB_NAME[k]);
    const top = jobKeys.sort((a, b) => c[b] - c[a])[0];
    const NEXT = { sum: ['count', 'COUNTIFS（件数集計）'], lookup: ['error', 'IFERRORの正しい使い方'], judge: ['sum', 'SUMIFS（条件集計）'], count: ['clean', 'TRIM（データ整形）'], clean: ['lookup', 'XLOOKUP（マスター検索）'], error: ['c5', 'AI REVIEW'] };
    return { c, groups, top, next: top ? NEXT[top] : ['c1', 'FIRST SUCCESS'] };
  }
  function weekHTML(w) {
    return `<section class="panel week"><h3>YOUR WEEK</h3><div class="week-rows">` +
      Object.keys(w.c).map((k) => `<div class="wr"><span>${esc(w.groups[k])}</span><span class="checks">${'✓'.repeat(Math.min(w.c[k], 12))}</span></div>`).join('') +
      `</div>${w.top ? `<p>よく使った仕事：<b>${esc(JOB_NAME[w.top])}</b></p>` : ''}
      <div class="row gap center-v"><span>次のおすすめ：<b>${esc(w.next[1])}</b> <span class="muted">約30秒〜</span></span>
      <button class="btn sm week-next" data-m="${w.next[0]}">やってみる ▶</button></div></section>`;
  }

  // ---------- QUEST MAP ----------
  function questMap() {
    stopRun(); setTop('QUEST MAP');
    const desc = { c1: 'かんたんな成功 → 小さな挑戦 → AI CHECK', c2: 'WHAT → LOGIC → CHOOSE → BUILD → X-RAY → READ → PREDICT', c3: 'サイレントエラー発見 → DETECT → REPAIR → 10秒VERIFY', c4: '#N/A の原因診断 → RECOVERY', c5: 'AIの式を ACCEPT / CHECK / REPAIR / REJECT', c6: '15:40 → 16:00 数字が合わない会議資料を救え', c7: '実務に近い汚れたデータ → FORMULA CARD → MY WORK' };
    app().innerHTML = `<section class="qmap"><h2>QUEST MAP</h2><p class="muted">Easy Success → Small Challenge → Trap → Discovery → Recovery → Hard Challenge</p><ol class="chapters">` +
      D.chapters.map((c, i) => { const m = D.missions[c]; return `<li><button class="chapter ${S.chapters[c] ? 'done' : ''}" data-c="${c}"><span class="ch-n">${i + 1}</span><span class="ch-t"><b>${esc(m.title)}</b><small>${esc(desc[c])}</small></span><span class="ch-s">${S.chapters[c] ? '✓ CLEAR' : '▶'}</span></button></li>`; }).join('') +
      `</ol><button class="btn ghost" id="mapBack">← トップへ</button></section>`;
    $$('.chapter').forEach((b) => (b.onclick = () => startMission(b.dataset.c)));
    $('#mapBack').onclick = home;
  }

  // ---------- 60 SEC WORK CHALLENGE ----------
  function challenge() {
    stopRun(); setTop('60 SEC WORK CHALLENGE');
    app().innerHTML = `<section class="panel challenge-intro"><img src="assets/img/chars/navi.jpg" alt="" class="avatar lg">
      <div><h2>60秒 WORK CHALLENGE</h2><p><b>正解数を測るテストではありません。</b><br>あなたに役立つMISSIONを探します。</p>
      <p class="muted">7問・わからなければ「わからない」を押してOK。時間は目安です。</p>
      <button class="btn primary" id="chStart">はじめる</button> <button class="btn ghost" id="chBack">やめる</button></div></section>`;
    $('#chBack').onclick = home;
    $('#chStart').onclick = () => {
      const ans = [];
      let i = 0;
      const t0 = Date.now();
      const tick = setInterval(() => { const el = $('#chTime'); if (!el) return clearInterval(tick); const left = Math.max(0, 60 - Math.floor((Date.now() - t0) / 1000)); el.style.width = (left / 60 * 100) + '%'; $('#chSec').textContent = left ? left + '秒' : '時間は目安。続けてOK'; }, 250);
      const q = () => {
        const it = D.challenge[i];
        app().innerHTML = `<section class="panel challenge"><div class="ch-top"><span class="badge blue">${esc(it.skill)}</span><span>${i + 1} / ${D.challenge.length}</span><span id="chSec" class="muted"></span></div>
          <div class="timebar"><i id="chTime"></i></div><h3>${esc(it.q)}</h3>
          <div class="opts">${it.options.map((o, k) => `<button class="opt" data-k="${k}">${esc(o)}</button>`).join('')}<button class="opt idk" data-k="-1">わからない</button></div></section>`;
        $$('.opt').forEach((b) => (b.onclick = () => { sfx('click'); ans.push({ skill: it.skill, r: +b.dataset.k === it.a ? 1 : +b.dataset.k === -1 ? 0 : -1 }); i++; if (i < D.challenge.length) q(); else { clearInterval(tick); result(ans); } }));
      };
      q();
    };
    function result(ans) {
      const map = {};
      ans.forEach((a) => (map[a.skill] = a.r === 1 ? 4 : 2));
      S.map = map; save();
      const weak = ans.filter((a) => a.r !== 1).map((a) => D.challengeMap[a.skill]);
      const recs = [...new Set(weak)].slice(0, 3);
      if (!recs.length) recs.push('c3', 'c5', 'c6');
      const label = { c1: 'FIRST SUCCESS：まず1式を完成', c2: 'X-RAY MISSION：何を・どの条件で', c3: 'SILENT ERROR：正常な式を疑う', c4: 'DEBUG：エラーの原因を特定', c5: 'AI REVIEW：AIの式を検証', c6: 'THE BROKEN REPORT：総合実戦' };
      app().innerHTML = `<section class="panel challenge-result"><h2>YOUR FORMULA MAP</h2>
        <div class="map">${ans.map((a) => `<div class="mrow"><span>${esc(a.skill)}</span><span class="stars">${starHTML(map[a.skill])}</span></div>`).join('')}</div>
        <h3>今日おすすめのMISSION</h3><div class="recs">${recs.map((r) => `<button class="btn rec" data-m="${r}">${esc(label[r] || D.missions[r].title)} ▶</button>`).join('')}</div>
        <p class="muted">★は「今の得意・不得意の目安」。ここから伸ばしていきます。</p>
        <button class="btn ghost" id="chHome">トップへ</button></section>`;
      $$('.rec').forEach((b) => (b.onclick = () => startMission(b.dataset.m)));
      $('#chHome').onclick = home;
    }
  }
  const starHTML = (n) => '<span class="on">' + '★'.repeat(n) + '</span><span class="off">' + '★'.repeat(5 - n) + '</span>';

  // ===================== MISSION RUNNER =====================
  function startMission(id) {
    stopRun();
    const m = D.missions[id];
    run = { id, m, i: 0, wb: clone(D.workbook), fails: 0, gained: [], reportFix: {}, t0: Date.now() };
    if (m.final && S.timer) run.deadline = Date.now() + 5 * 60 * 1000;
    session.jobs.add(m.job);
    logWeek(m.job);
    renderStep();
  }
  function stopRun() {
    run = null;
    if (finalTimer) { clearInterval(finalTimer); finalTimer = null; }
  }
  function flowChips(step) {
    const flow = ['REQUEST', 'WHAT', 'LOGIC', 'CHOOSE', 'BUILD', 'X-RAY', 'RESULT', 'VERIFY'];
    const tag = step.tag || ({ build: 'BUILD', xray: 'X-RAY', verify: 'VERIFY', repair: 'BUILD', talk: 'REQUEST' })[step.type];
    const idx = flow.indexOf(tag);
    if (idx < 0) return '';
    return '<ol class="flow">' + flow.map((f, i) => '<li class="' + (i === idx ? 'on' : i < idx ? 'done' : '') + '">' + f + '</li>').join('') + '</ol>';
  }
  const GUIDE = {
    WHAT: 'まず「何を求めるか」だけ考えよう。関数名はまだいらないよ。', LOGIC: '依頼文の中から「条件」を探そう。', CHOOSE: '道具は「何をしたいか」で選ぶよ。',
    BUILD: '迷ったらヒントを1つずつ開いてね。答えをすぐ見なくて大丈夫。', READ: '式を左から「意味」に翻訳してみよう。', PREDICT: 'ざっくりでOK。予想してから実行すると、検算の力がつくよ。',
    DETECT: 'エラーが出ていない＝正しい、とは限らないよ。', REPAIR: '直したら、もう一度結果を確かめよう。', VERIFY: '10秒でできる検証。対象・条件・数件チェック！',
    DEBUG: 'エラーは手がかり。消す前に、原因を特定しよう。', 'AI REVIEW': 'AIは使ってOK。でも確認せずに信用しない。', TRANSFER: '見た目が違っても、JobとLogicは同じ。',
  };
  function renderStep() {
    if (!run) return;
    const step = run.m.steps[run.i];
    if (!step) return home();
    run.hint = 0; run.stepFails = 0; run.shownAnswer = false; run.bridgeHints = null;
    const total = run.m.steps.length;
    const dots = run.m.steps.map((_, i) => '<i class="' + (i < run.i ? 'done' : i === run.i ? 'on' : '') + '"></i>').join('');
    const clock = step.clock ? '<span class="clock" id="clock">' + esc(step.clock) + '</span>' : '';
    setTop((run.m.chapter ? 'CH.' + run.m.chapter + ' ' : '') + run.m.title, '<span class="dots" title="' + (run.i + 1) + '/' + total + '">' + dots + '</span>');
    const skill = step.skill || step.tag;
    const level = currentLevel(step);
    app().innerHTML = `<section class="mission ${run.m.final ? 'final' : ''}">
      <div class="mhead">${clock}<div class="mtitle"><span class="badge ${step.recovery ? 'green' : run.m.final ? 'red' : 'blue'}">${esc(step.recovery ? 'RECOVERY' : step.title || (step.type === 'talk' ? 'MISSION' : step.type.toUpperCase()))}</span>${run.deadline ? '<span class="countdown" id="countdown"></span>' : ''}</div>${flowChips(step)}</div>
      <div class="mgrid">
        <div class="work" id="work"></div>
        <aside class="guide" id="guide">
          <div class="g-char"><img src="${D.chars.navi.img}" alt="ナビゲーター ミライ" class="avatar"><div class="bubble" id="gsay">${esc(GUIDE[skill] || '一緒に解決しよう！')}</div></div>
          ${['build', 'repair', 'bridge', 'ai'].includes(step.type) || (step.type === 'detect' && step.repairTo) ? `<div class="hints" id="hints"${step.type === 'ai' ? ' hidden' : ''}><div class="h-head"><b>PROGRESSIVE HINT</b><button class="btn sm ghost" id="hintBtn">💡 HINT 1</button></div><ol id="hintList"></ol></div>` : ''}
          ${step.type === 'build' && step.template && step.forceLevel === undefined ? `<div class="support"><b>SUPPORT</b> <span class="lv">${LEVELS[level].k}｜${LEVELS[level].name}</span><small>${LEVELS[level].d}</small>
            <div class="row gap"><button class="btn sm ghost" id="supUp" ${level === 0 ? 'disabled' : ''}>支援を増やす</button><button class="btn sm ghost" id="supDown" ${level === 3 ? 'disabled' : ''}>自分で書く</button></div></div>` : ''}
          <button class="btn sm ghost rescue-mini" id="gRescue">⚡ 10秒RESCUE</button>
        </aside>
      </div></section>`;
    $('#gRescue').onclick = () => openRescue();
    if ($('#supUp')) { $('#supUp').onclick = () => { S.support = Math.max(0, level - 1); save(); renderStep(); }; $('#supDown').onclick = () => { S.support = Math.min(3, level + 1); save(); renderStep(); }; }
    if (step.voice && ['talk', 'xray'].includes(step.type)) speak(step.voice);
    if (run.deadline && !finalTimer) finalTimer = setInterval(updateCountdown, 500);
    updateCountdown();
    const work = $('#work');
    const R = RENDER[step.type];
    R(work, step);
    bindBlockHover(work);
    window.scrollTo(0, 0);
  }
  function updateCountdown() {
    const el = $('#countdown'); if (!el || !run || !run.deadline) return;
    const left = Math.max(0, run.deadline - Date.now());
    el.textContent = left ? '残り ' + Math.floor(left / 60000) + ':' + String(Math.floor(left / 1000) % 60).padStart(2, '0') : '時間切れ…でも大丈夫。最後まで確認しよう';
    el.classList.toggle('over', !left);
  }
  function guideSay(text, who) {
    const g = $('#gsay'); if (!g) return;
    g.textContent = text;
    const img = $('.g-char img');
    if (img && who && D.chars[who]) img.src = D.chars[who].img;
  }
  function currentLevel(step) {
    if (step.forceLevel !== undefined) return step.forceLevel;
    const job = step.job || (run && run.m.job);
    const m = MASTERY.indexOf(S.mastery[job] || 'NEW');
    // 初めての仕事では支援を残す（A/B）、ヒントなし成功までは C まで
    const cap = m < 2 ? 1 : m < 3 ? 2 : 3;
    if (step.recovery) return Math.min(S.support, 1, cap);
    return Math.min(S.support, cap);
  }
  function next() {
    if (!run) return;
    sfx('click');
    run.i++;
    if (run.i >= run.m.steps.length) return finishMission();
    renderStep();
  }
  function finishMission() {
    const m = run.m;
    S.chapters[run.id] = true; save();
    if (m.next) return startMission(m.next);
    home();
  }
  // 失敗の記録と感情適応
  function fail(skill) {
    rec(skill, false);
    run.fails++; run.stepFails++;
    sfx('warning');
    if (run.fails >= 3) {
      run.fails = 0;
      const prev = S.support;
      S.support = 0; save();
      openHints(1);
      modal(`<div class="adaptive"><img src="${D.chars.navi.img}" class="avatar lg" alt=""><div><p class="say">式を覚えなくて大丈夫。<br>まず<b>“何を求めたいか”</b>を確認しよう。</p>
        <p class="muted">サポートを ${LEVELS[prev].k} → A（FULL GUIDE）に戻し、ヒントを開きました。うまくいったら、また少しずつ減らしていきます。</p>
        <button class="btn primary" data-close>OK、もう一度</button></div></div>`, () => { const st = run && run.m.steps[run.i]; if (st && st.type === 'build' && st.template && prev > 0) renderStep(); });
    } else if (run.stepFails === 2) {
      guideSay('ヒントを1つ開いてみよう。考え方からでOK。');
      openHints(Math.max(1, run.hint));
    }
  }
  function succeed(skill, assisted) {
    rec(skill, true, assisted || run.hint >= 3 || run.shownAnswer);
    run.fails = 0;
  }
  function openHints(n) {
    const step = run && run.m.steps[run.i]; if (!step) return;
    const hints = run.bridgeHints || step.hints || (step.repairTo && step.repairTo.hints);
    const list = $('#hintList'); if (!list || !hints) return;
    run.hint = Math.max(run.hint, Math.min(n, 5));
    const names = ['考え方', '関数カテゴリー', '関数名', '構文', '完成式'];
    list.innerHTML = hints.slice(0, run.hint).map((h, i) => '<li class="h' + i + '"><small>HINT ' + (i + 1) + '｜' + names[i] + '</small>' + (i >= 3 ? '<code>' + esc(h) + '</code>' : esc(h)) + '</li>').join('');
    const b = $('#hintBtn');
    if (b) { b.textContent = run.hint >= 5 ? '✓ 全HINT表示' : '💡 HINT ' + (run.hint + 1); b.disabled = run.hint >= 5; }
    if (run.hint >= 5) run.shownAnswer = true;
  }
  function bindHints() {
    const b = $('#hintBtn'); if (!b) return;
    b.onclick = () => { sfx('hint'); openHints(run.hint + 1); if (run.hint === 5) guideSay('完成式を見ても大丈夫。ただし「見ただけ」では習得にしないよ。次の問題で自力に挑戦しよう。'); };
  }

  // ===================== STEP RENDERERS =====================
  const RENDER = {};

  RENDER.talk = (w, s) => {
    const c = D.chars[s.who];
    w.innerHTML = `<div class="talk ${s.who}"><img src="${c.img}" alt="${esc(c.name)}" class="talk-img">
      <div class="talk-body"><div class="talk-name">${esc(c.name)}<small>${esc(c.role)}</small></div><p class="talk-text">${esc(s.text)}</p>
      ${s.finalIntro ? `<div class="final-opts"><label class="switch"><input type="checkbox" id="timerToggle" ${S.timer ? 'checked' : ''}> 時間制限あり（5分・演出用）</label><small class="muted">初心者は OFF のままでOK。時間切れでも止まりません。</small></div>` : ''}
      <button class="btn primary" id="go">${s.finalIntro ? 'レポートを確認する' : '次へ ▶'}</button></div></div>`;
    if (s.who !== 'navi') guideSay(s.who === 'boss' ? 'まず依頼を読もう。「何を」「どの条件で」がカギ。' : s.who === 'ai' ? 'AIは使ってOK。でも確認せずに信用しない。' : '一緒に解決しよう！');
    if (s.finalIntro) $('#timerToggle').onchange = (e) => { S.timer = e.target.checked; save(); run.deadline = S.timer ? Date.now() + 5 * 60 * 1000 : null; if (!S.timer) { clearInterval(finalTimer); finalTimer = null; } };
    $('#go').onclick = () => { if (s.finalIntro && S.timer && !run.deadline) run.deadline = Date.now() + 5 * 60 * 1000; next(); };
  };

  function requestHTML(s) {
    return s.request ? `<div class="request"><span class="rq-tag">REQUEST</span><p>${esc(s.request)}</p></div>` : '';
  }
  function reportHTML(s) {
    if (s.report === undefined && !run.m.final) return '';
    const rep = run.wb['月次レポート'];
    return `<div class="report"><div class="rp-head">📊 月次レポート（16:00 会議用）</div>` + rep.rows.map((r, i) => {
      const f = run.reportFix[i] || r[2];
      const v = F.evaluate(f, run.wb, '月次レポート').value;
      const st = run.reportFix[i] ? 'fixed' : i === s.report ? 'cur' : '';
      return `<div class="rp-row ${st}"><span>${esc(r[0])}</span><b>${esc(fmt(v))}</b><span class="rp-st">${run.reportFix[i] ? '✓ 修正済' : i === s.report ? '🔎 確認中' : ''}</span></div>`;
    }).join('') + '</div>';
  }

  // ---------- CHOICE (WHAT / LOGIC / CHOOSE / READ / PREDICT) ----------
  RENDER.choice = (w, s) => {
    const showSheet = s.sheet && s.tag !== 'DETECT';
    const res = s.formula ? F.evaluate(s.formula, run.wb, s.sheet) : null;
    const hideResult = s.tag === 'PREDICT' || s.tag === 'READ';
    w.innerHTML = reportHTML(s) + requestHTML(s) +
      (s.formula ? formulaBarHTML(s.formula, run.wb, s.sheet, { plain: true, result: hideResult ? undefined : res.value }) : '') +
      `<h3 class="prompt">${esc(s.prompt)}</h3>
      <div class="opts ${s.options.some((o) => o.code) ? 'tools' : ''}">${s.options.map((o, k) => `<button class="opt" data-k="${k}">${o.code ? `<span class="tool-t">${esc(o.t)}</span><code>${esc(o.code)}</code>` : esc(o.t)}</button>`).join('')}</div>
      ${s.multi ? '<button class="btn primary" id="multiOk">これで決定</button>' : ''}
      <div class="feedback" id="fb" aria-live="polite"></div><div id="after"></div>
      ${showSheet ? '<div class="sheet-area">' + tabsHTML([s.sheet], s.sheet) + sheetHTML(run.wb, s.sheet) + '</div>' : ''}`;
    const fb = $('#fb', w);
    const done = (msg) => {
      fb.className = 'feedback good'; fb.innerHTML = '✓ ' + esc(msg || '');
      sfx('correct'); succeed(s.skill || s.tag);
      $$('.opt', w).forEach((b) => (b.disabled = true));
      if ($('#multiOk', w)) $('#multiOk', w).disabled = true;
      const after = $('#after', w);
      if (s.reveal) after.innerHTML += '<div class="reveal">🔍 ' + esc(s.reveal) + '</div>';
      if (s.report !== undefined && s.report === 2) { run.reportFix[2] = '=XLOOKUP("B-003",マスター!A:A,マスター!E:E,"マスター未登録")'; }
      if (s.formula && hideResult) after.innerHTML += formulaBarHTML(s.formula, run.wb, s.sheet, { result: res.value });
      if (s.xrayAfter) { const x = document.createElement('div'); after.appendChild(x); mountXray(x, run.wb, s.sheet, s.formula, {}); }
      after.insertAdjacentHTML('beforeend', '<button class="btn primary next" id="nx">次へ ▶</button>');
      bindBlockHover(after);
      $('#nx', w).onclick = next;
    };
    if (s.multi) {
      $$('.opt', w).forEach((b) => (b.onclick = () => { sfx('click'); b.classList.toggle('sel'); }));
      $('#multiOk', w).onclick = () => {
        const sel = $$('.opt', w).map((b) => b.classList.contains('sel'));
        const ok = s.options.every((o, k) => !!o.ok === sel[k]);
        if (ok) done(s.fbOk); else { fb.className = 'feedback bad'; fb.textContent = s.fbNg; fail(s.skill || s.tag); }
      };
    } else {
      $$('.opt', w).forEach((b) => (b.onclick = () => {
        const o = s.options[+b.dataset.k];
        if (o.ok) { b.classList.add('right'); done(o.fb); if (o.code) useFns([o.code]); }
        else { b.classList.add('wrong'); b.disabled = true; fb.className = 'feedback bad'; fb.textContent = o.fb || 'もう一度考えてみよう。'; fail(s.skill || s.tag); }
      }));
    }
  };

  // ---------- BUILD ----------
  RENDER.build = (w, s) => {
    const level = currentLevel(s);
    const tpl = s.template;
    const useSlots = tpl && level <= 1;
    const slots = useSlots ? tpl.slots.map((sl, i) => (level === 1 && i < Math.floor(tpl.slots.length / 2) ? sl.ans[0] : '')) : null;
    const locked = useSlots ? tpl.slots.map((_, i) => level === 1 && i < Math.floor(tpl.slots.length / 2)) : null;
    let active = useSlots ? slots.findIndex((x) => !x) : -1;
    const startText = s.type === 'repair' ? s.formula : level === 2 && tpl ? '=' + tpl.fn + '(' : '=';
    w.innerHTML = reportHTML(s) + requestHTML(s) +
      (s.type === 'repair' ? `<div class="broken">${formulaBarHTML(s.formula, run.wb, s.sheet, { result: F.evaluate(s.formula, run.wb, s.sheet).value })}</div><h3 class="prompt">式を直して、正しい結果にしよう。</h3>` : '') +
      `<div class="builder" id="builder"></div>
      <div class="row gap"><button class="btn primary" id="runBtn">▶ 実行</button><span class="muted small">Enterでも実行。列の文字（A・B・C…）をクリックすると範囲を入力できます。</span></div>
      <div class="feedback" id="fb" aria-live="polite"></div><div id="after"></div>
      <div class="sheet-area">${tabsHTML([s.sheet], s.sheet)}${sheetHTML(run.wb, s.sheet, { showSpaces: s.sheet === '受入データ' })}</div>`;
    const builder = $('#builder', w);
    let input = null;
    function drawSlots() {
      const preview = '=' + tpl.fn + '(' + slots.map((x, i) => (x ? '<span class="arg a' + (i % BLOCK_COLORS) + '">' + esc(x) + '</span>' : '<span class="hole' + (i === active ? ' act' : '') + '">　?　</span>')).join(',') + ')';
      builder.innerHTML = `<div class="lv-tag">SUPPORT ${LEVELS[level].k}｜${LEVELS[level].name}</div>
        <div class="fbar big"><span class="fx">fx</span><code class="fcode">${preview}</code></div>
        <div class="slots">${tpl.slots.map((sl, i) => `<div class="slot a${i % BLOCK_COLORS} ${i === active ? 'act' : ''} ${locked[i] ? 'locked' : ''}" data-i="${i}">
          <small>${esc(sl.label)}</small>
          ${locked[i] ? `<b>${esc(slots[i])}</b><em>${esc(chipMeaning(run.wb, s.sheet, slots[i]))}</em>` :
            level === 0 ? `<b>${slots[i] ? esc(slots[i]) : '—'}</b><em>${esc(slots[i] ? chipMeaning(run.wb, s.sheet, slots[i]) : 'タップして選ぶ')}</em>` :
            `<input class="slot-in" data-i="${i}" value="${esc(slots[i])}" placeholder="?" aria-label="${esc(sl.label)}" autocomplete="off" spellcheck="false"><em>${esc(chipMeaning(run.wb, s.sheet, slots[i]))}</em>`}
        </div>`).join('')}</div>
        ${level === 0 && active >= 0 ? `<div class="chips"><span class="muted small">「${esc(tpl.slots[active].label)}」の候補：</span>${shuffle(tpl.slots[active].chips, active).map((c) => `<button class="chip" data-c="${esc(c)}">${esc(c)}<small>${esc(chipMeaning(run.wb, s.sheet, c))}</small></button>`).join('')}</div>` : ''}`;
      $$('.slot', builder).forEach((el) => (el.onclick = () => { const i = +el.dataset.i; if (locked[i] || level !== 0) return; active = i; drawSlots(); }));
      $$('.chip', builder).forEach((b) => (b.onclick = () => { sfx('click'); slots[active] = b.dataset.c; const nx = slots.findIndex((x, j) => !x && !locked[j]); active = nx; drawSlots(); }));
      $$('.slot-in', builder).forEach((inp) => {
        inp.oninput = () => { slots[+inp.dataset.i] = inp.value; const em = inp.nextElementSibling; em.textContent = chipMeaning(run.wb, s.sheet, inp.value); $('.fcode', builder).innerHTML = '=' + tpl.fn + '(' + slots.map((x) => esc(x || '?')).join(',') + ')'; };
        inp.onfocus = () => (input = inp);
        inp.onkeydown = (e) => { if (e.key === 'Enter') exec(); };
      });
    }
    function shuffle(arr, seed) { const a = arr.slice(); for (let i = a.length - 1; i > 0; i--) { const j = (i * 7 + seed * 3 + 1) % (i + 1); [a[i], a[j]] = [a[j], a[i]]; } return a; }
    function drawText() {
      builder.innerHTML = `<div class="lv-tag">${s.type === 'repair' ? 'REPAIR' : 'SUPPORT ' + LEVELS[level].k + '｜' + LEVELS[level].name}</div>
        ${level === 2 && tpl ? `<div class="syntax">構文：<code>${esc(s.hints[3])}</code></div>` : ''}
        <div class="fbar big edit"><span class="fx">fx</span><input id="fin" class="fin" value="${esc(startText)}" spellcheck="false" autocomplete="off" aria-label="式を入力"></div>
        <div id="live" class="live"></div>`;
      input = $('#fin', builder);
      const live = () => { const ex = F.explain(input.value, run.wb, s.sheet); $('#live', builder).innerHTML = ex && ex.fn ? ex.blocks.map((b, i) => '<span class="blk a' + (i % BLOCK_COLORS) + '"><small>' + esc(b.label) + '</small><b>' + esc(b.meaning) + '</b></span>').join('') : '<span class="muted small">入力すると、式の意味ブロックがここに表示されます。</span>'; };
      input.oninput = live; live();
      input.onkeydown = (e) => { if (e.key === 'Enter') exec(); };
      input.focus({ preventScroll: true }); input.setSelectionRange(input.value.length, input.value.length);
    }
    if (useSlots) drawSlots(); else drawText();
    bindHints();
    // 列見出しクリック → 範囲を挿入
    $$('.sheet th.colh', w).forEach((th) => (th.onclick = () => {
      const ref = th.dataset.col + ':' + th.dataset.col;
      if (useSlots && level === 0 && active >= 0) { slots[active] = ref; active = slots.findIndex((x, j) => !x && !locked[j]); drawSlots(); return; }
      const inp = input || $('.slot-in', w); if (!inp) return;
      const p = inp.selectionStart ?? inp.value.length;
      inp.value = inp.value.slice(0, p) + ref + inp.value.slice(inp.selectionEnd ?? p);
      inp.dispatchEvent(new Event('input')); inp.focus();
    }));
    function current() { return useSlots ? '=' + tpl.fn + '(' + slots.map((x) => x || '').join(',') + ')' : input.value; }
    function exec() {
      const f = current();
      const fb = $('#fb', w);
      if (useSlots && slots.some((x) => !String(x).trim())) { fb.className = 'feedback info'; fb.textContent = 'まだ空いているブロックがあります。'; return; }
      const res = F.evaluate(f, run.wb, s.sheet);
      const fnsOk = res.fns && res.fns.some((x) => s.fns.includes(x));
      if (res.syntax) { fb.className = 'feedback bad'; fb.innerHTML = '🧩 式の形を確認しよう：' + esc(res.message) + '<br><small>例：カンマ , やカッコ ( ) 、文字は "B" のようにダブルクォートで囲む。</small>'; fail(s.skill); return; }
      if (F.sameValue(res.value, s.expect) && res.refs.length === 0) { fb.className = 'feedback info'; fb.textContent = '結果は合っています。でも数字を直接入れると、データが変わったとき追従できません。セル範囲を使って式にしよう。'; return; }
      if (F.sameValue(res.value, s.expect) && fnsOk) return win(f, res);
      if (F.sameValue(res.value, s.expect)) { fb.className = 'feedback info'; fb.innerHTML = '結果は ' + esc(fmt(res.value)) + ' で正解！ ただ、この仕事には <code>' + esc(s.fns[0]) + '</code> のほうが読みやすく、条件の追加にも強いです。'; return win(f, res); }
      const mk = (s.mistakes || []).find((m) => F.sameValue(m.v, res.value));
      fb.className = 'feedback bad';
      fb.innerHTML = '結果：<b>' + esc(fmt(res.value)) + '</b>' + (F.isErr(res.value) && res.value.why ? '（' + esc(res.value.why) + '）' : '') + '<br>' + esc(mk ? mk.msg : F.isErr(res.value) ? 'エラーは手がかり。どの部分が原因か見てみよう。' : '期待している結果と違うみたい。X-RAYで、式がどの行を使ったか見てみよう。') +
        (!F.isErr(res.value) && res.trace.length ? '<br><button class="btn sm ghost" id="myXray">🔬 自分の式をX-RAY</button>' : '');
      const mx = $('#myXray', w);
      if (mx) mx.onclick = () => { const a = $('#after', w); a.innerHTML = '<div></div>'; mountXray(a.firstChild, run.wb, s.sheet, f, {}); };
      fail(s.skill);
    }
    function win(f, res) {
      const assisted = level <= 1 || run.hint >= 3 || run.shownAnswer;
      succeed(s.skill, assisted);
      useFns(res.fns);
      if (s.job) raise(s.job, assisted ? 'PRACTICED' : 'INDEPENDENT');
      if (s.type === 'repair') rec('REPAIR', true, assisted);
      if (s.report !== undefined) run.reportFix[s.report] = f;
      if (!S.kpi.firstFormulaMs) { S.kpi.firstFormulaMs = Date.now() - sessionStart; save(); }
      // FADING SUPPORT：支援なしに近い成功なら次は支援を減らす
      if (s.forceLevel === undefined && !s.recovery && !run.shownAnswer && run.hint <= 1 && run.stepFails === 0 && S.support < 3) { S.support++; save(); }
      sfx('correct');
      $('#runBtn', w).disabled = true;
      const fb = $('#fb', w); fb.className = 'feedback good';
      const sc = s.success || { title: 'FORMULA WORKED!' };
      fb.innerHTML = `<div class="worked"><span class="big-num">${esc(fmt(res.value))}</span><span class="check">✓</span><span class="w-title">${esc(sc.title)}</span></div>` +
        (sc.meaning ? '<div class="xr-meaning">' + esc(sc.meaning) + '</div>' : '') +
        (run.shownAnswer ? '<p class="small muted">完成式を見て成功。次は自力でできたら「習得」です。</p>' : '') +
        (!S.kpi.shown && S.kpi.firstFormulaMs ? '<p class="small muted">⏱ 最初の役に立つ式まで ' + Math.round(S.kpi.firstFormulaMs / 1000) + ' 秒</p>' : '');
      if (!S.kpi.shown) { S.kpi.shown = true; save(); }
      const after = $('#after', w);
      const nextStep = run.m.steps[run.i + 1];
      if (res.trace.length && s.compact) {
        after.insertAdjacentHTML('beforeend', '<button class="btn sm ghost" id="okXray">🔬 直した式をX-RAY</button><div id="okXrayA"></div>');
        $('#okXray', w).onclick = () => { $('#okXray', w).remove(); mountXray($('#okXrayA', w), run.wb, s.sheet, f, { meaning: sc.meaning }); };
      } else if (res.trace.length && (!nextStep || nextStep.type !== 'xray')) { const x = document.createElement('div'); after.appendChild(x); mountXray(x, run.wb, s.sheet, f, { meaning: sc.meaning, startAt: 0 }); raise(s.job, 'DISCOVERED'); }
      after.insertAdjacentHTML('beforeend', '<button class="btn primary next" id="nx">次へ ▶</button>');
      $('#nx', w).onclick = next;
      if (S.support > 0 && level < S.support && s.template && !s.recovery && s.forceLevel === undefined) guideSay('いい調子！ 次は支援を少し減らすよ（SUPPORT ' + LEVELS[S.support].k + '）。');
    }
    $('#runBtn', w).onclick = exec;
  };
  RENDER.repair = (w, s) => RENDER.build(w, Object.assign({}, s, { type: 'repair', template: null }));

  // ---------- X-RAY ----------
  RENDER.xray = (w, s) => {
    w.innerHTML = '<div id="xr"></div><div id="after"></div>';
    mountXray($('#xr', w), run.wb, s.sheet, s.formula, { meaning: s.meaning, onDone: () => {
      succeed('READ'); raise(run.m.job, 'DISCOVERED');
      if (!$('#nx', w)) { $('#after', w).innerHTML = '<p class="aha">💡 式 → データ → 結果 がつながった！ 行をクリックすると、対象/対象外の理由も見られます。</p><button class="btn primary next" id="nx">次へ ▶</button>'; $('#nx', w).onclick = next; }
    } });
  };

  // ---------- DETECT ----------
  RENDER.detect = (w, s) => {
    const res = F.evaluate(s.formula, run.wb, s.sheet);
    w.innerHTML = reportHTML(s) + requestHTML(s) +
      `<div class="excel-result"><div class="er-cell"><small>提示された式</small>${formulaBarHTML(s.formula, run.wb, s.sheet, { plain: true })}</div>
        <div class="er-val"><small>Excelの結果</small><b>${esc(fmt(res.value))}</b><span class="noerr">エラーなし ✓</span></div></div>
      <div class="row gap" id="choice"><button class="btn" id="submit">📤 このまま提出</button><button class="btn primary" id="doubt">🔎 疑ってみる</button></div>
      <div id="stage"></div>
      <div class="sheet-area">${tabsHTML([s.sheet], s.sheet)}${sheetHTML(run.wb, s.sheet, { showSpaces: !!s.showSpaces })}</div>`;
    $('#submit', w).onclick = () => {
      fail('DETECT');
      $('#choice', w).remove();
      $('#stage', w).innerHTML = `<div class="wait"><div class="wait-t">🚨 WAIT</div><div class="wait-cmp"><div><small>あなたが提出した値</small><b class="bad">${esc(fmt(res.value))}</b></div><div><small>正解</small><b class="good">${esc(fmt(s.correct))}</b></div></div>
        <p>Excelは<b>エラーを出していません</b>。でも、正しくありません。原因を探そう。</p></div>`;
      causes();
    };
    $('#doubt', w).onclick = () => { sfx('click'); $('#choice', w).remove(); causes(true); };
    function causes(doubted) {
      const st = $('#stage', w);
      st.insertAdjacentHTML('beforeend', `<h3 class="prompt">${doubted ? 'ナイス判断。' : ''}何がおかしい？</h3>
        <div class="opts">${s.causes.map((c, k) => `<button class="opt" data-k="${k}">${esc(c.t)}</button>`).join('')}</div>
        <button class="btn sm ghost" id="dx">🔬 X-RAYで調べる</button><div id="dxa"></div><div class="feedback" id="fb"></div><div id="after"></div>`);
      $('#dx', w).onclick = () => { $('#dx', w).remove(); mountXray($('#dxa', w), run.wb, s.sheet, s.formula, {}); };
      $$('.opt', st).forEach((b) => (b.onclick = () => {
        const c = s.causes[+b.dataset.k];
        const fb = $('#fb', w);
        if (!c.ok) { b.classList.add('wrong'); b.disabled = true; fb.className = 'feedback bad'; fb.textContent = c.fb || 'もう少し調べてみよう。'; fail('DETECT'); return; }
        b.classList.add('right'); $$('.opt', st).forEach((x) => (x.disabled = true));
        succeed('DETECT', !doubted); sfx('discovery'); logWeek('verify');
        if (s.voice) speak(s.voice);
        fb.className = 'feedback good';
        fb.innerHTML = `<div class="noerr-banner">NO ERROR ≠ CORRECT</div><p>🔍 ${esc(s.explain)}</p><p class="small">正しい結果：<b>${esc(fmt(s.correct))}</b></p>`;
        if (s.repairTo) {
          const a = $('#after', w);
          a.innerHTML = '<h3 class="prompt">では、直そう。</h3><div id="rp"></div>';
          const fake = Object.assign({}, s, s.repairTo, { type: 'repair', template: null, request: null, report: s.report, compact: true });
          const holder = $('#rp', w);
          RENDER.build(holder, fake);
          holder.querySelector('.sheet-area')?.remove();
          bindHints();
        } else {
          $('#after', w).innerHTML = '<button class="btn primary next" id="nx">次へ ▶</button>';
          $('#nx', w).onclick = next;
        }
      }));
    }
  };

  // ---------- VERIFY ----------
  RENDER.verify = (w, s) => {
    const res = F.evaluate(s.formula, run.wb, s.sheet);
    const t0 = Date.now();
    const sh = run.wb[s.sheet];
    const ex = F.explain(s.formula, run.wb, s.sheet);
    w.innerHTML = reportHTML(s) + requestHTML(s) +
      `<div class="verify"><div class="v-head"><span class="v-title">🎯 この結果、本物？</span><span class="v-res">${esc(fmt(res.value))}</span><span class="v-time" id="vt">0.0秒</span></div>
      ${formulaBarHTML(s.formula, run.wb, s.sheet, { blocks: false })}
      <ol class="v-checks">
        <li id="v1" class="on"><b>① 対象は正しい？</b><p>この式が合計しているのは、どの列？</p><div class="opts small">${sh.headers.map((h, i) => `<button class="opt" data-h="${esc(h)}">${esc(F.numToCol(i + 1))}：${esc(h)}</button>`).join('')}</div></li>
        <li id="v2"><b>② 条件は全部ある？</b><p>依頼に必要な条件をすべて選ぶと、式の条件と照合します。</p><div class="opts small">${[...s.conds, ...s.decoys].sort().map((c) => `<button class="opt" data-c="${esc(c)}">${esc(c)}</button>`).join('')}</div><button class="btn sm" id="v2ok">照合する</button></li>
        <li id="v3"><b>③ 数件確認すると一致する？</b><p>表の中で、依頼に合う行をクリックして選ぼう（フィルターの代わり）。</p><div class="v-sum">選んだ行の合計：<b id="vsum">0</b> ／ 式の結果：<b>${esc(fmt(res.value))}</b></div><button class="btn sm" id="v3ok">一致を確認</button></li>
      </ol><div class="feedback" id="fb"></div><div id="after"></div>
      <p class="muted small">他の検証方法：概算 / 手計算 / フィルター / 元データ / 別の式で計算</p></div>
      <div class="sheet-area">${tabsHTML([s.sheet], s.sheet)}${sheetHTML(run.wb, s.sheet)}</div>`;
    const tick = setInterval(() => { const el = $('#vt', w); if (!el) return clearInterval(tick); el.textContent = ((Date.now() - t0) / 1000).toFixed(1) + '秒'; }, 100);
    const fb = $('#fb', w);
    let phase = 1;
    const setPhase = (p) => { phase = p; ['v1', 'v2', 'v3'].forEach((id, i) => { $('#' + id, w).className = i + 1 < p ? 'done' : i + 1 === p ? 'on' : ''; }); fb.className = 'feedback'; fb.textContent = ''; };
    $$('#v1 .opt', w).forEach((b) => (b.onclick = () => {
      if (phase !== 1) return;
      if (b.dataset.h === s.target.header) { b.classList.add('right'); sfx('correct'); setPhase(2); }
      else { b.classList.add('wrong'); fb.className = 'feedback bad'; fb.textContent = '式の最初の範囲（合計するもの）を見てみよう。列の見出しは？'; rec('VERIFY', false); }
    }));
    $$('#v2 .opt', w).forEach((b) => (b.onclick = () => { if (phase === 2) b.classList.toggle('sel'); }));
    $('#v2ok', w).onclick = () => {
      if (phase !== 2) return;
      const sel = $$('#v2 .opt.sel', w).map((b) => b.dataset.c);
      const okSet = sel.length === s.conds.length && s.conds.every((c) => sel.includes(c));
      if (!okSet) { fb.className = 'feedback bad'; fb.textContent = '依頼文に出てくる条件だけを選ぼう。'; rec('VERIFY', false); return; }
      const inFormula = s.conds.every((c) => (ex.pairs || []).includes(c));
      $('#v2', w).insertAdjacentHTML('beforeend', '<div class="match">' + s.conds.map((c) => '<span class="pair">' + esc(c) + ' ' + ((ex.pairs || []).includes(c) ? '✓ 式にあり' : '✖ 式にない') + '</span>').join('') + '</div>');
      if (inFormula) { sfx('correct'); setPhase(3); } else { fb.className = 'feedback bad'; fb.textContent = '式に足りない条件があります！'; }
    };
    const picked = new Set();
    $$('.sheet tbody tr[data-row]', w).forEach((tr) => (tr.onclick = () => {
      if (phase !== 3) return;
      const r = +tr.dataset.row; if (r < 2) return;
      if (picked.has(r)) picked.delete(r); else picked.add(r);
      tr.classList.toggle('picked');
      const col = s.target.col;
      $('#vsum', w).textContent = fmt([...picked].reduce((a, x) => a + (Number(F.cellValue(sh, col, x)) || 0), 0));
    }));
    $('#v3ok', w).onclick = () => {
      if (phase !== 3) return;
      const ok = picked.size === s.rowsOk.length && s.rowsOk.every((r) => picked.has(r));
      if (!ok) {
        const extra = [...picked].filter((r) => !s.rowsOk.includes(r));
        fb.className = 'feedback bad';
        fb.textContent = extra.length ? '行' + extra.join('・') + ' は条件に合わないかも。ラインと月をもう一度見てみよう。' : 'まだ足りない行があります。条件に合う行を全部選ぼう。';
        rec('VERIFY', false); return;
      }
      clearInterval(tick);
      const sec = ((Date.now() - t0) / 1000).toFixed(1);
      setPhase(4);
      succeed('VERIFY'); raise(run.m.job, 'VERIFIED'); logWeek('verify');
      sfx('verified');
      if (s.voice) speak(s.voice);
      if (s.finalClear) run.wb['月次レポート'].rows.forEach((row, i) => { if (!run.reportFix[i]) run.reportFix[i] = i === 0 ? s.formula : row[2]; });
      $('#after', w).innerHTML = `<div class="verified"><span class="shield">🛡</span><div><b>VERIFIED</b><small>${sec}秒で検証完了。対象・条件・元データ、すべて一致。</small></div></div><button class="btn primary next" id="nx">次へ ▶</button>`;
      if (s.finalClear) { const c = $('#clock'); if (c) c.textContent = '16:00'; }
      $('#nx', w).onclick = next;
    };
  };

  // ---------- DEBUG ----------
  RENDER.debug = (w, s) => {
    const sheet = s.sheet;
    const row = +s.cell.slice(1);
    const vCol = 'A';
    const variant = s.variant;
    const computed = { B: (r) => '=XLOOKUP(A' + r + ',マスター!A:A,マスター!B:B)' };
    const draw = () => sheetHTML(run.wb, sheet, { computed, showWide: true, rowClass: (r) => (r === row ? 'focus' : '') });
    const res = F.evaluate(s.formula, run.wb, sheet);
    const val = F.cellValue(run.wb[sheet], vCol, row);
    w.innerHTML = `<div class="debug"><div class="dbg-head"><span class="err-chip">${esc(fmt(res.value))}</span><b>${esc(s.cell)} の原因を調べましょう</b><span class="muted small">❌ FAILED ではなく 🔎 DEBUG MISSION</span></div>
      ${formulaBarHTML(s.formula, run.wb, sheet)}
      <div class="tools"><span class="small muted">調査ツール：</span>
        <button class="btn sm ghost" data-t="value">🔤 検索値を1文字ずつ見る</button>
        <button class="btn sm ghost" data-t="master">📒 マスターを検索</button>
        <button class="btn sm ghost" data-t="range">📐 範囲を確認</button>
        <button class="btn sm ghost" data-t="type">🔢 データ型を確認</button></div>
      <div class="clues" id="clues"></div>
      <h3 class="prompt">原因はどれ？</h3>
      <div class="opts" id="causes">
        <button class="opt" data-c="missing">検索値がマスターに存在しない</button>
        <button class="opt" data-c="fullwidth">表記が違う（全角/半角・空白）</button>
        <button class="opt" data-c="range">参照範囲が違う</button>
        <button class="opt" data-c="type">データ型が違う（文字/数値）</button></div>
      <div class="feedback" id="fb"></div><div id="after"></div></div>
      <div class="sheet-area" id="dsheet">${tabsHTML([sheet, 'マスター'], sheet)}${draw()}</div>`;
    bindTabs(w, '#dsheet', [sheet, 'マスター'], (n) => (n === sheet ? draw() : sheetHTML(run.wb, n, { showWide: true })));
    const clues = $('#clues', w);
    const tools = {
      value: () => '検索値 ' + esc(vCol + row) + ' = ' + [...String(val)].map((ch) => { const wide = /[！-～]/.test(ch); return '<span class="ch ' + (wide ? 'wide' : '') + '">' + esc(ch) + '<small>' + (wide ? '全角' : '半角') + '</small></span>'; }).join(''),
      master: () => { const has = run.wb['マスター'].rows.some((r) => r[0] === val); const half = String(val).replace(/[！-～]/g, (m) => String.fromCharCode(m.charCodeAt(0) - 0xfee0)); const hasHalf = run.wb['マスター'].rows.some((r) => r[0] === half); return 'マスター品番に「' + esc(val) + '」：<b>' + (has ? 'あり' : 'なし') + '</b>' + (half !== val ? '　／　半角の「' + esc(half) + '」：<b>' + (hasHalf ? 'あり' : 'なし') + '</b>' : ''); },
      range: () => '探す場所：マスター!A:A（品番） ／ 持ってくる：マスター!B:B（品名） → <b>範囲は正しい</b>',
      type: () => '検索値の型：<b>文字</b> ／ マスター品番の型：<b>文字</b> → 型は一致',
    };
    $$('[data-t]', w).forEach((b) => (b.onclick = () => { sfx('click'); b.disabled = true; clues.insertAdjacentHTML('beforeend', '<div class="clue">' + tools[b.dataset.t]() + '</div>'); }));
    $$('#causes .opt', w).forEach((b) => (b.onclick = () => {
      const fb = $('#fb', w);
      if (b.dataset.c !== variant) { b.classList.add('wrong'); b.disabled = true; fb.className = 'feedback bad'; fb.textContent = '調査ツールで手がかりを集めてみよう。'; fail('DEBUG'); return; }
      b.classList.add('right'); $$('#causes .opt', w).forEach((x) => (x.disabled = true));
      sfx('discovery'); succeed('DEBUG');
      fb.className = 'feedback good';
      fb.innerHTML = variant === 'fullwidth' ? '🔍 原因特定：「Ａ」が<b>全角</b>。Excelは「Ａ-001」と「A-001」を別の文字として扱います。' : '🔍 原因特定：「B-003」が<b>マスターに未登録</b>（不完全なマスター）。';
      const fixes = variant === 'fullwidth'
        ? [{ t: 'A5 を半角「A-001」に直す（元データを整える）', ok: true }, { t: 'IFERRORで囲んで空白にする', fb: 'エラーを消しただけ。品名が出ないまま出荷されます。エラーを消しただけではCLEARにしません。' }, { t: 'マスターに「Ａ-001」を追加する', fb: '同じ品番が2つになり、マスターが汚れます。' }]
        : [{ t: '「マスター未登録」と表示して、担当者にB-003の登録を依頼する', ok: true }, { t: 'IFERRORで0にする', fb: '問題が隠れて、誰も気づけなくなります。' }, { t: 'A6 を A-003 に書き換える', fb: 'データの改ざんになります。B-003 は実在する注文です。' }];
      $('#after', w).innerHTML = '<h3 class="prompt">どう直す？</h3><div class="opts" id="fixes">' + fixes.map((f, k) => `<button class="opt" data-k="${k}">${esc(f.t)}</button>`).join('') + '</div><div class="feedback" id="fb2"></div><div id="after2"></div>';
      $$('#fixes .opt', w).forEach((fbn) => (fbn.onclick = () => {
        const f = fixes[+fbn.dataset.k]; const fb2 = $('#fb2', w);
        if (!f.ok) { fbn.classList.add('wrong'); fbn.disabled = true; fb2.className = 'feedback bad'; fb2.textContent = f.fb; fail('REPAIR'); return; }
        fbn.classList.add('right'); $$('#fixes .opt', w).forEach((x) => (x.disabled = true));
        succeed('REPAIR'); sfx('correct'); raise('error', 'PRACTICED'); logWeek('error');
        if (variant === 'fullwidth') run.wb[sheet].rows[row - 2][0] = 'A-001';
        else computed.B = (r) => (r === row ? '=XLOOKUP(A' + r + ',マスター!A:A,マスター!B:B,"マスター未登録")' : '=XLOOKUP(A' + r + ',マスター!A:A,マスター!B:B)');
        $('#dsheet', w).innerHTML = tabsHTML([sheet, 'マスター'], sheet) + draw();
        bindTabs(w, '#dsheet', [sheet, 'マスター'], (n) => (n === sheet ? draw() : sheetHTML(run.wb, n, { showWide: true })));
        fb2.className = 'feedback good';
        fb2.innerHTML = '🛠 修正完了：' + esc(s.cell) + ' → <b>' + esc(fmt(F.evaluate(computed.B(row), run.wb, sheet).value)) + '</b>。原因を特定して直せました。';
        $('#after2', w).innerHTML = '<button class="btn primary next" id="nx">次へ ▶</button>'; $('#nx', w).onclick = next;
      }));
    }));
  };
  function bindTabs(w, sel, names, drawFn) {
    $$(sel + ' .tab', w).forEach((t) => (t.onclick = () => { const area = $(sel, w); area.innerHTML = tabsHTML(names, t.dataset.tab) + drawFn(t.dataset.tab); bindTabs(w, sel, names, drawFn); }));
  }

  // ---------- AI REVIEW ----------
  RENDER.ai = (w, s) => {
    let ri = 0;
    const round = () => {
      const r = s.rounds[ri];
      const sheet = r.sheet || s.sheet;
      const res = F.evaluate(r.formula, run.wb, sheet);
      let checked = false;
      if (r.clock) { const c = $('#clock'); if (c) c.textContent = r.clock; }
      sfx('ai');
      w.innerHTML = reportHTML(s) + (s.rounds.length > 1 ? `<div class="muted small">AI提案 ${ri + 1} / ${s.rounds.length}</div>` : '') + requestHTML(r) +
        `<div class="ai-box"><img src="${D.chars.ai.img}" alt="AIアシスタント" class="avatar"><div class="ai-say"><p>${esc(r.say)}</p><code class="ai-code">${esc(r.formula)}</code><span class="ai-res">→ ${esc(fmt(res.value))}</span></div></div>
        <div class="ai-actions"><button class="btn act-accept" data-a="ACCEPT">ACCEPT</button><button class="btn act-check" data-a="CHECK">CHECK</button><button class="btn act-repair" data-a="REPAIR">REPAIR</button><button class="btn act-reject" data-a="REJECT">REJECT</button></div>
        <div id="checks"></div><div class="feedback" id="fb"></div><div id="after"></div>
        <div class="sheet-area">${tabsHTML([sheet], sheet)}${sheetHTML(run.wb, sheet)}</div>`;
      guideSay(ri === 0 ? 'AIは使ってOK。でも確認せずに信用しない。まずCHECKしてみよう。' : 'この提案はどうかな？ CHECKの観点：何を計算？条件は全部？範囲は？データ型は？', 'navi');
      $$('.ai-actions .btn', w).forEach((b) => (b.onclick = () => {
        const a = b.dataset.a; const fb = $('#fb', w);
        if (a === 'CHECK') {
          if (checked) return; checked = true; b.disabled = true; sfx('click');
          $('#checks', w).innerHTML = '<div class="checklist"><b>🔎 CHECK</b>' + r.checks.map((c) => '<div class="ck">☑ ' + esc(c) + '</div>').join('') + '<button class="btn sm ghost" id="aix">🔬 X-RAYで見る</button><div id="aixa"></div></div>';
          $('#aix', w).onclick = () => { $('#aix', w).remove(); mountXray($('#aixa', w), run.wb, sheet, r.formula, {}); };
          logWeek('ai');
          return;
        }
        if (!r.good.includes(a)) {
          b.classList.add('wrong'); fb.className = 'feedback bad';
          fb.textContent = a === 'ACCEPT' ? (checked ? 'もう一度チェック項目を見てみよう。正常に見えるけど…' : '確認せずに信用しない。まずCHECKしてみよう。') : 'この式は正しいかも。CHECKで確かめてみよう。';
          fail('AI REVIEW'); return;
        }
        $$('.ai-actions .btn', w).forEach((x) => (x.disabled = true)); b.classList.add('right');
        if (a === 'REPAIR' && r.fix) {
          fb.className = 'feedback good'; fb.innerHTML = '🔍 ' + esc(r.explain) + '<br>AIの式を直してみよう。';
          const exp = F.evaluate(r.fix, run.wb, sheet).value;
          const holder = $('#after', w);
          const fixFns = F.evaluate(r.fix, run.wb, sheet).fns;
          run.hint = 0; run.shownAnswer = false;
          run.bridgeHints = ['何が足りない？ CHECKの結果を見よう。', 'AIの式をベースに、足りない所だけ直す。', fixFns[0] + ' を使う。', r.fix.replace(/"[^"]*"/g, '"…"'), r.fix];
          const hp = $('#hints'); if (hp) { hp.hidden = false; $('#hintList').innerHTML = ''; $('#hintBtn').disabled = false; $('#hintBtn').textContent = '💡 HINT 1'; }
          RENDER.build(holder, { type: 'repair', skill: 'REPAIR', formula: r.formula, sheet, expect: exp, fns: [fixFns[0]], hints: run.bridgeHints, template: null, compact: true });
          holder.querySelector('.sheet-area')?.remove();
          // 修正成功後の「次へ」をAIラウンド進行に差し替え
          const obs = new MutationObserver(() => { const nx = $('#nx', holder); if (nx) { obs.disconnect(); succeed('AI REVIEW', !checked); nx.onclick = adv; } });
          obs.observe(holder, { childList: true, subtree: true });
          return;
        }
        succeed('AI REVIEW', !checked); sfx('correct');
        fb.className = 'feedback good';
        fb.innerHTML = (a === 'ACCEPT' && !checked ? '✓ 正しい式でした。次はCHECKしてからACCEPTすると完璧。<br>' : '✓ ') + esc(r.explain) + (r.fix && a === 'REJECT' ? '<br>正しくは <code>' + esc(r.fix) + '</code>' : '') + (r.best && a !== r.best ? '<br><small>REPAIR（読みやすく直す）がベストでした。</small>' : '');
        $('#after', w).innerHTML = '<button class="btn primary next" id="nx">' + (ri + 1 < s.rounds.length ? '次の提案 ▶' : '次へ ▶') + '</button>';
        $('#nx', w).onclick = adv;
      }));
    };
    const adv = () => { ri++; if (ri < s.rounds.length) round(); else next(); };
    round();
  };

  // ---------- BRIDGE ----------
  RENDER.bridge = (w, s) => {
    const sheet = s.sheet;
    const tabs = [sheet, '10月実績(旧)', 'マスター'];
    let phase = 1, userF = null, userV = null;
    const picked = new Set();
    run.bridgeHints = ['依頼は前と同じ。「生産数」を「ライン＝B」かつ「月＝10月」で合計。列名が違うだけ。', '複数条件の合計。', 'SUMIFS。', '=SUMIFS(生産数の列, ﾗｲﾝ名の列, "Bライン", 月の列, "10月")', '=SUMIFS(F:F,C:C,"Bライン",E:E,"10月")'];
    const draw = (opts = {}) => sheetHTML(run.wb, sheet, Object.assign({ showSpaces: phase >= 3, rowClass: (r) => (picked.has(r) ? 'picked' : '') }, opts));
    w.innerHTML = requestHTML(s) +
      `<div class="bridge-note"><b>🌉 BRIDGE</b>：実務風の列名・途中の空白行・古い式の #REF!・複数シート・不要な列。<br>でも <b>JobとLogicは同じ</b>。関数名の指定はありません。</div>
      <div id="bphase"></div><div class="feedback" id="fb"></div><div id="after"></div>
      <div class="sheet-area" id="bsheet">${tabsHTML(tabs, sheet)}${draw()}</div>`;
    const bind = () => bindTabs(w, '#bsheet', tabs, (n) => (n === sheet ? draw() : sheetHTML(run.wb, n)));
    bind();
    const fb = $('#fb', w);
    // phase 1: BUILD (JOB ONLY)
    $('#bphase', w).innerHTML = `<div class="lv-tag">SUPPORT D｜JOB ONLY</div><div class="fbar big edit"><span class="fx">fx</span><input id="fin" class="fin" value="=" spellcheck="false" autocomplete="off" aria-label="式を入力"></div><div id="live" class="live"></div><button class="btn primary" id="runBtn">▶ 実行</button>`;
    const input = $('#fin', w);
    input.oninput = () => { const ex = F.explain(input.value, run.wb, sheet); $('#live', w).innerHTML = ex && ex.fn ? ex.blocks.map((b, i) => '<span class="blk a' + (i % BLOCK_COLORS) + '"><small>' + esc(b.label) + '</small><b>' + esc(b.meaning) + '</b></span>').join('') : ''; };
    input.onkeydown = (e) => { if (e.key === 'Enter') build(); };
    $('#runBtn', w).onclick = build;
    bindHints();
    w.addEventListener('click', (e) => {
      const tr = e.target.closest('#bsheet tbody tr[data-row]');
      if (tr && phase === 2 && $('#bsheet .tab.on', w).dataset.tab === sheet) {
        const r = +tr.dataset.row; if (r < 2) return;
        picked.has(r) ? picked.delete(r) : picked.add(r); tr.classList.toggle('picked');
        $('#vsum', w).textContent = fmt([...picked].reduce((a, x) => a + (Number(F.cellValue(run.wb[sheet], 'F', x)) || 0), 0));
        return;
      }
      const th = e.target.closest('th.colh'); if (!th || phase !== 1) return; const ref = th.dataset.col + ':' + th.dataset.col; const p = input.selectionStart ?? input.value.length; input.value = input.value.slice(0, p) + ref + input.value.slice(p); input.dispatchEvent(new Event('input')); input.focus(); });
    function build() {
      const res = F.evaluate(input.value, run.wb, sheet);
      if (res.syntax) { fb.className = 'feedback bad'; fb.textContent = '🧩 ' + res.message; fail('TRANSFER'); return; }
      const v = res.value;
      const okVals = [320, 738, 1248];
      if (!F.isErr(v) && okVals.includes(v) && res.fns.some((f) => /SUMIF/.test(f))) {
        userF = input.value; userV = v; useFns(res.fns); succeed('BUILD', run.hint >= 3);
        $('#runBtn', w).disabled = true; input.disabled = true; sfx('correct');
        fb.className = 'feedback good';
        fb.innerHTML = '式は動きました：<b>' + esc(fmt(v)) + '</b>。…でも実務データです。<b>10秒VERIFY</b>で本物か確かめよう。';
        verifyPhase();
        return;
      }
      fb.className = 'feedback bad';
      fb.innerHTML = '結果：<b>' + esc(fmt(v)) + '</b>。' + (v === 500 || v === 918 ? '9月分も入っていない？ 月の条件は？' : F.isErr(v) ? esc(v.why || 'エラーは手がかり。') : '列名を見直そう：生産数は「生産数(個)」、ラインは「ﾗｲﾝ名」、月は「月」の列。値は「Bライン」のように入っています。');
      fail('TRANSFER');
    }
    function verifyPhase() {
      phase = 2;
      $('#after', w).innerHTML = `<div class="verify"><div class="v-head"><span class="v-title">🎯 この結果、本物？ ③ 数件確認</span></div><p>表の中で「Bライン・10月」の行をクリックして選ぼう。</p>
        <div class="v-sum">選んだ行の合計：<b id="vsum">0</b> ／ 式の結果：<b>${esc(fmt(userV))}</b></div><button class="btn sm" id="vok">一致を確認</button></div><div id="after2"></div>`;
      $('#vok', w).onclick = () => {
        const want = [3, 7, 10];
        const ok = picked.size === 3 && want.every((r) => picked.has(r));
        if (!ok) { fb.className = 'feedback bad'; fb.textContent = 'ﾗｲﾝ名がBラインで、月が10月の行は3行あるはず。見た目がそっくりな行にも注意。'; rec('VERIFY', false); return; }
        if (userV === s.correct) { done(true); return; }
        sfx('warning'); phase = 3;
        fb.className = 'feedback bad';
        fb.innerHTML = '<div class="wait-t">🚨 MISMATCH</div>手で数えると <b>1,248</b>、式は <b>' + esc(fmt(userV)) + '</b>。Excelはエラーを出していません。';
        $('#bsheet', w).innerHTML = tabsHTML(tabs, sheet) + draw(); bind();
        detectPhase();
      };
    }
    function detectPhase() {
      const a2 = $('#after2', w);
      a2.innerHTML = `<div id="bx"></div><h3 class="prompt">原因は？（当てはまるものを全部）</h3>
        <div class="opts" id="bc">${['ﾗｲﾝ名に見えない空白がある', '生産数が「文字」として入っている', '途中に空白行がある', '前回比の列が #REF!', '列名が半角ｶﾅ'].map((t, k) => `<button class="opt" data-k="${k}">${esc(t)}</button>`).join('')}</div><button class="btn sm" id="bcok">決定</button><div id="after3"></div>`;
      mountXray($('#bx', w), run.wb, sheet, userF, { auto: false, startAt: 0 });
      $$('#bc .opt', w).forEach((b) => (b.onclick = () => b.classList.toggle('sel')));
      $('#bcok', w).onclick = () => {
        const sel = $$('#bc .opt', w).map((b) => b.classList.contains('sel'));
        const need = userV === 320 ? [0, 1] : userV === 738 ? [1] : [0];
        const ok = sel.every((x, k) => x === need.includes(k));
        if (!ok) { fb.className = 'feedback bad'; fb.textContent = 'X-RAYで、対象外になった行（クリックで理由が出る）を調べよう。空白行・#REF!・列名は、この式の結果には影響していません。'; fail('DETECT'); return; }
        succeed('DETECT'); sfx('discovery');
        fb.className = 'feedback good'; fb.innerHTML = '🔍 発見！ 行7の「Bライン␣」（末尾に空白）と、行10の「510」（文字の数字）が、静かに集計から漏れていました。';
        $('#after3', w).innerHTML = `<h3 class="prompt">データをどう整える？（全部選ぶ）</h3><div class="opts" id="bf">
          <button class="opt" data-k="trim">ﾗｲﾝ名の前後の空白を取る（TRIM）</button><button class="opt" data-k="num">文字の数字を数値に変換（VALUE / 区切り位置）</button><button class="opt" data-k="iferror">式をIFERRORで囲む</button><button class="opt" data-k="del">Bライン以外の行を削除する</button></div>
          <button class="btn primary" id="bfok">データを整えて再計算</button>`;
        $$('#bf .opt', w).forEach((b) => (b.onclick = () => b.classList.toggle('sel')));
        $('#bfok', w).onclick = () => {
          const sel2 = new Set($$('#bf .opt.sel', w).map((b) => b.dataset.k));
          const needs = new Set(need.map((k) => (k === 0 ? 'trim' : 'num')));
          if (sel2.has('iferror') || sel2.has('del') || ![...needs].every((k) => sel2.has(k))) { fb.className = 'feedback bad'; fb.textContent = sel2.has('iferror') ? 'IFERRORではエラーでない問題は直りません。' : sel2.has('del') ? '元データを消すと、他の集計が壊れます。' : '見つけた原因それぞれに対応しよう。'; fail('REPAIR'); return; }
          const sh = run.wb[sheet];
          sh.rows.forEach((r) => { if (typeof r[2] === 'string') r[2] = r[2].trim(); if (typeof r[5] === 'string' && r[5].trim() !== '' && !isNaN(Number(r[5]))) r[5] = Number(r[5]); });
          userV = F.evaluate(userF, run.wb, sheet).value;
          $('#bsheet', w).innerHTML = tabsHTML(tabs, sheet) + draw(); bind();
          succeed('REPAIR');
          if (userV === s.correct) done(false);
          else { fb.className = 'feedback bad'; fb.textContent = '再計算：' + fmt(userV) + '。まだ一致しません。'; }
        };
      };
    }
    function done(clean) {
      phase = 9;
      succeed('TRANSFER'); succeed('VERIFY'); raise('sum', 'TRANSFERRED'); logWeek('verify');
      sfx('verified');
      $('#after', w).innerHTML = `<div class="verified"><span class="shield">🌉</span><div><b>BRIDGE CLEAR — TRANSFERRED</b><small>${clean ? '最初から正しい値を出せた！ ' : ''}再計算 1,248 ＝ 手計算 1,248。汚れた実務データでも、JobとLogicは同じでした。</small></div></div>
        <ul class="bridge-learn"><li>列名が違っても「何を・どの条件で」は同じ</li><li>見えない空白・文字の数字は、エラーを出さずに数字を変える</li><li>だから最後は必ず VERIFY</li></ul>
        <button class="btn primary next" id="nx">FORMULA CARD へ ▶</button>`;
      $('#nx', w).onclick = next;
    }
  };

  // ---------- CHAPTER END ----------
  RENDER.chapterEnd = (w, s) => {
    S.chapters[run.id] = true; save();
    run.gained.push(...s.gained);
    const isFinal = run.m.final;
    sfx(isFinal ? 'clear' : 'verified');
    const nextM = run.m.next && D.missions[run.m.next];
    w.innerHTML = `<div class="chapter-end ${isFinal ? 'final' : ''}">
      ${isFinal ? `<div class="final-clock">16:00</div><img src="${D.chars.boss.img}" class="avatar lg" alt=""><p class="boss-say">「…助かった。数字、合ってるな。」</p>` : ''}
      <div class="ce-title">${esc(s.title)}</div>
      <h3>持ち帰れる力</h3><ul class="gained">${s.gained.map((g) => '<li>✓ ' + esc(g) + '</li>').join('')}</ul>
      <div class="row gap center">${nextM ? `<button class="btn primary" id="goNext">次の章へ：${esc(nextM.title)} ▶</button>` : ''}<button class="btn ghost" id="toCard">ここで終える（FORMULA CARD）</button></div></div>`;
    if (isFinal) speak('clear');
    if ($('#goNext', w)) $('#goNext', w).onclick = () => startMission(run.m.next);
    $('#toCard', w).onclick = () => { const id = run.id; renderCard(w, { fromChapter: id }); };
  };

  // ---------- FORMULA CARD ----------
  function renderCard(w, o = {}) {
    sfx('clear');
    if (run && !run.m.mini && !run.m.next) { S.chapters[run.id] = true; save(); }
    const jobsDone = D.jobs.filter((j) => MASTERY.indexOf(S.mastery[j.id] || 'NEW') >= 2);
    const fns = S.fnsUsed.filter((f) => !['MAX', 'MIN'].includes(f));
    w.innerHTML = `<div class="fcard">
      <div class="fc-head"><div><span class="fc-title">YOUR FORMULA CARD</span><span class="fc-date">${esc(todayKey())}</span></div><img src="${D.chars.navi.img}" alt="" class="avatar"></div>
      <div class="fc-grid">
        <div><h4>今日できるようになった仕事</h4><ul class="fc-jobs">${(jobsDone.length ? jobsDone : D.jobs.filter((j) => session.jobs.has(j.id))).map((j) => `<li><span class="ok">✓</span>${esc(j.skill)}<small>${esc(MASTERY_JA[S.mastery[j.id] || 'NEW'])}</small></li>`).join('')}
          ${(S.skills.DETECT && S.skills.DETECT.ok + S.skills.DETECT.as) ? '<li><span class="ok">✓</span>間違い発見<small>サイレントエラー</small></li>' : ''}
          ${(S.skills['AI REVIEW'] && S.skills['AI REVIEW'].ok + S.skills['AI REVIEW'].as) ? '<li><span class="ok">✓</span>AI式確認<small>VERIFY BEFORE TRUST</small></li>' : ''}</ul></div>
        <div><h4>能力</h4><div class="map">${COMPETENCY.map((k) => `<div class="mrow"><span>${esc(k)}</span><span class="stars">${starHTML(stars(k))}</span></div>`).join('')}</div></div>
      </div>
      <p class="fc-used">Used: ${esc(fns.join(' / ') || '—')}</p>
      ${S.kpi.firstFormulaMs ? `<p class="fc-kpi">最初の役に立つ式まで ${Math.round(S.kpi.firstFormulaMs / 1000)}秒</p>` : ''}
    </div>
    <div class="row gap center">${o.standalone ? '<button class="btn primary" id="cHome">トップへ</button>' : '<button class="btn primary" id="cWork">明日の仕事を決める（MY WORK BRIDGE）▶</button><button class="btn ghost" id="cHome">トップへ</button>'}</div>`;
    if ($('#cWork', w)) $('#cWork', w).onclick = () => renderMyWork(w);
    $('#cHome', w).onclick = home;
  }
  RENDER.card = (w) => renderCard(w);
  function renderMyWork(w) {
    w.innerHTML = `<div class="mywork"><h2>明日の仕事で何に使えそう？</h2><p class="muted">1つだけ選ぼう。小さく1回、が定着のコツ。</p>
      <div class="opts">${D.myWork.map((m) => `<button class="opt" data-id="${m.id}">□ ${esc(m.t)}</button>`).join('')}</div><div id="tm"></div></div>`;
    $$('.opt', w).forEach((b) => (b.onclick = () => {
      $$('.opt', w).forEach((x) => { x.classList.remove('sel'); x.textContent = x.textContent.replace('■', '□'); });
      b.classList.add('sel'); b.textContent = b.textContent.replace('□', '■');
      const m = D.myWork.find((x) => x.id === b.dataset.id);
      S.tomorrow = { job: m.id, mission: m.mission, date: todayKey() }; save();
      $('#tm', w).innerHTML = `<div class="tomorrow-card"><span class="badge yellow">TOMORROW MISSION</span><p>${esc(m.mission)}</p><small>困ったら ⚡10秒RESCUE。使えたら、トップ画面で「使えた！」を押してね。</small></div>
        <p class="clear-msg">学習の終了をCLEARとしない。<br><b>実務で使えた瞬間を、本当のCLEARとする。</b></p><button class="btn primary" id="tmOk">トップへ</button>`;
      $('#tmOk', w).onclick = home;
    }));
  }
  RENDER.mywork = (w) => renderMyWork(w);

  // ===================== RESCUE =====================
  function openRescue(focus) {
    sfx('click'); logWeek('rescue');
    const ov = $('#overlay');
    ov.hidden = false;
    ov.innerHTML = `<div class="rescue" role="dialog" aria-label="10秒RESCUE"><div class="rs-head"><b>⚡ 10 SEC RESCUE</b><span>今、Excelで何をしたい？</span><button class="x" data-close aria-label="閉じる">✕</button></div>
      <div class="rs-jobs">${D.rescue.map((r) => `<button class="rs-job ${r.id === focus ? 'on' : ''}" data-id="${r.id}"><span>${esc(r.job)}</span><b>${esc(r.fn)}</b></button>`).join('')}</div><div id="rsBody"></div></div>`;
    $$('[data-close]', ov).forEach((b) => (b.onclick = closeOverlay));
    ov.onclick = (e) => { if (e.target === ov) closeOverlay(); };
    $$('.rs-job', ov).forEach((b) => (b.onclick = () => { $$('.rs-job', ov).forEach((x) => x.classList.remove('on')); b.classList.add('on'); show(b.dataset.id); }));
    function show(id) {
      const r = D.rescue.find((x) => x.id === id);
      const body = $('#rsBody', ov);
      body.innerHTML = `<div class="rs-card"><div class="rs-fn">${esc(r.fn)}</div><code class="rs-syn">${esc(r.syntax)}</code>
        <div class="rs-ex"><code>${esc(r.example)}</code><button class="btn sm primary" id="copy">COPY</button></div>
        <button class="btn sm ghost" id="und">🧠 30 SEC UNDERSTAND</button><div id="undBody"></div></div>`;
      $('#copy', body).onclick = () => copy(r.example);
      $('#und', body).onclick = () => {
        $('#und', body).remove();
        $('#undBody', body).innerHTML = `<div class="und"><p><b>意味：</b>${esc(r.meaning)}</p><div id="rx"></div><p><b>⚠ Trap：</b>${esc(r.trap)}</p><p><b>🛡 Verify：</b>${esc(r.verify)}</p></div>`;
        const wb = clone(D.workbook);
        if (/XLOOKUP|SUMIFS|COUNTIFS/.test(r.example)) mountXray($('#rx', body), wb, r.sheet, r.example, {});
        else { const v = F.evaluate(r.example, wb, r.sheet).value; $('#rx', body).innerHTML = formulaBarHTML(r.example, wb, r.sheet, { result: v }); }
      };
    }
    show(focus || 'sumifs');
    if (!focus) $('.rs-job', ov).classList.add('on');
  }
  function copy(text) {
    const ok = () => toast('コピーしました：' + text, 'good');
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(text).then(ok, fb); else fb();
    function fb() { const t = document.createElement('textarea'); t.value = text; document.body.appendChild(t); t.select(); try { document.execCommand('copy'); ok(); } catch (e) { toast('コピーできませんでした。手で選択してください。'); } t.remove(); }
  }
  function closeOverlay() { const ov = $('#overlay'); ov.hidden = true; ov.innerHTML = ''; if (modal.cb) { const cb = modal.cb; modal.cb = null; cb(); } }
  function modal(html, cb) {
    const ov = $('#overlay');
    ov.hidden = false;
    ov.innerHTML = '<div class="modal" role="dialog">' + html + '</div>';
    modal.cb = cb || null;
    $$('[data-close]', ov).forEach((b) => (b.onclick = closeOverlay));
  }

  // ===================== EXIT (途中離脱でも1技能を持ち帰る) =====================
  function exitRun() {
    if (!run) return home();
    const gained = new Set(run.gained);
    Object.entries(session.skills).forEach(([k, v]) => { if (v.ok + v.as > 0) gained.add({ WHAT: '求めるものを見極める', LOGIC: '条件を分解する', CHOOSE: '仕事に合う道具を選ぶ', BUILD: '式を組み立てる', READ: '式を読む', PREDICT: '結果を予想する', DETECT: '間違いを発見する', REPAIR: '式を直す', VERIFY: '結果を検証する', DEBUG: 'エラー原因を調べる', 'AI REVIEW': 'AIの式を確認する', TRANSFER: '実務データに移す' }[k] || k); });
    modal(`<div class="exit"><h2>ここで終わっても大丈夫。</h2>
      ${gained.size ? '<p>今日持ち帰れるもの：</p><ul class="gained">' + [...gained].map((g) => '<li>✓ ' + esc(g) + '</li>').join('') + '</ul>' : '<p>今日の1つ：<b>困ったら ⚡10秒RESCUE で式をコピーできる</b>。</p>'}
      <p class="muted small">続きは QUEST MAP からいつでも再開できます。</p>
      <div class="row gap center"><button class="btn ghost" data-close>続ける</button><button class="btn primary" id="exitYes">終える</button></div></div>`);
    $('#exitYes').onclick = () => { modal.cb = null; closeOverlay(); home(); };
  }

  // ===================== SETTINGS =====================
  function settings() {
    modal(`<div class="settings"><h2>設定</h2>
      <label class="switch"><input type="checkbox" id="stSound" ${S.sound ? 'checked' : ''}> 効果音（職場ではOFF推奨）</label>
      <label class="switch"><input type="checkbox" id="stNarr" ${S.narration ? 'checked' : ''}> ナレーション音声（字幕は常に表示）</label>
      <label class="switch"><input type="checkbox" id="stTimer" ${S.timer ? 'checked' : ''}> FINAL MISSION の時間制限（演出用）</label>
      <div class="support-set"><b>SUPPORT（式づくりの支援）</b><div class="row gap">${LEVELS.map((l, i) => `<button class="btn sm ${S.support === i ? 'primary' : 'ghost'}" data-lv="${i}">${l.k} ${l.name}</button>`).join('')}</div><small class="muted">上達すると自動で減り、連続でつまずくと自動で戻ります。</small></div>
      <div class="row gap"><button class="btn ghost" id="stReset">学習データをリセット</button><button class="btn primary" data-close>閉じる</button></div></div>`);
    $('#stSound').onchange = (e) => { S.sound = e.target.checked; save(); syncSoundBtn(); };
    $('#stNarr').onchange = (e) => { S.narration = e.target.checked; if (S.narration && !S.sound) { S.sound = true; $('#stSound').checked = true; } save(); syncSoundBtn(); };
    $('#stTimer').onchange = (e) => { S.timer = e.target.checked; save(); };
    $$('[data-lv]').forEach((b) => (b.onclick = () => { S.support = +b.dataset.lv; save(); settings(); }));
    $('#stReset').onclick = () => { if (confirm('学習データ（進捗・FORMULA CARD・1週間ログ）を消去します。よろしいですか？')) { S = fresh(); save(); closeOverlay(); home(); } };
  }
  function syncSoundBtn() { const b = $('#soundBtn'); b.textContent = S.sound ? '🔊' : '🔇'; b.setAttribute('aria-pressed', S.sound); b.title = S.sound ? '効果音 ON' : '効果音 OFF'; }

  // ===================== BOOT =====================
  function boot() {
    $('#homeBtn').onclick = () => (run ? exitRun() : home());
    $('#exitBtn').onclick = exitRun;
    $('#rescueBtn').onclick = () => openRescue();
    $('#soundBtn').onclick = () => { S.sound = !S.sound; save(); syncSoundBtn(); if (S.sound) sfx('click'); };
    $('#settingsBtn').onclick = settings;
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('#overlay').hidden) closeOverlay(); });
    syncSoundBtn();
    const hash = location.hash.replace('#', '');
    if (hash === 'rescue') { home(); openRescue(); } else if (D.missions[hash]) startMission(hash); else home();
  }
  window.EFQApp = { startMission, home, openRescue, state: () => S, where: () => (run ? { id: run.id, i: run.i } : null) };
  boot();
})();
