// End-to-end playthrough: NODE_PATH=$(npm root -g) node tests/e2e.js [outDir]
// 全章＋全Jobミッションを自動で正答しながら通しプレイし、JSエラーがないことを確認する。
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const out = process.argv[2] || path.join(__dirname, '..', '.e2e-shots');
fs.mkdirSync(out, { recursive: true });
const url = 'file://' + path.resolve(__dirname, '..', 'index.html');

(async () => {
  const browser = await chromium.launch(process.env.PW_CHROMIUM ? { executablePath: process.env.PW_CHROMIUM } : {});
  const page = await browser.newPage({ viewport: { width: 1360, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));
  page.on('console', (m) => { if (m.type() === 'error' && !/Failed to load resource|ERR_FILE_NOT_FOUND|net::/.test(m.text())) errors.push('console: ' + m.text()); });
  await page.goto(url);
  await page.evaluate(() => localStorage.clear());
  await page.reload();
  let shot = 0;
  const snap = async (name) => page.screenshot({ path: path.join(out, String(++shot).padStart(3, '0') + '_' + name + '.png'), fullPage: true });
  await snap('home');

  const click = async (sel) => { await page.locator(sel).first().waitFor({ state: 'visible', timeout: 15000 }); await page.locator(sel).first().click(); };
  const step = async () => page.evaluate(() => { const w = window.EFQApp.where(); if (!w) return null; const st = window.EFQData.missions[w.id].steps[w.i]; return { w, st, level: window.EFQApp.state().support }; });

  async function solve(firstWrong) {
    const info = await step(); if (!info) return false;
    const { st } = info;
    const tag = info.w.id + '_' + info.w.i + '_' + st.type;
    switch (st.type) {
      case 'talk': await snap(tag); await click('#go'); break;
      case 'choice': {
        if (st.multi) { for (const [k, o] of st.options.entries()) if (o.ok) await click(`.opt[data-k="${k}"]`); await click('#multiOk'); }
        else {
          if (firstWrong) { const bad = st.options.findIndex((o) => !o.ok); await click(`.opt[data-k="${bad}"]`); }
          await click(`.opt[data-k="${st.options.findIndex((o) => o.ok)}"]`);
        }
        await page.waitForTimeout(200); await snap(tag); await click('#nx'); break;
      }
      case 'build': case 'repair': {
        if (await page.locator('.chip').count()) {
          for (const sl of st.template.slots) await click(`.chip[data-c='${sl.ans[0]}']`);
        } else if (await page.locator('.slot-in').count()) {
          const ins = page.locator('.slot-in');
          for (let i = 0; i < await ins.count(); i++) { const idx = +(await ins.nth(i).getAttribute('data-i')); await ins.nth(i).fill(st.template.slots[idx].ans[0]); }
        } else {
          if (firstWrong) { await page.fill('#fin', '=SUM(D:D'); await click('#runBtn'); await page.waitForTimeout(150); }
          await page.fill('#fin', st.hints[4]);
        }
        await click('#runBtn'); await page.waitForTimeout(400); await snap(tag); await click('#nx'); break;
      }
      case 'xray': await page.waitForTimeout(500); await snap(tag + '_a'); await page.locator('#nx').waitFor({ timeout: 15000 }); await snap(tag + '_b'); await click('#nx'); break;
      case 'detect': {
        await snap(tag + '_a');
        await click(firstWrong ? '#submit' : '#doubt');
        await click(`#stage .opt[data-k="${st.causes.findIndex((c) => c.ok)}"]`);
        if (st.repairTo) { await page.fill('#rp #fin', st.repairTo.hints[4]); await click('#rp #runBtn'); }
        await page.waitForTimeout(200); await snap(tag + '_b'); await click('#nx'); break;
      }
      case 'verify': {
        await click(`#v1 .opt[data-h="${st.target.header}"]`);
        for (const c of st.conds) await click(`#v2 .opt[data-c="${c}"]`);
        await click('#v2ok');
        for (const r of st.rowsOk) await click(`.sheet-area tr[data-row="${r}"]`);
        await click('#v3ok'); await page.waitForTimeout(200); await snap(tag); await click('#nx'); break;
      }
      case 'debug': {
        for (const t of ['value', 'master']) await click(`[data-t="${t}"]`);
        await click(`#causes .opt[data-c="${st.variant}"]`);
        if (firstWrong) await click('#fixes .opt[data-k="1"]');
        await click('#fixes .opt[data-k="0"]'); await page.waitForTimeout(150); await snap(tag); await click('#nx'); break;
      }
      case 'ai': {
        for (const [ri, r] of st.rounds.entries()) {
          await click('.act-check');
          if (r.good.includes('REPAIR') && (r.best === 'REPAIR' || !r.good.includes('REJECT') || ri % 2 === 1)) {
            await click('.act-repair'); await page.fill('#after #fin', r.fix); await click('#after #runBtn');
          } else await click(r.good.includes('ACCEPT') ? '.act-accept' : '.act-reject');
          await page.waitForTimeout(200); await snap(tag + '_r' + ri); await click('#nx');
        }
        break;
      }
      case 'bridge': {
        await snap(tag + '_a');
        await page.fill('#fin', '=SUMIFS(F:F,C:C,"Bライン",E:E,"10月")'); await click('#runBtn');
        for (const r of [3, 7, 10]) await click(`#bsheet tr[data-row="${r}"]`);
        await click('#vok'); await page.waitForTimeout(200); await snap(tag + '_b');
        await click('#bc .opt[data-k="0"]'); await click('#bc .opt[data-k="1"]'); await click('#bcok');
        await click('#bf .opt[data-k="trim"]'); await click('#bf .opt[data-k="num"]'); await click('#bfok');
        await page.waitForTimeout(200); await snap(tag + '_c'); await click('#nx'); break;
      }
      case 'chapterEnd':
        await snap(tag);
        if (info.w.id === 'c1') { // 章末から MY WORK BRIDGE を設定できる
          await click('#ceWork'); await click('.mywork-modal .opt'); await click('#tmOk');
          const tm = await page.evaluate(() => window.EFQApp.state().tomorrow);
          if (!tm) errors.push('tomorrow mission not set from chapter end');
          if (!(await page.locator('#goNext').count())) errors.push('chapter end lost after MY WORK modal');
        }
        if (await page.locator('#goNext').count()) await click('#goNext'); else await click('#toCard');
        break;
      case 'card':
        await snap(tag);
        if (info.w.id === 'c7') {
          const head = await page.locator('.fc-grid h4').first().innerText();
          const jobsTxt = await page.locator('.fc-jobs').innerText();
          if (!/今日/.test(head) || !/→/.test(jobsTxt)) errors.push('card is not today-based: ' + head + ' / ' + jobsTxt.slice(0, 60));
        }
        await click('#cWork'); await click('.mywork .opt'); await snap(tag + '_mywork'); await click('#tmOk'); return false;
      default: throw new Error('unknown step ' + st.type);
    }
    return true;
  }

  // main quest c1..c7 (with some deliberate mistakes to exercise fail paths)
  await click('.job[data-job="sum"]');
  // 初回だけ60秒CHALLENGEの提案が出る（既定はこのまま始める）
  if (!(await page.locator('.offer').count())) errors.push('first-visit challenge offer not shown');
  await snap('offer');
  await click('#offerGo');
  let n = 0;
  while (await solve(n % 5 === 2)) { n++; if (n > 200) throw new Error('loop'); }
  await page.waitForTimeout(200);
  if (!(await page.locator('.entry').count())) {
    // c7 ends via card/mywork -> home. If on chapterEnd card path, go home.
    await snap('end_state');
  }
  await snap('home_after');
  // YOUR WEEK は「ミッション開始」ではなくクリア等だけを記録する
  const weekLog = await page.evaluate(() => window.EFQApp.state().log);
  const badLog = weekLog.filter((x) => x.v !== 2 || !['clear', 'verify', 'real', 'rescue'].includes(x.k));
  const chapterClears = weekLog.filter((x) => x.k === 'clear').length;
  if (badLog.length) errors.push('week log has non-v2/unknown entries: ' + JSON.stringify(badLog.slice(0, 3)));
  if (chapterClears !== 7) errors.push('expected 7 chapter clears, got ' + chapterClears);
  for (const job of ['lookup', 'judge', 'count', 'clean', 'error']) {
    await page.evaluate(() => window.EFQApp.home());
    await click(`.job[data-job="${job}"]`);
    let k = 0; while (await solve(k === 1)) { k++; if (k > 50) throw new Error('loop ' + job); }
  }
  // challenge + rescue + adaptive
  await page.evaluate(() => window.EFQApp.home());
  await click('#goChallenge'); await click('#chStart');
  for (let i = 0; i < 7; i++) await click(i % 3 ? '.opt[data-k="1"]' : '.opt[data-k="-1"]');
  await snap('challenge_result');
  await click('#chHome');
  await click('#rescueBtn'); await click('.rs-job[data-id="xlookup"]'); await click('#und'); await page.waitForTimeout(4000); await snap('rescue');
  await click('.rescue [data-close]');
  await click('#tmDone').catch(() => {}); await page.waitForTimeout(200); await snap('realwin');
  if (await page.locator('.modal [data-close]').count()) await click('.modal [data-close]');
  // adaptive support: 3 failures
  await page.evaluate(() => window.EFQApp.startMission('c2'));
  await click('#go');
  for (let i = 1; i <= 2; i++) await click(`.opt[data-k="${i}"]`);
  // ✕で終えるときも明日の1回を決められる
  await page.evaluate(() => { const st = window.EFQApp.state(); st.tomorrow = null; localStorage.setItem('efq.v1', JSON.stringify(st)); });
  await page.reload();
  await page.evaluate(() => window.EFQApp.startMission('c2'));
  await click('#go'); await click('#exitBtn'); await click('#exitWork'); await click('.mywork .opt'); await snap('exit_mywork');
  if (!(await page.evaluate(() => window.EFQApp.state().tomorrow))) errors.push('tomorrow mission not set from exit');
  await click('#tmOk');
  // 「提出」は失敗に数えない（WAITは発見イベント）
  await page.evaluate(() => { localStorage.clear(); });
  await page.reload();
  await page.evaluate(() => window.EFQApp.startMission('c3'));
  await click('#go'); await click('#submit'); await click('#stage .opt[data-k="0"]');
  const detNg = await page.evaluate(() => (window.EFQApp.state().skills.DETECT || {}).ng || 0);
  if (detNg !== 0) errors.push('submit counted as DETECT failure: ng=' + detNg);
  // 3回の本当の誤答で、段階に合った声かけが出る
  await page.evaluate(() => window.EFQApp.startMission('c3'));
  await click('#go'); await click('#submit');
  for (const k of [1, 2, 3]) await click(`#stage .opt[data-k="${k}"]`);
  await page.waitForTimeout(200);
  const adaptive = await page.locator('.adaptive').count();
  const adaptiveText = adaptive ? await page.locator('.adaptive .say').innerText() : '';
  if (adaptive && !/X-RAY/.test(adaptiveText)) errors.push('adaptive message not DETECT-specific: ' + adaptiveText);
  await snap('adaptive');
  await page.setViewportSize({ width: 390, height: 844 });
  await page.evaluate(() => { document.querySelector('#overlay').hidden = true; window.EFQApp.home(); });
  await snap('mobile_home');
  await page.evaluate(() => window.EFQApp.startMission('c2'));
  await click('#go'); await click('.opt[data-k="0"]'); await snap('mobile_choice');
  // スマホ：BUILD の最初の操作が1画面目に入る
  await page.evaluate(() => window.EFQApp.startMission('c2'));
  await click('#go');
  for (const sel of ['.opt[data-k="0"]', '#nx']) await click(sel);
  for (const sel of ['.opt[data-k="0"]', '.opt[data-k="1"]', '#multiOk', '#nx', '.opt[data-k="2"]', '#nx']) await click(sel);
  await page.evaluate(() => scrollTo(0, 0));
  const mobFirstY = await page.evaluate(() => { const b = [...document.querySelectorAll('#work button:not([disabled]),#work input')].find((x) => x.offsetParent); return b ? b.getBoundingClientRect().top : 9999; });
  await snap('mobile_build');
  if (mobFirstY > 844) errors.push('mobile build: first control below the fold at ' + Math.round(mobFirstY));
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth + 1);
  await browser.close();
  console.log('screenshots:', shot, 'adaptive modal:', adaptive, 'mobile h-overflow:', overflow);
  if (errors.length) { console.error(errors.join('\n')); process.exit(1); }
  if (!adaptive) { console.error('adaptive modal not shown'); process.exit(1); }
  if (overflow) { console.error('horizontal overflow on mobile'); process.exit(1); }
  console.log('E2E OK');
})().catch((e) => { console.error(e); process.exit(1); });
