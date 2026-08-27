/* Sync should be quiet: once at open, on every record you make, and whenever
   someone taps refresh — not on a fast timer and not on every tab switch. */
const { BASE, ok, done, launch, resetBackend } = require('./helpers');

const stays = async (fn, ms) => {                 // true if fn stays false for ms
  const until = Date.now() + ms;
  while (Date.now() < until) {
    if (await fn()) return false;
    await new Promise(r => setTimeout(r, 500));
  }
  return true;
};
const becomes = async (fn, ms = 15000) => {
  const until = Date.now() + ms;
  while (Date.now() < until) {
    if (await fn()) return true;
    await new Promise(r => setTimeout(r, 400));
  }
  return false;
};

(async () => {
  await resetBackend();
  const browser = await launch();
  const errs = [];
  const calls = { A: 0, B: 0 };
  const mk = async tag => {
    const ctx = await browser.newContext({ viewport: { width: 430, height: 900 } });
    const pg = await ctx.newPage();
    pg.on('pageerror', e => errs.push(tag + ' PAGEERROR: ' + e.message));
    pg.on('dialog', d => d.accept());
    pg.on('request', r => { if (r.url().includes('/exec')) calls[tag]++; });
    await pg.goto(BASE);
    await pg.waitForTimeout(400);
    return pg;
  };
  const dlg = async (pg, i) => {
    if (await pg.locator('#modal.on').count()) { await pg.click(`[data-dlg="${i}"]`); await pg.waitForTimeout(250); }
  };
  const connect = async pg => {
    await pg.click('.tab[data-tab="settings"]');
    await pg.fill('#syncUrl', BASE + '/exec');
    await pg.fill('#syncToken', 'khaneh-1404');
    await pg.click('#btnSyncSave');
    await pg.waitForTimeout(500);
    await dlg(pg, 0);
    await pg.waitForTimeout(900);
  };

  const A = await mk('A'), B = await mk('B');
  await A.click('.person:has-text("فاطیما")');
  await A.click('#btnCheckin');
  await A.waitForTimeout(250);
  await A.click('#btnGate');
  await A.waitForTimeout(300);
  await dlg(A, 0);
  await A.locator('#posList .chk').nth(1).click();      // P2
  await A.waitForTimeout(250);
  await connect(A);
  await connect(B);

  const chore = 'چک اتاق‌ها';
  await B.click('.tab[data-tab="history"]');
  await B.waitForTimeout(600);
  ok('B starts out seeing the chore',
     (await B.locator('#historyList').innerText()).includes(chore));

  // moving around the app must not trigger a sync any more
  calls.B = 0;
  for (const t of ['tasks', 'scores', 'settings', 'history']) {
    await B.click(`.tab[data-tab="${t}"]`);
    await B.waitForTimeout(400);
  }
  ok('switching tabs no longer syncs', calls.B === 0, 'requests while browsing: ' + calls.B);

  // A deletes it; B is left completely alone
  await A.click('.tab[data-tab="history"]');
  await A.waitForTimeout(300);
  await A.locator('#historyList li', { hasText: chore }).locator('[data-del]').click();
  await A.waitForTimeout(300);
  await dlg(A, 1);
  await A.waitForTimeout(1500);

  await B.click('.tab[data-tab="history"]');
  await B.waitForTimeout(300);
  calls.B = 0;
  const quiet = await stays(async () =>
    !(await B.locator('#historyList').innerText()).includes(chore), 20000);
  ok('an idle phone no longer polls every few seconds', quiet && calls.B === 0,
     'requests in 20s: ' + calls.B);

  // the refresh button is how you pull it in
  await B.click('#syncChip');
  ok('tapping refresh brings the change in', await becomes(async () =>
     !(await B.locator('#historyList').innerText()).includes(chore)),
     'requests after tap: ' + calls.B);

  // and simply recording something also brings others' changes along
  await resetBackend();
  const C = await mk('A');
  await connect(C);
  await C.click('.tab[data-tab="checkin"]');
  await C.waitForTimeout(300);
  calls.A = 0;
  await C.click('.person:has-text("محمد")');
  await C.click('#btnCheckin');
  await C.waitForTimeout(1500);
  ok('recording something syncs on its own', calls.A > 0);

  /* one busy answer must not pin the indicator on "busy" for good */
  await C.waitForTimeout(1500);                      // let any pending sync settle first
  await fetch(BASE + '/busy?n=1');                   // one transient refusal
  await C.click('#syncChip');
  const seen = new Set();
  for (let i = 0; i < 40; i++) {                     // watch the indicator as it goes
    seen.add((await C.locator('#syncChip').innerText()).trim());
    await C.waitForTimeout(100);
  }
  ok('a passing refusal shows as queued while it retries',
     [...seen].some(t => /شلوغ|صف|همگام‌سازی/.test(t)), [...seen].join(' → '));
  ok('and clears itself once the retry succeeds',
     await becomes(async () => /🔄 همگام/.test(await C.locator('#syncChip').innerText()), 20000),
     'chip: ' + (await C.locator('#syncChip').innerText()).trim());

  /* a refusal that keeps coming must surface, not hide behind "busy" forever */
  await fetch(BASE + '/busy?n=5');
  await C.click('#syncChip');
  const surfaced = await becomes(async () =>
    /آفلاین|خطا/.test(await C.locator('#syncChip').innerText()), 20000);
  ok('a persistent refusal stops hiding and reports itself', surfaced,
     'chip: ' + (await C.locator('#syncChip').innerText()).trim());
  const why = (await C.locator('#syncStatus').innerText()).replace(/\s+/g, ' ');
  ok('and the settings page says what the server said', /شلوغ|قفل|خطا/.test(why), why.slice(0, 80));
  await fetch(BASE + '/busy?n=0');
  await C.click('#syncChip');
  await C.waitForTimeout(1500);

  /* the connection test must report the truth, good or bad */
  await C.click('.tab[data-tab="settings"]');
  await C.click('#btnSyncTest');
  await C.waitForTimeout(1500);
  const good = (await C.locator('#mBody').innerText()).replace(/\s+/g, ' ');
  ok('connection test reports a healthy sheet and its version',
     /سالم/.test(good) && /v\d/.test(good), good.slice(0, 110));
  await C.click('#tOk');

  await fetch(BASE + '/busy?n=1');
  await C.click('#btnSyncTest');
  await C.waitForTimeout(1500);
  const bad = (await C.locator('#mBody').innerText()).replace(/\s+/g, ' ');
  ok('and reports a refusal with the server\'s own words',
     /قفل آزاد نشد|خطا/.test(bad), bad.slice(0, 110));
  await C.click('#tOk');

  /* a request that never comes back must not wedge sync for good */
  await fetch(BASE + '/hang');
  await C.click('.tab[data-tab="tasks"]');
  await C.click('#btnGate');
  await C.waitForTimeout(300);
  await dlg(C, 0);                                  // this write's request hangs
  await C.waitForTimeout(2500);
  const stuckChip = (await C.locator('#syncChip').innerText()).trim();

  await C.click('#syncChip');                       // manual refresh must break it
  const recovered = await becomes(async () =>
    /همگام|صف|شلوغ/.test(await C.locator('#syncChip').innerText()), 20000);
  ok('a hung request does not wedge sync', recovered, 'chip while hung: ' + stuckChip);

  const posted = await becomes(async () => {
    const d = await (await fetch(BASE + '/dump')).json();
    return (d['logs'] || []).slice(1).some(r => String(r[6]) === 'HOME');
  }, 30000);
  ok('the record still reaches the sheet afterwards', posted);

  if (errs.length) console.log('ERRORS:\n' + errs.join('\n'));
  ok('no page errors', errs.length === 0);
  await browser.close();
  done();
})();
