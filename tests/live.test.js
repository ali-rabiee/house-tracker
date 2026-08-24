/* One phone deletes; a second phone that nobody touches — app open, sitting on
   a tab mid-session — must see it, and see a restore, on its own. */
const { BASE, ok, done, launch, resetBackend } = require('./helpers');

const waitFor = async (fn, ms = 40000) => {
  const until = Date.now() + ms;
  while (Date.now() < until) {
    if (await fn()) return (Date.now() - (until - ms)) / 1000;
    await new Promise(r => setTimeout(r, 500));
  }
  return null;
};

(async () => {
  await resetBackend();
  const browser = await launch();
  const errs = [];
  const mk = async () => {
    const ctx = await browser.newContext({ viewport: { width: 430, height: 900 } });
    const pg = await ctx.newPage();
    pg.on('pageerror', e => errs.push('PAGEERROR: ' + e.message));
    pg.on('dialog', d => d.accept());
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

  const A = await mk(), B = await mk();

  // A records a visit with one chore
  await A.click('.person:has-text("فاطیما")');
  await A.click('#btnCheckin');
  await A.waitForTimeout(250);
  await A.click('#btnGate');
  await A.waitForTimeout(300);
  await dlg(A, 0);
  await A.locator('#posList .chk').nth(1).click();     // P2
  await A.waitForTimeout(250);
  await connect(A);
  await connect(B);

  // B is parked on the history tab and will not be touched again
  await B.click('.tab[data-tab="history"]');
  await B.waitForTimeout(600);
  const chore = 'چک اتاق‌ها';
  ok('B sees the chore to begin with',
     (await B.locator('#historyList').innerText()).includes(chore));

  // A deletes it (choosing "delete without backup")
  await A.click('.tab[data-tab="history"]');
  await A.waitForTimeout(300);
  await A.locator('#historyList li', { hasText: chore }).locator('[data-del]').click();
  await A.waitForTimeout(300);
  await dlg(A, 1);
  await A.waitForTimeout(1800);

  const goneIn = await waitFor(async () =>
    !(await B.locator('#historyList').innerText()).includes(chore));
  ok('the deletion reaches the idle phone by itself', goneIn !== null,
     goneIn === null ? 'still there after 40s' : 'after ~' + goneIn.toFixed(0) + 's');

  // the row really left the logs sheet and sits in the bin
  const dump = await (await fetch(BASE + '/dump')).json();
  const live = (dump['logs'] || []).slice(1).map(r => r[7]);
  const binned = (dump['حذف‌شده‌ها'] || []).slice(1);
  ok('the row is out of the logs sheet', !live.some(t => String(t).includes(chore)));
  ok('the row is in the bin with its details', binned.some(r => String(r[10]).includes(chore)),
     binned.map(r => r[10]).join(' | '));

  // restoring it from the Sheet must reach the same untouched phone
  const head = (dump['حذف‌شده‌ها'] || [])[0] || [];
  await (await fetch(BASE + '/restore?col=' + (head.indexOf('بازیابی؟') + 1))).text();
  const backIn = await waitFor(async () =>
    (await B.locator('#historyList').innerText()).includes(chore));
  ok('a restore from the Sheet reaches it too', backIn !== null,
     backIn === null ? 'never came back' : 'after ~' + backIn.toFixed(0) + 's');

  if (errs.length) console.log('ERRORS:\n' + errs.join('\n'));
  ok('no page errors', errs.length === 0);
  await browser.close();
  done();
})();
