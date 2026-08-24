/* Feature checks against a plain page: gate confirmation, the checkout
   checklist, the per-person pages and the Excel export. */
const { BASE, ok, done, launch } = require('./helpers');

(async () => {
  const browser = await launch();
  const page = await browser.newPage({ viewport: { width: 430, height: 932 } });
  let native = 0;
  page.on('dialog', d => { native++; d.dismiss(); });
  const errs = [];
  page.on('pageerror', e => errs.push('PAGEERROR: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errs.push('CONSOLE: ' + m.text().slice(0, 110)); });
  const dlg = async i => { await page.click(`[data-dlg="${i}"]`); await page.waitForTimeout(250); };
  const modalText = () => page.locator('#mBody').innerText();

  await page.goto(BASE);
  await page.waitForTimeout(400);
  await page.click('.person:has-text("فاطیما")');
  await page.click('#btnCheckin');
  await page.waitForTimeout(300);

  // the gate asks before it counts, and a refusal keeps the chores locked
  await page.click('#btnGate');
  await page.waitForTimeout(300);
  ok('gate asks for confirmation', /مطمئنی/.test(await modalText()),
     (await modalText()).replace(/\s+/g, ' ').slice(0, 42));
  await dlg(1);
  ok('cancelling leaves chores locked', await page.locator('#posList .chk').first().isDisabled());
  await page.click('#btnGate');
  await page.waitForTimeout(300);
  await dlg(0);
  ok('confirming unlocks the chores', !(await page.locator('#posList .chk').first().isDisabled()));

  // the checkout list starts fully negative and clears one item at a time
  const count = async () => (await page.locator('#coCount').innerText()).replace(/\s+/g, ' ');
  ok('all checkout items start as penalties', /7 امتیاز منفی/.test(await count()), await count());
  ok('every row renders as pending', (await page.locator('#negList li.pending').count()) === 7);
  for (const i of [0, 1, 2, 3]) { await page.locator('#negList .chk').nth(i).click(); await page.waitForTimeout(120); }
  ok('clearing items lowers the penalty count', /3 امتیاز منفی/.test(await count()), await count());
  ok('progress bar reflects 4 of 7',
     (await page.locator('#coFill').getAttribute('style')).includes('57%'));

  await page.locator('#posList .chk').nth(0).click();
  await page.waitForTimeout(200);
  await page.click('#btnCheckout');
  await page.waitForTimeout(300);
  await dlg(0);
  const state = await page.evaluate(() => ({
    neg: L().filter(l => l.type === 'neg').length,
    co: L().filter(l => l.type === 'checkout').length,
    ok: L().filter(l => l.type === 'co_ok').length
  }));
  ok('leftover items become negative points', state.neg === 3 && state.co === 1 && state.ok === 4,
     JSON.stringify(state));

  // scores are per person, never a leaderboard
  await page.click('.tab[data-tab="scores"]');
  await page.waitForTimeout(400);
  const panel = (await page.locator('#personPanel').innerText()).replace(/\s+/g, ' ');
  ok('person page shows their own stats', panel.includes('فاطیما') && panel.includes('امتیاز مثبت'));
  ok('no cross-person leaderboard', !(await page.locator('#scoreBody').count()));
  ok('breakdown bars render', (await page.locator('.barfill').count()) > 0);
  ok('trend chart renders', (await page.locator('#personPanel svg').count()) === 1);
  await page.click('.person[data-sp="محمد"]');
  await page.waitForTimeout(300);
  ok('switching person switches the page', (await page.locator('#personPanel').innerText()).includes('محمد'));

  // excel export
  await page.click('.tab[data-tab="settings"]');
  const dl = page.waitForEvent('download', { timeout: 8000 });
  await page.click('#btnExcel');
  const file = await dl;
  ok('excel file downloads', /\.xlsx$/.test(file.suggestedFilename()), file.suggestedFilename());

  ok('no native browser dialog was used', native === 0, 'native=' + native);
  if (errs.length) console.log('ERRORS:\n' + errs.join('\n'));
  ok('no page errors', errs.length === 0);
  await browser.close();
  done();
})();
