/* The artifact viewer runs the page in a sandboxed iframe, where the browser
   ignores confirm()/alert()/prompt() entirely. Everything must still work, and
   a stay that spans days must not demand a fresh check-in. */
const { BASE, ok, done, launch, resetBackend } = require('./helpers');

(async () => {
  const browser = await launch();
  const page = await browser.newPage({ viewport: { width: 430, height: 932 } });
  let native = 0;
  page.on('dialog', d => { native++; d.accept(); });
  const errs = [];
  page.on('pageerror', e => errs.push('PAGEERROR: ' + e.message));

  await resetBackend();
  await page.goto(BASE + '/sandbox');
  await page.waitForTimeout(800);
  const f = page.frames().find(x => x.url().replace(/\/$/, '') === BASE);
  const modalOpen = () => f.evaluate(() => document.getElementById('modal').classList.contains('on'));
  const dlg = async i => { await f.click(`[data-dlg="${i}"]`); await page.waitForTimeout(300); };

  // checked in three days ago, then the site was closed and reopened
  const sessId = await f.evaluate(() => {
    S.ui.person = 'فاطیما';
    const l = addLog({ type: 'checkin', person: 'فاطیما' });
    l.ts = Date.now() - 3 * 86400000;
    save(); render();
    return l.id;
  });
  await page.waitForTimeout(300);
  ok('a 3-day-old check-in is still an open session',
     (await f.innerText('#checkinState')).includes('در خانه است'));

  await f.click('.tab[data-tab="tasks"]');
  await f.click('#btnGate');
  await page.waitForTimeout(300);
  ok('the gate question is drawn in-page', await modalOpen());
  await dlg(1);
  ok('cancelling leaves chores locked', await f.isDisabled('#posList .chk >> nth=0'));
  await f.click('#btnGate');
  await page.waitForTimeout(300);
  await dlg(0);
  ok('confirming unlocks chores with no new check-in',
     !(await f.isDisabled('#posList .chk >> nth=0')));

  await f.click('#posList .chk >> nth=0');
  await page.waitForTimeout(300);
  const chore = await f.evaluate(() =>
    L().filter(x => x.type === 'pos' && x.taskId !== 'HOME').slice(-1)[0].sess);
  ok('a chore logged today joins the earlier check-in', chore === sessId, 'sess=' + chore);

  await f.click('#btnCheckout');
  await page.waitForTimeout(300);
  ok('checkout confirmation is in-page', await modalOpen());
  await dlg(0);
  const co = await f.evaluate(() => ({
    sess: L().find(l => l.type === 'checkout').sess,
    open: !!openSession('فاطیما')
  }));
  ok('checkout binds to that check-in and closes it', co.sess === sessId && !co.open, JSON.stringify(co));

  // and it reaches the Sheet as one multi-day visit
  await f.click('.tab[data-tab="settings"]');
  await f.fill('#syncUrl', BASE + '/exec');
  await f.fill('#syncToken', 'khaneh-1404');
  await f.click('#btnSyncSave');
  await page.waitForTimeout(500);
  if (await modalOpen()) await dlg(0);
  await page.waitForTimeout(1200);
  ok('sync works from the sandboxed frame', (await f.innerText('#syncChip')).includes('همگام'));

  const dump = await (await fetch(BASE + '/dump')).json();
  const row = (dump['نوبت‌ها'] || []).find(r => r[0] === 'فاطیما') || [];
  ok('sheet shows one multi-day session with its chores', Number(row[4]) > 4000 && row[5] >= 1,
     JSON.stringify(row.slice(0, 7)));
  const logs = dump['logs'] || [];
  const col = (logs[0] || []).indexOf('نوبت (چک‌این)');
  const choreRow = logs.find(r => String(r[13]) === 'pos' && r[6] !== 'HOME') || [];
  ok('each log row names its check-in', col > 0 && /^\d{4}-\d{2}-\d{2}/.test(String(choreRow[col])),
     'نوبت=' + choreRow[col]);

  ok('no native browser dialog was needed', native === 0, 'native=' + native);
  if (errs.length) console.log('ERRORS:\n' + errs.join('\n'));
  ok('no page errors', errs.length === 0);
  await browser.close();
  done();
})();
