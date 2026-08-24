/* Two isolated browser contexts stand in for two phones sharing one Sheet. */
const { BASE, ok, done, launch, resetBackend } = require('./helpers');

const nums = t => (t.match(/\d+/g) || []).join(',');   // pos,neg,checkin,checkout

(async () => {
  const browser = await launch();
  const errs = [];
  const mk = async name => {
    const ctx = await browser.newContext({ viewport: { width: 430, height: 900 } });   // its own storage
    const pg = await ctx.newPage();
    pg.on('pageerror', e => errs.push(name + ' PAGEERROR: ' + e.message));
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
  const syncNow = async pg => {
    await pg.click('.tab[data-tab="settings"]');
    await pg.click('#btnSyncNow');
    await pg.waitForTimeout(900);
  };
  const stats = async (pg, who) => {
    await pg.click('.tab[data-tab="scores"]');
    await pg.waitForTimeout(250);
    await pg.click(`.person[data-sp="${who}"]`);
    await pg.waitForTimeout(350);
    return (await pg.locator('#personPanel .stat').innerText()).replace(/\s+/g, ' ').trim();
  };

  await resetBackend();
  const A = await mk('A'), B = await mk('B');

  await A.click('.person:has-text("فاطیما")');
  await A.click('#btnCheckin');
  await A.waitForTimeout(250);
  await A.click('#btnGate');
  await A.waitForTimeout(300);
  await dlg(A, 0);
  await A.locator('#posList .chk').nth(0).click();
  await A.waitForTimeout(250);
  await connect(A);
  ok('phone A connects', (await A.locator('#syncChip').innerText()).includes('همگام'));

  await connect(B);
  ok('phone B pulls A history', nums(await stats(B, 'فاطیما')).startsWith('2,0,1'), await stats(B, 'فاطیما'));

  await B.click('.tab[data-tab="checkin"]');
  await B.click('.person:has-text("محمد")');
  await B.click('#btnCheckin');
  await B.waitForTimeout(300);
  for (const i of [0, 1]) { await B.locator('#negList .chk').nth(i).click(); await B.waitForTimeout(150); }
  await B.click('#btnCheckout');
  await B.waitForTimeout(300);
  await dlg(B, 0);
  await B.waitForTimeout(1800);
  await syncNow(A);
  ok('A sees B checkout with its penalties', nums(await stats(A, 'محمد')) === '0,5,1,1', await stats(A, 'محمد'));

  const dump = await (await fetch(BASE + '/dump')).json();
  ok('sheet built the sessions report',
     (dump['نوبت‌ها'] || []).some(r => r[0] === 'محمد' && r[6] === 5));
  ok('sheet built the per-person summary',
     (dump['خلاصهٔ افراد'] || []).some(r => r[0] === 'فاطیما' && r[1] === 2));

  await B.click('.tab[data-tab="history"]');
  await B.waitForTimeout(300);
  const before = await B.locator('#historyList li').count();
  await B.locator('#historyList li', { hasText: 'آب دادن به گل' }).locator('[data-del]').click();
  await B.waitForTimeout(300);
  await dlg(B, 0);
  await B.waitForTimeout(1800);
  await syncNow(A);
  await A.click('.tab[data-tab="history"]');
  await A.waitForTimeout(300);
  const afterA = await A.locator('#historyList li').count();
  const afterB = await B.locator('#historyList li').count();
  ok('a delete travels and stays deleted', afterB === before - 1 && afterA === afterB,
     `B ${before}->${afterB}, A ${afterA}`);

  await A.click('.tab[data-tab="settings"]');
  await A.fill('#newPerson', 'سارا');
  await A.click('#btnAddPerson');
  await A.waitForTimeout(1800);
  await syncNow(B);
  await B.click('.tab[data-tab="checkin"]');
  await B.waitForTimeout(300);
  ok('the shared people list travels', (await B.locator('#peopleGrid').innerText()).includes('سارا'));

  await B.context().setOffline(true);
  await B.click('.person:has-text("علی")');
  await B.click('#btnCheckin');
  await B.waitForTimeout(1600);
  const chip = (await B.locator('#syncChip').innerText()).trim();
  await B.context().setOffline(false);
  await syncNow(B);
  await syncNow(A);
  ok('an offline record reaches the other phone', nums(await stats(A, 'علی')) === '0,0,1,0',
     'offline chip: ' + chip);

  for (const who of ['فاطیما', 'محمد', 'علی']) {
    ok('both phones agree on ' + who, (await stats(A, who)) === (await stats(B, who)));
  }
  if (errs.length) console.log('ERRORS:\n' + errs.join('\n'));
  ok('no page errors', errs.length === 0);
  await browser.close();
  done();
})();
