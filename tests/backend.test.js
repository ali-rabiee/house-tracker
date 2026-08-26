/* Runs google-apps-script/Code.gs directly, without a browser. */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { createGas } = require('./gas-emu');
const { ok, done } = require('./helpers');

const FILE = path.join(__dirname, '..', 'google-apps-script', 'Code.gs');
const BIN = 'حذف‌شده‌ها';
const src = fs.readFileSync(FILE, 'utf8');

/* Apps Script's Run button defaults to the first function in the file, so that
   function must be the one that actually sets the sheet up. */
ok('the Run dropdown defaults to setup', src.match(/^function (\w+)/m)[1] === 'setup',
   src.match(/^function (\w+)/m)[1]);

const g = createGas(FILE);
const made = g.run('setup');
ok('setup creates the sheets and reports what it made', /logs/.test(made) && /config/.test(made), made);
ok('the logs header is written', ((g.rows('logs') || [])[0] || []).length >= 18);
ok('rebuild runs on its own', g.run('rebuild') === 'OK');

let msg = '(no error)';
try { g.run('doGet'); } catch (e) { msg = e.message; }
ok('running a web-app entry point by hand explains itself', /setup/.test(msg), msg.slice(0, 55));

/* a standalone script (not created from inside a Sheet) must say why it can't work */
const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => null, getActive: () => null, flush(){}, getUi(){ throw new Error('x'); } },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: t => ({ setMimeType(){ return this; }, getContent(){ return t; } }) },
  LockService: { getScriptLock: () => ({ waitLock(){}, releaseLock(){} }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty(){} }) },
  Utilities: { formatDate: () => '' }, Session: { getScriptTimeZone: () => 'x' }, Logger: { log(){} }, console
};
vm.createContext(sandbox);
vm.runInContext(src, sandbox);
let unbound = '(no error)';
try { sandbox.setup(); } catch (e) { unbound = e.message; }
ok('an unbound script names the real problem', /وصل نیست/.test(unbound), unbound.slice(0, 45));

/* a chore and its check-in arriving together still resolve the session link */
const g2 = createGas(FILE);
const t = Date.now();
const res = g2.post({ token: 'khaneh-1404', since: 0, logs: [
  { id: 'c1', ts: t - 3 * 86400000, person: 'فاطیما', type: 'checkin', sess: 'c1' },
  { id: 'p1', ts: t, person: 'فاطیما', type: 'pos', taskId: 'P1', code: 'P1', title: 'گل‌ها', sess: 'c1' },
  { id: 'o1', ts: t + 1000, person: 'فاطیما', type: 'checkout', sess: 'c1' }
]});
ok('the backend accepts a batch', res.ok && res.logs.length === 3);
const logs = g2.rows('logs');
const col = logs[0].indexOf('نوبت (چک‌این)');
ok('a chore batched with its check-in still names it',
   /^\d{4}-\d{2}-\d{2}/.test(String(logs.find(r => r[0] === 'p1')[col])),
   String(logs.find(r => r[0] === 'p1')[col]));
const sess = g2.rows('نوبت‌ها').find(r => r[0] === 'فاطیما');
ok('the visit is reported as one multi-day session', Number(sess[4]) > 4000, 'minutes=' + sess[4]);


/* ---- readable sheets: nothing overflows, widths follow the content ---- */
const g3 = createGas(FILE);
const now = Date.now();
const long = 'یک یادداشت خیلی طولانی که اگر ستون کوتاه بماند از باکس خودش بیرون می‌زند و جدول را به‌هم می‌ریزد';
g3.post({ token: 'khaneh-1404', since: 0, logs: [
  { id: 'k1', ts: now - 86400000, person: 'فاطیما', type: 'checkin', sess: 'k1' },
  { id: 'k2', ts: now, person: 'فاطیما', type: 'pos', taskId: 'P1', code: 'P1',
    title: 'آب دادن به گل‌ها', note: long, sess: 'k1' },
  { id: 'k3', ts: now + 1, person: 'فاطیما', type: 'checkout', sess: 'k1' }
], config: { people: ['علی', 'فاطیما', 'محمد'], pos: [], neg: [] } });

const REPORTS = ['نوبت‌ها', 'خلاصهٔ افراد', 'روزانه', 'آمار کارها'];
for (const name of ['logs'].concat(REPORTS)) {
  const sh = g3.sheet(name);
  const cols = (g3.rows(name)[0] || []).filter(c => c !== '').length;
  const wrapped = sh.fmt.wrap.some(w => w.v === 'WRAP' && w.nc >= cols);
  const widths = Object.keys(sh.widths).map(k => sh.widths[k]);
  const inRange = widths.length >= cols && widths.every(w => w >= 72 && w <= 320);
  ok(`«${name}» wraps text and never overflows`, wrapped);
  ok(`«${name}» column widths sized and capped`, inRange,
     'min=' + Math.min.apply(null, widths) + ' max=' + Math.max.apply(null, widths));
  ok(`«${name}» header frozen and RTL`, sh.fmt.frozen === 1 && sh.fmt.rtl);
  ok(`«${name}» rows banded for readability`, sh.bandings.length === 1);
}

/* the long note must have widened its own column, up to the cap */
const logs3 = g3.rows('logs');
const noteCol = logs3[0].indexOf('یادداشت') + 1;
ok('a long note widens its column to the cap', g3.sheet('logs').widths[noteCol] === 320,
   'width=' + g3.sheet('logs').widths[noteCol]);
const dateCol = logs3[0].indexOf('تاریخ') + 1;
ok('a short column stays narrow', g3.sheet('logs').widths[dateCol] < 200,
   'width=' + g3.sheet('logs').widths[dateCol]);

/* machine columns are hidden on logs, and only there */
const hiddenLogs = Object.keys(g3.sheet('logs').hidden).map(Number).sort((a, b) => a - b);
ok('logs hides its technical columns', hiddenLogs.join(',') === '14,15,16,17,18,19,20', hiddenLogs.join(','));
ok('reports hide nothing', REPORTS.every(n => Object.keys(g3.sheet(n).hidden).length === 0));
ok('showTechColumns brings them back', (g3.run('showTechColumns'),
   Object.keys(g3.sheet('logs').hidden).length === 0));

/* rebuilding must not lose formatting or leave stale rows */
g3.run('tidy');
ok('tidy re-applies formatting', g3.sheet('نوبت‌ها').bandings.length === 1);


/* ---- deleting moves the row out of logs and into the recoverable bin ---- */
const g4 = createGas(FILE);
const base = Date.now();
g4.post({ token: 'khaneh-1404', since: 0, logs: [
  { id: 'd1', ts: base - 7200000, person: 'محمد', type: 'checkin', sess: 'd1' },
  { id: 'd2', ts: base - 3600000, person: 'محمد', type: 'pos', taskId: 'P5', code: 'P5',
    title: 'بیرون گذاشتن آشغال', icon: '🗑️', note: 'سطل حیاط', by: 'محمد', sess: 'd1' }
]});
const beforeRows = g4.rows('logs').length;
const afterDelete = g4.post({ token: 'khaneh-1404', since: 0,
  logs: [{ id: 'd2', ts: base - 3600000, person: 'محمد', type: 'pos', taskId: 'P5', code: 'P5',
           title: 'بیرون گذاشتن آشغال', sess: 'd1', deleted: true }] });

const liveIds = g4.rows('logs').slice(1).map(r => r[0]);
ok('the deleted row leaves the logs sheet', !liveIds.includes('d2') && liveIds.includes('d1'),
   'logs=' + liveIds.join(','));
ok('logs shrank by exactly one row', g4.rows('logs').length === beforeRows - 1);

const bin = g4.rows(BIN);
const binHead = bin[0];
const gone = bin.slice(1).find(r => r[0] === 'd2') || [];
ok('the bin keeps it with the deletion date and weekday',
   /^\d{4}-\d{2}-\d{2}$/.test(String(gone[binHead.indexOf('تاریخ حذف')])) &&
   String(gone[binHead.indexOf('روز حذف')]).length > 2,
   gone[binHead.indexOf('تاریخ حذف')] + ' ' + gone[binHead.indexOf('روز حذف')]);
ok('the bin keeps the original date and weekday too',
   /^\d{4}-\d{2}-\d{2}$/.test(String(gone[binHead.indexOf('تاریخ ثبت')])) &&
   String(gone[binHead.indexOf('روز ثبت')]).length > 2);
/* the deleting phone sent no note here — the bin fills it from the stored row */
ok('the bin keeps who, what and the note',
   gone[binHead.indexOf('نام')] === 'محمد' &&
   gone[binHead.indexOf('عنوان کار')] === 'بیرون گذاشتن آشغال' &&
   gone[binHead.indexOf('یادداشت')] === 'سطل حیاط',
   JSON.stringify([gone[binHead.indexOf('نام')], gone[binHead.indexOf('یادداشت')]]));
ok('a restored record keeps its icon and task id',
   String(gone[binHead.indexOf('taskId')]) === 'P5' && String(gone[binHead.indexOf('icon')]) === '🗑️',
   gone[binHead.indexOf('taskId')] + ' ' + gone[binHead.indexOf('icon')]);
ok('the deletion is reported to the phones', afterDelete.logs.some(l => l.id === 'd2' && l.deleted));
ok('reports drop the deleted record',
   !(g4.rows('نوبت‌ها').slice(1).some(r => String(r[7]).includes('آشغال'))));

/* ---- restoring puts it back and tells every phone ---- */
const binSh = g4.sheet(BIN);
const restoreCol = binHead.indexOf('بازیابی؟') + 1;
const rowIdx = bin.findIndex(r => r[0] === 'd2') + 1;
binSh.getRange(rowIdx, restoreCol).setValues([['بله']]);
const restored = g4.run('restoreDeleted');
ok('restore reports what it brought back', /1 رکورد/.test(restored), restored);
ok('the record is back in logs', g4.rows('logs').slice(1).some(r => r[0] === 'd2'));
ok('the bin no longer holds it', !g4.rows(BIN).slice(1).some(r => r[0] === 'd2'));

const afterRestore = g4.post({ token: 'khaneh-1404', since: afterDelete.maxRev, logs: [] });
const back = afterRestore.logs.find(l => l.id === 'd2');
ok('phones are told to bring it back', !!back && !back.deleted,
   back ? 'deleted=' + back.deleted : 'not sent');
ok('reports include it again',
   g4.rows('نوبت‌ها').slice(1).some(r => String(r[7]).includes('آشغال')));


/* ---- the Sheet must not go "busy" under normal use ---- */
const g5 = createGas(FILE);
const t5 = Date.now();
g5.run('setup');

/* a read is the common case: three phones polling. it must not take the lock,
   or every poll queues behind whoever is writing and eventually times out. */
const locksBefore = g5.stats.locks;
for (let i = 0; i < 5; i++) g5.post({ token: 'khaneh-1404', since: 0 });
ok('polling reads take no lock at all', g5.stats.locks === locksBefore,
   'locks taken by 5 reads: ' + (g5.stats.locks - locksBefore));

/* writes still take it, so two phones never interleave a write */
g5.post({ token: 'khaneh-1404', since: 0,
          logs: [{ id: 'w1', ts: t5, person: 'علی', type: 'checkin', sess: 'w1' }] });
ok('a write does take the lock', g5.stats.locks === locksBefore + 1);

/* auto-fit is the slow part; it must not run on every write */
const resizesBefore = g5.stats.resizes;
for (let i = 0; i < 6; i++) {
  g5.post({ token: 'khaneh-1404', since: 0,
            logs: [{ id: 'w' + i + 'x', ts: t5 + i, person: 'علی', type: 'pos',
                     taskId: 'P1', code: 'P1', title: 'گل‌ها', sess: 'w1' }] });
}
ok('six writes do not re-fit the columns each time', g5.stats.resizes === resizesBefore,
   'extra auto-fits: ' + (g5.stats.resizes - resizesBefore));
ok('the tidy action still fits them on demand',
   (g5.run('tidy'), g5.stats.resizes > resizesBefore));

/* when the lock really is held, the answer is a retryable "busy", not a failure */
g5.stats.lockBusy = true;
const busy = g5.post({ token: 'khaneh-1404', since: 0,
                       logs: [{ id: 'w9', ts: t5, person: 'علی', type: 'checkin', sess: 'w9' }] });
ok('a contended write answers busy, and says so as retryable',
   busy.ok === false && busy.busy === true, JSON.stringify(busy).slice(0, 70));
g5.stats.lockBusy = false;
const after = g5.post({ token: 'khaneh-1404', since: 0,
                        logs: [{ id: 'w9', ts: t5, person: 'علی', type: 'checkin', sess: 'w9' }] });
ok('the retry then succeeds', after.ok === true && after.logs.some(l => l.id === 'w9'));
done();
