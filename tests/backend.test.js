/* Runs google-apps-script/Code.gs directly, without a browser. */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { createGas } = require('./gas-emu');
const { ok, done } = require('./helpers');

const FILE = path.join(__dirname, '..', 'google-apps-script', 'Code.gs');
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
done();
