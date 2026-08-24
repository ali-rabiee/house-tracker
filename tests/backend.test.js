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
done();
