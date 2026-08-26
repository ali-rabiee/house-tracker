/* Test host: serves the app, and runs the real google-apps-script/Code.gs behind /exec
   so the sync tests exercise the actual backend rather than a stand-in. */
const http = require('http');
const fs = require('fs');
const path = require('path');
const { createGas } = require('./gas-emu');

const ROOT = path.join(__dirname, '..');
const GAS_FILE = path.join(ROOT, 'google-apps-script', 'Code.gs');
let gas = createGas(GAS_FILE);
let hangNext = false;
let busyNext = 0;

/* the artifact viewer runs the page in a sandboxed iframe without allow-modals */
const SANDBOX_PAGE = `<!doctype html><meta charset="utf-8"><title>sandboxed</title>
<iframe id="f" sandbox="allow-scripts allow-same-origin" style="width:430px;height:900px;border:0"
        src="/"></iframe>`;

const server = http.createServer((req, res) => {
  if (req.url.startsWith('/busy')) {           /* answer the next /exec as "busy" once */
    busyNext = Number(new URL(req.url, 'http://x').searchParams.get('n') || 1);
    res.end('ok');
    return;
  }
  if (req.url.startsWith('/hang')) {           /* next /exec never answers, like a dead mobile link */
    hangNext = true;
    res.end('ok');
    return;
  }
  if (req.url.startsWith('/exec')) {
    if (hangNext) { hangNext = false; return; }   /* deliberately leave it open */
    if (busyNext > 0) {
      busyNext--;
      res.setHeader('Content-Type', 'application/json');
      res.end(JSON.stringify({ ok: false, busy: true, error: 'شلوغ است' }));
      return;
    }
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => {
      res.setHeader('Content-Type', 'application/json');
      try { res.end(JSON.stringify(gas.post(JSON.parse(body || '{}')))); }
      catch (e) { res.end(JSON.stringify({ ok: false, error: String(e.message) })); }
    });
    return;
  }
  if (req.url.startsWith('/reset')) {          /* a fresh, empty Sheet per suite */
    gas = createGas(GAS_FILE);
    res.setHeader('Content-Type', 'application/json');
    res.end('{"ok":true}');
    return;
  }
  if (req.url.startsWith('/restore')) {        /* mark every binned row and run the Sheet action */
    const col = Number(new URL(req.url, 'http://x').searchParams.get('col'));
    const sh = gas.sheet('حذف‌شده‌ها');
    const rows = gas.rows('حذف‌شده‌ها') || [];
    for (let r = 2; r <= rows.length; r++) sh.getRange(r, col).setValues([['بله']]);
    res.end(String(gas.run('restoreDeleted')));
    return;
  }
  if (req.url.startsWith('/dump')) {
    const out = {};
    gas.sheetNames().forEach(n => out[n] = gas.rows(n));
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.end(JSON.stringify(out, null, 1));
    return;
  }
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.end(req.url.startsWith('/sandbox') ? SANDBOX_PAGE : fs.readFileSync(path.join(ROOT, 'index.html')));
});

const PORT = Number(process.env.PORT || 877);
server.listen(PORT, () => console.log('test server on http://localhost:' + PORT));
