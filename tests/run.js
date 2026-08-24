/* Starts the test host, runs every suite against it, then shuts it down. */
const { spawn, spawnSync } = require('child_process');
const path = require('path');

const SUITES = ['backend.test.js', 'app.test.js', 'sandbox.test.js', 'sync.test.js'];
const only = process.argv[2];

const server = spawn(process.execPath, [path.join(__dirname, 'server.js')], { stdio: 'inherit' });
const stop = () => { try { server.kill('SIGKILL'); } catch (e) {} };
process.on('exit', stop);
process.on('SIGINT', () => { stop(); process.exit(1); });

setTimeout(() => {
  let failed = 0;
  for (const suite of SUITES) {
    if (only && !suite.includes(only)) continue;
    console.log('\n=== ' + suite + ' ===');
    const r = spawnSync(process.execPath, [path.join(__dirname, suite)], { stdio: 'inherit' });
    if (r.status !== 0) failed++;
  }
  stop();
  console.log(failed ? `\n${failed} suite(s) failed` : '\nall suites passed');
  process.exit(failed ? 1 : 0);
}, 1500);
