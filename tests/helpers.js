const path = require('path');

const BROWSER = process.env.CHROMIUM_PATH || '/opt/pw-browsers/chromium-1194/chrome-linux/chrome';
const BASE = 'http://localhost:' + (process.env.PORT || 877);

let failures = 0;
const ok = (label, cond, extra = '') => {
  if (!cond) failures++;
  console.log((cond ? 'PASS  ' : 'FAIL  ') + label + (extra ? '  [' + extra + ']' : ''));
};
const done = () => { if (failures) { console.log(`\n${failures} check(s) failed`); process.exitCode = 1; } };

function launch() {
  const { chromium } = require('playwright-core');
  return chromium.launch({ executablePath: BROWSER });
}

/* Every question the app asks is drawn in-page; a native dialog would mean a bug. */
async function dlg(scope, index) {
  if (await scope.locator('#modal.on').count()) {
    await scope.click(`[data-dlg="${index}"]`);
    await scope.waitForTimeout ? await scope.waitForTimeout(250) : await new Promise(r => setTimeout(r, 250));
  }
}

const resetBackend = () => fetch(BASE + '/reset');

module.exports = { BASE, ok, done, launch, dlg, resetBackend };
