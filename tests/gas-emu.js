/* Minimal Apps Script emulator: runs the real Code.gs against in-memory sheets. */
const fs = require('fs');
const vm = require('vm');

function p2(n){ return String(n).padStart(2,'0'); }
function formatDate(d, tz, pat){
  const map = {
    'yyyy-MM-dd HH:mm': () => `${d.getFullYear()}-${p2(d.getMonth()+1)}-${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}`,
    'yyyy-MM-dd': () => `${d.getFullYear()}-${p2(d.getMonth()+1)}-${p2(d.getDate())}`,
    'HH:mm': () => `${p2(d.getHours())}:${p2(d.getMinutes())}`,
    'u': () => String(d.getDay() === 0 ? 7 : d.getDay())
  };
  if (!map[pat]) throw new Error('unsupported pattern ' + pat);
  return map[pat]();
}

function makeSheet(name){
  const cells = [];
  const at = (r, c) => { while (cells.length < r) cells.push([]); const row = cells[r-1]; while (row.length < c) row.push(''); return row; };
  const sheet = {
    name, cells,
    widths: {}, hidden: {}, bandings: [],
    fmt: { wrap: [], valign: [], halign: [], fontSize: null, frozen: 0, rtl: false },
    getName(){ return name; },
    getMaxRows(){ return Math.max(cells.length, 1); },
    getMaxColumns(){ return Math.max.apply(null, [1].concat(cells.map(r => r.length))); },
    getLastColumn(){
      let last = 0;
      cells.forEach(r => r.forEach((v, i) => { if (v !== '' && v != null) last = Math.max(last, i + 1); }));
      return last;
    },
    /* width follows the longest text in the column, the way autofit does */
    autoResizeColumns(start, count){
      for (let c = start; c < start + count; c++){
        let widest = 0;
        cells.forEach(r => { const v = r[c-1]; if (v != null) widest = Math.max(widest, String(v).length); });
        sheet.widths[c] = Math.max(24, widest * 7 + 12);
      }
      return sheet;
    },
    getColumnWidth(c){ return sheet.widths[c] || 100; },
    setColumnWidth(c, w){ sheet.widths[c] = w; return sheet; },
    hideColumns(c, n){ for (let i = 0; i < (n || 1); i++) sheet.hidden[c + i] = true; return sheet; },
    showColumns(c, n){ for (let i = 0; i < (n || 1); i++) delete sheet.hidden[c + i]; return sheet; },
    getBandings(){ return sheet.bandings.slice(); },
    clearContents(){ cells.length = 0; return sheet; },
    getLastRow(){ let last = 0; cells.forEach((r, i) => { if (r && r.some(c => c !== '' && c != null)) last = i + 1; }); return last; },
    getRange(row, col, nr = 1, nc = 1){
      const range = {
        setValues(vals){
          vals.forEach((rv, i) => rv.forEach((v, j) => { const r = at(row + i, col + j); r[col + j - 1] = v; }));
          return range;
        },
        setWrapStrategy(v){ sheet.fmt.wrap.push({ row, col, nr, nc, v }); return range; },
        setVerticalAlignment(v){ sheet.fmt.valign.push(v); return range; },
        setHorizontalAlignment(v){ sheet.fmt.halign.push(v); return range; },
        setFontSize(v){ sheet.fmt.fontSize = v; return range; },
        applyRowBanding(theme, header, footer){
          const b = {
            theme, header, footer,
            setHeaderRowColor(c){ b.headerColor = c; return b; },
            setFirstRowColor(c){ b.first = c; return b; },
            setSecondRowColor(c){ b.second = c; return b; },
            setFooterRowColor(c){ b.footerColor = c; return b; },
            remove(){ sheet.bandings = sheet.bandings.filter(x => x !== b); }
          };
          sheet.bandings.push(b);
          return b;
        },
        getValues(){
          const out = [];
          for (let i = 0; i < nr; i++){
            const row2 = [];
            for (let j = 0; j < nc; j++){
              const r = cells[row + i - 1] || [];
              const v = r[col + j - 1];
              row2.push(v === undefined ? '' : v);
            }
            out.push(row2);
          }
          return out;
        },
        setFontWeight(){ return range; }, setBackground(){ return range; }, setFontColor(){ return range; }
      };
      return range;
    },
    clear(){ cells.length = 0; return sheet; },
    setFrozenRows(n){ sheet.fmt.frozen = n; return sheet; },
    setRightToLeft(v){ sheet.fmt.rtl = v !== false; return sheet; }
  };
  return sheet;
}

function createGas(codePath){
  const sheets = new Map();
  const props = new Map();
  const ss = {
    getSheetByName: n => sheets.get(n) || null,
    insertSheet: n => { const s = makeSheet(n); sheets.set(n, s); return s; },
    getSheets: () => [...sheets.values()],
    getSpreadsheetTimeZone: () => 'local',
    toast: () => {}
  };
  const sandbox = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getActive: () => ss,
      flush: () => {},
      getUi: () => { throw new Error('no UI in this context'); },
      WrapStrategy: { WRAP: 'WRAP', OVERFLOW: 'OVERFLOW', CLIP: 'CLIP' },
      BandingTheme: { LIGHT_GREY: 'LIGHT_GREY' }
    },
    Logger: { log: () => {} },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => (props.has(k) ? props.get(k) : null),
      setProperty: (k, v) => props.set(k, v),
      deleteProperty: k => props.delete(k)
    })},
    LockService: { getScriptLock: () => ({ waitLock(){}, releaseLock(){} }) },
    ContentService: {
      MimeType: { JSON: 'json' },
      createTextOutput: t => ({ _t: t, setMimeType(){ return this; }, getContent(){ return this._t; } })
    },
    Utilities: { formatDate },
    Session: { getScriptTimeZone: () => 'local' },
    console
  };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(codePath, 'utf8'), sandbox, { filename: codePath });
  return {
    run(name, ...args){ return sandbox[name](...args); },
    post(body){
      const res = sandbox.doPost({ postData: { contents: JSON.stringify(body) }, parameter: {} });
      return JSON.parse(res.getContent());
    },
    sheetNames(){ return [...sheets.keys()]; },
    sheet(name){ return sheets.get(name) || null; },
    rows(name){
      const s = sheets.get(name);
      if (!s) return null;
      const last = s.getLastRow();
      return last ? s.getRange(1, 1, last, 20).getValues().map(r => r.map(c => (c === undefined ? '' : c))) : [];
    }
  };
}
module.exports = { createGas };
