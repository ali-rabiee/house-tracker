/**
 * مراقبت از خانه — بک‌اند Google Sheets
 * ------------------------------------------------------------------
 * این فایل را در Extensions → Apps Script همان Google Sheet بچسبان،
 * TOKEN را عوض کن، و با Deploy → New deployment → Web app منتشرش کن.
 * راهنمای کامل در SETUP.md
 *
 * ⚠️ برای ساخت برگه‌ها، در منوی کشویی بالای صفحه تابع «setup» انتخاب شده باشد
 *    و بعد Run را بزن. اجرای بقیهٔ توابع هیچ برگه‌ای نمی‌سازد.
 */

/** رمز مشترک — همین را در تنظیمات اپ هم وارد کن. حتماً عوضش کن. */
var TOKEN = 'khaneh-1404';

var LOGS_SHEET = 'logs';
var CONFIG_SHEET = 'config';
var BIN_SHEET = 'حذف‌شده‌ها';

/** ستون‌های برگهٔ logs — ستون‌های اول برای خواندن آدم، بقیه برای برنامه */
var HEAD = ['id','تاریخ','ساعت','روز هفته','نام','نوع','کد کار','عنوان کار','یادداشت',
            'ثبت‌کننده','امتیاز مثبت','امتیاز منفی','نوبت (چک‌این)',
            'type','taskId','icon','ts','rev','deleted','sess'];
var C = {id:0, date:1, time:2, wd:3, person:4, typeFa:5, code:6, title:7, note:8,
         by:9, pos:10, neg:11, sessAt:12,
         type:13, taskId:14, icon:15, ts:16, rev:17, deleted:18, sess:19};

var WD = ['یکشنبه','دوشنبه','سه‌شنبه','چهارشنبه','پنجشنبه','جمعه','شنبه'];

/** ستون‌های فنی برگهٔ logs — پنهان می‌شوند تا جدول خوانا بماند */
var TECH = [C.type, C.taskId, C.icon, C.ts, C.rev, C.deleted, C.sess];

/** سطل حذف‌شده‌ها: رکورد پاک‌شده از logs بیرون می‌رود ولی ردش اینجا می‌ماند
    تا حذف به گوشی‌های دیگر هم برسد و دوباره برنگردد. */
var BIN_HEAD = ['id','تاریخ حذف','روز حذف','ساعت حذف',
                'تاریخ ثبت','روز ثبت','ساعت ثبت','نام','نوع','کد کار','عنوان کار','یادداشت','ثبت‌کننده',
                'بازیابی؟','rev','ts','type','taskId','icon','sess'];
var B = {id:0, delDate:1, delWd:2, delTime:3,
         date:4, wd:5, time:6, person:7, typeFa:8, code:9, title:10, note:11, by:12,
         restore:13, rev:14, ts:15, type:16, taskId:17, icon:18, sess:19};
var BIN_TECH = [B.rev, B.ts, B.type, B.taskId, B.icon, B.sess];

var STYLE = {
  headerBg: '#8A5A34',
  headerFg: '#FFF8EE',
  bandBg:   '#FAF5EC',
  minWidth: 72,
  maxWidth: 320,       /* جلوی ستون‌های عریضِ بی‌انتها را می‌گیرد */
  padding:  18,
  bigSheet: 3000       /* از این تعداد سطر بیشتر، تنظیم خودکار عرض انجام نمی‌شود */
};


/* ================================================================
   ⬇️ این تابع را اجرا کن (setup) — اولین گزینهٔ منوی Run همین است
   ================================================================ */

/** برگه‌ها را می‌سازد و گزارش‌ها را بازسازی می‌کند. */
function setup(){
  var ss = book();
  logsSheet();
  configSheet();
  binSheet();
  buildReports();
  tidy();
  SpreadsheetApp.flush();                  /* تا تغییرات فوراً در شیت دیده شود */
  var names = [];
  ss.getSheets().forEach(function(sh){ names.push(sh.getName()); });
  var msg = 'برگه‌ها آماده شد: ' + names.join('، ');
  try{ ss.toast(msg, 'مراقبت از خانه', 10); }catch(err){}
  Logger.log(msg + '\n(اگر در شیت چیزی نمی‌بینی، صفحهٔ شیت را یک بار refresh کن.)');
  return msg;
}

/** گزارش‌ها را از روی logs دوباره می‌سازد (بدون دست زدن به ثبت‌ها). */
function rebuild(){
  buildReports();
  SpreadsheetApp.flush();
  try{ book().toast('گزارش‌ها بازسازی شد', 'مراقبت از خانه', 6); }catch(err){}
  return 'OK';
}

/**
 * هر سطری از «حذف‌شده‌ها» که در ستون «بازیابی؟» علامت خورده باشد را برمی‌گرداند.
 * رکورد با یک rev تازه به logs برمی‌گردد، پس همهٔ گوشی‌ها در همگام‌سازی بعدی
 * دوباره آن را می‌بینند — حتی اگر اپشان همان لحظه باز باشد.
 */
function restoreDeleted(){
  var sh = binSheet();
  var last = sh.getLastRow();
  if(last < 2) return note('چیزی برای بازیابی نیست');

  var vals = sh.getRange(2, 1, last - 1, BIN_HEAD.length).getValues();
  var back = [], rows = [];
  for(var i = 0; i < vals.length; i++){
    var v = vals[i];
    if(!v[B.id] || !isYes(v[B.restore])) continue;
    back.push({
      id: String(v[B.id]),
      ts: Number(v[B.ts]) || 0,
      person: String(v[B.person] || ''),
      type: String(v[B.type] || ''),
      taskId: String(v[B.taskId] || ''),
      code: String(v[B.code] || ''),
      title: String(v[B.title] || ''),
      icon: String(v[B.icon] || ''),
      note: String(v[B.note] || ''),
      by: String(v[B.by] || ''),
      sess: String(v[B.sess] || ''),
      deleted: false
    });
    rows.push(i + 2);
  }
  if(!back.length) return note('هیچ سطری علامت «بله» در ستون «بازیابی؟» ندارد');

  upsertLogs(back);                                  /* با rev تازه به logs برمی‌گردد */
  rows.sort(function(a,b){ return b - a; });
  for(var d = 0; d < rows.length; d++) sh.deleteRow(rows[d]);   /* دیگر tombstone نیست */
  buildReports();
  tidy();
  SpreadsheetApp.flush();
  return note(back.length + ' رکورد بازیابی شد');
}

function isYes(v){
  var t = String(v == null ? '' : v).trim().toLowerCase();
  return t === 'بله' || t === 'بلی' || t === 'آری' || t === 'yes' || t === 'y'
      || t === 'true' || t === '1' || t === '✓' || t === '✔';
}

function note(msg){
  Logger.log(msg);
  try{ book().toast(msg, 'مراقبت از خانه', 8); }catch(err){}
  return msg;
}

/** عرض ستون‌ها و قالب همهٔ برگه‌ها را دوباره مرتب می‌کند. */
function tidy(){
  var ss = book();
  styleSheet(logsSheet(), HEAD.length, TECH);
  styleSheet(binSheet(), BIN_HEAD.length, BIN_TECH);
  ['نوبت‌ها','خلاصهٔ افراد','روزانه','آمار کارها'].forEach(function(name){
    var sh = ss.getSheetByName(name);
    if(sh) styleSheet(sh, Math.max(sh.getLastColumn(), 1));
  });
  SpreadsheetApp.flush();
  try{ PropertiesService.getScriptProperties().setProperty('tidyAt', String(Date.now())); }catch(err){}
  try{ ss.toast('ستون‌ها مرتب شد', 'مراقبت از خانه', 6); }catch(err){}
  return 'OK';
}

/** منوی «مراقبت از خانه» را به خود شیت اضافه می‌کند (بعد از refresh دیده می‌شود). */
function onOpen(){
  try{
    SpreadsheetApp.getUi()
      .createMenu('مراقبت از خانه')
      .addItem('ساخت / بازسازی برگه‌ها', 'setup')
      .addItem('بازسازی گزارش‌ها', 'rebuild')
      .addItem('بازیابی رکوردهای علامت‌خورده', 'restoreDeleted')
      .addItem('مرتب کردن ستون‌ها', 'tidy')
      .addItem('نمایش ستون‌های فنی', 'showTechColumns')
      .addToUi();
  }catch(err){}
}


/* ---------------- نقطهٔ ورود وب‌اپ ---------------- */
/* این دو تا را دستی Run نکن — فقط از طریق آدرس Web App صدا زده می‌شوند. */

function doGet(e){
  if(!e) throw new Error(MANUAL_RUN);
  return handle(e, null);
}
function doPost(e){
  if(!e) throw new Error(MANUAL_RUN);
  var body = null;
  try{ body = JSON.parse(e.postData.contents); }catch(err){ body = null; }
  return handle(e, body);
}

var MANUAL_RUN = 'این تابع برای درخواست‌های اپ است و دستی اجرا نمی‌شود. '
  + 'در منوی کشویی بالای صفحه «setup» را انتخاب کن و دوباره Run بزن.';

function handle(e, body){
  var token = (body && body.token) || (e && e.parameter && e.parameter.token) || '';
  if(TOKEN && token !== TOKEN) return json({ok:false, error:'رمز اشتباه است'});

  var since = Number((body && body.since) || (e && e.parameter && e.parameter.since) || 0);
  var writes = !!(body && ((body.logs && body.logs.length) || body.config));

  /* گوشی‌ها بیشتر وقت‌ها فقط می‌خوانند. خواندن قفل نمی‌گیرد، وگرنه پشت
     نوشتن‌ها صف می‌بندد و بعد از مدتی «سرور مشغول است» می‌گیرد. */
  if(!writes){
    try{ return json(snapshot(since)); }
    catch(err){ return json({ok:false, error:String(err && err.message || err)}); }
  }

  var lock = LockService.getScriptLock();
  if(!lock.tryLock(10000)){
    /* صف شلوغ است — اپ خودش کمی بعد دوباره تلاش می‌کند و چیزی از دست نمی‌رود */
    return json({ok:false, busy:true, error:'شلوغ است، چند لحظهٔ دیگر دوباره تلاش می‌شود'});
  }
  try{
    if(body.logs && body.logs.length) upsertLogs(body.logs);
    if(body.config) writeConfig(body.config);
    buildReports();
    return json(snapshot(since));
  }catch(err){
    return json({ok:false, error:String(err && err.message || err)});
  }finally{
    lock.releaseLock();
  }
}

function snapshot(since){
  return {
    ok: true,
    logs: readLogs(since).concat(readBin(since)),
    config: readConfig(since),
    maxRev: currentRev(),
    serverTime: Date.now()
  };
}

function json(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- revisions ---------------- */
/* یک شمارندهٔ صعودی سمت سرور: ساعت گوشی‌ها ملاک نیست */

function currentRev(){
  var v = PropertiesService.getScriptProperties().getProperty('rev');
  return v ? Number(v) : 0;
}
function nextRev(n){
  var rev = currentRev() + (n || 1);
  PropertiesService.getScriptProperties().setProperty('rev', String(rev));
  return rev;
}

/* ---------------- sheets ---------------- */

function book(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if(!ss) throw new Error('این اسکریپت به هیچ Google Sheet وصل نیست. '
    + 'از داخل خود شیت با Extensions → Apps Script بازش کن و کد را همان‌جا بچسبان.');
  return ss;
}

function logsSheet(){
  var ss = book(), sh = ss.getSheetByName(LOGS_SHEET);
  if(!sh){
    sh = ss.insertSheet(LOGS_SHEET);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD])
      .setFontWeight('bold').setBackground('#8A5A34').setFontColor('#FFF8EE');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}
function binSheet(){
  var ss = book(), sh = ss.getSheetByName(BIN_SHEET);
  if(!sh){
    sh = ss.insertSheet(BIN_SHEET);
    sh.getRange(1, 1, 1, BIN_HEAD.length).setValues([BIN_HEAD]);
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}
function configSheet(){
  var ss = book(), sh = ss.getSheetByName(CONFIG_SHEET);
  if(!sh){
    sh = ss.insertSheet(CONFIG_SHEET);
    sh.getRange(1, 1, 1, 2).setValues([['config JSON','rev']])
      .setFontWeight('bold').setBackground('#8A5A34').setFontColor('#FFF8EE');
  }
  return sh;
}

function tz(){ return book().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }

/* the check-in time a record points at. tsById must already include the
   rows arriving in this same request, not only what the sheet holds. */
function sessionStamp(log, tsById){
  if(log.type === 'checkin') return fmt(log.ts, 'yyyy-MM-dd HH:mm');
  if(!log.sess) return 'بدون چک‌این';
  var at = tsById[String(log.sess)];
  return at ? fmt(at, 'yyyy-MM-dd HH:mm') : 'بدون چک‌این';
}

function typeLabel(t){
  return t === 'checkin' ? 'ورود'
       : t === 'pos' ? 'کار مثبت'
       : t === 'neg' ? 'امتیاز منفی'
       : t === 'co_ok' ? 'مورد چک‌اوت (انجام شد)'
       : t === 'checkout' ? 'چک‌اوت'
       : String(t || '');
}

function rowOf(log, rev, tsById){
  var d = new Date(Number(log.ts) || Date.now());
  var typeFa = typeLabel(log.type);
  var r = [];
  r[C.id] = log.id;
  r[C.date] = Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
  r[C.time] = Utilities.formatDate(d, tz(), 'HH:mm');
  r[C.wd] = WD[Number(Utilities.formatDate(d, tz(), 'u')) % 7];
  r[C.person] = log.person || '';
  r[C.typeFa] = typeFa;
  r[C.code] = log.code || '';
  r[C.title] = log.title || (log.type === 'checkin' ? 'ورود به خانه' : log.type === 'checkout' ? 'چک‌اوت' : '');
  r[C.note] = log.note || '';
  r[C.by] = log.by || '';
  r[C.pos] = (!log.deleted && log.type === 'pos') ? 1 : 0;
  r[C.neg] = (!log.deleted && log.type === 'neg') ? 1 : 0;
  r[C.type] = log.type || '';
  r[C.taskId] = log.taskId || '';
  r[C.icon] = log.icon || '';
  r[C.ts] = Number(log.ts) || Date.now();
  r[C.rev] = rev;
  r[C.deleted] = log.deleted ? 1 : 0;
  r[C.sess] = log.type === 'checkin' ? log.id : (log.sess || '');
  r[C.sessAt] = sessionStamp(log, tsById || {});
  return r;
}

function upsertLogs(logs){
  var sh = logsSheet();
  var last = sh.getLastRow();
  var existing = last > 1 ? sh.getRange(2, 1, last - 1, HEAD.length).getValues() : [];
  var at = {}, tsById = {};
  for(var i = 0; i < existing.length; i++){
    var eid = String(existing[i][C.id] || '');
    if(!eid) continue;
    at[eid] = i + 2;
    tsById[eid] = Number(existing[i][C.ts]) || 0;
  }
  /* rows in this batch can reference each other (a chore and its check-in) */
  for(var j = 0; j < logs.length; j++){
    if(logs[j] && logs[j].id) tsById[String(logs[j].id)] = Number(logs[j].ts) || 0;
  }

  var base = nextRev(logs.length);          /* یک بلوک rev برای کل این درخواست */
  var rev = base - logs.length;
  var appends = [];

  var drops = [], binned = [];

  for(var k = 0; k < logs.length; k++){
    var log = logs[k];
    if(!log || !log.id) continue;
    rev++;
    if(log.deleted){
      var known = at[String(log.id)];
      if(known) drops.push(known);          /* سطرش از logs برداشته می‌شود */
      /* هرچه گوشی نفرستاده از خودِ سطر موجود برداشته می‌شود تا سطل کامل بماند */
      binned.push(binRowOf(known ? merged(log, existing[known - 2]) : log, rev));
      continue;
    }
    var row = rowOf(log, rev, tsById);
    var r = at[String(log.id)];
    if(r) sh.getRange(r, 1, 1, HEAD.length).setValues([row]);
    else appends.push(row);
  }

  if(appends.length){
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, HEAD.length).setValues(appends);
  }
  /* از پایین به بالا، تا شمارهٔ سطرهای بعدی جابه‌جا نشود */
  drops.sort(function(a,b){ return b - a; });
  for(var d = 0; d < drops.length; d++) sh.deleteRow(drops[d]);
  if(binned.length) writeBin(binned);
}

/** فیلدهای جاافتادهٔ یک حذف را از سطر موجود در logs پر می‌کند */
function merged(log, row){
  if(!row) return log;
  return {
    id: log.id,
    ts: Number(log.ts) || Number(row[C.ts]) || Date.now(),
    person: log.person || String(row[C.person] || ''),
    type: log.type || String(row[C.type] || ''),
    taskId: log.taskId || String(row[C.taskId] || ''),
    code: log.code || String(row[C.code] || ''),
    title: log.title || String(row[C.title] || ''),
    icon: log.icon || String(row[C.icon] || ''),
    note: log.note || String(row[C.note] || ''),
    by: log.by || String(row[C.by] || ''),
    sess: log.sess || String(row[C.sess] || ''),
    deleted: true
  };
}

function binRowOf(log, rev){
  var now = new Date();
  var when = Number(log.ts) || Date.now();
  var was = new Date(when);
  var r = [];
  r[B.id] = log.id;
  r[B.delDate] = fmt(now.getTime(), 'yyyy-MM-dd');
  r[B.delWd] = weekdayFa(now.getTime());
  r[B.delTime] = fmt(now.getTime(), 'HH:mm');
  r[B.date] = fmt(when, 'yyyy-MM-dd');
  r[B.wd] = weekdayFa(when);
  r[B.time] = fmt(when, 'HH:mm');
  r[B.person] = log.person || '';
  r[B.typeFa] = typeLabel(log.type);
  r[B.code] = log.code || '';
  r[B.title] = log.title || (log.type === 'checkin' ? 'ورود به خانه' : log.type === 'checkout' ? 'چک‌اوت' : '');
  r[B.note] = log.note || '';
  r[B.by] = log.by || '';
  r[B.restore] = '';                 /* اینجا «بله» بنویس و بازیابی را اجرا کن */
  r[B.rev] = rev;
  r[B.ts] = when;
  r[B.type] = log.type || '';
  r[B.taskId] = log.taskId || '';
  r[B.icon] = log.icon || '';
  r[B.sess] = log.sess || '';
  return r;
}

function weekdayFa(ts){ return WD[Number(fmt(ts, 'u')) % 7]; }

function writeBin(rows){
  var sh = binSheet();
  var last = sh.getLastRow();
  var at = {};
  if(last > 1){
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for(var i = 0; i < ids.length; i++) at[String(ids[i][0])] = i + 2;
  }
  var appends = [];
  for(var k = 0; k < rows.length; k++){
    var r = at[String(rows[k][B.id])];
    if(r) sh.getRange(r, 1, 1, BIN_HEAD.length).setValues([rows[k]]);
    else appends.push(rows[k]);
  }
  if(appends.length){
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, BIN_HEAD.length).setValues(appends);
  }
}

/** حذف‌ها هم مثل ثبت‌ها rev دارند تا به گوشی‌های دیگر برسند */
function readBin(since){
  var sh = binSheet();
  var last = sh.getLastRow();
  if(last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, BIN_HEAD.length).getValues();
  var out = [];
  for(var i = 0; i < vals.length; i++){
    var v = vals[i];
    if(!v[B.id]) continue;
    var rev = Number(v[B.rev]) || 0;
    if(rev <= since) continue;
    out.push({
      id: String(v[B.id]),
      ts: Number(v[B.ts]) || 0,
      person: String(v[B.person] || ''),
      title: String(v[B.title] || ''),
      code: String(v[B.code] || ''),
      sess: String(v[B.sess] || ''),
      type: String(v[B.type] || ''),
      deleted: true,
      rev: rev
    });
  }
  return out;
}

function readLogs(since){
  var sh = logsSheet();
  var last = sh.getLastRow();
  if(last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
  var out = [];
  for(var i = 0; i < vals.length; i++){
    var v = vals[i];
    if(!v[C.id]) continue;
    var rev = Number(v[C.rev]) || 0;
    if(rev <= since) continue;
    out.push({
      id: String(v[C.id]),
      ts: Number(v[C.ts]) || 0,
      person: String(v[C.person] || ''),
      type: String(v[C.type] || ''),
      taskId: String(v[C.taskId] || ''),
      code: String(v[C.code] || ''),
      title: String(v[C.title] || ''),
      icon: String(v[C.icon] || ''),
      note: String(v[C.note] || ''),
      by: String(v[C.by] || ''),
      sess: String(v[C.sess] || ''),
      deleted: Number(v[C.deleted]) ? true : false,
      rev: rev
    });
  }
  return out;
}

function writeConfig(cfg){
  var sh = configSheet();
  sh.getRange(2, 1, 1, 2).setValues([[JSON.stringify(cfg), nextRev()]]);
}
function readConfig(since){
  var sh = configSheet();
  if(sh.getLastRow() < 2) return null;
  var v = sh.getRange(2, 1, 1, 2).getValues()[0];
  var rev = Number(v[1]) || 0;
  if(!v[0] || rev <= since) return null;
  try{ return {data: JSON.parse(v[0]), rev: rev}; }
  catch(err){ return null; }
}

/* ---------------- خوانا و مرتب نگه داشتن برگه‌ها ---------------- */

/** سطر عنوان، شکستن متن، رنگ‌بندی یک‌درمیان و عرض ستون‌ها بر اساس محتوا */
function styleSheet(sh, cols, hideCols){
  try{
    var last = Math.max(sh.getLastRow(), 1);
    sh.setRightToLeft(true);
    sh.setFrozenRows(1);

    /* هیچ متنی از باکس خودش بیرون نمی‌زند: به‌جای سرریز، در همان خانه می‌شکند */
    sh.getRange(1, 1, last, cols)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('middle');

    sh.getRange(1, 1, 1, cols)
      .setFontWeight('bold')
      .setBackground(STYLE.headerBg)
      .setFontColor(STYLE.headerFg)
      .setHorizontalAlignment('center');

    bandRows(sh, last, cols);
    fitColumns(sh, cols, hideCols);
  }catch(err){
    Logger.log('styleSheet(' + sh.getName() + '): ' + err);   /* قالب‌بندی هرگز نباید جلوی ذخیرهٔ داده را بگیرد */
  }
}

/** رنگ یک‌درمیان سطرها با همان پالت کرم/قهوه‌ای اپ */
function bandRows(sh, last, cols){
  try{
    var old = sh.getBandings();
    for(var i = 0; i < old.length; i++) old[i].remove();
    if(last < 2) return;
    var band = sh.getRange(1, 1, last, cols)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    band.setHeaderRowColor(STYLE.headerBg)
        .setFirstRowColor('#FFFFFF')
        .setSecondRowColor(STYLE.bandBg);
  }catch(err){ Logger.log('bandRows: ' + err); }
}

/** عرض هر ستون به اندازهٔ محتوایش، ولی محدود بین کمینه و بیشینه */
function fitColumns(sh, cols, hideCols){
  if(sh.getLastRow() > STYLE.bigSheet) return;      /* روی برگه‌های خیلی بزرگ وقت‌گیر است */
  sh.showColumns(1, cols);                          /* پنهان‌ها هم اندازه‌گیری شوند */
  sh.autoResizeColumns(1, cols);
  for(var c = 1; c <= cols; c++){
    var w = sh.getColumnWidth(c);
    var want = Math.max(STYLE.minWidth, Math.min(STYLE.maxWidth, w + STYLE.padding));
    if(want !== w) sh.setColumnWidth(c, want);
  }
  if(hideCols && hideCols.length){
    for(var i = 0; i < hideCols.length; i++) sh.hideColumns(hideCols[i] + 1);
  }
}

/** ستون‌های فنی را دوباره نشان می‌دهد (از منوی خود شیت) */
function showTechColumns(){
  var sh = logsSheet();
  sh.showColumns(1, HEAD.length);
  SpreadsheetApp.flush();
  try{ book().toast('ستون‌های فنی نمایش داده شد', 'مراقبت از خانه', 6); }catch(err){}
  return 'OK';
}

/* ---------------- برگه‌های گزارش (خودکار ساخته می‌شوند) ---------------- */
/* هر بار که چیزی ثبت شود، این چهار برگه از روی logs بازنویسی می‌شوند.
   دستی در آن‌ها چیزی ننویس — پاک می‌شود. */

function buildReports(){
  var logs = readLogs(0).filter(function(l){ return !l.deleted; });
  logs.sort(function(a,b){ return a.ts - b.ts; });
  var people = peopleFromConfig(logs);
  var sessions = buildSessions(logs);

  writeReport('نوبت‌ها',
    ['نام','تاریخ','ساعت ورود','ساعت خروج','مدت (دقیقه)','امتیاز مثبت','امتیاز منفی',
     'کارهای انجام‌شده','موارد انجام‌نشدهٔ چک‌اوت','وضعیت'],
    sessions.map(function(x){
      return [
        x.person,
        x.start ? fmt(x.start, 'yyyy-MM-dd') : (x.end ? fmt(x.end, 'yyyy-MM-dd') + ' (بدون چک‌این)' : 'بدون چک‌این'),
        x.start ? fmt(x.start, 'HH:mm') : '',
        x.end ? fmt(x.end, 'HH:mm') : '',
        (x.start && x.end) ? Math.max(0, Math.round((x.end - x.start)/60000)) : '',
        x.pos.length, x.neg.length,
        x.pos.map(label).join('، '),
        x.neg.map(label).join('، '),
        !x.start ? 'ثبت بدون چک‌این' : !x.end ? 'باز — چک‌اوت نشده' : (x.neg.length ? 'بسته با منفی' : 'بسته و کامل')
      ];
    }));

  writeReport('خلاصهٔ افراد',
    ['نام','امتیاز مثبت','امتیاز منفی','تعداد چک‌این','تعداد چک‌اوت','نوبت باز',
     'چک‌اوت کامل','درصد چک‌اوت کامل','آخرین ورود','آخرین فعالیت'],
    people.map(function(p){
      var mine = logs.filter(function(l){ return l.person === p; });
      var ms = sessions.filter(function(x){ return x.person === p; });
      var closed = ms.filter(function(x){ return x.end; });
      var clean = closed.filter(function(x){ return !x.neg.length; }).length;
      var ci = lastOf(mine, 'checkin');
      var last = mine.length ? mine[mine.length-1] : null;
      return [
        p,
        countType(mine, 'pos'), countType(mine, 'neg'),
        countType(mine, 'checkin'), countType(mine, 'checkout'),
        ms.filter(function(x){ return x.start && !x.end; }).length,
        clean,
        closed.length ? Math.round(clean/closed.length*100) : 0,
        ci ? fmt(ci.ts, 'yyyy-MM-dd HH:mm') : '—',
        last ? fmt(last.ts, 'yyyy-MM-dd HH:mm') : '—'
      ];
    }));

  var daily = {};
  logs.forEach(function(l){
    var k = fmt(l.ts, 'yyyy-MM-dd') + '|' + l.person;
    if(!daily[k]) daily[k] = {d: fmt(l.ts, 'yyyy-MM-dd'), p: l.person, pos:0, neg:0, ci:0, co:0};
    if(l.type === 'pos') daily[k].pos++;
    else if(l.type === 'neg') daily[k].neg++;
    else if(l.type === 'checkin') daily[k].ci++;
    else if(l.type === 'checkout') daily[k].co++;
  });
  writeReport('روزانه', ['تاریخ','نام','امتیاز مثبت','امتیاز منفی','چک‌این','چک‌اوت'],
    Object.keys(daily).sort().map(function(k){
      var x = daily[k]; return [x.d, x.p, x.pos, x.neg, x.ci, x.co];
    }));

  var tasks = {};
  logs.forEach(function(l){
    if(l.type !== 'pos' && l.type !== 'neg') return;
    var k = l.taskId || l.code;
    if(!k) return;
    if(!tasks[k]) tasks[k] = {code: l.code || k, title: l.title || k, type: l.type, per: {}};
    tasks[k].per[l.person] = (tasks[k].per[l.person] || 0) + 1;
  });
  writeReport('آمار کارها', ['کد','عنوان','نوع'].concat(people).concat(['مجموع']),
    Object.keys(tasks).map(function(k){
      var t = tasks[k], sum = 0;
      var per = people.map(function(p){ var v = t.per[p] || 0; sum += v; return v; });
      return [t.code, t.title, t.type === 'pos' ? 'مثبت' : 'چک‌اوت — منفی'].concat(per).concat([sum]);
    }));

  maybeTidy();      /* عرض ستون‌ها هر چند دقیقه یک‌بار، نه در هر ثبت */
}

function label(l){ return (l.code ? l.code + ' ' : '') + (l.title || ''); }
function countType(list, t){
  return list.filter(function(l){ return l.type === t; }).length;
}
function lastOf(list, t){
  var f = list.filter(function(l){ return l.type === t; });
  return f.length ? f[f.length-1] : null;
}
function fmt(ts, pattern){ return Utilities.formatDate(new Date(Number(ts)), tz(), pattern); }

function peopleFromConfig(logs){
  var cfg = readConfig(-1);
  if(cfg && cfg.data && cfg.data.people && cfg.data.people.length) return cfg.data.people;
  var seen = [];
  logs.forEach(function(l){ if(l.person && seen.indexOf(l.person) < 0) seen.push(l.person); });
  return seen;
}

/** یک نوبت = یک چک‌این و هر ثبتی که به آن ارجاع می‌دهد */
function buildSessions(logs){
  var byPerson = {};
  logs.forEach(function(l){
    if(!byPerson[l.person]) byPerson[l.person] = [];
    byPerson[l.person].push(l);
  });
  var out = [];
  Object.keys(byPerson).forEach(function(p){
    var mine = byPerson[p];
    var map = {};
    var orphan = {person:p, id:'', start:null, end:null, pos:[], neg:[], ok:[]};

    mine.forEach(function(l){
      if(l.type === 'checkin') map[l.id] = {person:p, id:l.id, start:l.ts, end:null, pos:[], neg:[], ok:[]};
    });
    mine.forEach(function(l){
      if(l.type === 'checkin') return;
      var key = l.sess || nearestCheckin(mine, l.ts);
      var t = map[key] || orphan;
      if(l.type === 'pos') t.pos.push(l);
      else if(l.type === 'neg') t.neg.push(l);
      else if(l.type === 'co_ok') t.ok.push(l);
      else if(l.type === 'checkout') t.end = l.ts;
    });

    Object.keys(map).forEach(function(k){ out.push(map[k]); });
    if(orphan.pos.length || orphan.neg.length || orphan.ok.length || orphan.end) out.push(orphan);
  });
  return out.sort(function(a,b){ return (a.start||a.end||0) - (b.start||b.end||0); });
}

/** ثبت‌های قدیمی که هنوز ارجاع ندارند، با تاریخ به نزدیک‌ترین چک‌این قبلی بسته می‌شوند */
function nearestCheckin(mine, ts){
  var best = '';
  mine.forEach(function(l){
    if(l.type === 'checkin' && l.ts <= ts) best = l.id;
  });
  return best;
}

function writeReport(name, head, rows){
  var ss = book(), sh = ss.getSheetByName(name), fresh = false;
  if(!sh){ sh = ss.insertSheet(name); fresh = true; }
  sh.clearContents();                    /* محتوا پاک می‌شود، قالب‌بندی می‌ماند */
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  if(rows.length){
    var width = head.length;
    var padded = rows.map(function(r){
      var out = r.slice(0, width);
      while(out.length < width) out.push('');
      return out;
    });
    sh.getRange(2, 1, padded.length, width).setValues(padded);
  }
  /* قالب‌بندی روی برگه می‌ماند، پس فقط بار اول لازم است؛ تنظیم دوره‌ای با maybeTidy */
  if(fresh) styleSheet(sh, head.length);
}

/**
 * تنظیم عرض ستون‌ها گران‌ترین کار این اسکریپت است و قفل را نگه می‌دارد.
 * پس حداکثر هر چند دقیقه یک‌بار اجرا می‌شود — نه در هر ثبت.
 */
var TIDY_EVERY_MS = 5 * 60 * 1000;
function maybeTidy(){
  try{
    var props = PropertiesService.getScriptProperties();
    var last = Number(props.getProperty('tidyAt') || 0);
    var now = Date.now();
    if(now - last < TIDY_EVERY_MS) return false;
    props.setProperty('tidyAt', String(now));
    tidy();
    return true;
  }catch(err){ Logger.log('maybeTidy: ' + err); return false; }
}

/* ---------------- ابزار دستی ---------------- */

/** اگر خواستی همه‌چیز را پاک کنی (اطلاعات از بین می‌رود). */
function resetAll(){
  logsSheet().clear(); configSheet().clear(); binSheet().clear();
  PropertiesService.getScriptProperties().deleteProperty('rev');
  book().getSheetByName(LOGS_SHEET).getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
}
