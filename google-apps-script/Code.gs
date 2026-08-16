/**
 * مراقبت از خانه — بک‌اند Google Sheets
 * ------------------------------------------------------------------
 * این فایل را در Extensions → Apps Script همان Google Sheet بچسبان،
 * TOKEN را عوض کن، و با Deploy → New deployment → Web app منتشرش کن.
 * راهنمای کامل در SETUP-SYNC.md
 */

/** رمز مشترک — همین را در تنظیمات اپ هم وارد کن. حتماً عوضش کن. */
var TOKEN = 'khaneh-1404';

var LOGS_SHEET = 'logs';
var CONFIG_SHEET = 'config';

/** ستون‌های برگهٔ logs — ستون‌های اول برای خواندن آدم، بقیه برای برنامه */
var HEAD = ['id','تاریخ','ساعت','روز هفته','نام','نوع','کد کار','عنوان کار','یادداشت',
            'ثبت‌کننده','امتیاز مثبت','امتیاز منفی','type','taskId','icon','ts','rev','deleted'];
var C = {id:0, date:1, time:2, wd:3, person:4, typeFa:5, code:6, title:7, note:8,
         by:9, pos:10, neg:11, type:12, taskId:13, icon:14, ts:15, rev:16, deleted:17};

var WD = ['یکشنبه','دوشنبه','سه‌شنبه','چهارشنبه','پنجشنبه','جمعه','شنبه'];

function doGet(e){ return handle(e, null); }
function doPost(e){
  var body = null;
  try{ body = JSON.parse(e.postData.contents); }catch(err){ body = null; }
  return handle(e, body);
}

function handle(e, body){
  var token = (body && body.token) || (e && e.parameter && e.parameter.token) || '';
  if(TOKEN && token !== TOKEN) return json({ok:false, error:'رمز اشتباه است'});

  var lock = LockService.getScriptLock();
  try{ lock.waitLock(25000); }
  catch(err){ return json({ok:false, error:'سرور مشغول است، دوباره تلاش کن'}); }

  try{
    var since = Number((body && body.since) || (e && e.parameter && e.parameter.since) || 0);
    if(body && body.logs && body.logs.length) upsertLogs(body.logs);
    if(body && body.config) writeConfig(body.config);
    return json({
      ok: true,
      logs: readLogs(since),
      config: readConfig(since),
      maxRev: currentRev(),
      serverTime: Date.now()
    });
  }catch(err){
    return json({ok:false, error:String(err && err.message || err)});
  }finally{
    lock.releaseLock();
  }
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

function book(){ return SpreadsheetApp.getActiveSpreadsheet(); }

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

function rowOf(log, rev){
  var d = new Date(Number(log.ts) || Date.now());
  var typeFa = log.type === 'checkin' ? 'ورود' : log.type === 'pos' ? 'کار مثبت' : 'امتیاز منفی';
  var r = [];
  r[C.id] = log.id;
  r[C.date] = Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
  r[C.time] = Utilities.formatDate(d, tz(), 'HH:mm');
  r[C.wd] = WD[Number(Utilities.formatDate(d, tz(), 'u')) % 7];
  r[C.person] = log.person || '';
  r[C.typeFa] = typeFa + (log.deleted ? ' (حذف‌شده)' : '');
  r[C.code] = log.code || '';
  r[C.title] = log.title || (log.type === 'checkin' ? 'ورود به خانه' : '');
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
  return r;
}

function upsertLogs(logs){
  var sh = logsSheet();
  var last = sh.getLastRow();
  var ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  var at = {};
  for(var i = 0; i < ids.length; i++) at[String(ids[i][0])] = i + 2;

  var base = nextRev(logs.length);          /* یک بلوک rev برای کل این درخواست */
  var rev = base - logs.length;
  var appends = [];

  for(var k = 0; k < logs.length; k++){
    var log = logs[k];
    if(!log || !log.id) continue;
    rev++;
    var row = rowOf(log, rev);
    var r = at[String(log.id)];
    if(r) sh.getRange(r, 1, 1, HEAD.length).setValues([row]);
    else appends.push(row);
  }
  if(appends.length){
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, HEAD.length).setValues(appends);
  }
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

/* ---------------- ابزار دستی ---------------- */

/** یک بار از منوی Run اجرا کن تا برگه‌ها ساخته شوند (اختیاری). */
function setup(){
  logsSheet(); configSheet();
  SpreadsheetApp.getActive().toast('برگه‌ها آماده شد');
}

/** اگر خواستی همه‌چیز را پاک کنی (اطلاعات از بین می‌رود). */
function resetAll(){
  logsSheet().clear(); configSheet().clear();
  PropertiesService.getScriptProperties().deleteProperty('rev');
  book().getSheetByName(LOGS_SHEET).getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
}
