/**
 * مراقبت از خانه — بک‌اند Google Sheets
 * ------------------------------------------------------------------
 * این فایل را در Extensions → Apps Script همان Google Sheet بچسبان،
 * TOKEN را عوض کن، و با Deploy → New deployment → Web app منتشرش کن.
 * راهنمای کامل در SETUP.md
 */

/** رمز مشترک — همین را در تنظیمات اپ هم وارد کن. حتماً عوضش کن. */
var TOKEN = 'khaneh-1404';

var LOGS_SHEET = 'logs';
var CONFIG_SHEET = 'config';

/** ستون‌های برگهٔ logs — ستون‌های اول برای خواندن آدم، بقیه برای برنامه */
var HEAD = ['id','تاریخ','ساعت','روز هفته','نام','نوع','کد کار','عنوان کار','یادداشت',
            'ثبت‌کننده','امتیاز مثبت','امتیاز منفی','نوبت (چک‌این)',
            'type','taskId','icon','ts','rev','deleted','sess'];
var C = {id:0, date:1, time:2, wd:3, person:4, typeFa:5, code:6, title:7, note:8,
         by:9, pos:10, neg:11, sessAt:12,
         type:13, taskId:14, icon:15, ts:16, rev:17, deleted:18, sess:19};

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
    var changed = false;
    if(body && body.logs && body.logs.length){ upsertLogs(body.logs); changed = true; }
    if(body && body.config){ writeConfig(body.config); changed = true; }
    if(changed) buildReports();
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

/* the check-in time a record points at. tsById must already include the
   rows arriving in this same request, not only what the sheet holds. */
function sessionStamp(log, tsById){
  if(log.type === 'checkin') return fmt(log.ts, 'yyyy-MM-dd HH:mm');
  if(!log.sess) return 'بدون چک‌این';
  var at = tsById[String(log.sess)];
  return at ? fmt(at, 'yyyy-MM-dd HH:mm') : 'بدون چک‌این';
}

function rowOf(log, rev, tsById){
  var d = new Date(Number(log.ts) || Date.now());
  var typeFa = log.type === 'checkin' ? 'ورود'
             : log.type === 'pos' ? 'کار مثبت'
             : log.type === 'neg' ? 'امتیاز منفی'
             : log.type === 'co_ok' ? 'مورد چک‌اوت (انجام شد)'
             : log.type === 'checkout' ? 'چک‌اوت'
             : String(log.type || '');
  var r = [];
  r[C.id] = log.id;
  r[C.date] = Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
  r[C.time] = Utilities.formatDate(d, tz(), 'HH:mm');
  r[C.wd] = WD[Number(Utilities.formatDate(d, tz(), 'u')) % 7];
  r[C.person] = log.person || '';
  r[C.typeFa] = typeFa + (log.deleted ? ' (حذف‌شده)' : '');
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

  for(var k = 0; k < logs.length; k++){
    var log = logs[k];
    if(!log || !log.id) continue;
    rev++;
    var row = rowOf(log, rev, tsById);
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
  var ss = book(), sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); sh.setRightToLeft(true); }
  sh.clear();
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#8A5A34').setFontColor('#FFF8EE');
  if(rows.length){
    var width = head.length;
    var padded = rows.map(function(r){
      var out = r.slice(0, width);
      while(out.length < width) out.push('');
      return out;
    });
    sh.getRange(2, 1, padded.length, width).setValues(padded);
  }
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, head.length);
}

/* ---------------- ابزار دستی ---------------- */

/** یک بار از منوی Run اجرا کن تا برگه‌ها ساخته شوند (اختیاری). */
function setup(){
  logsSheet(); configSheet(); buildReports();
  SpreadsheetApp.getActive().toast('برگه‌ها آماده شد');
}

/** اگر خواستی همه‌چیز را پاک کنی (اطلاعات از بین می‌رود). */
function resetAll(){
  logsSheet().clear(); configSheet().clear();
  PropertiesService.getScriptProperties().deleteProperty('rev');
  book().getSheetByName(LOGS_SHEET).getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
}
