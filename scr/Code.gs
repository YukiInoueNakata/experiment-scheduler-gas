/** ========= 基本ユーティリティ ========= */
function getSS_() {
  if (SS_ID) return SpreadsheetApp.openById(SS_ID);
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('スプレッドシートに紐づいていません。SS_ID を設定してください。');
  return ss;
}

/** ========= シート定義 ========= */
const SHEETS = {
  SLOTS: 'Slots',
  RESP: 'Responses',
  CONF: 'Confirmed',
  ARCH: 'Archive',
  AS: 'AddSlots',
  CO: 'CancelOps',
  MQ: 'MailQueue'
};
const SLOT_HEADERS = ['SlotID','Date','Start','End','Capacity','Location','Status','ConfirmedCount','Timezone'];
const RESP_HEADERS = ['Timestamp','Name','Email','SlotID','Date','Start','End','Status','NotifiedConfirm','NotifiedWait','NotifiedRemind','Notes'];
const CONF_HEADERS = ['SlotID','Date','Start','End','Location','ConfirmedAt','Subject1Name','Subject1Email','Subject2Name','Subject2Email'];
const ARCH_HEADERS = ['ArchivedAt','Timestamp','Name','Email','SlotID','Date','Start','End','Status','Notes'];
const MQ_HEADERS   = ['CreatedAt','Type','To','Subject','Body','ICSText','MetaJson','Status','LastTriedAt','Error'];

/** ========= 正規化（NaN対策） ========= */
function normDateStr_(v, zone) {
  var tz = zone || CONFIG.tz || 'Asia/Tokyo';
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  var d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  throw new Error('Invalid date: ' + s);
}
function normTimeStr_(v, zone) {
  var tz = zone || CONFIG.tz || 'Asia/Tokyo';
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'HH:mm');
  var s = String(v || '').trim();
  if (/^\d{1,2}:\d{2}$/.test(s)) return s;
  if (/^\d{4}$/.test(s)) return s.slice(0,2)+':'+s.slice(2);
  var d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, tz, 'HH:mm');
  throw new Error('Invalid time: ' + s);
}

/** ========= 初期化＆各シート ========= */
function ensureSheets_() {
  var ss = getSS_();
  if (!ss.getSheetByName(SHEETS.SLOTS)) ss.insertSheet(SHEETS.SLOTS).appendRow(SLOT_HEADERS);
  if (!ss.getSheetByName(SHEETS.RESP))  ss.insertSheet(SHEETS.RESP).appendRow(RESP_HEADERS);
  ensureConfirmedSheet_();
  ensureArchiveSheet_();
  ensureMailQueueSheet_();
  ensureAddSlotsSheet_();
  ensureCancelOpsSheet_();
  removeDefaultSheet_();
}
function removeDefaultSheet_(){
  var ss = getSS_();
  ['シート1','Sheet1'].forEach(function(n){
    var sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}
function ensureConfirmedSheet_(){
  var ss = getSS_(); var sh = ss.getSheetByName(SHEETS.CONF);
  if (!sh) { sh = ss.insertSheet(SHEETS.CONF); sh.appendRow(CONF_HEADERS); }
  return sh;
}
function ensureArchiveSheet_(){
  var ss = getSS_(); var sh = ss.getSheetByName(SHEETS.ARCH);
  if (!sh) { sh = ss.insertSheet(SHEETS.ARCH); sh.appendRow(ARCH_HEADERS); }
  return sh;
}
function ensureMailQueueSheet_(){
  var ss = getSS_(); var sh = ss.getSheetByName(SHEETS.MQ);
  if (!sh) { sh = ss.insertSheet(SHEETS.MQ); sh.appendRow(MQ_HEADERS); }
  return sh;
}
function ensureAddSlotsSheet_(){
  var ss = getSS_(); var sh = ss.getSheetByName(SHEETS.AS);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.AS);
    sh.appendRow(['Mode','Date','Start','End','FromDate','ToDate','TimeWindows','ExcludeWeekends','Capacity','Location','Timezone','RespectConfigExcludes','Status','Result']);
    sh.appendRow(['date','2025-09-10','','','','','','FALSE',CONFIG.capacity,CONFIG.location,CONFIG.tz,'TRUE','example','← この行は見本です']);
    sh.setFrozenRows(1);
    var ruleMode = SpreadsheetApp.newDataValidation().requireValueInList(['datetime','date','range'], true).setAllowInvalid(false).build();
    sh.getRange('A2:A1000').setDataValidation(ruleMode);
    var ruleBool = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE','FALSE'], true).setAllowInvalid(false).build();
    sh.getRange('H2:H1000').setDataValidation(ruleBool);
    sh.getRange('L2:L1000').setDataValidation(ruleBool);
    sh.setColumnWidths(1, 12, 140);
  }
  return sh;
}
function ensureCancelOpsSheet_(){
  var ss = getSS_(); var sh = ss.getSheetByName(SHEETS.CO);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.CO);
    sh.appendRow(['Email','Scope','SlotPolicy','Reason','Status','Result']);
    sh.appendRow(['user@example.com','confirmed','refill-slot','本人都合','example','← この行は見本です']);
    sh.setFrozenRows(1);
    var ruleScope = SpreadsheetApp.newDataValidation().requireValueInList(['confirmed','all'], true).setAllowInvalid(false).build();
    sh.getRange('B2:B1000').setDataValidation(ruleScope);
    var rulePolicy = SpreadsheetApp.newDataValidation().requireValueInList(['drop-slot','refill-slot'], true).setAllowInvalid(false).build();
    sh.getRange('C2:C1000').setDataValidation(rulePolicy);
    sh.setColumnWidths(1, 6, 160);
  }
  return sh;
}

/** ========= 枠生成 ========= */
function setup() {
  ensureSheets_();
  generateSlotsFromConfig_();
  setupTriggers(); // トリガーも同時作成
}
function clearSlots_(){
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  sh.clear(); sh.appendRow(SLOT_HEADERS);
}
function generateSlotsFromConfig_(){
  clearSlots_();
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  var start = new Date(CONFIG.startDate+'T00:00:00'), end = new Date(CONFIG.endDate+'T00:00:00');
  var isExcludedDate = function(s){ return (CONFIG.excludeDates||[]).indexOf(s)>=0; };
  var isExcludedDT   = function(d,st,en){ return (CONFIG.excludeDateTimes||[]).indexOf(d+' '+st+'-'+en)>=0; };
  for (var d=new Date(start); d<=end; d=new Date(d.getTime()+86400000)){
    if (CONFIG.excludeWeekends && (d.getDay()===0 || d.getDay()===6)) continue;
    var y=d.getFullYear(), m=('0'+(d.getMonth()+1)).slice(-2), da=('0'+d.getDate()).slice(-2);
    var dateStr = y+'-'+m+'-'+da;
    if (isExcludedDate(dateStr)) continue;
    CONFIG.timeWindows.forEach(function(win){
      var p=win.split('-'); var st=p[0], en=p[1];
      if (isExcludedDT(dateStr, st, en)) return;
      createSlotRowIfNotExists_(dateStr, st, en, CONFIG.capacity, CONFIG.location, CONFIG.tz);
    });
  }
}
function createSlotRowIfNotExists_(dateStr, st, en, cap, loc, tz){
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  var id = dateStr + '_' + st.replace(':','');
  var vals = sh.getDataRange().getValues();
  for (var i=1;i<vals.length;i++){ if (vals[i][0]===id) return false; }
  sh.appendRow([id, dateStr, st, en, cap, loc, 'open', 0, tz]);
  return true;
}

/** ========= Webアプリ ========= */
function doGet() {
  var t = HtmlService.createTemplateFromFile('Index');
  t.title = CONFIG.title;
  t.consentHtml = TEMPLATES.consentHtml;
  t.capacity = CONFIG.capacity;
  return t.evaluate().setTitle(CONFIG.title);
}
function include(filename){ return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function getSlots() {
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  var values = sh.getDataRange().getValues(); var head = values.shift();
  var resp = getResponses_(), bySlot = groupBy_(resp, function(r){ return r.SlotID; });

  var tomorrowStr = null;
  if (CONFIG.showOnlyFromTomorrow) {
    var n=new Date(), t=new Date(n.getFullYear(), n.getMonth(), n.getDate()+1);
    var y=t.getFullYear(), m=('0'+(t.getMonth()+1)).slice(-2), d=('0'+t.getDate()).slice(-2);
    tomorrowStr = y+'-'+m+'-'+d;
  }
  var out = values.map(function(row){
    var rec = asObj_(head,row);
    var ds = normDateStr_(rec.Date), st=normTimeStr_(rec.Start), en=normTimeStr_(rec.End);
    var confirmed = (bySlot[rec.SlotID]||[]).filter(function(r){ return r.Status==='confirmed'; }).length;
    var label = (function(){ var w='日月火水木金土'[ new Date(ds+'T00:00:00+09:00').getDay() ]; return ds+' ('+w+')'; })();
    return { slotId:rec.SlotID, date:ds, dateLabel:label, start:st, end:en, capacity:Number(rec.Capacity),
             status:rec.Status, remaining:Math.max(0, Number(rec.Capacity)-confirmed), tz:rec.Timezone };
  }).filter(function(s){ return !tomorrowStr || s.date >= tomorrowStr; })
    .sort(function(a,b){ return (a.date+a.start).localeCompare(b.date+b.start); });

  return { title: CONFIG.title, slots: out, capacity: CONFIG.capacity };
}

/** ========= 申込処理 ========= */
function register(name, email, slotIds) {
  if (!name || !email || !slotIds || !slotIds.length) throw new Error('入力が不足しています。');
  email = String(email).trim().toLowerCase();

  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var now=new Date(), ss=getSS_(), respSh=ss.getSheetByName(SHEETS.RESP), slotSh=ss.getSheetByName(SHEETS.SLOTS);
    var slotsAll = readSheetAsObjects_(slotSh);

    var existing = getResponses_().filter(function(r){ return String(r.Email).toLowerCase()===email; });
    var already = new Set(existing.map(function(r){ return r.SlotID; }));

    var created=[];
    slotIds.forEach(function(id){
      if (already.has(id)) return;
      var slot = slotsAll.find(function(s){ return s.SlotID===id; });
      if (!slot) return;
      respSh.appendRow([now, name, email, id, slot.Date, slot.Start, slot.End, 'pending', false, false, false, '']);
      created.push({date:slot.Date, start:slot.Start, end:slot.End});
    });

    // 先に受付メール（確実に先に届く）
    if (created.length){
      var lines = created.map(function(s){
        var ds=normDateStr_(s.date), st=normTimeStr_(s.start), en=normTimeStr_(s.end);
        return '・'+fmtJPDateTime_(ds,st)+' - '+en+'（'+CONFIG.tz+'）';
      }).join('\n');
      var subject = renderTemplate_(TEMPLATES.participant.receiptSubject, {});
      var body = renderTemplate_(TEMPLATES.participant.receiptBody, { name:name, lines:lines, fromName:CONFIG.mailFromName });
      MailApp.sendEmail(email, subject, body, {name:CONFIG.mailFromName});
    }

    // 影響スロットの確定処理
    var results=[];
    created.forEach(function(c){
      var slotId = normDateStr_(c.date)+'_'+normTimeStr_(c.start).replace(':','');
      results.push(confirmIfCapacityReached_(slotId));
    });

    // 確定メールは遅延（受付より後に届く）
    scheduleQueueFlush_(2);

    return { ok:true, message:'受付しました。確定の可否はメールでお知らせします。', results:results };
  } finally { lock.releaseLock(); }
}

/** ========= 確定処理 ========= */
function canConfirmMore_(){
  var conf = ensureConfirmedSheet_();
  var rows = conf.getLastRow()-1;
  if (rows <= 0) return true;
  var currentPeople = rows * CONFIG.capacity;
  return currentPeople < CONFIG.totalConfirmCap;
}
function updateSlotAggregate_(slotId, confirmedCount, filled){
  var sh=getSS_().getSheetByName(SHEETS.SLOTS), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head), rowIndex=-1;
  for (var i=0;i<vals.length;i++){ if (vals[i][idx.SlotID]===slotId){ rowIndex=i+2; break; } }
  if (rowIndex>0){
    sh.getRange(rowIndex, idx.ConfirmedCount+1).setValue(confirmedCount);
    sh.getRange(rowIndex, idx.Status+1).setValue(filled ? 'filled' : 'open');
  }
}

function confirmIfCapacityReached_(slotId) {
  const ss = getSS_();
  const slotSh = ss.getSheetByName(SHEETS.SLOTS);
  const respSh = ss.getSheetByName(SHEETS.RESP);

  const slots = readSheetAsObjects_(slotSh);
  const slot = slots.find(s => s.SlotID === slotId);
  if (!slot) return { slotId, status: 'notfound' };

  const cap = parseInt(slot.Capacity, 10);

  // このスロットの全申込（先着順）
  let all = getResponses_().filter(r => r.SlotID === slotId);
  all.sort((a, b) => new Date(a.Timestamp).getTime() - new Date(b.Timestamp).getTime());

  // 既に他スロットで確定済みのメール（1人1枠ポリシー用）
  const allConfirmed = getResponses_().filter(r => r.Status === 'confirmed');
  const confirmedEmails = new Set(allConfirmed.map(r => String(r.Email).toLowerCase()));

  // このスロット内の同一メールの連投は1席に
  const seenEmailsInThisSlot = new Set();

  // 先着 cap 名の候補
  const winners = [];
  for (const r of all) {
    if (winners.length >= cap) break;
    const email = String(r.Email).toLowerCase();
    if (seenEmailsInThisSlot.has(email)) continue;
    if (!CONFIG.allowMultipleConfirmationPerEmail && confirmedEmails.has(email)) continue;
    winners.push(r);
    seenEmailsInThisSlot.add(email);
  }

  const readyToConfirm = all.length >= cap; // 定員に達したか？

  const respValues = respSh.getDataRange().getValues();
  const head = respValues.shift();
  const idx = colIndex_(head);

  const keyOf = (x) => `${String(x.Email).toLowerCase()}|${new Date(x.Timestamp).getTime()}`;
  const winnerKeys = new Set(winners.map(keyOf));

  const newlyConfirmed = [];

  // 1) 定員未満なら「全員 pending」に揃え直して終了
  if (CONFIG.requireFullCapacityToConfirm && !readyToConfirm) {
    respValues.forEach((row, i) => {
      const obj = asObj_(head, row);
      if (obj.SlotID !== slotId) return;
      if (obj.Status !== 'pending') {
        row[idx.Status] = 'pending';
        row[idx.NotifiedConfirm] = false; // 念のため
        respSh.getRange(i + 2, 1, 1, row.length).setValues([row]);
      }
    });
    // スロット集計もリセット
    const slotRow = slots.findIndex(s => s.SlotID === slotId);
    slotSh.getRange(slotRow + 2, SLOT_HEADERS.indexOf('ConfirmedCount') + 1).setValue(0);
    slotSh.getRange(slotRow + 2, SLOT_HEADERS.indexOf('Status') + 1).setValue('open');
    return { slotId, filled: false, confirmedCount: 0, newlyConfirmed: [] };
  }

  // 2) 定員達成：勝者は confirmed、非勝者は必ず waitlist（←★ここを強制降格に修正）
  respValues.forEach((row, i) => {
    const obj = asObj_(head, row);
    if (obj.SlotID !== slotId) return;

    const rowKey = keyOf(obj);
    const isWinner = winnerKeys.has(rowKey);

    if (isWinner) {
      if (obj.Status !== 'confirmed') {
        row[idx.Status] = 'confirmed';
        row[idx.NotifiedWait] = false;
        newlyConfirmed.push({
          rowIndex: i + 2,
          name: obj.Name, email: obj.Email,
          date: obj.Date, start: obj.Start, end: obj.End
        });
      }
    } else {
      // ← ここがポイント：すでに confirmed でも waitlist に降格
      if (obj.Status !== 'waitlist') {
        row[idx.Status] = 'waitlist';
      }
    }
    respSh.getRange(i + 2, 1, 1, row.length).setValues([row]);
  });

  // confirmed人数を再集計
  const confirmedNowCount = getResponses_().filter(r => r.SlotID === slotId && r.Status === 'confirmed').length;

  // スロット側の集計
  const slotRowIdx = slots.findIndex(s => s.SlotID === slotId);
  slotSh.getRange(slotRowIdx + 2, SLOT_HEADERS.indexOf('ConfirmedCount') + 1).setValue(confirmedNowCount);
  slotSh.getRange(slotRowIdx + 2, SLOT_HEADERS.indexOf('Status') + 1).setValue(confirmedNowCount >= cap ? 'filled' : 'open');

  // 参加者宛の確定メール
  newlyConfirmed.forEach(nc => {
    sendConfirmMail_(nc.name, nc.email, nc.date, nc.start, nc.end, slot.Location, slot.Timezone);
    markNotified_(nc.rowIndex, 'NotifiedConfirm', true);
  });

  // Confirmedシートの upsert（常に先着 cap 名を書き戻す）
  const winnersLimited = winners.slice(0, cap);
  const upsert = upsertConfirmedRow_(slot, winnersLimited);
  if (upsert.created) {
    sendAdminConfirmMail_(slot, winnersLimited);
  }

  return {
    slotId,
    filled: confirmedNowCount >= cap,
    confirmedCount: confirmedNowCount,
    newlyConfirmed: newlyConfirmed.map(n => n.email)
  };
}


/** ========= アーカイブ ========= */
function archiveOtherChoicesForEmail_(emailLower, keepSlotId){
  var ss=getSS_(), resp=ss.getSheetByName(SHEETS.RESP), arch=ensureArchiveSheet_();
  var vals=resp.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  for (var i=vals.length-1;i>=0;i--){
    var row=vals[i], obj=asObj_(head,row);
    if (String(obj.Email).toLowerCase()!==emailLower) continue;
    if (obj.SlotID===keepSlotId && obj.Status==='confirmed') continue;
    // pending/waitlist をアーカイブ
    if (obj.Status!=='confirmed'){
      arch.appendRow([new Date(), obj.Timestamp, obj.Name, obj.Email, obj.SlotID, obj.Date, obj.Start, obj.End, obj.Status, 'auto-archived-after-confirm']);
      resp.deleteRow(i+2);
    }
  }
}

/** ========= メール（確定・管理者・リマインド・キュー） ========= */
function makeICS_({title, date, start, end, location, description, tz}) {
  var zone=tz||CONFIG.tz, ds=normDateStr_(date,zone), st=normTimeStr_(start,zone), en=normTimeStr_(end,zone);
  function z(n){ return ('0'+n).slice(-2); }
  var sy=+ds.slice(0,4), sm=+ds.slice(5,7), sd=+ds.slice(8,10), sh=+st.slice(0,2), smin=+st.slice(3,5), eh=+en.slice(0,2), emin=+en.slice(3,5);
  var dtStart = ''+sy+z(sm)+z(sd)+'T'+z(sh)+z(smin)+'00';
  var dtEnd   = ''+sy+z(sm)+z(sd)+'T'+z(eh)+z(emin)+'00';
  return [
    'BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//Experiment Scheduler//JP','CALSCALE:GREGORIAN','METHOD:PUBLISH','BEGIN:VEVENT',
    'DTSTART;TZID='+zone+':'+dtStart,'DTEND;TZID='+zone+':'+dtEnd,'SUMMARY:'+title,'DESCRIPTION:'+description,'LOCATION:'+location,'END:VEVENT','END:VCALENDAR'
  ].join('\r\n');
}
function sendConfirmMail_(name, email, date, start, end, location, tz) {
  var zone=tz||CONFIG.tz, ds=normDateStr_(date,zone), st=normTimeStr_(start,zone), en=normTimeStr_(end,zone);
  var when=fmtJPDateTime_(ds,st)+' - '+en;
  var subject=renderTemplate_(TEMPLATES.participant.confirmSubject,{when:when});
  var body=renderTemplate_(TEMPLATES.participant.confirmBody,{name:name, when:when, tz:zone, location:location, fromName:CONFIG.mailFromName});
  var ics=makeICS_({title:'実験参加', date:ds, start:st, end:en, location:location, description:'実験参加の予約（確定）', tz:zone});
  sendMailSmart_({type:'confirm', to:email, subject:subject, body:body, icsText:ics}); // 確定はキューへ
}
function sendAdminConfirmMail_(slot, winners) {
  if (!CONFIG.adminEmails||!CONFIG.adminEmails.length) return;
  var zone=CONFIG.tz, ds=normDateStr_(slot.Date,zone), st=normTimeStr_(slot.Start,zone), en=normTimeStr_(slot.End,zone);
  var when=fmtJPDateTime_(ds,st)+' - '+en;
  var participants = winners.map(function(w){ return '・'+w.Name+' <'+w.Email+'>'; }).join('\n');
  var subject=renderTemplate_(TEMPLATES.admin.confirmSubject,{when:when, count:winners.length});
  var body=renderTemplate_(TEMPLATES.admin.confirmBody,{when:when, tz:zone, location:CONFIG.location, participants:participants});
  CONFIG.adminEmails.forEach(function(addr){ sendMailSmart_({type:'admin', to:addr, subject:subject, body:body}); });
}
function sendReminders() {
  var tz=CONFIG.tz, now=new Date(), next=new Date(now.getFullYear(), now.getMonth(), now.getDate()+1);
  var yyyy=next.getFullYear(), mm=('0'+(next.getMonth()+1)).slice(-2), dd=('0'+next.getDate()).slice(-2), targetDate=yyyy+'-'+mm+'-'+dd;
  var confirmed=getResponses_().filter(function(r){ return r.Status==='confirmed' && String(r.Date)===targetDate; });
  confirmed.forEach(function(r){
    if (String(r.NotifiedRemind)==='true') return;
    var ds=normDateStr_(r.Date,tz), st=normTimeStr_(r.Start,tz), en=normTimeStr_(r.End,tz), when=fmtJPDateTime_(ds,st)+' - '+en;
    var subject=renderTemplate_(TEMPLATES.participant.remindSubject,{when:when});
    var body=renderTemplate_(TEMPLATES.participant.remindBody,{name:r.Name, when:when, tz:tz, location:CONFIG.location, fromName:CONFIG.mailFromName});
    var res=sendMailSmart_({type:'reminder', to:r.Email, subject:subject, body:body, meta:{timestamp:r.Timestamp, email:String(r.Email).toLowerCase(), slotId:r.SlotID}});
    if (res.sent) markNotifiedByFind_(r, 'NotifiedRemind', true);
  });
}

/** ========= メールキュー＆上限ケア ========= */
function ensureMailQuota_(){ return MailApp.getRemainingDailyQuota(); }
function sendMailSmart_(opt){
  var reserve = (CONFIG.mail && CONFIG.mail.reserveForReminders) || 0;
  var remain = ensureMailQuota_();
  var isReminder = opt.type==='reminder';
  var canUse = Math.max(0, remain - reserve) + (isReminder ? reserve : 0);
  var sendNow = (opt.type==='confirm') ? false : (canUse > 0);

  if (sendNow){
    try{
      if (opt.icsText) GmailApp.sendEmail(opt.to, opt.subject, opt.body, {name:CONFIG.mailFromName, attachments:[Utilities.newBlob(opt.icsText,'text/calendar','invite.ics')]});
      else MailApp.sendEmail(opt.to, opt.subject, opt.body, {name:CONFIG.mailFromName});
      return {sent:true};
    }catch(e){ /* fallthrough to queue */ }
  }
  var sh=ensureMailQueueSheet_();
  sh.appendRow([new Date(), opt.type, opt.to, opt.subject, opt.body, opt.icsText||'', JSON.stringify(opt.meta||{}), 'pending', '', '']);
  return {sent:false, queued:true};
}
function processMailQueue_(){
  var sh=ensureMailQueueSheet_(); var vals=sh.getDataRange().getValues(); if (vals.length<2) return;
  var head=vals[0]; var idx=colIndex_(head);
  var rows=[];
  for (var i=1;i<vals.length;i++){
    if (String(vals[i][idx.Status]).toLowerCase()!=='pending') continue;
    rows.push({row:i+1, arr:vals[i]});
  }
  for (var k=0;k<rows.length;k++){
    var remain=ensureMailQuota_(), type=rows[k].arr[idx.Type], isReminder=(type==='reminder');
    var reserve=(CONFIG.mail && CONFIG.mail.reserveForReminders) || 0;
    var canUse=Math.max(0, remain - reserve) + (isReminder ? reserve : 0);
    if (canUse<=0) break;

    var to=rows[k].arr[idx.To], sub=rows[k].arr[idx.Subject], body=rows[k].arr[idx.Body], ics=rows[k].arr[idx.ICSText];
    try{
      if (ics) GmailApp.sendEmail(to, sub, body, {name:CONFIG.mailFromName, attachments:[Utilities.newBlob(ics,'text/calendar','invite.ics')]});
      else MailApp.sendEmail(to, sub, body, {name:CONFIG.mailFromName});
      sh.getRange(rows[k].row, idx.Status+1).setValue('sent');
      sh.getRange(rows[k].row, idx.LastTriedAt+1).setValue(new Date());
      sh.getRange(rows[k].row, idx.Error+1).setValue('');
    }catch(e){
      sh.getRange(rows[k].row, idx.Status+1).setValue('error');
      sh.getRange(rows[k].row, idx.LastTriedAt+1).setValue(new Date());
      sh.getRange(rows[k].row, idx.Error+1).setValue(String(e));
    }
  }
}
function scheduleQueueFlush_(minutes){ ScriptApp.newTrigger('processMailQueue_').timeBased().after(minutes*60*1000).create(); }

/** ========= Confirmed 反映 ========= */
function upsertConfirmedRow_(slot, winners) {
  const sh = ensureConfirmedSheet_();
  const values = sh.getDataRange().getValues();
  const head = values.shift();
  const idx = colIndex_(head);

  // winners[0], winners[1] を列化（不足分は空欄）
  const s1 = winners[0] || {};
  const s2 = winners[1] || {};
  const now = new Date();

  const rowData = [
    slot.SlotID,
    normDateStr_(slot.Date),
    normTimeStr_(slot.Start),
    normTimeStr_(slot.End),
    slot.Location,
    now,
    s1.Name || '', s1.Email || '',
    s2.Name || '', s2.Email || '',
  ];

  // 既に該当SlotIDがあるなら更新、無ければ追記
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[idx.SlotID] === slot.SlotID) {
      sh.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
      return { created: false, rowIndex: i + 2 };
    }
  }
  sh.appendRow(rowData);
  return { created: true, rowIndex: sh.getLastRow() };
}


/** ========= AddSlots / CancelOps ========= */
function applyAddSlots(){
  var sh=ensureAddSlotsSheet_(), data=sh.getDataRange().getValues(); if (data.length<2) return;
  var head=data.shift(), idx=colIndex_(head);
  var put=function(row, st, msg){ row[idx.Status]=st; row[idx.Result]=msg; };

  for (var i=0;i<data.length;i++){
    var row=data[i], statusNow=String(row[idx.Status]||'').toLowerCase();
    if (statusNow==='done' || statusNow==='example') continue;
    try{
      var mode=String(row[idx.Mode]||'').toLowerCase().trim();
      if (['datetime','date','range'].indexOf(mode)<0) throw new Error('Mode は datetime/date/range');

      var cap=row[idx.Capacity]?Number(row[idx.Capacity]):Number(CONFIG.capacity);
      var loc=row[idx.Location]?String(row[idx.Location]):String(CONFIG.location);
      var tz =row[idx.Timezone]?String(row[idx.Timezone]):String(CONFIG.tz);
      var respect=String(row[idx.RespectConfigExcludes]||'').toUpperCase()==='TRUE';
      var exWk=String(row[idx.ExcludeWeekends]||'').toUpperCase()==='TRUE';

      var added=0, skipped=0;
      var addOne=function(ds, st, en){
        if (respect) {
          if ((CONFIG.excludeDates||[]).indexOf(ds)>=0) { skipped++; return; }
          if ((CONFIG.excludeDateTimes||[]).indexOf(ds+' '+st+'-'+en)>=0) { skipped++; return; }
        }
        createSlotRowIfNotExists_(ds, st, en, cap, loc, tz) ? added++ : skipped++;
      };
      var normDate=function(v){ return normDateStr_(v, tz); };
      var normTime=function(v){ return normTimeStr_(v, tz); };
      var parseTW=function(txt){ return String(txt||'').split(',').map(function(x){return x.trim();}).filter(Boolean).map(function(w){var p=w.split('-'); return [normTime(p[0]), normTime(p[1])];}); };
      var twOrDefault=function(cell){
        var s=String(cell||'').trim().toUpperCase();
        if (!s || s==='DEFAULT') return (CONFIG.timeWindows||[]).map(function(w){var p=w.split('-'); return [normTime(p[0]), normTime(p[1])];});
        return parseTW(s);
      };

      if (mode==='datetime'){
        var ds=normDate(row[idx.Date]), st=row[idx.Start], en=row[idx.End];
        if (!st || !en){ var list=parseTW(row[idx.TimeWindows]); if (list.length!==1) throw new Error('TimeWindows は1つだけ'); st=list[0][0]; en=list[0][1]; }
        else { st=normTime(st); en=normTime(en); }
        addOne(ds, st, en);
      } else if (mode==='date'){
        var ds2=normDate(row[idx.Date]); if (!ds2) throw new Error('Date が必要');
        twOrDefault(row[idx.TimeWindows]).forEach(function(p){ addOne(ds2, p[0], p[1]); });
      } else {
        var from=normDate(row[idx.FromDate]), to=normDate(row[idx.ToDate]); if(!from||!to) throw new Error('From/To が必要');
        var tws=twOrDefault(row[idx.TimeWindows]);
        for (var d=new Date(from+'T00:00:00'); d<=new Date(to+'T00:00:00'); d=new Date(d.getTime()+86400000)){
          if (exWk && (d.getDay()===0 || d.getDay()===6)) continue;
          var y=d.getFullYear(), m=('0'+(d.getMonth()+1)).slice(-2), dd=('0'+d.getDate()).slice(-2);
          var ds3=y+'-'+m+'-'+dd; tws.forEach(function(p){ addOne(ds3, p[0], p[1]); });
        }
      }
      put(row,'done','added='+added+', skipped='+skipped);
    }catch(e){ put(row,'error', String(e)); }
    sh.getRange(i+2,1,1,row.length).setValues([row]);
  }
  SpreadsheetApp.getActive().toast('AddSlots 完了', '追加枠', 5);
}

function applyCancelOps(){
  var sh=ensureCancelOpsSheet_(), values=sh.getDataRange().getValues(); if (values.length<2) return;
  var head=values.shift(), idx=colIndex_(head), ui=SpreadsheetApp.getUi(), notes=[];
  for (var i=0;i<values.length;i++){
    var row=values[i], status=String(row[idx.Status]||'').toLowerCase();
    if (status==='done' || status==='example') continue;
    var put=function(st,msg){ row[idx.Status]=st; row[idx.Result]=msg; sh.getRange(i+2,1,1,row.length).setValues([row]); };
    try{
      var email=String(row[idx.Email]||'').trim().toLowerCase();
      if (!email){ put('error','Email必須'); continue; }
      var scope=String(row[idx.Scope]||'confirmed').trim().toLowerCase();
      var policy=String(row[idx.SlotPolicy]||'refill-slot').trim().toLowerCase();
      var reason=String(row[idx.Reason]||'').trim() || (scope==='all'?'delete-all':'cancel-confirmed');

      var res=performCancellationForEmail_(email, scope, policy, reason);
      res.noCandidateSlots.forEach(function(sid){ notes.push(sid); });
      put('done','removed='+res.removedCount+', slotDrop='+res.slotDropCount+', refilled='+res.refilledCount+', noCandidate='+res.noCandidateSlots.length);
    }catch(e){ put('error', String(e)); }
  }
  if (notes.length) ui.alert('補充できない枠があります','候補不足の枠:\n'+notes.join('\n'), ui.ButtonSet.OK);
  SpreadsheetApp.getActive().toast('CancelOps 完了', 'キャンセル', 5);
}

// 簡易キャンセル骨子
function performCancellationForEmail_(emailLower, scope, policy, reason){
  var ss=getSS_(), resp=ss.getSheetByName(SHEETS.RESP), confSh=ensureConfirmedSheet_();
  var vals=resp.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  var removed=0, slotDrop=0, refilled=0, noCand=[]; var toRemoveIdx=[], confirmedSlots=new Set();

  vals.forEach(function(row,i){
    var obj=asObj_(head,row);
    if (String(obj.Email).toLowerCase()!==emailLower) return;
    if (scope==='confirmed' && obj.Status!=='confirmed') return;
    toRemoveIdx.push(i+2);
    if (obj.Status==='confirmed') confirmedSlots.add(obj.SlotID);
  });

  for (var n=toRemoveIdx.length-1;n>=0;n--){
    var irow=toRemoveIdx[n];
    var row=resp.getRange(irow,1,1,resp.getLastColumn()).getValues()[0];
    var o=asObj_(head,row);
    ensureArchiveSheet_().appendRow([new Date(), o.Timestamp, o.Name, o.Email, o.SlotID, o.Date, o.Start, o.End, o.Status, 'cancel:'+reason]);
    resp.deleteRow(irow); removed++;
  }

  confirmedSlots.forEach(function(slotId){
    if (policy==='drop-slot'){
      var R=getResponses_().filter(function(r){ return r.SlotID===slotId && r.Status==='confirmed'; });
      R.forEach(function(r){
        if (String(r.Email).toLowerCase()===emailLower) return;
        ensureArchiveSheet_().appendRow([new Date(), r.Timestamp, r.Name, r.Email, r.SlotID, r.Date, r.Start, r.End, r.Status, 'slot-canceled']);
        deleteResponseRow_(r);
        var ds=normDateStr_(r.Date), st=normTimeStr_(r.Start), en=normTimeStr_(r.End), when=fmtJPDateTime_(ds,st)+' - '+en;
        var subject=renderTemplate_(TEMPLATES.participant.slotCanceledSubject,{when:when});
        var body=renderTemplate_(TEMPLATES.participant.slotCanceledBody,{name:r.Name, when:when, tz:CONFIG.tz, location:CONFIG.location, fromName:CONFIG.mailFromName});
        sendMailSmart_({type:'admin', to:r.Email, subject:subject, body:body});
      });
      deleteConfirmedRow_(slotId); slotDrop++; updateSlotAggregate_(slotId,0,false);
    } else {
      var waitOrPending=getResponses_().filter(function(r){ return r.SlotID===slotId && (r.Status==='waitlist'||r.Status==='pending'); })
        .sort(function(a,b){ return new Date(a.Timestamp)-new Date(b.Timestamp); });
      if (waitOrPending.length < CONFIG.capacity){ noCand.push(slotId); return; }
      waitOrPending.slice(0, CONFIG.capacity).forEach(function(c){ setResponseStatus_(c,'confirmed'); });
      updateSlotAggregate_(slotId, CONFIG.capacity, true);
      var slot=readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS)).find(function(s){ return s.SlotID===slotId; });
      var winners=getResponses_().filter(function(r){ return r.SlotID===slotId && r.Status==='confirmed'; }).slice(0,CONFIG.capacity);
      upsertConfirmedRow_(slot, winners);
      winners.forEach(function(w){ sendConfirmMail_(w.Name, w.Email, w.Date, w.Start, w.End, slot.Location, slot.Timezone); });
      sendAdminConfirmMail_(slot, winners);
      refilled++;
    }
  });

  return {removedCount:removed, slotDropCount:slotDrop, refilledCount:refilled, noCandidateSlots:noCand};
}
function deleteResponseRow_(rec){
  var sh=getSS_().getSheetByName(SHEETS.RESP), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  for (var i=0;i<vals.length;i++){
    var r=asObj_(head, vals[i]);
    if (r.Timestamp==rec.Timestamp && r.Email==rec.Email && r.SlotID==rec.SlotID){ sh.deleteRow(i+2); return true; }
  } return false;
}
function deleteConfirmedRow_(slotId){
  var sh=ensureConfirmedSheet_(), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  for (var i=0;i<vals.length;i++){ if (vals[i][idx.SlotID]===slotId){ sh.deleteRow(i+2); return true; } }
  return false;
}
function setResponseStatus_(rec, status){
  var sh=getSS_().getSheetByName(SHEETS.RESP), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  for (var i=0;i<vals.length;i++){
    var r=asObj_(head, vals[i]);
    if (r.Timestamp==rec.Timestamp && r.Email==rec.Email && r.SlotID==rec.SlotID){
      vals[i][idx.Status]=status; sh.getRange(i+2,1,1,vals[i].length).setValues([vals[i]]); return true;
    }
  } return false;
}

/** ========= トリガー＆UI ========= */
function setupTriggers() {
  // 既存の同名トリガは掃除
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (['sendReminders','sendDailyAdminDigest','processMailQueue_','onOpenUi_'].includes(fn)) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 前日9:00に参加者へ
  ScriptApp.newTrigger('sendReminders')
    .timeBased().atHour(9).nearMinute(0).everyDays(1).create();

  // 毎日0:00に管理者へダイジェスト
  ScriptApp.newTrigger('sendDailyAdminDigest')
    .timeBased().atHour(0).nearMinute(0).everyDays(1).create();

  // 毎時（xx:10）にメールキュー
  const min = (CONFIG.mail && CONFIG.mail.hourlyQueueTriggerMinute) || 10;
  ScriptApp.newTrigger('processMailQueue_')
    .timeBased().everyHours(1).nearMinute(min).create();

  // ここがポイント：インストール型 onOpen（スタンドアロンでもメニューが出る）
  if (typeof SS_ID === 'string' && SS_ID) {
    ScriptApp.newTrigger('onOpenUi_')
      .forSpreadsheet(SS_ID)
      .onOpen()
      .create();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('スケジューラ')
    .addItem('操作パネル（ボタン）', 'openControlPanel')
    .addItem('追加枠を反映（AddSlots）', 'applyAddSlots')
    .addItem('キャンセル処理（CancelOps）', 'applyCancelOps')
    .addSeparator()
    .addItem('枠生成（setup）', 'setup')
    .addItem('トリガー設定（setupTriggers）', 'setupTriggers')
    .addToUi();
}

function addSchedulerMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('スケジューラ')
    .addItem('操作パネルを開く', 'openControlPanel')
    .addSeparator()
    .addItem('setup（枠生成）', 'setup')
    .addItem('setupTriggers（トリガー作成）', 'setupTriggers')
    .addToUi();
}

// コンテナバインド時に効く
function onOpen() {
  addSchedulerMenu_();
}

// スタンドアロンでも効く（インストール型トリガが呼ぶ）
function onOpenUi_() {
  addSchedulerMenu_();
}

function openControlPanel(){
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui; padding:12px; width:280px;">' +
      '<h3 style="margin:0 0 12px;">スケジューラ 操作</h3>' +
      '<p style="color:#444">各シートに入力してから、該当ボタンを押してください。</p>' +
      '<button onclick="google.script.run.applyAddSlots();this.disabled=true;this.innerText=\'実行中…\';" style="padding:8px 12px;margin-bottom:8px;width:100%;">AddSlots を実行</button>' +
      '<button onclick="google.script.run.applyCancelOps();this.disabled=true;this.innerText=\'実行中…\';" style="padding:8px 12px;margin-bottom:8px;width:100%;">CancelOps を実行</button>' +
      '<hr>' +
      '<button onclick="google.script.run.setup();google.script.run.setupTriggers();this.innerText=\'セットアップ完了\';" style="padding:8px 12px;width:100%;">初期セットアップ（枠生成＆トリガー）</button>' +
    '</div>'
  ).setTitle('スケジューラ操作');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** ========= 共通ヘルパ ========= */
function getResponses_(){ var sh=getSS_().getSheetByName(SHEETS.RESP), vals=sh.getDataRange().getValues(), head=vals.shift(); return vals.map(function(r){ return asObj_(head,r); }); }
function readSheetAsObjects_(sh){ var vals=sh.getDataRange().getValues(), head=vals.shift(); return vals.map(function(r){ return asObj_(head,r); }); }
function asObj_(head,row){ var o={}; head.forEach(function(h,i){ o[h]=row[i]; }); return o; }
function groupBy_(arr,keyFn){ return arr.reduce(function(m,x){ var k=keyFn(x); (m[k]||(m[k]=[])).push(x); return m; },{}); }
function colIndex_(head){ var o={}; head.forEach(function(h,i){ o[h]=i; }); return o; }
function markNotified_(rowIndex, colName, val) {
  var sh=getSS_().getSheetByName(SHEETS.RESP), head=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var idx=head.indexOf(colName)+1; sh.getRange(rowIndex, idx).setValue(val);
}
function markNotifiedByFind_(rec, colName, val) {
  const sh = getSS_().getSheetByName(SHEETS.RESP);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return false;
  const head = data[0], idx = colIndex_(head);
  const recTime = rec.Timestamp instanceof Date ? rec.Timestamp.getTime() : new Date(rec.Timestamp).getTime();
  const recEmail = String(rec.Email).toLowerCase();
  const recSlot  = String(rec.SlotID);
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowTime = row[idx.Timestamp] instanceof Date ? row[idx.Timestamp].getTime() : new Date(row[idx.Timestamp]).getTime();
    const rowEmail = String(row[idx.Email]).toLowerCase();
    const rowSlot  = String(row[idx.SlotID]);
    if (rowTime === recTime && rowEmail === recEmail && rowSlot === recSlot) {
      row[idx[colName]] = !!val;
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return true;
    }
  }
  return false;
}

/** ========= バックフィル（任意で1回） ========= */
function backfillNotifiedConfirm_() {
  const sh = getSS_().getSheetByName(SHEETS.RESP);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return;
  const head = vals[0];
  const idx = colIndex_(head);
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const status = String(row[idx.Status]);
    const flag = String(row[idx.NotifiedConfirm]);
    if (status === 'confirmed' && flag !== 'true') {
      row[idx.NotifiedConfirm] = true;
      sh.getRange(r+1, 1, 1, row.length).setValues([row]);
    }
  }
}
