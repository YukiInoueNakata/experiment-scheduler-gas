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

function getSlotHeaders() {
  return ['SlotID','Date','Start','End','Capacity','Location','Status','ConfirmedCount','Timezone'];
}

function getResponseHeaders() {
  return ['Timestamp','Name','Email','SlotID','Date','Start','End','Status','NotifiedConfirm','NotifiedWait','NotifiedRemind','Notes'];
}

function getConfirmedHeaders() {
  const base = ['SlotID','Date','Start','End','Location','ConfirmedAt'];
  for (let i = 1; i <= CONFIG.capacity; i++) {
    base.push(`Subject${i}Name`, `Subject${i}Email`);
  }
  base.push('ActualCount');
  return base;
}

function getArchiveHeaders() {
  return ['ArchivedAt','Timestamp','Name','Email','SlotID','Date','Start','End','Status','Notes','NotifiedConfirm','NotifiedWait','NotifiedRemind','RestoredAt'];
}

const MQ_HEADERS = ['CreatedAt','Type','To','Subject','Body','ICSText','MetaJson','Status','LastTriedAt','Error'];

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
  if (!ss.getSheetByName(SHEETS.SLOTS)) {
    const sh = ss.insertSheet(SHEETS.SLOTS);
    sh.appendRow(getSlotHeaders());
  }
  if (!ss.getSheetByName(SHEETS.RESP)) {
    const sh = ss.insertSheet(SHEETS.RESP);
    sh.appendRow(getResponseHeaders());
  }
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
  var ss = getSS_(); 
  var sh = ss.getSheetByName(SHEETS.CONF);
  if (!sh) { 
    sh = ss.insertSheet(SHEETS.CONF); 
    sh.appendRow(getConfirmedHeaders()); 
  }
  return sh;
}

function ensureArchiveSheet_(){
  var ss = getSS_(); 
  var sh = ss.getSheetByName(SHEETS.ARCH);
  if (!sh) { 
    sh = ss.insertSheet(SHEETS.ARCH); 
    sh.appendRow(getArchiveHeaders()); 
  }
  return sh;
}

function ensureMailQueueSheet_(){
  var ss = getSS_(); 
  var sh = ss.getSheetByName(SHEETS.MQ);
  if (!sh) { 
    sh = ss.insertSheet(SHEETS.MQ); 
    sh.appendRow(MQ_HEADERS); 
  }
  return sh;
}

function ensureAddSlotsSheet_(){
  var ss = getSS_(); 
  var sh = ss.getSheetByName(SHEETS.AS);
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
  var ss = getSS_(); 
  var sh = ss.getSheetByName(SHEETS.CO);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.CO);
    sh.appendRow(['Email','Scope','SlotPolicy','FillPolicy','Reason','Status','Result']);
    sh.appendRow(['user@example.com','confirmed','refill-slot','try-fill','本人都合','example','← この行は見本です']);
    sh.setFrozenRows(1);
    var ruleScope = SpreadsheetApp.newDataValidation().requireValueInList(['confirmed','all'], true).setAllowInvalid(false).build();
    sh.getRange('B2:B1000').setDataValidation(ruleScope);
    var rulePolicy = SpreadsheetApp.newDataValidation().requireValueInList(['drop-slot','refill-slot'], true).setAllowInvalid(false).build();
    sh.getRange('C2:C1000').setDataValidation(rulePolicy);
    var ruleFill = SpreadsheetApp.newDataValidation().requireValueInList(['try-fill','keep-partial','to-pending','cancel-all'], true).setAllowInvalid(false).build();
    sh.getRange('D2:D1000').setDataValidation(ruleFill);
    sh.setColumnWidths(1, 7, 160);
  }
  return sh;
}

/** ========= 枠生成 ========= */
function setup() {
  ensureSheets_();
  generateSlotsFromConfig_();
  setupTriggers();
}

function clearSlots_(){
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  sh.clear(); 
  sh.appendRow(getSlotHeaders());
}

function generateSlotsFromConfig_(){
  clearSlots_();
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  var start = new Date(CONFIG.startDate+'T00:00:00'), end = new Date(CONFIG.endDate+'T00:00:00');
  var isExcludedDate = function(s){ return (CONFIG.excludeDates||[]).indexOf(s)>=0; };
  var isExcludedDT = function(d,st,en){ return (CONFIG.excludeDateTimes||[]).indexOf(d+' '+st+'-'+en)>=0; };
  
  for (var d=new Date(start); d<=end; d=new Date(d.getTime()+86400000)){
    if (CONFIG.excludeWeekends && (d.getDay()===0 || d.getDay()===6)) continue;
    var y=d.getFullYear(), m=('0'+(d.getMonth()+1)).slice(-2), da=('0'+d.getDate()).slice(-2);
    var dateStr = y+'-'+m+'-'+da;
    if (isExcludedDate(dateStr)) continue;
    CONFIG.timeWindows.forEach(function(win){
      var p=win.split('-'); 
      var st=p[0], en=p[1];
      if (isExcludedDT(dateStr, st, en)) return;
      createSlotRowIfNotExists_(dateStr, st, en, CONFIG.capacity, CONFIG.location, CONFIG.tz);
    });
  }
}

function createSlotRowIfNotExists_(dateStr, st, en, cap, loc, tz){
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  var id = dateStr + '_' + st.replace(':','');
  var vals = sh.getDataRange().getValues();
  for (var i=1;i<vals.length;i++){ 
    if (vals[i][0]===id) return false; 
  }
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

function include(filename){ 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

function getSlots() {
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  var values = sh.getDataRange().getValues(); 
  var head = values.shift();
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
    var label = (function(){ 
      var w='日月火水木金土'[ new Date(ds+'T00:00:00+09:00').getDay() ]; 
      return ds+' ('+w+')'; 
    })();
    return { 
      slotId:rec.SlotID, 
      date:ds, 
      dateLabel:label, 
      start:st, 
      end:en, 
      capacity:Number(rec.Capacity),
      status:rec.Status, 
      remaining:Math.max(0, Number(rec.Capacity)-confirmed), 
      tz:rec.Timezone 
    };
  }).filter(function(s){ 
    return !tomorrowStr || s.date >= tomorrowStr; 
  }).sort(function(a,b){ 
    return (a.date+a.start).localeCompare(b.date+b.start); 
  });

  return { title: CONFIG.title, slots: out, capacity: CONFIG.capacity };
}

/** ========= 申込処理（改修版） ========= */
function register(name, email, slotIds) {
  if (!name || !email || !slotIds || !slotIds.length) throw new Error('入力が不足しています。');
  email = String(email).trim().toLowerCase();

  var lock = LockService.getScriptLock(); 
  lock.waitLock(30000);
  
  try {
    var now=new Date(), ss=getSS_(), respSh=ss.getSheetByName(SHEETS.RESP), slotSh=ss.getSheetByName(SHEETS.SLOTS);
    var slotsAll = readSheetAsObjects_(slotSh);

    var existing = getResponses_().filter(function(r){ 
      return String(r.Email).toLowerCase()===email; 
    });
    var already = new Set(existing.map(function(r){ return r.SlotID; }));

    var created=[];
    slotIds.forEach(function(id){
      if (already.has(id)) return;
      var slot = slotsAll.find(function(s){ return s.SlotID===id; });
      if (!slot) return;
      respSh.appendRow([now, name, email, id, slot.Date, slot.Start, slot.End, 'pending', false, false, false, '']);
      created.push({slotId:id, date:slot.Date, start:slot.Start, end:slot.End});
    });

    // 受付メール送信
    if (created.length){
      var lines = created.map(function(s){
        var ds=normDateStr_(s.date), st=normTimeStr_(s.start), en=normTimeStr_(s.end);
        return '・'+fmtJPDateTime_(ds,st)+' - '+en+'（'+CONFIG.tz+'）';
      }).join('\n');
      var subject = renderTemplate_(TEMPLATES.participant.receiptSubject, {});
      var body = renderTemplate_(TEMPLATES.participant.receiptBody, { name:name, lines:lines, fromName:CONFIG.mailFromName });
      MailApp.sendEmail(email, subject, body, {name:CONFIG.mailFromName});
    }

    // 30秒後にバッチ処理をスケジュール
    scheduleDelayedBatch_(CONFIG.batchProcessDelaySeconds || 30);

    return { 
      ok:true, 
      message:'受付しました。確定の可否はメールでお知らせします。', 
      created:created.length 
    };
    
  } finally { 
    lock.releaseLock(); 
  }
}

/** ========= バッチ処理（新規追加） ========= */
function scheduleDelayedBatch_(seconds) {
  ScriptApp.newTrigger('processPendingBatch_')
    .timeBased()
    .after(seconds * 1000)
    .create();
}

function processPendingBatch_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    // 1. 過剰登録のクリーンアップ
    cleanupOverflowedPending_();
    
    // 2. 空き枠への追加登録
    fillRemainingSlots_();
    
    // 3. pendingのスロットごとの処理
    processAllPendingSlots_();
    
    // 4. 確定者の他選択肢をArchive
    archiveConfirmedAlternatives_();
    
    // 5. メール送信
    processMailQueue_();
    
  } finally {
    lock.releaseLock();
  }
}

function cleanupOverflowedPending_() {
  const conf = readSheetAsObjects_(ensureConfirmedSheet_());
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS));
  
  slots.forEach(slot => {
    const slotId = slot.SlotID;
    const capacity = Number(slot.Capacity);
    const confirmed = conf.find(c => c.SlotID === slotId);
    const actualCount = confirmed ? Number(confirmed.ActualCount || 0) : 0;
    
    if (actualCount >= capacity) {
      // 満員のスロットのpending/waitlistをArchive
      const responses = getResponses_().filter(r => 
        r.SlotID === slotId && 
        (r.Status === 'pending' || r.Status === 'waitlist')
      );
      
      responses.forEach(r => {
        moveToArchive_(r, 'slot-already-full');
        deleteResponseRow_(r);
      });
    }
  });
}

function fillRemainingSlots_() {
  const conf = readSheetAsObjects_(ensureConfirmedSheet_());
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS));
  
  conf.forEach(confirmed => {
    const actualCount = Number(confirmed.ActualCount || 0);
    const slot = slots.find(s => s.SlotID === confirmed.SlotID);
    if (!slot) return;
    
    const capacity = Number(slot.Capacity);
    const availableSeats = capacity - actualCount;
    
    if (availableSeats > 0 && actualCount >= CONFIG.minCapacityToConfirm) {
      // 空き枠があり、最小人数を満たしている
      const candidates = getResponses_()
        .filter(r => r.SlotID === confirmed.SlotID && 
                    (r.Status === 'pending' || r.Status === 'waitlist'))
        .sort((a,b) => new Date(a.Timestamp) - new Date(b.Timestamp));
      
      const toAdd = candidates.slice(0, availableSeats).filter(c => 
        !hasConfirmedElsewhere_(c.Email, confirmed.SlotID)
      );
      
      toAdd.forEach(c => {
        setResponseStatus_(c, 'confirmed');
        sendConfirmMail_(c.Name, c.Email, c.Date, c.Start, c.End, slot.Location, slot.Timezone);
      });
      
      if (toAdd.length > 0) {
        updateConfirmedSheet_(confirmed.SlotID);
      }
    }
  });
}

function processAllPendingSlots_() {
  const pendingResponses = getResponses_().filter(r => r.Status === 'pending');
  const bySlot = groupBy_(pendingResponses, r => r.SlotID);
  
  Object.keys(bySlot).forEach(slotId => {
    confirmIfCapacityReached_(slotId);
  });
}

/** ========= 確定処理（改修版） ========= */
function confirmIfCapacityReached_(slotId) {
  const ss = getSS_();
  const slotSh = ss.getSheetByName(SHEETS.SLOTS);
  const respSh = ss.getSheetByName(SHEETS.RESP);

  const slots = readSheetAsObjects_(slotSh);
  const slot = slots.find(s => s.SlotID === slotId);
  if (!slot) return { slotId, status: 'notfound' };

  const cap = parseInt(slot.Capacity, 10);
  const minCap = CONFIG.minCapacityToConfirm;

  // このスロットの全申込（先着順）
  let all = getResponses_().filter(r => r.SlotID === slotId);
  all.sort((a, b) => new Date(a.Timestamp).getTime() - new Date(b.Timestamp).getTime());

  // 既に他スロットで確定済みのメール
  const allConfirmed = getResponses_().filter(r => r.Status === 'confirmed');
  const confirmedEmails = new Set(allConfirmed.map(r => String(r.Email).toLowerCase()));

  // このスロット内の重複を排除
  const seenEmailsInThisSlot = new Set();
  const candidates = [];
  
  for (const r of all) {
    if (candidates.length >= cap) break;
    const email = String(r.Email).toLowerCase();
    if (seenEmailsInThisSlot.has(email)) continue;
    if (!CONFIG.allowMultipleConfirmationPerEmail && confirmedEmails.has(email)) continue;
    candidates.push(r);
    seenEmailsInThisSlot.add(email);
  }

  const canConfirm = candidates.length >= minCap;

  const respValues = respSh.getDataRange().getValues();
  const head = respValues.shift();
  const idx = colIndex_(head);

  const newlyConfirmed = [];

  if (!canConfirm) {
    // 最小人数未満 → 全員pending維持
    respValues.forEach((row, i) => {
      const obj = asObj_(head, row);
      if (obj.SlotID !== slotId) return;
      if (obj.Status !== 'pending') {
        row[idx.Status] = 'pending';
        row[idx.NotifiedConfirm] = false;
        respSh.getRange(i + 2, 1, 1, row.length).setValues([row]);
      }
    });
    updateSlotAggregate_(slotId, 0, false);
    return { slotId, filled: false, confirmedCount: 0, newlyConfirmed: [] };
  }

  // 確定処理
  const winners = candidates.slice(0, cap);
  const winnerEmails = new Set(winners.map(w => String(w.Email).toLowerCase()));

  respValues.forEach((row, i) => {
    const obj = asObj_(head, row);
    if (obj.SlotID !== slotId) return;

    const email = String(obj.Email).toLowerCase();
    const isWinner = winnerEmails.has(email);

    if (isWinner) {
      if (obj.Status !== 'confirmed') {
        row[idx.Status] = 'confirmed';
        row[idx.NotifiedWait] = false;
        newlyConfirmed.push({
          rowIndex: i + 2,
          name: obj.Name, 
          email: obj.Email,
          date: obj.Date, 
          start: obj.Start, 
          end: obj.End
        });
      }
    } else {
      if (obj.Status !== 'waitlist') {
        row[idx.Status] = 'waitlist';
      }
    }
    respSh.getRange(i + 2, 1, 1, row.length).setValues([row]);
  });

  // 確定人数を集計
  const confirmedNowCount = winners.length;

  // スロット集計更新
  updateSlotAggregate_(slotId, confirmedNowCount, confirmedNowCount >= cap);

  // 確定メール送信
  newlyConfirmed.forEach(nc => {
    sendConfirmMail_(nc.name, nc.email, nc.date, nc.start, nc.end, slot.Location, slot.Timezone);
    markNotified_(nc.rowIndex, 'NotifiedConfirm', true);
  });

  // Confirmedシート更新
  updateConfirmedSheet_(slotId);
  
  // 管理者メール
  if (newlyConfirmed.length > 0) {
    sendAdminConfirmMail_(slot, winners);
  }

  return {
    slotId,
    filled: confirmedNowCount >= cap,
    confirmedCount: confirmedNowCount,
    newlyConfirmed: newlyConfirmed.map(n => n.email)
  };
}

function updateConfirmedSheet_(slotId) {
  const sh = ensureConfirmedSheet_();
  const slot = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .find(s => s.SlotID === slotId);
  if (!slot) return;
  
  const confirmed = getResponses_()
    .filter(r => r.SlotID === slotId && r.Status === 'confirmed')
    .sort((a,b) => new Date(a.Timestamp) - new Date(b.Timestamp));
  
  const rowData = [
    slotId,
    normDateStr_(slot.Date),
    normTimeStr_(slot.Start),
    normTimeStr_(slot.End),
    slot.Location,
    new Date()
  ];
  
  // 各参加者の情報を追加
  for (let i = 0; i < CONFIG.capacity; i++) {
    if (i < confirmed.length) {
      rowData.push(confirmed[i].Name, confirmed[i].Email);
    } else {
      rowData.push('', '');
    }
  }
  
  rowData.push(confirmed.length); // ActualCount
  
  // upsert処理
  const values = sh.getDataRange().getValues();
  const head = values.shift();
  const idx = colIndex_(head);
  
  let found = false;
  for (let i = 0; i < values.length; i++) {
    if (values[i][idx.SlotID] === slotId) {
      sh.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
      found = true;
      break;
    }
  }
  
  if (!found) {
    sh.appendRow(rowData);
  }
}

/** ========= アーカイブ処理（改修版） ========= */
function moveToArchive_(record, reason) {
  const archSh = ensureArchiveSheet_();
  const archiveData = [
    new Date(),                 // ArchivedAt
    record.Timestamp,
    record.Name,
    record.Email,
    record.SlotID,
    record.Date,
    record.Start,
    record.End,
    record.Status,
    reason,
    record.NotifiedConfirm || false,
    record.NotifiedWait || false,
    record.NotifiedRemind || false,
    ''                          // RestoredAt (empty initially)
  ];
  archSh.appendRow(archiveData);
}

function archiveConfirmedAlternatives_() {
  const confirmed = getResponses_().filter(r => r.Status === 'confirmed');
  const confirmedByEmail = groupBy_(confirmed, r => String(r.Email).toLowerCase());
  
  Object.keys(confirmedByEmail).forEach(email => {
    const slots = confirmedByEmail[email];
    if (slots.length === 0) return;
    
    const confirmedSlotId = slots[0].SlotID;
    
    // この人の他の申込みをArchive
    const others = getResponses_().filter(r => 
      String(r.Email).toLowerCase() === email && 
      r.SlotID !== confirmedSlotId &&
      r.Status !== 'confirmed'
    );
    
    others.forEach(r => {
      moveToArchive_(r, 'auto-archived-confirmed-elsewhere');
      deleteResponseRow_(r);
    });
  });
}

/** ========= Archive復元機能（新規追加） ========= */
function restoreFromArchiveIfEligible(email, slotId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    // 1. 他で確定していないことを確認
    const currentConfirmed = getResponses_()
      .filter(r => 
        String(r.Email).toLowerCase() === email.toLowerCase() && 
        r.Status === 'confirmed'
      );
    
    if (!CONFIG.allowMultipleConfirmationPerEmail && currentConfirmed.length > 0) {
      return {
        restored: false,
        reason: 'already-confirmed-elsewhere',
        confirmedSlot: currentConfirmed[0].SlotID
      };
    }
    
    // 2. Archiveから該当レコードを検索
    const archSh = ensureArchiveSheet_();
    const archData = archSh.getDataRange().getValues();
    const archHead = archData.shift();
    const archIdx = colIndex_(archHead);
    
    let targetRow = -1;
    let targetRecord = null;
    
    for (let i = archData.length - 1; i >= 0; i--) {
      const row = archData[i];
      if (String(row[archIdx.Email]).toLowerCase() === email.toLowerCase() &&
          row[archIdx.SlotID] === slotId &&
          String(row[archIdx.Notes]).includes('auto-archived')) {
        targetRow = i + 2;
        targetRecord = row;
        break;
      }
    }
    
    if (!targetRecord) {
      return { restored: false, reason: 'not-found-in-archive' };
    }
    
    // 3. 既に登録済みでないか確認
    const existing = getResponses_().filter(r => 
      String(r.Email).toLowerCase() === email.toLowerCase() && 
      r.SlotID === slotId
    );
    
    if (existing.length > 0) {
      return { restored: false, reason: 'already-registered' };
    }
    
    // 4. Responsesシートに復元
    const respSh = getSS_().getSheetByName(SHEETS.RESP);
    const restoredData = [
      targetRecord[archIdx.Timestamp],
      targetRecord[archIdx.Name],
      targetRecord[archIdx.Email],
      targetRecord[archIdx.SlotID],
      targetRecord[archIdx.Date],
      targetRecord[archIdx.Start],
      targetRecord[archIdx.End],
      'waitlist',  // 復元時はwaitlistとして
      false,       // NotifiedConfirm リセット
      false,       // NotifiedWait リセット
      false,       // NotifiedRemind リセット
      'restored-from-archive'
    ];
    
    respSh.appendRow(restoredData);
    
    // 5. Archiveの復元日時を更新
    archSh.getRange(targetRow, archIdx.RestoredAt + 1)
      .setValue(new Date());
    
    return {
      restored: true,
      email: email,
      slotId: slotId,
      newStatus: 'waitlist'
    };
    
  } finally {
    lock.releaseLock();
  }
}

/** ========= メール関連 ========= */
function updateSlotAggregate_(slotId, confirmedCount, filled){
  var sh=getSS_().getSheetByName(SHEETS.SLOTS), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head), rowIndex=-1;
  for (var i=0;i<vals.length;i++){ 
    if (vals[i][idx.SlotID]===slotId){ 
      rowIndex=i+2; 
      break; 
    } 
  }
  if (rowIndex>0){
    sh.getRange(rowIndex, idx.ConfirmedCount+1).setValue(confirmedCount);
    sh.getRange(rowIndex, idx.Status+1).setValue(filled ? 'filled' : 'open');
  }
}

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
  sendMailSmart_({type:'confirm', to:email, subject:subject, body:body, icsText:ics});
}

function sendAdminConfirmMail_(slot, winners) {
  if (!CONFIG.adminEmails||!CONFIG.adminEmails.length) return;
  var zone=CONFIG.tz, ds=normDateStr_(slot.Date,zone), st=normTimeStr_(slot.Start,zone), en=normTimeStr_(slot.End,zone);
  var when=fmtJPDateTime_(ds,st)+' - '+en;
  var participants = winners.map(function(w){ return '・'+w.Name+' <'+w.Email+'>'; }).join('\n');
  var subject=renderTemplate_(TEMPLATES.admin.confirmSubject,{when:when, count:winners.length});
  var body=renderTemplate_(TEMPLATES.admin.confirmBody,{when:when, tz:zone, location:CONFIG.location, participants:participants});
  CONFIG.adminEmails.forEach(function(addr){ 
    sendMailSmart_({type:'admin', to:addr, subject:subject, body:body}); 
  });
}

function sendReminders() {
  var tz=CONFIG.tz, now=new Date(), next=new Date(now.getFullYear(), now.getMonth(), now.getDate()+1);
  var yyyy=next.getFullYear(), mm=('0'+(next.getMonth()+1)).slice(-2), dd=('0'+next.getDate()).slice(-2), targetDate=yyyy+'-'+mm+'-'+dd;
  var confirmed=getResponses_().filter(function(r){ 
    return r.Status==='confirmed' && String(r.Date)===targetDate; 
  });
  confirmed.forEach(function(r){
    if (String(r.NotifiedRemind)==='true') return;
    var ds=normDateStr_(r.Date,tz), st=normTimeStr_(r.Start,tz), en=normTimeStr_(r.End,tz), when=fmtJPDateTime_(ds,st)+' - '+en;
    var subject=renderTemplate_(TEMPLATES.participant.remindSubject,{when:when});
    var body=renderTemplate_(TEMPLATES.participant.remindBody,{name:r.Name, when:when, tz:tz, location:CONFIG.location, fromName:CONFIG.mailFromName});
    var res=sendMailSmart_({type:'reminder', to:r.Email, subject:subject, body:body, meta:{timestamp:r.Timestamp, email:String(r.Email).toLowerCase(), slotId:r.SlotID}});
    if (res.sent) markNotifiedByFind_(r, 'NotifiedRemind', true);
  });
}

/** ========= メールキュー処理 ========= */
function ensureMailQuota_(){ 
  return MailApp.getRemainingDailyQuota(); 
}

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
  var sh=ensureMailQueueSheet_(); 
  var vals=sh.getDataRange().getValues(); 
  if (vals.length<2) return;
  var head=vals[0]; 
  var idx=colIndex_(head);
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

/** ========= キャンセル処理（改修版） ========= */
function applyCancelOps(){
  var sh=ensureCancelOpsSheet_(), values=sh.getDataRange().getValues(); 
  if (values.length<2) return;
  var head=values.shift(), idx=colIndex_(head), ui=SpreadsheetApp.getUi(), notes=[];
  for (var i=0;i<values.length;i++){
    var row=values[i], status=String(row[idx.Status]||'').toLowerCase();
    if (status==='done' || status==='example') continue;
    var put=function(st,msg){ 
      row[idx.Status]=st; 
      row[idx.Result]=msg; 
      sh.getRange(i+2,1,1,row.length).setValues([row]); 
    };
    try{
      var email=String(row[idx.Email]||'').trim().toLowerCase();
      if (!email){ put('error','Email必須'); continue; }
      var scope=String(row[idx.Scope]||'confirmed').trim().toLowerCase();
      var policy=String(row[idx.SlotPolicy]||'refill-slot').trim().toLowerCase();
      var fillPolicy=String(row[idx.FillPolicy]||'try-fill').trim().toLowerCase();
      var reason=String(row[idx.Reason]||'').trim() || 'cancel';

      var res=performCancellationForEmail_(email, scope, policy, fillPolicy, reason);
      res.noCandidateSlots.forEach(function(sid){ notes.push(sid); });
      put('done','removed='+res.removedCount+', refilled='+res.refilledCount+', dropped='+res.droppedCount);
    }catch(e){ 
      put('error', String(e)); 
    }
  }
  if (notes.length) ui.alert('補充できない枠があります','候補不足の枠:\n'+notes.join('\n'), ui.ButtonSet.OK);
  SpreadsheetApp.getActive().toast('CancelOps 完了', 'キャンセル', 5);
}

function performCancellationForEmail_(emailLower, scope, policy, fillPolicy, reason){
  var ss=getSS_(), resp=ss.getSheetByName(SHEETS.RESP);
  var vals=resp.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  var removed=0, refilled=0, dropped=0, noCand=[]; 
  var toRemoveIdx=[], confirmedSlots=new Set();

  // 削除対象の特定
  vals.forEach(function(row,i){
    var obj=asObj_(head,row);
    if (String(obj.Email).toLowerCase()!==emailLower) return;
    if (scope==='confirmed' && obj.Status!=='confirmed') return;
    toRemoveIdx.push(i+2);
    if (obj.Status==='confirmed') confirmedSlots.add(obj.SlotID);
  });

  // Archive & 削除
  for (var n=toRemoveIdx.length-1;n>=0;n--){
    var irow=toRemoveIdx[n];
    var row=resp.getRange(irow,1,1,resp.getLastColumn()).getValues()[0];
    var o=asObj_(head,row);
    moveToArchive_(o, 'cancel:'+reason);
    resp.deleteRow(irow); 
    removed++;
  }

  // 各確定枠の処理
  confirmedSlots.forEach(function(slotId){
    if (policy==='drop-slot'){
      // スロット全体をキャンセル
      dropEntireSlot_(slotId);
      dropped++;
    } else {
      // 補充を試みる
      const result = tryRefillSlot_(slotId, fillPolicy);
      if (result.refilled) {
        refilled++;
      } else {
        noCand.push(slotId);
      }
    }
  });

  return {removedCount:removed, refilledCount:refilled, droppedCount:dropped, noCandidateSlots:noCand};
}

function tryRefillSlot_(slotId, fillPolicy) {
  const slot = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .find(s => s.SlotID === slotId);
  if (!slot) return {refilled: false};
  
  const capacity = Number(slot.Capacity);
  const minCap = CONFIG.minCapacityToConfirm;
  
  // 現在の確定者数
  const currentConfirmed = getResponses_()
    .filter(r => r.SlotID === slotId && r.Status === 'confirmed');
  const currentCount = currentConfirmed.length;
  
  // 補充必要数
  const needed = capacity - currentCount;
  
  // 1. Responsesから候補を探す
  let candidates = getResponses_()
    .filter(r => r.SlotID === slotId && 
            (r.Status === 'waitlist' || r.Status === 'pending'))
    .filter(r => !hasConfirmedElsewhere_(r.Email, slotId))
    .sort((a,b) => new Date(a.Timestamp) - new Date(b.Timestamp));
  
  // 2. 不足の場合はArchiveから復元
  if (candidates.length < needed) {
    const additionalNeeded = needed - candidates.length;
    const restored = restoreFromArchiveForSlot_(slotId, additionalNeeded);
    candidates = candidates.concat(restored);
  }
  
  // 3. 補充後の人数チェック
  const afterFillCount = currentCount + Math.min(candidates.length, needed);
  
  if (afterFillCount < minCap) {
    // 最小人数未満
    switch(fillPolicy) {
      case 'keep-partial':
        // 人数不足でも維持
        break;
      case 'to-pending':
        // 全員pendingに戻す
        currentConfirmed.forEach(r => setResponseStatus_(r, 'pending'));
        updateSlotAggregate_(slotId, 0, false);
        return {refilled: false};
      case 'cancel-all':
        // 全員キャンセル
        dropEntireSlot_(slotId);
        return {refilled: false};
      default: // try-fill
        if (currentCount > 0) {
          // 現在の確定者は維持
        } else {
          return {refilled: false};
        }
    }
  }
  
  // 4. 補充実行
  const toPromote = candidates.slice(0, needed);
  toPromote.forEach(c => {
    setResponseStatus_(c, 'confirmed');
    sendConfirmMail_(c.Name, c.Email, c.Date, c.Start, c.End, slot.Location, slot.Timezone);
  });
  
  // 5. Confirmedシート更新
  if (toPromote.length > 0) {
    updateConfirmedSheet_(slotId);
    const newConfirmed = getResponses_()
      .filter(r => r.SlotID === slotId && r.Status === 'confirmed');
    sendAdminConfirmMail_(slot, newConfirmed);
  }
  
  return {refilled: toPromote.length > 0};
}

function restoreFromArchiveForSlot_(slotId, maxCount) {
  const archSh = ensureArchiveSheet_();
  const archData = archSh.getDataRange().getValues();
  if (archData.length < 2) return [];
  
  const archHead = archData.shift();
  const archIdx = colIndex_(archHead);
  const restored = [];
  
  for (let i = archData.length - 1; i >= 0 && restored.length < maxCount; i--) {
    const row = archData[i];
    if (row[archIdx.SlotID] !== slotId) continue;
    if (!String(row[archIdx.Notes]).includes('auto-archived')) continue;
    if (row[archIdx.RestoredAt]) continue;
    
    const email = String(row[archIdx.Email]).toLowerCase();
    const result = restoreFromArchiveIfEligible(email, slotId);
    
    if (result.restored) {
      restored.push({
        Name: row[archIdx.Name],
        Email: row[archIdx.Email],
        Date: row[archIdx.Date],
        Start: row[archIdx.Start],
        End: row[archIdx.End]
      });
    }
  }
  
  return restored;
}

function dropEntireSlot_(slotId) {
  const confirmed = getResponses_()
    .filter(r => r.SlotID === slotId && r.Status === 'confirmed');
  
  confirmed.forEach(r => {
    moveToArchive_(r, 'slot-canceled');
    deleteResponseRow_(r);
    
    // キャンセル通知
    const ds=normDateStr_(r.Date), st=normTimeStr_(r.Start), en=normTimeStr_(r.End);
    const when=fmtJPDateTime_(ds,st)+' - '+en;
    const subject=renderTemplate_(TEMPLATES.participant.slotCanceledSubject,{when:when});
    const body=renderTemplate_(TEMPLATES.participant.slotCanceledBody,{
      name:r.Name, when:when, tz:CONFIG.tz, location:CONFIG.location, fromName:CONFIG.mailFromName
    });
    sendMailSmart_({type:'admin', to:r.Email, subject:subject, body:body});
  });
  
  // Confirmedシートから削除
  deleteConfirmedRow_(slotId);
  updateSlotAggregate_(slotId, 0, false);
}

/** ========= AddSlots処理 ========= */
function applyAddSlots(){
  var sh=ensureAddSlotsSheet_(), data=sh.getDataRange().getValues(); 
  if (data.length<2) return;
  var head=data.shift(), idx=colIndex_(head);
  var put=function(row, st, msg){ 
    row[idx.Status]=st; 
    row[idx.Result]=msg; 
  };

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
      var parseTW=function(txt){ 
        return String(txt||'').split(',').map(function(x){return x.trim();}).filter(Boolean)
          .map(function(w){var p=w.split('-'); return [normTime(p[0]), normTime(p[1])];});
      };
      var twOrDefault=function(cell){
        var s=String(cell||'').trim().toUpperCase();
        if (!s || s==='DEFAULT') return (CONFIG.timeWindows||[])
          .map(function(w){var p=w.split('-'); return [normTime(p[0]), normTime(p[1])];});
        return parseTW(s);
      };

      if (mode==='datetime'){
        var ds=normDate(row[idx.Date]), st=row[idx.Start], en=row[idx.End];
        if (!st || !en){ 
          var list=parseTW(row[idx.TimeWindows]); 
          if (list.length!==1) throw new Error('TimeWindows は1つだけ'); 
          st=list[0][0]; en=list[0][1]; 
        }
        else { st=normTime(st); en=normTime(en); }
        addOne(ds, st, en);
      } else if (mode==='date'){
        var ds2=normDate(row[idx.Date]); 
        if (!ds2) throw new Error('Date が必要');
        twOrDefault(row[idx.TimeWindows]).forEach(function(p){ addOne(ds2, p[0], p[1]); });
      } else {
        var from=normDate(row[idx.FromDate]), to=normDate(row[idx.ToDate]); 
        if(!from||!to) throw new Error('From/To が必要');
        var tws=twOrDefault(row[idx.TimeWindows]);
        for (var d=new Date(from+'T00:00:00'); d<=new Date(to+'T00:00:00'); d=new Date(d.getTime()+86400000)){
          if (exWk && (d.getDay()===0 || d.getDay()===6)) continue;
          var y=d.getFullYear(), m=('0'+(d.getMonth()+1)).slice(-2), dd=('0'+d.getDate()).slice(-2);
          var ds3=y+'-'+m+'-'+dd; 
          tws.forEach(function(p){ addOne(ds3, p[0], p[1]); });
        }
      }
      put(row,'done','added='+added+', skipped='+skipped);
    }catch(e){ 
      put(row,'error', String(e)); 
    }
    sh.getRange(i+2,1,1,row.length).setValues([row]);
  }
  SpreadsheetApp.getActive().toast('AddSlots 完了', '追加枠', 5);
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

  // インストール型 onOpen
  if (typeof SS_ID === 'string' && SS_ID) {
    ScriptApp.newTrigger('onOpenUi_')
      .forSpreadsheet(SS_ID)
      .onOpen()
      .create();
  }
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

function onOpen() {
  addSchedulerMenu_();
}

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

/** ========= 日次ダイジェスト ========= */
function sendDailyAdminDigest() {
  if (!CONFIG.adminEmails || !CONFIG.adminEmails.length) return;
  
  const today = new Date();
  const todayStr = normDateStr_(today);
  const confirmed = getResponses_()
    .filter(r => r.Status === 'confirmed' && r.Date >= todayStr)
    .sort((a,b) => (a.Date+a.Start).localeCompare(b.Date+b.Start));
  
  if (confirmed.length === 0) return;
  
  const bySlot = groupBy_(confirmed, r => `${r.Date}_${r.Start}`);
  let body = renderTemplate_(TEMPLATES.admin.dailyDigestBodyIntro, {date: todayStr});
  
  Object.keys(bySlot).sort().forEach(key => {
    const group = bySlot[key];
    const first = group[0];
    const when = fmtJPDateTime_(first.Date, first.Start) + ' - ' + first.End;
    body += '\n■ ' + when + '\n';
    group.forEach(p => {
      body += '  ・' + p.Name + ' <' + p.Email + '>\n';
    });
  });
  
  const subject = renderTemplate_(TEMPLATES.admin.dailyDigestSubject, {date: todayStr});
  CONFIG.adminEmails.forEach(addr => {
    sendMailSmart_({type:'admin', to:addr, subject:subject, body:body});
  });
}

/** ========= 共通ヘルパ ========= */
function getResponses_(){ 
  var sh=getSS_().getSheetByName(SHEETS.RESP), vals=sh.getDataRange().getValues(), head=vals.shift(); 
  return vals.map(function(r){ return asObj_(head,r); }); 
}

function readSheetAsObjects_(sh){ 
  var vals=sh.getDataRange().getValues(), head=vals.shift(); 
  return vals.map(function(r){ return asObj_(head,r); }); 
}

function asObj_(head,row){ 
  var o={}; 
  head.forEach(function(h,i){ o[h]=row[i]; }); 
  return o; 
}

function groupBy_(arr,keyFn){ 
  return arr.reduce(function(m,x){ 
    var k=keyFn(x); 
    (m[k]||(m[k]=[])).push(x); 
    return m; 
  },{}); 
}

function colIndex_(head){ 
  var o={}; 
  head.forEach(function(h,i){ o[h]=i; }); 
  return o; 
}

function markNotified_(rowIndex, colName, val) {
  var sh=getSS_().getSheetByName(SHEETS.RESP), head=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var idx=head.indexOf(colName)+1; 
  sh.getRange(rowIndex, idx).setValue(val);
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

function deleteResponseRow_(rec){
  var sh=getSS_().getSheetByName(SHEETS.RESP), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  for (var i=0;i<vals.length;i++){
    var r=asObj_(head, vals[i]);
    if (r.Timestamp==rec.Timestamp && r.Email==rec.Email && r.SlotID==rec.SlotID){ 
      sh.deleteRow(i+2); 
      return true; 
    }
  } 
  return false;
}

function deleteConfirmedRow_(slotId){
  var sh=ensureConfirmedSheet_(), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  for (var i=0;i<vals.length;i++){ 
    if (vals[i][idx.SlotID]===slotId){ 
      sh.deleteRow(i+2); 
      return true; 
    } 
  }
  return false;
}

function setResponseStatus_(rec, status){
  var sh=getSS_().getSheetByName(SHEETS.RESP), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  for (var i=0;i<vals.length;i++){
    var r=asObj_(head, vals[i]);
    if (r.Timestamp==rec.Timestamp && r.Email==rec.Email && r.SlotID==rec.SlotID){
      vals[i][idx.Status]=status; 
      sh.getRange(i+2,1,1,vals[i].length).setValues([vals[i]]); 
      return true;
    }
  } 
  return false;
}

function hasConfirmedElsewhere_(email, excludeSlotId) {
  const confirmed = getResponses_().filter(r => 
    String(r.Email).toLowerCase() === email.toLowerCase() && 
    r.Status === 'confirmed' && 
    r.SlotID !== excludeSlotId
  );
  return confirmed.length > 0;
