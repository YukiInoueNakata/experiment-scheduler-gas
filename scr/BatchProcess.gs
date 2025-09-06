/** ========= バッチ処理 ========= */
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
    // 0. 過去日付のデータを除外（今日以前をすべて除外）
    archivePastDatePending_();
    
    // 1. 過剰登録のクリーンアップ
    cleanupOverflowedPending_();
    
    // 2. 空き枠への追加登録
    fillRemainingSlots_();
    
    // 3. pendingのスロットごとの処理
    processAllPendingSlots_();
    
    // 4. 確定後のデータ整理
    cleanupAfterConfirmation();
    
    // 5. メール送信
    processMailQueue_();
    
  } finally {
    lock.releaseLock();
  }
}

/** ========= 過去日付のpending/waitlistをArchive ========= */
function archivePastDatePending_() {
  // 明日の日付を取得（明日より前 = 今日以前をアーカイブ）
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const tomorrowStr = normDateStr_(tomorrow);
  
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const responses = respSh.getDataRange().getValues();
  
  if (responses.length <= 1) return;
  
  const head = responses[0];
  let archivedCount = 0;
  
  // 後ろから処理（インデックスのズレを防ぐ）
  for (let i = responses.length - 1; i > 0; i--) {
    const row = responses[i];
    const obj = asObj_(head, row);
    
    // 日付を確実に正規化して比較
    const objDateStr = normDateStr_(obj.Date);
    
    // 明日より前（今日以前）の日付のpending/waitlistをArchive
    // confirmedは除外（過去の確定データは別処理で管理）
    if (objDateStr < tomorrowStr && (obj.Status === 'pending' || obj.Status === 'waitlist')) {
      moveToArchive_(obj, 'past-date-pending');
      respSh.deleteRow(i + 1);
      archivedCount++;
      console.log(`過去日付をArchive: ${obj.Name} - ${objDateStr} ${obj.SlotID}`);
    }
  }
  
  if (archivedCount > 0) {
    console.log(`今日以前のpending/waitlist ${archivedCount}件をArchiveに移動`);
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
  // 明日以降のスロットのみ処理対象とする
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const tomorrowStr = normDateStr_(tomorrow);
  
  const pendingResponses = getResponses_()
    .filter(r => {
      const dateStr = normDateStr_(r.Date);
      return r.Status === 'pending' && dateStr >= tomorrowStr;
    });
    
  const bySlot = groupBy_(pendingResponses, r => r.SlotID);
  
  Object.keys(bySlot).forEach(slotId => {
    confirmIfCapacityReached_(slotId);
  });
}

/** ========= 確定処理 ========= */
function confirmIfCapacityReached_(slotId) {
  const ss = getSS_();
  const slotSh = ss.getSheetByName(SHEETS.SLOTS);
  const respSh = ss.getSheetByName(SHEETS.RESP);

  const slots = readSheetAsObjects_(slotSh);
  const slot = slots.find(s => s.SlotID === slotId);
  if (!slot) return { slotId, status: 'notfound' };
  
  // スロットの日付が明日以降かチェック
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const tomorrowStr = normDateStr_(tomorrow);
  const slotDateStr = normDateStr_(slot.Date);
  
  if (slotDateStr < tomorrowStr) {
    console.log(`スロット ${slotId} は過去日付のためスキップ: ${slotDateStr}`);
    return { slotId, status: 'past-date', filled: false, confirmedCount: 0 };
  }

  const cap = parseInt(slot.Capacity, 10);
  const minCap = CONFIG.minCapacityToConfirm;

  let all = getResponses_().filter(r => r.SlotID === slotId);
  all.sort((a, b) => new Date(a.Timestamp).getTime() - new Date(b.Timestamp).getTime());

  const allConfirmed = getResponses_().filter(r => r.Status === 'confirmed');
  const confirmedEmails = new Set(allConfirmed.map(r => String(r.Email).toLowerCase()));

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

  const confirmedNowCount = winners.length;

  updateSlotAggregate_(slotId, confirmedNowCount, confirmedNowCount >= cap);

  newlyConfirmed.forEach(nc => {
    sendConfirmMail_(nc.name, nc.email, nc.date, nc.start, nc.end, slot.Location, slot.Timezone);
    markNotified_(nc.rowIndex, 'NotifiedConfirm', true);
  });

  updateConfirmedSheet_(slotId);
  
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
  
  for (let i = 0; i < CONFIG.capacity; i++) {
    if (i < confirmed.length) {
      rowData.push(confirmed[i].Name, confirmed[i].Email);
    } else {
      rowData.push('', '');
    }
  }
  
  rowData.push(confirmed.length);
  
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
