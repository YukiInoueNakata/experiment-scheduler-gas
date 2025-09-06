/** ========= データ整理機能 ========= */
function cleanupAfterConfirmation() {
  archivePastConfirmedData();
  archiveConfirmedAlternatives_();
  archiveConfirmedSlotData();
}

function archivePastConfirmedData() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const yesterdayStr = normDateStr_(new Date(today.getTime() - 24 * 60 * 60 * 1000));
  
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const responses = respSh.getDataRange().getValues();
  
  if (responses.length <= 1) return;
  
  const head = responses[0];
  let archivedCount = 0;
  
  for (let i = responses.length - 1; i > 0; i--) {
    const row = responses[i];
    const obj = asObj_(head, row);
    
    if (obj.Status === 'confirmed' && obj.Date <= yesterdayStr) {
      moveToArchive_(obj, 'past-confirmed-cleanup');
      respSh.deleteRow(i + 1);
      archivedCount++;
    }
  }
  
  console.log(`昨日までの確定データ ${archivedCount}件をArchiveに移動`);
}

function archiveConfirmedSlotData() {
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const responses = respSh.getDataRange().getValues();
  
  if (responses.length <= 1) return;
  
  const head = responses[0];
  const confirmedSlots = new Set();
  
  for (let i = 1; i < responses.length; i++) {
    const obj = asObj_(head, responses[i]);
    if (obj.Status === 'confirmed') {
      confirmedSlots.add(obj.SlotID);
    }
  }
  
  if (confirmedSlots.size === 0) return;
  
  let archivedCount = 0;
  
  for (let i = responses.length - 1; i > 0; i--) {
    const row = responses[i];
    const obj = asObj_(head, row);
    
    if (confirmedSlots.has(obj.SlotID)) {
      moveToArchive_(obj, 'confirmed-slot-cleanup');
      respSh.deleteRow(i + 1);
      archivedCount++;
    }
  }
  
  console.log(`確定済み枠のデータ ${archivedCount}件をArchiveに移動`);
}

function archiveConfirmedAlternatives_() {
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const responses = respSh.getDataRange().getValues();
  
  if (responses.length <= 1) return;
  
  const head = responses[0];
  const confirmedUsers = new Map();
  
  for (let i = 1; i < responses.length; i++) {
    const obj = asObj_(head, responses[i]);
    if (obj.Status === 'confirmed') {
      const email = String(obj.Email).toLowerCase();
      confirmedUsers.set(email, obj.SlotID);
    }
  }
  
  if (confirmedUsers.size === 0) return;
  
  let archivedCount = 0;
  
  for (let i = responses.length - 1; i > 0; i--) {
    const row = responses[i];
    const obj = asObj_(head, row);
    const email = String(obj.Email).toLowerCase();
    
    if (confirmedUsers.has(email)) {
      const confirmedSlotId = confirmedUsers.get(email);
      
      if (obj.SlotID !== confirmedSlotId) {
        moveToArchive_(obj, 'auto-archived-confirmed-elsewhere');
        respSh.deleteRow(i + 1);
        archivedCount++;
      }
    }
  }
  
  console.log(`確定者の他申込み ${archivedCount}件をArchiveに移動`);
}

function moveToArchive_(record, reason) {
  const archSh = ensureArchiveSheet_();
  const archiveData = [
    new Date(),
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
    ''
  ];
  archSh.appendRow(archiveData);
}

function restoreFromArchiveIfEligible(email, slotId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
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
    
    const existing = getResponses_().filter(r => 
      String(r.Email).toLowerCase() === email.toLowerCase() && 
      r.SlotID === slotId
    );
    
    if (existing.length > 0) {
      return { restored: false, reason: 'already-registered' };
    }
    
    const respSh = getSS_().getSheetByName(SHEETS.RESP);
    const restoredData = [
      targetRecord[archIdx.Timestamp],
      targetRecord[archIdx.Name],
      targetRecord[archIdx.Email],
      targetRecord[archIdx.SlotID],
      targetRecord[archIdx.Date],
      targetRecord[archIdx.Start],
      targetRecord[archIdx.End],
      'waitlist',
      false,
      false,
      false,
      'restored-from-archive'
    ];
    
    respSh.appendRow(restoredData);
    archSh.getRange(targetRow, archIdx.RestoredAt + 1).setValue(new Date());
    
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

function manualCleanupConfirmed() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '確定データの整理',
    '以下を実行します：\n' +
    '1. 昨日までの確定データをArchive\n' +
    '2. 確定者の他の申込みをArchive\n' +
    '3. 確定済み枠の全データをArchive\n\n' +
    '実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    cleanupAfterConfirmation();
    ui.alert('完了', 'データ整理が完了しました。', ui.ButtonSet.OK);
  }
}

function dailyDataCleanup() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    const cutoffDate = normDateStr_(sevenDaysAgo);
    
    const respSh = getSS_().getSheetByName(SHEETS.RESP);
    const responses = respSh.getDataRange().getValues();
    
    if (responses.length <= 1) return;
    
    const head = responses[0];
    let archivedCount = 0;
    
    for (let i = responses.length - 1; i > 0; i--) {
      const row = responses[i];
      const obj = asObj_(head, row);
      
      if (obj.Date < cutoffDate) {
        moveToArchive_(obj, 'daily-cleanup-7days');
        respSh.deleteRow(i + 1);
        archivedCount++;
      }
    }
    
    console.log(`日次クリーンアップ: ${archivedCount}件をArchive`);
    
  } finally {
    lock.releaseLock();
  }
}
