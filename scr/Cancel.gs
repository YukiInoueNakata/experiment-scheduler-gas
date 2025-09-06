/** ========= キャンセル処理 ========= */
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
    moveToArchive_(o, 'cancel:'+reason);
    resp.deleteRow(irow); 
    removed++;
  }

  confirmedSlots.forEach(function(slotId){
    if (policy==='drop-slot'){
      dropEntireSlot_(slotId);
      dropped++;
    } else {
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
  
  const currentConfirmed = getResponses_()
    .filter(r => r.SlotID === slotId && r.Status === 'confirmed');
  const currentCount = currentConfirmed.length;
  
  const needed = capacity - currentCount;
  
  let candidates = getResponses_()
    .filter(r => r.SlotID === slotId && 
            (r.Status === 'waitlist' || r.Status === 'pending'))
    .filter(r => !hasConfirmedElsewhere_(r.Email, slotId))
    .sort((a,b) => new Date(a.Timestamp) - new Date(b.Timestamp));
  
  if (candidates.length < needed) {
    const additionalNeeded = needed - candidates.length;
    const restored = restoreFromArchiveForSlot_(slotId, additionalNeeded);
    candidates = candidates.concat(restored);
  }
  
  const afterFillCount = currentCount + Math.min(candidates.length, needed);
  
  if (afterFillCount < minCap) {
    switch(fillPolicy) {
      case 'keep-partial':
        break;
      case 'to-pending':
        currentConfirmed.forEach(r => setResponseStatus_(r, 'pending'));
        updateSlotAggregate_(slotId, 0, false);
        return {refilled: false};
      case 'cancel-all':
        dropEntireSlot_(slotId);
        return {refilled: false};
      default:
        if (currentCount > 0) {
          // 現在の確定者は維持
        } else {
          return {refilled: false};
        }
    }
  }
  
  const toPromote = candidates.slice(0, needed);
  toPromote.forEach(c => {
    setResponseStatus_(c, 'confirmed');
    sendConfirmMail_(c.Name, c.Email, c.Date, c.Start, c.End, slot.Location, slot.Timezone);
  });
  
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
    
    const ds=normDateStr_(r.Date), st=normTimeStr_(r.Start), en=normTimeStr_(r.End);
    const when=fmtJPDateTime_(ds,st)+' - '+en;
    const subject=renderTemplate_(TEMPLATES.participant.slotCanceledSubject,{when:when});
    const body=renderTemplate_(TEMPLATES.participant.slotCanceledBody,{
      name:r.Name, when:when, tz:CONFIG.tz, location:CONFIG.location, fromName:CONFIG.mailFromName
    });
    sendMailSmart_({type:'admin', to:r.Email, subject:subject, body:body});
  });
  
  deleteConfirmedRow_(slotId);
  updateSlotAggregate_(slotId, 0, false);
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
