/** ========= 正規化 ========= */
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
}
