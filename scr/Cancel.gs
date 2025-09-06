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
