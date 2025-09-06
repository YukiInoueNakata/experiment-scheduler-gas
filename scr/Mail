/** ========= メール送信関連 ========= */
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

/** ========= メールキュー処理 ========= */
function ensureMailQuota_(){ 
  return MailApp.getRemainingDailyQuota(); 
}

function sendMailSmart_(opt){
  const testDomains = ['@example.com'];
  const isTestEmail = testDomains.some(domain => opt.to.includes(domain));
  
  // テストモード または テストメールアドレスの場合
  if ((typeof isTestMode === 'function' && isTestMode()) || isTestEmail) {
    console.log('テストモード/テストメール：送信スキップ', {to: opt.to, subject: opt.subject});
    
    // TestMailLogに記録
    let logSheet = getSS_().getSheetByName('TestMailLog');
    if (!logSheet) {
      logSheet = getSS_().insertSheet('TestMailLog');
      logSheet.appendRow(['Timestamp', 'Type', 'To', 'Subject', 'Queue']);
    }
    
    // confirmタイプはMailQueueに入れる（動作確認用）
    if (opt.type === 'confirm') {
      var sh = ensureMailQueueSheet_();
      sh.appendRow([
        new Date(), 
        opt.type, 
        opt.to, 
        opt.subject, 
        opt.body, 
        opt.icsText||'', 
        JSON.stringify(opt.meta||{}), 
        'test-mode',  // ← statusを'test-mode'に
        '', 
        ''
      ]);
      logSheet.appendRow([new Date(), opt.type, opt.to, opt.subject, 'queued']);
      return {sent: false, queued: true, testMode: true};
    } else {
      // その他は即座に完了扱い
      logSheet.appendRow([new Date(), opt.type, opt.to, opt.subject, 'skipped']);
      return {sent: true, testMode: true};
    }
  }
  
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

/** ========= メールキュー処理（テストモード対応版） ========= */
function processMailQueue_(){
  var sh = ensureMailQueueSheet_(); 
  var vals = sh.getDataRange().getValues(); 
  if (vals.length < 2) return;
  
  var head = vals[0]; 
  var idx = colIndex_(head);
  var rows = [];
  
  // 処理対象の行を収集（pending と test-mode）
  for (var i = 1; i < vals.length; i++){
    const status = String(vals[i][idx.Status]).toLowerCase();
    if (status === 'pending' || status === 'test-mode') {
      rows.push({
        row: i + 1, 
        arr: vals[i], 
        isTest: status === 'test-mode'
      });
    }
  }
  
  // 各メールを処理
  for (var k = 0; k < rows.length; k++){
    // テストモードのメールは送信せずに処理済みにする
    if (rows[k].isTest) {
      sh.getRange(rows[k].row, idx.Status + 1).setValue('test-sent');
      sh.getRange(rows[k].row, idx.LastTriedAt + 1).setValue(new Date());
      sh.getRange(rows[k].row, idx.Error + 1).setValue('');
      
      // TestMailLogに記録
      let logSheet = getSS_().getSheetByName('TestMailLog');
      if (logSheet) {
        logSheet.appendRow([
          new Date(), 
          rows[k].arr[idx.Type], 
          rows[k].arr[idx.To], 
          rows[k].arr[idx.Subject],
          'processed-from-queue'
        ]);
      }
      
      console.log(`テストメール処理: ${rows[k].arr[idx.To]} - ${rows[k].arr[idx.Subject]}`);
      continue;
    }
    
    // 通常のメール送信処理
    var remain = ensureMailQuota_();
    var type = rows[k].arr[idx.Type];
    var isReminder = (type === 'reminder');
    var reserve = (CONFIG.mail && CONFIG.mail.reserveForReminders) || 0;
    var canUse = Math.max(0, remain - reserve) + (isReminder ? reserve : 0);
    
    // クォータ不足の場合は処理を中断
    if (canUse <= 0) break;

    var to = rows[k].arr[idx.To];
    var sub = rows[k].arr[idx.Subject];
    var body = rows[k].arr[idx.Body];
    var ics = rows[k].arr[idx.ICSText];
    
    try {
      // テストドメインの場合は実際には送信しない
      const testDomains = ['@example.com'];
      const isTestEmail = testDomains.some(domain => to.includes(domain));
      
      if (isTestEmail) {
        // テストメールの場合は送信せずに成功扱い
        console.log(`テストメール検出（送信スキップ）: ${to}`);
        sh.getRange(rows[k].row, idx.Status + 1).setValue('test-sent');
        sh.getRange(rows[k].row, idx.LastTriedAt + 1).setValue(new Date());
        sh.getRange(rows[k].row, idx.Error + 1).setValue('Test email - skipped');
        
        // TestMailLogに記録
        let logSheet = getSS_().getSheetByName('TestMailLog');
        if (!logSheet) {
          logSheet = getSS_().insertSheet('TestMailLog');
          logSheet.appendRow(['Timestamp', 'Type', 'To', 'Subject', 'Status']);
        }
        logSheet.appendRow([new Date(), type, to, sub, 'queue-processed']);
        
      } else {
        // 実際のメール送信
        if (ics) {
          GmailApp.sendEmail(to, sub, body, {
            name: CONFIG.mailFromName, 
            attachments: [Utilities.newBlob(ics, 'text/calendar', 'invite.ics')]
          });
        } else {
          MailApp.sendEmail(to, sub, body, {name: CONFIG.mailFromName});
        }
        
        sh.getRange(rows[k].row, idx.Status + 1).setValue('sent');
        sh.getRange(rows[k].row, idx.LastTriedAt + 1).setValue(new Date());
        sh.getRange(rows[k].row, idx.Error + 1).setValue('');
        
        console.log(`メール送信成功: ${to} - ${sub}`);
      }
      
    } catch(e) {
      // エラー処理
      sh.getRange(rows[k].row, idx.Status + 1).setValue('error');
      sh.getRange(rows[k].row, idx.LastTriedAt + 1).setValue(new Date());
      sh.getRange(rows[k].row, idx.Error + 1).setValue(String(e));
      
      console.error(`メール送信エラー: ${to} - ${e.toString()}`);
    }
  }
  
  // 処理完了後のログ
  if (rows.length > 0) {
    console.log(`MailQueue処理完了: ${rows.length}件処理`);
  }
}
