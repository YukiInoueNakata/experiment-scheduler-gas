/** ========= トリガー設定 ========= */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (['sendReminders','sendDailyAdminDigest','processMailQueue_','onOpenUi_','dailyDataCleanup'].includes(fn)) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('sendReminders')
    .timeBased().atHour(9).nearMinute(0).everyDays(1).create();

  ScriptApp.newTrigger('sendDailyAdminDigest')
    .timeBased().atHour(0).nearMinute(0).everyDays(1).create();

  const min = (CONFIG.mail && CONFIG.mail.hourlyQueueTriggerMinute) || 10;
  ScriptApp.newTrigger('processMailQueue_')
    .timeBased().everyHours(1).nearMinute(min).create();
    
  ScriptApp.newTrigger('dailyDataCleanup')
    .timeBased().atHour(2).nearMinute(0).everyDays(1).create();

  if (typeof SS_ID === 'string' && SS_ID) {
    ScriptApp.newTrigger('onOpenUi_')
      .forSpreadsheet(SS_ID)
      .onOpen()
      .create();
  }
}

/** ========= メニュー ========= */
function addSchedulerMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('スケジューラ')
    .addItem('操作パネルを開く', 'openControlPanel')
    .addSeparator()
    .addItem('setup（枠生成）', 'setup')
    .addItem('setupTriggers（トリガー作成）', 'setupTriggers')
    .addSeparator()
    .addSubMenu(ui.createMenu('データ整理')
      .addItem('確定データの整理', 'manualCleanupConfirmed')
      .addItem('7日以前のデータをArchive', 'dailyDataCleanup'))
    .addToUi();
}

function onOpen() {
  addSchedulerMenu_();
  
  try {
    if (typeof addTestMenu === 'function') {
      addTestMenu();
    }
  } catch(e) {
    // Test.gsが存在しない場合は無視
  }
}

function onOpenUi_() {
  addSchedulerMenu_();
  
  try {
    if (typeof addTestMenu === 'function') {
      addTestMenu();
    }
  } catch(e) {
    // Test.gsが存在しない場合は無視
  }
}

function openControlPanel(){
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui; padding:12px; width:280px;">' +
      '<h3 style="margin:0 0 12px;">スケジューラ 操作</h3>' +
      '<p style="color:#444">各シートに入力してから、該当ボタンを押してください。</p>' +
      '<button onclick="google.script.run.applyAddSlots();this.disabled=true;this.innerText=\'実行中…\';" style="padding:8px 12px;margin-bottom:8px;width:100%;">AddSlots を実行</button>' +
      '<button onclick="google.script.run.applyCancelOps();this.disabled=true;this.innerText=\'実行中…\';" style="padding:8px 12px;margin-bottom:8px;width:100%;">CancelOps を実行</button>' +
      '<hr>' +
      '<button onclick="google.script.run.manualCleanupConfirmed();this.disabled=true;this.innerText=\'実行中…\';" style="padding:8px 12px;margin-bottom:8px;width:100%;">確定データの整理</button>' +
      '<hr>' +
      '<button onclick="google.script.run.setup();google.script.run.setupTriggers();this.innerText=\'セットアップ完了\';" style="padding:8px 12px;width:100%;">初期セットアップ（枠生成＆トリガー）</button>' +
    '</div>'
  ).setTitle('スケジューラ操作');
  SpreadsheetApp.getUi().showSidebar(html);
}
