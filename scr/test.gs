/** ========= テスト専用関数（ファイル分割対応版） ========= */

// ========= テストモード管理 =========
function enableTestMode() {
  PropertiesService.getScriptProperties().setProperty('TEST_MODE', 'true');
  // テスト用にバッチ処理を高速化
  PropertiesService.getScriptProperties().setProperty('TEST_BATCH_DELAY', '5');
  SpreadsheetApp.getActive().toast('テストモード有効化（バッチ処理5秒）', 'テスト', 3);
}

function disableTestMode() {
  PropertiesService.getScriptProperties().deleteProperty('TEST_MODE');
  PropertiesService.getScriptProperties().deleteProperty('TEST_BATCH_DELAY');
  SpreadsheetApp.getActive().toast('テストモード無効化', 'テスト', 3);
}

function isTestMode() {
  return PropertiesService.getScriptProperties().getProperty('TEST_MODE') === 'true';
}

function getTestBatchDelay() {
  const delay = PropertiesService.getScriptProperties().getProperty('TEST_BATCH_DELAY');
  return delay ? parseInt(delay) : CONFIG.batchProcessDelaySeconds;
}

// ========= 現実的な20名テスト =========
function realisticTest20() {
  clearAllTestData();
  enableTestMode();
  
  const testUsers = [
    {name: '山田太郎', email: 'test.yamada@example.com'},
    {name: '佐藤花子', email: 'test.sato@example.com'},
    {name: '鈴木一郎', email: 'test.suzuki@example.com'},
    {name: '田中美咲', email: 'test.tanaka@example.com'},
    {name: '高橋健太', email: 'test.takahashi@example.com'},
    {name: '渡辺由美', email: 'test.watanabe@example.com'},
    {name: '伊藤大輔', email: 'test.ito@example.com'},
    {name: '中村愛子', email: 'test.nakamura@example.com'},
    {name: '小林修平', email: 'test.kobayashi@example.com'},
    {name: '加藤真理', email: 'test.kato@example.com'},
    {name: '木村光', email: 'test.kimura@example.com'},
    {name: '斎藤翔', email: 'test.saito@example.com'},
    {name: '松本優子', email: 'test.matsumoto@example.com'},
    {name: '井上健', email: 'test.inoue@example.com'},
    {name: '山口恵', email: 'test.yamaguchi@example.com'},
    {name: '福田正', email: 'test.fukuda@example.com'},
    {name: '森田愛', email: 'test.morita@example.com'},
    {name: '石田剛', email: 'test.ishida@example.com'},
    {name: '橋本舞', email: 'test.hashimoto@example.com'},
    {name: '清水誠', email: 'test.shimizu@example.com'}
  ];
  
  const allSlots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => s.Status === 'open' || s.Status === 'filled');
  
  if (allSlots.length < 10) {
    SpreadsheetApp.getUi().alert('エラー', '利用可能な枠が10個未満です。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  let totalApplications = 0;
  
  testUsers.forEach((user, index) => {
    const delaySeconds = index * 30 + Math.floor(Math.random() * 30);
    const timestamp = new Date();
    timestamp.setSeconds(timestamp.getSeconds() + delaySeconds);
    
    const numSlots = 3 + Math.floor(Math.random() * 5);
    const shuffled = [...allSlots].sort(() => Math.random() - 0.5);
    const selectedSlots = shuffled.slice(0, numSlots);
    
    selectedSlots.forEach(slot => {
      respSh.appendRow([
        timestamp,
        user.name,
        user.email,
        slot.SlotID,
        slot.Date,
        slot.Start,
        slot.End,
        'pending',
        false, false, false,
        'realistic-test'
      ]);
      totalApplications++;
    });
  });
  
  SpreadsheetApp.getActive().toast(
    `現実的テスト開始：20名、${totalApplications}件の申込みを生成しました。`,
    'テスト開始',
    10
  );
  
  for (let i = 1; i <= 3; i++) {
    ScriptApp.newTrigger('processPendingBatchForTest')
      .timeBased()
      .after(i * 60 * 1000)
      .create();
  }
}

// ========= シンプルな即時テスト =========
function simpleTestImmediate() {
  clearAllTestData();
  enableTestMode();
  
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => s.Status === 'open')
    .slice(0, 5);
  
  const testUsers = [
    {name: 'テストA', email: 'test.a@example.com'},
    {name: 'テストB', email: 'test.b@example.com'},
    {name: 'テストC', email: 'test.c@example.com'},
    {name: 'テストD', email: 'test.d@example.com'},
    {name: 'テストE', email: 'test.e@example.com'}
  ];
  
  slots.forEach(slot => {
    testUsers.slice(0, 3).forEach((user, index) => {
      const timestamp = new Date();
      timestamp.setMilliseconds(timestamp.getMilliseconds() + index * 100);
      
      respSh.appendRow([
        timestamp,
        user.name,
        user.email,
        slot.SlotID,
        slot.Date,
        slot.Start,
        slot.End,
        'pending',
        false, false, false,
        'simple-test'
      ]);
    });
  });
  
  SpreadsheetApp.getActive().toast('5秒後にバッチ処理を実行します', 'テスト', 3);
  Utilities.sleep(5000);
  
  processPendingBatch_();
  showTestStatus();
}

// ========= バッチ処理を今すぐ実行 =========
function runBatchNow() {
  processPendingBatch_();
  SpreadsheetApp.getActive().toast('バッチ処理を実行しました', '処理完了', 3);
}

// ========= テスト用バッチ処理 =========
function processPendingBatchForTest() {
  enableTestMode();
  processPendingBatch_();
  showTestStatus();
}

// ========= データクリア（完全版） =========
function clearAllTestData() {
  const testDomains = ['@example.com'];
  const sheets = [SHEETS.RESP, SHEETS.ARCH];
  let deletedCount = 0;
  
  // Responses と Archive のクリア
  sheets.forEach(sheetName => {
    const sh = getSS_().getSheetByName(sheetName);
    if (!sh) return;
    
    const data = sh.getDataRange().getValues();
    const emailCol = sheetName === SHEETS.RESP ? 2 : 3;
    
    for (let i = data.length - 1; i > 0; i--) {
      const email = String(data[i][emailCol] || '').toLowerCase();
      if (testDomains.some(domain => email.includes(domain))) {
        sh.deleteRow(i + 1);
        deletedCount++;
      }
    }
  });
  
  // Confirmedシートのクリア
  const confSh = ensureConfirmedSheet_();
  const confData = confSh.getDataRange().getValues();
  if (confData.length > 1) {
    const headers = getConfirmedHeaders();
    for (let i = confData.length - 1; i > 0; i--) {
      let hasTestData = false;
      for (let j = 1; j <= CONFIG.capacity; j++) {
        const emailColIndex = headers.indexOf(`Subject${j}Email`);
        if (emailColIndex >= 0) {
          const email = String(confData[i][emailColIndex] || '').toLowerCase();
          if (testDomains.some(domain => email.includes(domain))) {
            hasTestData = true;
            break;
          }
        }
      }
      if (hasTestData) {
        confSh.deleteRow(i + 1);
        deletedCount++;
      }
    }
  }
  
  // MailQueueのクリア
  clearMailQueueTestData();
  
  // Slotsのステータスリセット
  updateAllSlotStatuses();
  
  // TestMailLogシート削除
  const logSheet = getSS_().getSheetByName('TestMailLog');
  if (logSheet) {
    getSS_().deleteSheet(logSheet);
  }
  
  // テストトリガー削除
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === 'processPendingBatchForTest') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  SpreadsheetApp.getActive().toast(
    `テストデータを削除しました（${deletedCount}件）`,
    'クリア完了',
    5
  );
}

// ========= MailQueueのテストデータクリア =========
function clearMailQueueTestData() {
  const mqSh = getSS_().getSheetByName(SHEETS.MQ);
  if (!mqSh) return;
  
  const testDomains = ['@example.com'];
  let deletedCount = 0;
  
  const mqData = mqSh.getDataRange().getValues();
  if (mqData.length > 1) {
    const toIndex = 2; // To列は3列目（インデックス2）
    
    for (let i = mqData.length - 1; i > 0; i--) {
      const email = String(mqData[i][toIndex] || '').toLowerCase();
      if (testDomains.some(domain => email.includes(domain))) {
        mqSh.deleteRow(i + 1);
        deletedCount++;
      }
    }
  }
  
  console.log(`MailQueueから${deletedCount}件削除`);
}

// ========= スロット状態の更新 =========
function updateAllSlotStatuses() {
  const slotSh = getSS_().getSheetByName(SHEETS.SLOTS);
  const slotData = slotSh.getDataRange().getValues();
  const slotHead = slotData.shift();
  const slotIdx = colIndex_(slotHead);
  
  const responses = getResponses_();
  
  slotData.forEach((row, i) => {
    const slotId = row[slotIdx.SlotID];
    const capacity = Number(row[slotIdx.Capacity]);
    
    const confirmedCount = responses.filter(r => 
      r.SlotID === slotId && r.Status === 'confirmed'
    ).length;
    
    const newStatus = confirmedCount >= capacity ? 'filled' : 'open';
    
    slotSh.getRange(i + 2, slotIdx.ConfirmedCount + 1).setValue(confirmedCount);
    slotSh.getRange(i + 2, slotIdx.Status + 1).setValue(newStatus);
  });
}

// ========= テスト状況確認（詳細版） =========
function showTestStatus() {
  const testDomains = ['@example.com'];
  const responses = getResponses_();
  const testResponses = responses.filter(r => 
    testDomains.some(domain => String(r.Email).toLowerCase().includes(domain))
  );
  
  const statusCount = {
    confirmed: 0,
    pending: 0,
    waitlist: 0
  };
  
  const userStatus = {};
  
  testResponses.forEach(r => {
    statusCount[r.Status]++;
    
    const email = r.Email;
    if (!userStatus[email]) {
      userStatus[email] = {
        name: r.Name,
        confirmed: 0,
        pending: 0,
        waitlist: 0,
        total: 0
      };
    }
    userStatus[email][r.Status]++;
    userStatus[email].total++;
  });
  
  const archSh = getSS_().getSheetByName(SHEETS.ARCH);
  let archivedCount = 0;
  if (archSh) {
    const archData = archSh.getDataRange().getValues();
    for (let i = 1; i < archData.length; i++) {
      const email = String(archData[i][3] || '').toLowerCase();
      if (testDomains.some(domain => email.includes(domain))) {
        archivedCount++;
      }
    }
  }
  
  let message = `【テストデータ状況】\n\n`;
  message += `■ 全体統計\n`;
  message += `- Confirmed: ${statusCount.confirmed}件\n`;
  message += `- Pending: ${statusCount.pending}件\n`;
  message += `- Waitlist: ${statusCount.waitlist}件\n`;
  message += `- Archived: ${archivedCount}件\n\n`;
  
  message += `■ ユーザー別状況（確定者のみ）\n`;
  Object.keys(userStatus).forEach(email => {
    const user = userStatus[email];
    if (user.confirmed > 0) {
      message += `${user.name}: 確定${user.confirmed}/申込${user.total}\n`;
    }
  });
  
  message += `\nテストモード: ${isTestMode() ? '有効' : '無効'}`;
  message += `\nバッチ処理遅延: ${getTestBatchDelay()}秒`;
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('テスト状況', message, ui.ButtonSet.OK);
  
  console.log(message);
}

// ========= デバッグ用関数 =========
function debugCheckSheets() {
  const sheets = {
    'Responses': getSS_().getSheetByName(SHEETS.RESP),
    'Confirmed': getSS_().getSheetByName(SHEETS.CONF),
    'Archive': getSS_().getSheetByName(SHEETS.ARCH),
    'MailQueue': getSS_().getSheetByName(SHEETS.MQ)
  };
  
  let message = '【シート状況】\n\n';
  
  Object.keys(sheets).forEach(name => {
    const sh = sheets[name];
    if (sh) {
      const rows = sh.getLastRow() - 1; // ヘッダーを除く
      message += `${name}: ${rows}件\n`;
    } else {
      message += `${name}: シートなし\n`;
    }
  });
  
  SpreadsheetApp.getUi().alert('シート状況', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========= 10アカウント×20枠の大量テスト =========
function generateTestData10Accounts() {
  clearAllTestData();
  enableTestMode();
  
  const testAccounts = [
    {name: '山田太郎', email: 'test.yamada@example.com'},
    {name: '佐藤花子', email: 'test.sato@example.com'},
    {name: '鈴木一郎', email: 'test.suzuki@example.com'},
    {name: '田中美咲', email: 'test.tanaka@example.com'},
    {name: '高橋健太', email: 'test.takahashi@example.com'},
    {name: '渡辺由美', email: 'test.watanabe@example.com'},
    {name: '伊藤大輔', email: 'test.ito@example.com'},
    {name: '中村愛子', email: 'test.nakamura@example.com'},
    {name: '小林修平', email: 'test.kobayashi@example.com'},
    {name: '加藤真理', email: 'test.kato@example.com'}
  ];
  
  // 最初の30枠を取得
  const allSlots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => s.Status === 'open' || s.Status === 'filled')
    .slice(0, 30);
  
  if (allSlots.length < 30) {
    SpreadsheetApp.getUi().alert(
      'エラー', 
      `利用可能な枠が30個未満です（現在${allSlots.length}枠）。\n枠を追加してから実行してください。`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  let totalApplications = 0;
  const baseTime = new Date();
  
  // 各アカウントが20枠に申込み
  testAccounts.forEach((account, accountIndex) => {
    // 30枠からランダムに20枠選択
    const shuffled = [...allSlots].sort(() => Math.random() - 0.5);
    const selectedSlots = shuffled.slice(0, 20);
    
    selectedSlots.forEach((slot, slotIndex) => {
      // タイムスタンプを少しずつずらす（同じアカウントの申込みは連続的に）
      const timestamp = new Date(baseTime.getTime() + accountIndex * 1000 + slotIndex * 50);
      
      respSh.appendRow([
        timestamp,
        account.name,
        account.email,
        slot.SlotID,
        slot.Date,
        slot.Start,
        slot.End,
        'pending',
        false, false, false,
        'test10accounts'
      ]);
      totalApplications++;
    });
  });
  
  const expectedConfirmed = Math.min(
    allSlots.length * CONFIG.capacity,  // 全枠の最大収容人数
    testAccounts.length                  // または全アカウント数（1人1枠制限の場合）
  );
  
  SpreadsheetApp.getActive().toast(
    `テストデータ生成完了\n` +
    `・10アカウント × 20枠 = ${totalApplications}件の申込み\n` +
    `・5秒後にバッチ処理を実行します\n` +
    `・予想確定数: 最大${expectedConfirmed}名`,
    'テスト開始',
    10
  );
  
  // 5秒後にバッチ処理実行
  Utilities.sleep(5000);
  processPendingBatch_();
  
  // 結果表示
  Utilities.sleep(2000);
  showDetailedTestResults();
}

// ========= 詳細なテスト結果表示 =========
function showDetailedTestResults() {
  const testDomains = ['@example.com'];
  const responses = getResponses_();
  const testResponses = responses.filter(r => 
    testDomains.some(domain => String(r.Email).toLowerCase().includes(domain))
  );
  
  // 全体統計
  const statusCount = {
    confirmed: 0,
    pending: 0,
    waitlist: 0
  };
  
  // ユーザー別統計
  const userStatus = {};
  
  // スロット別統計
  const slotStatus = {};
  
  testResponses.forEach(r => {
    // 全体カウント
    statusCount[r.Status]++;
    
    // ユーザー別
    const email = r.Email;
    if (!userStatus[email]) {
      userStatus[email] = {
        name: r.Name,
        confirmed: 0,
        pending: 0,
        waitlist: 0,
        total: 0,
        confirmedSlot: null
      };
    }
    userStatus[email][r.Status]++;
    userStatus[email].total++;
    if (r.Status === 'confirmed') {
      userStatus[email].confirmedSlot = r.SlotID;
    }
    
    // スロット別
    const slotId = r.SlotID;
    if (!slotStatus[slotId]) {
      slotStatus[slotId] = {
        confirmed: 0,
        pending: 0,
        waitlist: 0,
        total: 0
      };
    }
    slotStatus[slotId][r.Status]++;
    slotStatus[slotId].total++;
  });
  
  // Archive件数
  const archSh = getSS_().getSheetByName(SHEETS.ARCH);
  let archivedCount = 0;
  if (archSh) {
    const archData = archSh.getDataRange().getValues();
    for (let i = 1; i < archData.length; i++) {
      const email = String(archData[i][3] || '').toLowerCase();
      if (testDomains.some(domain => email.includes(domain))) {
        archivedCount++;
      }
    }
  }
  
  // 結果メッセージ作成
  let message = `【テスト結果詳細】\n\n`;
  
  message += `■ 全体統計\n`;
  message += `- 総申込数: ${statusCount.confirmed + statusCount.pending + statusCount.waitlist}件\n`;
  message += `- Confirmed: ${statusCount.confirmed}件\n`;
  message += `- Pending: ${statusCount.pending}件\n`;
  message += `- Waitlist: ${statusCount.waitlist}件\n`;
  message += `- Archived: ${archivedCount}件\n\n`;
  
  message += `■ 確定状況\n`;
  let confirmedCount = 0;
  let noConfirmedCount = 0;
  Object.keys(userStatus).forEach(email => {
    const user = userStatus[email];
    if (user.confirmed > 0) {
      confirmedCount++;
      message += `✓ ${user.name}: ${user.confirmedSlot}\n`;
    } else {
      noConfirmedCount++;
    }
  });
  message += `\n確定: ${confirmedCount}名 / 未確定: ${noConfirmedCount}名\n\n`;
  
  message += `■ スロット充足率（上位5枠）\n`;
  const sortedSlots = Object.entries(slotStatus)
    .sort((a, b) => b[1].confirmed - a[1].confirmed)
    .slice(0, 5);
  
  sortedSlots.forEach(([slotId, stats]) => {
    const fillRate = `${stats.confirmed}/${CONFIG.capacity}`;
    const status = stats.confirmed >= CONFIG.capacity ? '満席' : '空席あり';
    message += `${slotId}: ${fillRate} (${status}) - 申込${stats.total}件\n`;
  });
  
  message += `\n設定: capacity=${CONFIG.capacity}, minConfirm=${CONFIG.minCapacityToConfirm}`;
  message += `\nallowMultiple=${CONFIG.allowMultipleConfirmationPerEmail}`;
  
  // 結果表示
  const ui = SpreadsheetApp.getUi();
  ui.alert('テスト結果', message, ui.ButtonSet.OK);
  
  // ログにも出力
  console.log(message);
  
  // Confirmedシートの状況も確認
  logConfirmedSheet();
}

// ========= Confirmedシートのログ出力 =========
function logConfirmedSheet() {
  const confSh = ensureConfirmedSheet_();
  const data = confSh.getDataRange().getValues();
  
  if (data.length <= 1) {
    console.log('Confirmedシート: データなし');
    return;
  }
  
  console.log(`Confirmedシート: ${data.length - 1}枠確定`);
  
  const headers = data[0];
  const actualCountIdx = headers.indexOf('ActualCount');
  
  data.slice(1, 6).forEach(row => {  // 最初の5件のみ表示
    const slotId = row[0];
    const actualCount = row[actualCountIdx];
    console.log(`  ${slotId}: ${actualCount}名確定`);
  });
}

// ========= generateTestData10を置き換え =========
function generateTestData10() {
  generateTestData10Accounts();
}

// ========= メニュー更新 =========
function addTestMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🧪テスト機能')
    .addItem('✅ テストモード有効化', 'enableTestMode')
    .addItem('❌ テストモード無効化', 'disableTestMode')
    .addSeparator()
    .addItem('📊 10アカウント×20枠テスト', 'generateTestData10Accounts')
    .addItem('🚀 現実的な20名テスト', 'realisticTest20')
    .addItem('⚡ シンプル即時テスト', 'simpleTestImmediate')
    .addItem('▶️ バッチ処理を今すぐ実行', 'runBatchNow')
    .addSeparator()
    .addItem('📈 詳細な結果表示', 'showDetailedTestResults')
    .addItem('📊 テスト状況確認', 'showTestStatus')
    .addItem('📋 シート状況確認', 'debugCheckSheets')
    .addSeparator()
    .addItem('🗑️ 全テストデータ削除', 'clearAllTestData')
    .addItem('🔄 スロット状態の再計算', 'updateAllSlotStatuses')
    .addToUi();
}
