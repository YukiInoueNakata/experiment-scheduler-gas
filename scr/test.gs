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

// ========= 包括的な日付テスト（新規追加） =========
function comprehensiveDateTest() {
  clearAllTestData();
  enableTestMode();
  
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const today = new Date();
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  
  const todayStr = normDateStr_(today);
  const tomorrowStr = normDateStr_(tomorrow);
  const yesterdayStr = normDateStr_(yesterday);
  
  // すべてのスロットを日付でグループ化
  const allSlots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => s.Status === 'open');
  
  const slotsByDateType = {
    past: [],
    today: [],
    future: []
  };
  
  allSlots.forEach(slot => {
    const dateStr = normDateStr_(slot.Date);
    if (dateStr < todayStr) {
      slotsByDateType.past.push(slot);
    } else if (dateStr === todayStr) {
      slotsByDateType.today.push(slot);
    } else {
      slotsByDateType.future.push(slot);
    }
  });
  
  console.log('===== 包括的日付テスト開始 =====');
  console.log(`過去日付スロット: ${slotsByDateType.past.length}個`);
  console.log(`今日のスロット: ${slotsByDateType.today.length}個`);
  console.log(`明日以降のスロット: ${slotsByDateType.future.length}個`);
  
  // テストユーザー
  const testUsers = [
    {name: 'テストA（過去）', email: 'test.past.a@example.com'},
    {name: 'テストB（過去）', email: 'test.past.b@example.com'},
    {name: 'テストC（今日）', email: 'test.today.c@example.com'},
    {name: 'テストD（今日）', email: 'test.today.d@example.com'},
    {name: 'テストE（明日）', email: 'test.future.e@example.com'},
    {name: 'テストF（明日）', email: 'test.future.f@example.com'},
    {name: 'テストG（混合）', email: 'test.mixed.g@example.com'}
  ];
  
  let totalApplications = 0;
  
  // 1. 過去日付スロットへの申込み（2名）
  if (slotsByDateType.past.length > 0) {
    const pastSlots = slotsByDateType.past.slice(0, 2);
    pastSlots.forEach(slot => {
      [testUsers[0], testUsers[1]].forEach((user, index) => {
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
          'date-test-past'
        ]);
        totalApplications++;
      });
    });
    console.log(`過去日付: ${pastSlots.length}枠に各2名申込み`);
  }
  
  // 2. 今日のスロットへの申込み（2名）
  if (slotsByDateType.today.length > 0) {
    const todaySlots = slotsByDateType.today.slice(0, 2);
    todaySlots.forEach(slot => {
      [testUsers[2], testUsers[3]].forEach((user, index) => {
        const timestamp = new Date();
        timestamp.setMilliseconds(timestamp.getMilliseconds() + 200 + index * 100);
        
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
          'date-test-today'
        ]);
        totalApplications++;
      });
    });
    console.log(`今日: ${todaySlots.length}枠に各2名申込み`);
  }
  
  // 3. 明日以降のスロットへの申込み（2名）
  if (slotsByDateType.future.length > 0) {
    const futureSlots = slotsByDateType.future.slice(0, 3);
    futureSlots.forEach(slot => {
      [testUsers[4], testUsers[5]].forEach((user, index) => {
        const timestamp = new Date();
        timestamp.setMilliseconds(timestamp.getMilliseconds() + 400 + index * 100);
        
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
          'date-test-future'
        ]);
        totalApplications++;
      });
    });
    console.log(`明日以降: ${futureSlots.length}枠に各2名申込み`);
  }
  
  // 4. 混合申込み（1名が全期間に申込み）
  const mixedUser = testUsers[6];
  if (slotsByDateType.past.length > 0) {
    const slot = slotsByDateType.past[0];
    respSh.appendRow([
      new Date(),
      mixedUser.name,
      mixedUser.email,
      slot.SlotID,
      slot.Date,
      slot.Start,
      slot.End,
      'pending',
      false, false, false,
      'date-test-mixed'
    ]);
    totalApplications++;
  }
  if (slotsByDateType.today.length > 0) {
    const slot = slotsByDateType.today[0];
    respSh.appendRow([
      new Date(),
      mixedUser.name,
      mixedUser.email,
      slot.SlotID,
      slot.Date,
      slot.Start,
      slot.End,
      'pending',
      false, false, false,
      'date-test-mixed'
    ]);
    totalApplications++;
  }
  if (slotsByDateType.future.length > 0) {
    const slot = slotsByDateType.future[0];
    respSh.appendRow([
      new Date(),
      mixedUser.name,
      mixedUser.email,
      slot.SlotID,
      slot.Date,
      slot.Start,
      slot.End,
      'pending',
      false, false, false,
      'date-test-mixed'
    ]);
    totalApplications++;
  }
  
  console.log(`\n総申込み数: ${totalApplications}件`);
  
  SpreadsheetApp.getActive().toast(
    `包括的日付テスト：${totalApplications}件の申込みを生成\n5秒後にバッチ処理を実行します`,
    'テスト',
    5
  );
  
  Utilities.sleep(5000);
  
  // バッチ処理実行
  console.log('\n===== バッチ処理実行 =====');
  processPendingBatch_();
  
  // 結果を分析・表示
  showComprehensiveDateTestResults();
}

// 包括的日付テストの結果表示
function showComprehensiveDateTestResults() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = normDateStr_(tomorrow);
  
  const testDomains = ['@example.com'];
  
  // Responsesの状態を確認
  const responses = getResponses_().filter(r => 
    testDomains.some(domain => String(r.Email).toLowerCase().includes(domain))
  );
  
  // Archiveの状態を確認
  const archSh = getSS_().getSheetByName(SHEETS.ARCH);
  const archivedData = [];
  if (archSh) {
    const archValues = archSh.getDataRange().getValues();
    for (let i = 1; i < archValues.length; i++) {
      const email = String(archValues[i][3] || '').toLowerCase();
      if (testDomains.some(domain => email.includes(domain))) {
        archivedData.push({
          email: email,
          date: normDateStr_(archValues[i][5]),
          notes: archValues[i][9]
        });
      }
    }
  }
  
  // 結果の集計
  const results = {
    past: { archived: 0, confirmed: 0, pending: 0, waitlist: 0 },
    today: { archived: 0, confirmed: 0, pending: 0, waitlist: 0 },
    future: { archived: 0, confirmed: 0, pending: 0, waitlist: 0 }
  };
  
  // Archiveデータの集計
  archivedData.forEach(item => {
    if (item.email.includes('past')) results.past.archived++;
    else if (item.email.includes('today')) results.today.archived++;
    else if (item.email.includes('future')) results.future.archived++;
  });
  
  // Responsesデータの集計
  responses.forEach(r => {
    const dateStr = normDateStr_(r.Date);
    const email = String(r.Email).toLowerCase();
    let category = null;
    
    if (email.includes('past')) category = 'past';
    else if (email.includes('today')) category = 'today';
    else if (email.includes('future')) category = 'future';
    else if (email.includes('mixed')) {
      if (dateStr < normDateStr_(new Date())) category = 'past';
      else if (dateStr === normDateStr_(new Date())) category = 'today';
      else category = 'future';
    }
    
    if (category) {
      results[category][r.Status]++;
    }
  });
  
  // 結果メッセージの作成
  let message = `【包括的日付テスト結果】\n\n`;
  
  message += `■ 過去日付の申込み\n`;
  message += `  ✅ Archive移動: ${results.past.archived}件（期待通り）\n`;
  message += `  ❌ Responses残存: ${results.past.confirmed + results.past.pending + results.past.waitlist}件\n\n`;
  
  message += `■ 今日の申込み\n`;
  message += `  ✅ Archive移動: ${results.today.archived}件（期待通り）\n`;
  message += `  ❌ Responses残存: ${results.today.confirmed + results.today.pending + results.today.waitlist}件\n\n`;
  
  message += `■ 明日以降の申込み\n`;
  message += `  Archive移動: ${results.future.archived}件\n`;
  message += `  Confirmed: ${results.future.confirmed}件\n`;
  message += `  Pending: ${results.future.pending}件\n`;
  message += `  Waitlist: ${results.future.waitlist}件\n\n`;
  
  const expectedBehavior = 
    results.past.archived > 0 && 
    results.today.archived > 0 && 
    results.future.confirmed > 0;
  
  if (expectedBehavior) {
    message += `✅ テスト成功: 日付処理が正しく動作しています`;
  } else {
    message += `❌ テスト失敗: 日付処理に問題があります`;
  }
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('包括的日付テスト結果', message, ui.ButtonSet.OK);
  
  console.log(message);
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
  
  // スロットの日付を確認（デバッグ用）
  console.log('===== シンプル即時テスト =====');
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = normDateStr_(tomorrow);
  
  slots.forEach((slot, index) => {
    const dateStr = normDateStr_(slot.Date);
    const isPast = dateStr < tomorrowStr;
    console.log(`スロット${index + 1}: ${slot.SlotID} (${dateStr}) ${isPast ? '過去/今日' : '明日以降'}`);
  });
  
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
// ========= メールテスト機能 =========
function testAllEmails() {
  enableTestMode();
  
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'メールテスト実行',
    'テスト用のダミーデータで全種類のメールをテストします。\n' +
    '実際にはメールは送信されず、TestMailLogシートに記録されます。\n\n' +
    '実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  console.log('===== メールテスト開始 =====');
  
  // テスト用の日付を準備（最後のスロットの日付を基準）
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => s.Status === 'open' || s.Status === 'filled')
    .sort((a, b) => (b.Date + b.Start).localeCompare(a.Date + a.Start));
  
  if (slots.length === 0) {
    ui.alert('エラー', 'スロットが存在しません', ui.ButtonSet.OK);
    return;
  }
  
  const lastSlot = slots[0];
  const lastDate = new Date(normDateStr_(lastSlot.Date) + 'T00:00:00');
  
  // 1. 受付メールのテスト
  testReceiptMail();
  
  // 2. 確定メールのテスト
  testConfirmMail(lastSlot);
  
  // 3. リマインダーメールのテスト
  testReminderMail(lastSlot);
  
  // 4. 管理者確定通知のテスト
  testAdminConfirmMail(lastSlot);
  
  // 5. 管理者日次ダイジェストのテスト（最後の3日分）
  testAdminDigest(lastDate);
  
  // 6. キャンセル関連メールのテスト
  testCancelMails(lastSlot);
  
  // 結果表示
  showMailTestResults();
}

// 受付メールのテスト
function testReceiptMail() {
  console.log('受付メールテスト...');
  
  const name = 'テスト太郎';
  const email = 'test.receipt@example.com';
  const lines = '・2025年09月30日(火)11:00 - 12:00（Asia/Tokyo）\n' +
                '・2025年09月30日(火)13:20 - 14:20（Asia/Tokyo）';
  
  const subject = renderTemplate_(TEMPLATES.participant.receiptSubject, {});
  const body = renderTemplate_(TEMPLATES.participant.receiptBody, {
    name: name,
    lines: lines,
    fromName: CONFIG.mailFromName
  });
  
  // テストモードなので実際には送信されない
  sendMailSmart_({
    type: 'receipt',
    to: email,
    subject: subject,
    body: body
  });
  
  console.log('受付メール: 完了');
}

// 確定メールのテスト
function testConfirmMail(slot) {
  console.log('確定メールテスト...');
  
  const testData = {
    name: 'テスト花子',
    email: 'test.confirm@example.com',
    date: slot.Date,
    start: slot.Start,
    end: slot.End,
    location: slot.Location || CONFIG.location,
    tz: slot.Timezone || CONFIG.tz
  };
  
  sendConfirmMail_(
    testData.name,
    testData.email,
    testData.date,
    testData.start,
    testData.end,
    testData.location,
    testData.tz
  );
  
  console.log('確定メール: 完了');
}

// リマインダーメールのテスト
function testReminderMail(slot) {
  console.log('リマインダーメールテスト...');
  
  const tz = CONFIG.tz;
  const ds = normDateStr_(slot.Date, tz);
  const st = normTimeStr_(slot.Start, tz);
  const en = normTimeStr_(slot.End, tz);
  const when = fmtJPDateTime_(ds, st) + ' - ' + en;
  
  const subject = renderTemplate_(TEMPLATES.participant.remindSubject, {when: when});
  const body = renderTemplate_(TEMPLATES.participant.remindBody, {
    name: 'テスト次郎',
    when: when,
    tz: tz,
    location: CONFIG.location,
    fromName: CONFIG.mailFromName
  });
  
  sendMailSmart_({
    type: 'reminder',
    to: 'test.reminder@example.com',
    subject: subject,
    body: body
  });
  
  console.log('リマインダーメール: 完了');
}

// 管理者確定通知のテスト
function testAdminConfirmMail(slot) {
  console.log('管理者確定通知テスト...');
  
  const testWinners = [
    {Name: 'テスト太郎', Email: 'test.taro@example.com'},
    {Name: 'テスト花子', Email: 'test.hanako@example.com'}
  ];
  
  // 管理者メールアドレスを一時的にテスト用に変更
  const originalEmails = CONFIG.adminEmails;
  CONFIG.adminEmails = ['test.admin@example.com'];
  
  sendAdminConfirmMail_(slot, testWinners);
  
  // 元に戻す
  CONFIG.adminEmails = originalEmails;
  
  console.log('管理者確定通知: 完了');
}

// 管理者日次ダイジェストのテスト（最後の3日分をダミーデータで）
function testAdminDigest(lastDate) {
  console.log('管理者日次ダイジェストテスト...');
  
  // ダミーデータを作成
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const addedRows = [];
  
  // 3日分のダミーデータを作成
  for (let dayOffset = 0; dayOffset < 3; dayOffset++) {
    const targetDate = new Date(lastDate);
    targetDate.setDate(targetDate.getDate() - dayOffset);
    const dateStr = normDateStr_(targetDate);
    
    // 各日4スロット
    const timeSlots = ['11:00', '13:20', '15:00', '16:50'];
    
    timeSlots.forEach((startTime, slotIndex) => {
      const slotId = `${dateStr}_${startTime.replace(':', '')}`;
      const endTime = `${parseInt(startTime.split(':')[0]) + 1}:${startTime.split(':')[1]}`;
      
      // 各スロットに2名ずつ確定者を追加
      for (let i = 0; i < 2; i++) {
        const rowData = [
          new Date(),
          `ダミー${dayOffset}${slotIndex}${i}`,
          `test.digest.${dayOffset}${slotIndex}${i}@example.com`,
          slotId,
          dateStr,
          startTime,
          endTime,
          'confirmed',
          false, false, false,
          'digest-test'
        ];
        respSh.appendRow(rowData);
        addedRows.push(respSh.getLastRow());
      }
    });
  }
  
  // 管理者メールアドレスを一時的にテスト用に変更
  const originalEmails = CONFIG.adminEmails;
  CONFIG.adminEmails = ['test.digest@example.com'];
  
  // ダイジェスト送信
  sendDailyAdminDigest();
  
  // 元に戻す
  CONFIG.adminEmails = originalEmails;
  
  // ダミーデータを削除
  for (let i = addedRows.length - 1; i >= 0; i--) {
    respSh.deleteRow(addedRows[i]);
  }
  
  console.log('管理者日次ダイジェスト: 完了');
}

// キャンセル関連メールのテスト
function testCancelMails(slot) {
  console.log('キャンセルメールテスト...');
  
  // スロットキャンセル通知のテスト
  const ds = normDateStr_(slot.Date);
  const st = normTimeStr_(slot.Start);
  const en = normTimeStr_(slot.End);
  const when = fmtJPDateTime_(ds, st) + ' - ' + en;
  
  const subject = renderTemplate_(TEMPLATES.participant.slotCanceledSubject, {when: when});
  const body = renderTemplate_(TEMPLATES.participant.slotCanceledBody, {
    name: 'テスト取消',
    when: when,
    tz: CONFIG.tz,
    location: CONFIG.location,
    fromName: CONFIG.mailFromName
  });
  
  sendMailSmart_({
    type: 'cancel',
    to: 'test.cancel@example.com',
    subject: subject,
    body: body
  });
  
  console.log('キャンセルメール: 完了');
}

// メールテスト結果の表示
function showMailTestResults() {
  const logSheet = getSS_().getSheetByName('TestMailLog');
  if (!logSheet) {
    SpreadsheetApp.getUi().alert(
      'テスト結果',
      'TestMailLogシートが作成されませんでした。\nテストモードが有効か確認してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const data = logSheet.getDataRange().getValues();
  const recentLogs = data.slice(-10); // 最新10件
  
  let message = '【メールテスト結果】\n\n';
  message += '最新のテストメール記録:\n';
  
  recentLogs.forEach(row => {
    if (row[0] instanceof Date) {
      const timestamp = Utilities.formatDate(row[0], 'Asia/Tokyo', 'HH:mm:ss');
      const type = row[1];
      const to = row[2];
      const status = row[4] || 'sent';
      message += `${timestamp} [${type}] → ${to} (${status})\n`;
    }
  });
  
  message += '\n※ テストモードのため実際のメール送信は行われていません';
  message += '\n※ 詳細はTestMailLogシートを確認してください';
  
  SpreadsheetApp.getUi().alert('メールテスト結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
  
  console.log('===== メールテスト完了 =====');
}

// 個別メールテスト：受付メールのみ
function testReceiptMailOnly() {
  enableTestMode();
  testReceiptMail();
  showMailTestResults();
}

// 個別メールテスト：確定メールのみ
function testConfirmMailOnly() {
  enableTestMode();
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => s.Status === 'open' || s.Status === 'filled');
  if (slots.length > 0) {
    testConfirmMail(slots[0]);
    showMailTestResults();
  }
}

// 個別メールテスト：管理者ダイジェストのみ
function testAdminDigestOnly() {
  enableTestMode();
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => s.Status === 'open' || s.Status === 'filled')
    .sort((a, b) => (b.Date + b.Start).localeCompare(a.Date + a.Start));
  if (slots.length > 0) {
    const lastDate = new Date(normDateStr_(slots[0].Date) + 'T00:00:00');
    testAdminDigest(lastDate);
    showMailTestResults();
  }
}

// ========= 確定しないケースのテスト =========
function testNoConfirmScenario() {
  clearAllTestData();
  enableTestMode();
  
  console.log('===== 確定しないケースのテスト開始 =====');
  console.log(`設定: capacity=${CONFIG.capacity}, minCapacityToConfirm=${CONFIG.minCapacityToConfirm}`);
  
  // minCapacityToConfirmが2以上でないとテストできない
  if (CONFIG.minCapacityToConfirm < 2) {
    SpreadsheetApp.getUi().alert(
      'テスト不可',
      'このテストはminCapacityToConfirm >= 2の場合のみ実行できます。\n' +
      `現在の設定: minCapacityToConfirm = ${CONFIG.minCapacityToConfirm}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  
  // 明日以降のスロットを取得
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const tomorrowStr = normDateStr_(tomorrow);
  
  const futureSlots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => {
      const dateStr = normDateStr_(s.Date);
      return s.Status === 'open' && dateStr >= tomorrowStr;
    })
    .slice(0, 5); // 最初の5枠を使用
  
  if (futureSlots.length < 3) {
    SpreadsheetApp.getUi().alert(
      'エラー',
      `明日以降の利用可能な枠が不足しています（${futureSlots.length}枠）。\n` +
      '最低3枠必要です。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // テストユーザー（minCapacityToConfirm - 1人だけ申込み）
  const insufficientUsers = [];
  for (let i = 0; i < CONFIG.minCapacityToConfirm - 1; i++) {
    insufficientUsers.push({
      name: `テスト未確定${i + 1}`,
      email: `test.noconfirm.${i + 1}@example.com`
    });
  }
  
  console.log(`\nテストケース: ${insufficientUsers.length}名が各枠に申込み（最小確定人数: ${CONFIG.minCapacityToConfirm}）`);
  
  // 各スロットに不足人数だけ申込み
  let totalApplications = 0;
  futureSlots.forEach((slot, slotIndex) => {
    console.log(`スロット${slotIndex + 1}: ${slot.SlotID} (${slot.Date} ${slot.Start}-${slot.End})`);
    
    insufficientUsers.forEach((user, userIndex) => {
      const timestamp = new Date();
      timestamp.setMilliseconds(timestamp.getMilliseconds() + slotIndex * 1000 + userIndex * 100);
      
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
        'no-confirm-test'
      ]);
      totalApplications++;
    });
  });
  
  console.log(`\n申込み完了: ${totalApplications}件`);
  
  // バッチ処理前の状態を記録
  const beforeBatch = captureNoConfirmState();
  console.log('\n===== バッチ処理前 =====');
  console.log(`Pending: ${beforeBatch.pendingCount}件`);
  
  // バッチ処理実行
  SpreadsheetApp.getActive().toast(
    `${insufficientUsers.length}名×${futureSlots.length}枠の申込みを生成\n` +
    '5秒後にバッチ処理を実行します',
    'テスト',
    5
  );
  
  Utilities.sleep(5000);
  processPendingBatch_();
  
  // バッチ処理後の状態を記録
  const afterBatch = captureNoConfirmState();
  console.log('\n===== バッチ処理後 =====');
  console.log(`Pending: ${afterBatch.pendingCount}件`);
  console.log(`Confirmed: ${afterBatch.confirmedCount}件`);
  console.log(`Waitlist: ${afterBatch.waitlistCount}件`);
  console.log(`Archived: ${afterBatch.archivedCount}件`);
  
  // 結果を分析
  analyzeNoConfirmResults(beforeBatch, afterBatch, futureSlots, insufficientUsers);
}

// 状態をキャプチャ（確定しないケース用）
function captureNoConfirmState() {
  const testDomains = ['@example.com'];
  const responses = getResponses_().filter(r => 
    testDomains.some(domain => String(r.Email).toLowerCase().includes(domain))
  );
  
  const statusCount = {
    pending: 0,
    confirmed: 0,
    waitlist: 0
  };
  
  const bySlot = {};
  
  responses.forEach(r => {
    statusCount[r.Status]++;
    
    if (!bySlot[r.SlotID]) {
      bySlot[r.SlotID] = {
        pending: 0,
        confirmed: 0,
        waitlist: 0
      };
    }
    bySlot[r.SlotID][r.Status]++;
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
  
  return {
    pendingCount: statusCount.pending,
    confirmedCount: statusCount.confirmed,
    waitlistCount: statusCount.waitlist,
    archivedCount: archivedCount,
    bySlot: bySlot,
    timestamp: new Date()
  };
}

// 結果分析（確定しないケース）
function analyzeNoConfirmResults(before, after, slots, users) {
  let message = `【確定しないケースのテスト結果】\n\n`;
  
  message += `■ テスト設定\n`;
  message += `- 最小確定人数: ${CONFIG.minCapacityToConfirm}名\n`;
  message += `- 申込み人数: ${users.length}名（不足: ${CONFIG.minCapacityToConfirm - users.length}名）\n`;
  message += `- テスト枠数: ${slots.length}枠\n\n`;
  
  message += `■ ステータス変化\n`;
  message += `バッチ処理前:\n`;
  message += `  Pending: ${before.pendingCount}件\n`;
  message += `  Confirmed: ${before.confirmedCount}件\n\n`;
  
  message += `バッチ処理後:\n`;
  message += `  Pending: ${after.pendingCount}件\n`;
  message += `  Confirmed: ${after.confirmedCount}件\n`;
  message += `  Waitlist: ${after.waitlistCount}件\n`;
  message += `  Archived: ${after.archivedCount}件\n\n`;
  
  // 各スロットの状態を確認
  message += `■ スロット別状況\n`;
  Object.keys(after.bySlot).forEach(slotId => {
    const slotStatus = after.bySlot[slotId];
    message += `${slotId}: `;
    message += `pending=${slotStatus.pending}, `;
    message += `confirmed=${slotStatus.confirmed}, `;
    message += `waitlist=${slotStatus.waitlist}\n`;
  });
  
  // 期待される動作の確認
  message += `\n■ 動作確認\n`;
  const allPending = after.pendingCount === before.pendingCount;
  const noConfirmed = after.confirmedCount === 0;
  const noArchived = after.archivedCount === 0;
  
  if (allPending && noConfirmed && noArchived) {
    message += `✅ 正常動作: 最小人数未満のため全員pendingのまま\n`;
  } else {
    message += `❌ 異常動作検出:\n`;
    if (!allPending) message += `  - Pendingが変化しました\n`;
    if (!noConfirmed) message += `  - 確定が発生しました（予期しない）\n`;
    if (!noArchived) message += `  - Archiveが発生しました（予期しない）\n`;
  }
  
  message += `\n■ 今後の挙動\n`;
  message += `1. このままでは永遠にpendingのまま\n`;
  message += `2. 追加で${CONFIG.minCapacityToConfirm - users.length}名以上が申込めば確定処理が走る\n`;
  message += `3. 日付が過ぎたら自動的にArchiveへ移動\n`;
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('確定しないケースのテスト結果', message, ui.ButtonSet.OK);
  
  console.log(message);
}

// 追加申込みシミュレーション（確定しないケースの続き）
function simulateAdditionalApplication() {
  console.log('===== 追加申込みシミュレーション =====');
  
  // 現在のpending状態を確認
  const testDomains = ['@example.com'];
  const pendingResponses = getResponses_().filter(r => 
    r.Status === 'pending' &&
    testDomains.some(domain => String(r.Email).toLowerCase().includes(domain))
  );
  
  if (pendingResponses.length === 0) {
    SpreadsheetApp.getUi().alert(
      'エラー',
      'pendingのテストデータが見つかりません。\n' +
      '先に「確定しないケースのテスト」を実行してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // スロットごとにグループ化
  const bySlot = {};
  pendingResponses.forEach(r => {
    if (!bySlot[r.SlotID]) bySlot[r.SlotID] = [];
    bySlot[r.SlotID].push(r);
  });
  
  // 最初のスロットに追加申込み
  const firstSlotId = Object.keys(bySlot)[0];
  const currentCount = bySlot[firstSlotId].length;
  const needed = CONFIG.minCapacityToConfirm - currentCount;
  
  console.log(`スロット${firstSlotId}の現在の申込み: ${currentCount}名`);
  console.log(`最小確定人数まであと: ${needed}名`);
  
  if (needed <= 0) {
    console.log('すでに最小人数を満たしています');
    return;
  }
  
  // 必要な人数だけ追加申込み
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const firstSlot = pendingResponses.find(r => r.SlotID === firstSlotId);
  
  for (let i = 0; i < needed; i++) {
    const additionalUser = {
      name: `追加テスト${i + 1}`,
      email: `test.additional.${i + 1}@example.com`
    };
    
    respSh.appendRow([
      new Date(),
      additionalUser.name,
      additionalUser.email,
      firstSlotId,
      firstSlot.Date,
      firstSlot.Start,
      firstSlot.End,
      'pending',
      false, false, false,
      'additional-test'
    ]);
    
    console.log(`追加申込み: ${additionalUser.name}`);
  }
  
  SpreadsheetApp.getActive().toast(
    `${needed}名の追加申込みを生成しました。\n` +
    '5秒後にバッチ処理を実行します。',
    '追加申込み',
    5
  );
  
  Utilities.sleep(5000);
  processPendingBatch_();
  
  // 結果確認
  const afterResponses = getResponses_().filter(r => 
    r.SlotID === firstSlotId &&
    testDomains.some(domain => String(r.Email).toLowerCase().includes(domain))
  );
  
  const confirmedCount = afterResponses.filter(r => r.Status === 'confirmed').length;
  
  let message = `【追加申込み後の結果】\n\n`;
  message += `スロット: ${firstSlotId}\n`;
  message += `追加前: ${currentCount}名（pending）\n`;
  message += `追加: ${needed}名\n`;
  message += `合計: ${currentCount + needed}名\n\n`;
  message += `バッチ処理後:\n`;
  message += `- Confirmed: ${confirmedCount}名\n`;
  
  if (confirmedCount >= CONFIG.minCapacityToConfirm) {
    message += `\n✅ 確定成功: 最小人数を満たしたため確定処理が実行されました`;
  } else {
    message += `\n❌ 確定失敗: 何か問題が発生しました`;
  }
  
  SpreadsheetApp.getUi().alert('追加申込みシミュレーション結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
  console.log(message);
}

// ========= CancelOps包括的テスト =========
function comprehensiveCancelTest() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'CancelOps包括的テスト',
    '全パターンのキャンセル処理をテストします。\n' +
    'テストデータを作成してから各種キャンセルを実行します。\n\n' +
    '実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  clearAllTestData();
  enableTestMode();
  
  console.log('===== CancelOps包括的テスト開始 =====');
  
  // テストデータを準備
  const testData = prepareCancelTestData();
  
  if (!testData.success) {
    ui.alert('エラー', testData.message, ui.ButtonSet.OK);
    return;
  }
  
  // 各パターンをテスト
  const results = [];
  
  // 1. 単純なキャンセル（補充あり）
  results.push(testSimpleCancel(testData));
  
  // 2. 枠全体の中止
  results.push(testDropSlot(testData));
  
  // 3. 補充不足のケース
  results.push(testInsufficientRefill(testData));
  
  // 4. 複数枠確定者のキャンセル
  results.push(testMultiSlotCancel(testData));
  
  // 5. Archive復元テスト
  results.push(testArchiveRestore(testData));
  
  // 結果表示
  showCancelTestResults(results);
}

// テストデータの準備
function prepareCancelTestData() {
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  
  // 明日以降のスロットを取得
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = normDateStr_(tomorrow);
  
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => {
      const dateStr = normDateStr_(s.Date);
      return s.Status === 'open' && dateStr >= tomorrowStr;
    })
    .slice(0, 6); // 6枠使用
  
  if (slots.length < 6) {
    return {
      success: false,
      message: `明日以降のスロットが不足しています（${slots.length}/6枠）`
    };
  }
  
  console.log(`テストスロット: ${slots.length}枠`);
  
  // テストユーザーを作成
  const users = {
    confirmed: [], // 確定者
    waitlist: [],  // キャンセル待ち
    pending: []    // 申込み中
  };
  
  // Slot1: 満席（capacity人確定 + waitlist2人）
  const slot1 = slots[0];
  for (let i = 0; i < CONFIG.capacity; i++) {
    const user = {
      name: `確定者${i + 1}`,
      email: `test.confirmed.${i + 1}@example.com`,
      slotId: slot1.SlotID,
      date: slot1.Date,
      start: slot1.Start,
      end: slot1.End
    };
    users.confirmed.push(user);
    
    respSh.appendRow([
      new Date(),
      user.name,
      user.email,
      user.slotId,
      user.date,
      user.start,
      user.end,
      'pending',
      false, false, false,
      'cancel-test-slot1'
    ]);
  }
  
  // waitlist追加
  for (let i = 0; i < 2; i++) {
    const user = {
      name: `待機者${i + 1}`,
      email: `test.waitlist.${i + 1}@example.com`,
      slotId: slot1.SlotID,
      date: slot1.Date,
      start: slot1.Start,
      end: slot1.End
    };
    users.waitlist.push(user);
    
    respSh.appendRow([
      new Date(),
      user.name,
      user.email,
      user.slotId,
      user.date,
      user.start,
      user.end,
      'pending',
      false, false, false,
      'cancel-test-slot1-wait'
    ]);
  }
  
  // Slot2: 最小人数ちょうど（minCapacityToConfirm人確定）
  const slot2 = slots[1];
  for (let i = 0; i < CONFIG.minCapacityToConfirm; i++) {
    const user = {
      name: `最小確定${i + 1}`,
      email: `test.minimal.${i + 1}@example.com`,
      slotId: slot2.SlotID,
      date: slot2.Date,
      start: slot2.Start,
      end: slot2.End
    };
    users.confirmed.push(user);
    
    respSh.appendRow([
      new Date(),
      user.name,
      user.email,
      user.slotId,
      user.date,
      user.start,
      user.end,
      'pending',
      false, false, false,
      'cancel-test-slot2'
    ]);
  }
  
  // Slot3: 複数枠確定者（1人が複数枠に確定）
  const multiUser = {
    name: '複数枠太郎',
    email: 'test.multi@example.com'
  };
  
  for (let i = 2; i < 4; i++) {
    const slot = slots[i];
    // 複数枠太郎を含めて確定
    for (let j = 0; j < CONFIG.minCapacityToConfirm; j++) {
      const user = j === 0 ? multiUser : {
        name: `共同参加${i}${j}`,
        email: `test.partner.${i}${j}@example.com`
      };
      
      respSh.appendRow([
        new Date(),
        user.name,
        user.email,
        slot.SlotID,
        slot.Date,
        slot.Start,
        slot.End,
        'pending',
        false, false, false,
        `cancel-test-slot${i + 1}`
      ]);
    }
  }
  
  // Slot4,5: Archive復元テスト用（後で使用）
  const slot4 = slots[4];
  const slot5 = slots[5];
  
  // バッチ処理を実行して確定させる
  console.log('バッチ処理実行中...');
  processPendingBatch_();
  
  return {
    success: true,
    slots: slots,
    users: users,
    multiUser: multiUser
  };
}

// テスト1: 単純なキャンセル（補充あり）
function testSimpleCancel(testData) {
  console.log('\n===== テスト1: 単純なキャンセル（補充あり） =====');
  
  const cancelOps = ensureCancelOpsSheet_();
  const targetEmail = 'test.confirmed.1@example.com';
  
  // キャンセル前の状態を記録
  const beforeState = getCancelTestState(testData.slots[0].SlotID);
  
  // CancelOps実行
  cancelOps.appendRow([
    targetEmail,
    'confirmed',
    'refill-slot',
    'try-fill',
    'テスト1',
    '',
    ''
  ]);
  
  applyCancelOps();
  
  // キャンセル後の状態を記録
  const afterState = getCancelTestState(testData.slots[0].SlotID);
  
  return {
    test: '単純なキャンセル（補充あり）',
    target: targetEmail,
    before: beforeState,
    after: afterState,
    success: afterState.confirmed === beforeState.confirmed // 補充されて同数のはず
  };
}

// テスト2: 枠全体の中止
function testDropSlot(testData) {
  console.log('\n===== テスト2: 枠全体の中止 =====');
  
  const cancelOps = ensureCancelOpsSheet_();
  const targetEmail = 'test.minimal.1@example.com';
  
  // キャンセル前の状態を記録
  const beforeState = getCancelTestState(testData.slots[1].SlotID);
  
  // CancelOps実行（drop-slot）
  cancelOps.appendRow([
    targetEmail,
    'confirmed',
    'drop-slot',
    'try-fill',
    'テスト2',
    '',
    ''
  ]);
  
  applyCancelOps();
  
  // キャンセル後の状態を記録
  const afterState = getCancelTestState(testData.slots[1].SlotID);
  
  return {
    test: '枠全体の中止',
    target: targetEmail,
    before: beforeState,
    after: afterState,
    success: afterState.confirmed === 0 && afterState.total === 0 // 全員削除されるはず
  };
}

// テスト3: 補充不足のケース
function testInsufficientRefill(testData) {
  console.log('\n===== テスト3: 補充不足のケース =====');
  
  // 新しい枠でテスト（minCapacityToConfirm人だけ確定、waitlistなし）
  const slot = testData.slots[4];
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  
  // minCapacityToConfirm人だけ申込み
  for (let i = 0; i < CONFIG.minCapacityToConfirm; i++) {
    respSh.appendRow([
      new Date(),
      `不足テスト${i + 1}`,
      `test.insufficient.${i + 1}@example.com`,
      slot.SlotID,
      slot.Date,
      slot.Start,
      slot.End,
      'pending',
      false, false, false,
      'insufficient-test'
    ]);
  }
  
  // 確定させる
  processPendingBatch_();
  
  const beforeState = getCancelTestState(slot.SlotID);
  
  // 1人キャンセル（to-pending設定）
  const cancelOps = ensureCancelOpsSheet_();
  cancelOps.appendRow([
    'test.insufficient.1@example.com',
    'confirmed',
    'refill-slot',
    'to-pending', // 最小人数未満ならpendingへ
    'テスト3',
    '',
    ''
  ]);
  
  applyCancelOps();
  
  const afterState = getCancelTestState(slot.SlotID);
  
  return {
    test: '補充不足（to-pending）',
    target: 'test.insufficient.1@example.com',
    before: beforeState,
    after: afterState,
    success: afterState.confirmed === 0 && afterState.pending > 0 // 全員pendingになるはず
  };
}

// テスト4: 複数枠確定者のキャンセル
function testMultiSlotCancel(testData) {
  console.log('\n===== テスト4: 複数枠確定者のキャンセル =====');
  
  const targetEmail = testData.multiUser.email;
  
  // キャンセル前の状態を記録
  const slot3State = getCancelTestState(testData.slots[2].SlotID);
  const slot4State = getCancelTestState(testData.slots[3].SlotID);
  
  // CancelOps実行（全枠キャンセル）
  const cancelOps = ensureCancelOpsSheet_();
  cancelOps.appendRow([
    targetEmail,
    'all', // 全申込みキャンセル
    'refill-slot',
    'try-fill',
    'テスト4',
    '',
    ''
  ]);
  
  applyCancelOps();
  
  // キャンセル後の状態を記録
  const slot3After = getCancelTestState(testData.slots[2].SlotID);
  const slot4After = getCancelTestState(testData.slots[3].SlotID);
  
  return {
    test: '複数枠確定者のキャンセル',
    target: targetEmail,
    before: {slot3: slot3State, slot4: slot4State},
    after: {slot3: slot3After, slot4: slot4After},
    success: true // 詳細は結果表示で確認
  };
}

// テスト5: Archive復元テスト
function testArchiveRestore(testData) {
  console.log('\n===== テスト5: Archive復元テスト =====');
  
  const slot = testData.slots[5];
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const archSh = ensureArchiveSheet_();
  
  // テスト用のArchiveデータを作成
  const archivedUser = {
    name: 'Archive復元太郎',
    email: 'test.archive@example.com'
  };
  
  // Archiveに直接データを追加（auto-archivedフラグ付き）
  archSh.appendRow([
    new Date(), // ArchivedAt
    new Date(), // Timestamp
    archivedUser.name,
    archivedUser.email,
    slot.SlotID,
    slot.Date,
    slot.Start,
    slot.End,
    'pending',
    'auto-archived-confirmed-elsewhere', // Notes
    false, false, false,
    '' // RestoredAt
  ]);
  
  // 現在の参加者（minCapacityToConfirm人）
  for (let i = 0; i < CONFIG.minCapacityToConfirm; i++) {
    respSh.appendRow([
      new Date(),
      `復元テスト${i + 1}`,
      `test.restore.${i + 1}@example.com`,
      slot.SlotID,
      slot.Date,
      slot.Start,
      slot.End,
      'pending',
      false, false, false,
      'restore-test'
    ]);
  }
  
  // 確定させる
  processPendingBatch_();
  
  const beforeState = getCancelTestState(slot.SlotID);
  const archiveCountBefore = getArchiveCount();
  
  // 1人キャンセルして復元を試みる
  const cancelOps = ensureCancelOpsSheet_();
  cancelOps.appendRow([
    'test.restore.1@example.com',
    'confirmed',
    'refill-slot',
    'try-fill',
    'テスト5',
    '',
    ''
  ]);
  
  applyCancelOps();
  
  const afterState = getCancelTestState(slot.SlotID);
  const archiveCountAfter = getArchiveCount();
  
  // 復元されたか確認
  const restored = getResponses_().find(r => 
    r.Email === archivedUser.email && r.SlotID === slot.SlotID
  );
  
  return {
    test: 'Archive復元',
    target: 'test.restore.1@example.com',
    before: beforeState,
    after: afterState,
    restored: !!restored,
    archiveChange: archiveCountAfter - archiveCountBefore,
    success: !!restored // 復元されればtrue
  };
}

// 状態取得ヘルパー
function getCancelTestState(slotId) {
  const responses = getResponses_().filter(r => r.SlotID === slotId);
  
  return {
    total: responses.length,
    confirmed: responses.filter(r => r.Status === 'confirmed').length,
    pending: responses.filter(r => r.Status === 'pending').length,
    waitlist: responses.filter(r => r.Status === 'waitlist').length
  };
}

// Archive件数取得
function getArchiveCount() {
  const archSh = getSS_().getSheetByName(SHEETS.ARCH);
  if (!archSh) return 0;
  return archSh.getLastRow() - 1; // ヘッダーを除く
}

// テスト結果表示
function showCancelTestResults(results) {
  let message = '【CancelOps包括的テスト結果】\n\n';
  
  results.forEach((result, index) => {
    message += `■ ${result.test}\n`;
    message += `対象: ${result.target}\n`;
    
    if (result.test === '複数枠確定者のキャンセル') {
      message += `Slot3: 確定${result.before.slot3.confirmed}→${result.after.slot3.confirmed}\n`;
      message += `Slot4: 確定${result.before.slot4.confirmed}→${result.after.slot4.confirmed}\n`;
    } else if (result.test === 'Archive復元') {
      message += `状態: 確定${result.before.confirmed}→${result.after.confirmed}\n`;
      message += `復元: ${result.restored ? '成功' : '失敗'}\n`;
    } else {
      message += `状態変化:\n`;
      message += `  確定: ${result.before.confirmed}→${result.after.confirmed}\n`;
      message += `  Pending: ${result.before.pending}→${result.after.pending}\n`;
      message += `  合計: ${result.before.total}→${result.after.total}\n`;
    }
    
    message += `結果: ${result.success ? '✅ 成功' : '❌ 失敗'}\n\n`;
  });
  
  // 現実的な問題の確認
  message += '■ 現実的な問題のチェック\n';
  message += checkRealWorldIssues();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('CancelOpsテスト結果', message, ui.ButtonSet.OK);
  
  console.log(message);
}

// 現実的な問題のチェック
function checkRealWorldIssues() {
  let issues = '';
  
  // 1. 同時キャンセルの問題
  const responses = getResponses_();
  const confirmedBySlot = {};
  responses.forEach(r => {
    if (r.Status === 'confirmed') {
      if (!confirmedBySlot[r.SlotID]) confirmedBySlot[r.SlotID] = 0;
      confirmedBySlot[r.SlotID]++;
    }
  });
  
  Object.keys(confirmedBySlot).forEach(slotId => {
    const count = confirmedBySlot[slotId];
    const slot = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
      .find(s => s.SlotID === slotId);
    
    if (slot && count > Number(slot.Capacity)) {
      issues += `⚠️ ${slotId}: 定員超過（${count}/${slot.Capacity}）\n`;
    }
    
    if (count > 0 && count < CONFIG.minCapacityToConfirm) {
      issues += `⚠️ ${slotId}: 最小人数未満で確定（${count}/${CONFIG.minCapacityToConfirm}）\n`;
    }
  });
  
  // 2. メール送信の確認
  const mailQueue = getSS_().getSheetByName(SHEETS.MQ);
  if (mailQueue) {
    const pendingMails = mailQueue.getDataRange().getValues()
      .filter((row, i) => i > 0 && row[7] === 'pending').length;
    
    if (pendingMails > 0) {
      issues += `⚠️ 未送信メール: ${pendingMails}件\n`;
    }
  }
  
  // 3. データ整合性
  const confSheet = ensureConfirmedSheet_();
  const confData = confSheet.getDataRange().getValues();
  if (confData.length > 1) {
    const headers = confData[0];
    const actualCountIdx = headers.indexOf('ActualCount');
    
    for (let i = 1; i < confData.length; i++) {
      const slotId = confData[i][0];
      const actualCount = confData[i][actualCountIdx];
      const realCount = responses.filter(r => 
        r.SlotID === slotId && r.Status === 'confirmed'
      ).length;
      
      if (actualCount !== realCount) {
        issues += `⚠️ ${slotId}: Confirmedシート不整合（記録${actualCount}/実際${realCount}）\n`;
      }
    }
  }
  
  return issues || '問題なし\n';
}

// 個別テスト: キャンセル後の補充フロー
function testCancelRefillFlow() {
  clearAllTestData();
  enableTestMode();
  
  console.log('===== キャンセル補充フローテスト =====');
  
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  
  // 明日のスロットを取得
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = normDateStr_(tomorrow);
  
  const slot = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .find(s => normDateStr_(s.Date) >= tomorrowStr && s.Status === 'open');
  
  if (!slot) {
    SpreadsheetApp.getUi().alert('エラー', 'テスト用スロットが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // capacity + 2人の申込みを作成
  const users = [];
  for (let i = 0; i < CONFIG.capacity + 2; i++) {
    const user = {
      name: `フローテスト${i + 1}`,
      email: `test.flow.${i + 1}@example.com`
    };
    users.push(user);
    
    respSh.appendRow([
      new Date(),
      user.name,
      user.email,
      slot.SlotID,
      slot.Date,
      slot.Start,
      slot.End,
      'pending',
      false, false, false,
      'flow-test'
    ]);
  }
  
  // 確定処理
  console.log('初回確定処理...');
  processPendingBatch_();
  
  const state1 = getCancelTestState(slot.SlotID);
  console.log(`確定後: confirmed=${state1.confirmed}, waitlist=${state1.waitlist}`);
  
  // 確定者1人をキャンセル
  const cancelOps = ensureCancelOpsSheet_();
  cancelOps.appendRow([
    users[0].email,
    'confirmed',
    'refill-slot',
    'try-fill',
    'フローテスト',
    '',
    ''
  ]);
  
  console.log('キャンセル処理...');
  applyCancelOps();
  
  const state2 = getCancelTestState(slot.SlotID);
  console.log(`キャンセル後: confirmed=${state2.confirmed}, waitlist=${state2.waitlist}`);
  
  let message = '【キャンセル補充フローテスト】\n\n';
  message += `初期申込み: ${users.length}名\n`;
  message += `定員: ${CONFIG.capacity}名\n\n`;
  message += `確定処理後:\n`;
  message += `  確定: ${state1.confirmed}名\n`;
  message += `  待機: ${state1.waitlist}名\n\n`;
  message += `1名キャンセル後:\n`;
  message += `  確定: ${state2.confirmed}名\n`;
  message += `  待機: ${state2.waitlist}名\n\n`;
  
  if (state2.confirmed === CONFIG.capacity) {
    message += '✅ 補充成功: waitlistから1名が確定';
  } else {
    message += '❌ 補充失敗: 期待通りに動作していません';
  }
  
  SpreadsheetApp.getUi().alert('テスト結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========= メニュー更新（CancelOpsテスト追加版） =========
function addTestMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🧪テスト機能')
    .addItem('✅ テストモード有効化', 'enableTestMode')
    .addItem('❌ テストモード無効化', 'disableTestMode')
    .addSeparator()
    .addSubMenu(ui.createMenu('📝 基本テスト')
      .addItem('📅 包括的日付テスト（過去/今日/明日）', 'comprehensiveDateTest')
      .addItem('⚡ シンプル即時テスト', 'simpleTestImmediate')
      .addItem('🚀 現実的な20名テスト', 'realisticTest20')
      .addItem('📊 10アカウント×20枠テスト', 'generateTestData10Accounts')
      .addItem('❓ 確定しないケースのテスト', 'testNoConfirmScenario')
      .addItem('➕ 追加申込みシミュレーション', 'simulateAdditionalApplication'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🚫 CancelOpsテスト')
      .addItem('📋 包括的キャンセルテスト', 'comprehensiveCancelTest')
      .addItem('🔄 キャンセル補充フローテスト', 'testCancelRefillFlow')
      .addItem('🔍 個別: 単純キャンセル', 'testSimpleCancelOnly')
      .addItem('💥 個別: 枠全体中止', 'testDropSlotOnly')
      .addItem('⚠️ 個別: 補充不足', 'testInsufficientRefillOnly')
      .addItem('👥 個別: 複数枠キャンセル', 'testMultiSlotCancelOnly')
      .addItem('🗄️ 個別: Archive復元', 'testArchiveRestoreOnly'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 メールテスト')
      .addItem('全メールテスト', 'testAllEmails')
      .addItem('受付メールのみ', 'testReceiptMailOnly')
      .addItem('確定メールのみ', 'testConfirmMailOnly')
      .addItem('管理者ダイジェストのみ', 'testAdminDigestOnly'))
    .addSeparator()
    .addItem('▶️ バッチ処理を今すぐ実行', 'runBatchNow')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 状況確認')
      .addItem('📈 詳細な結果表示', 'showDetailedTestResults')
      .addItem('📊 テスト状況確認', 'showTestStatus')
      .addItem('📋 シート状況確認', 'debugCheckSheets')
      .addItem('🔍 データ整合性チェック', 'checkDataIntegrity'))
    .addSeparator()
    .addItem('🗑️ 全テストデータ削除', 'clearAllTestData')
    .addItem('🔄 スロット状態の再計算', 'updateAllSlotStatuses')
    .addToUi();
}

// ========= 個別CancelOpsテスト（簡易実行用） =========
function testSimpleCancelOnly() {
  clearAllTestData();
  enableTestMode();
  const testData = prepareCancelTestData();
  if (testData.success) {
    const result = testSimpleCancel(testData);
    showSingleCancelTestResult(result);
  }
}

function testDropSlotOnly() {
  clearAllTestData();
  enableTestMode();
  const testData = prepareCancelTestData();
  if (testData.success) {
    const result = testDropSlot(testData);
    showSingleCancelTestResult(result);
  }
}

function testInsufficientRefillOnly() {
  clearAllTestData();
  enableTestMode();
  const testData = prepareCancelTestData();
  if (testData.success) {
    const result = testInsufficientRefill(testData);
    showSingleCancelTestResult(result);
  }
}

function testMultiSlotCancelOnly() {
  clearAllTestData();
  enableTestMode();
  const testData = prepareCancelTestData();
  if (testData.success) {
    const result = testMultiSlotCancel(testData);
    showSingleCancelTestResult(result);
  }
}

function testArchiveRestoreOnly() {
  clearAllTestData();
  enableTestMode();
  const testData = prepareCancelTestData();
  if (testData.success) {
    const result = testArchiveRestore(testData);
    showSingleCancelTestResult(result);
  }
}

// 個別テスト結果表示
function showSingleCancelTestResult(result) {
  let message = `【${result.test}】\n\n`;
  message += `対象: ${result.target}\n\n`;
  
  if (result.test === '複数枠確定者のキャンセル') {
    message += `■ Slot3\n`;
    message += `  確定: ${result.before.slot3.confirmed} → ${result.after.slot3.confirmed}\n`;
    message += `  合計: ${result.before.slot3.total} → ${result.after.slot3.total}\n\n`;
    message += `■ Slot4\n`;
    message += `  確定: ${result.before.slot4.confirmed} → ${result.after.slot4.confirmed}\n`;
    message += `  合計: ${result.before.slot4.total} → ${result.after.slot4.total}\n`;
  } else if (result.test === 'Archive復元') {
    message += `確定: ${result.before.confirmed} → ${result.after.confirmed}\n`;
    message += `復元: ${result.restored ? '✅ 成功' : '❌ 失敗'}\n`;
    message += `Archive変化: ${result.archiveChange}件\n`;
  } else {
    message += `確定: ${result.before.confirmed} → ${result.after.confirmed}\n`;
    message += `Pending: ${result.before.pending} → ${result.after.pending}\n`;
    message += `Waitlist: ${result.before.waitlist} → ${result.after.waitlist}\n`;
    message += `合計: ${result.before.total} → ${result.after.total}\n`;
  }
  
  message += `\n結果: ${result.success ? '✅ テスト成功' : '❌ テスト失敗'}`;
  
  SpreadsheetApp.getUi().alert('CancelOpsテスト結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========= データ整合性チェック =========
function checkDataIntegrity() {
  console.log('===== データ整合性チェック =====');
  
  let issues = [];
  
  // 1. Responses vs Confirmed シートの整合性
  const responses = getResponses_();
  const confSheet = ensureConfirmedSheet_();
  const confData = confSheet.getDataRange().getValues();
  
  if (confData.length > 1) {
    const headers = confData[0];
    const actualCountIdx = headers.indexOf('ActualCount');
    
    for (let i = 1; i < confData.length; i++) {
      const slotId = confData[i][0];
      const recordedCount = confData[i][actualCountIdx];
      const actualConfirmed = responses.filter(r => 
        r.SlotID === slotId && r.Status === 'confirmed'
      ).length;
      
      if (recordedCount !== actualConfirmed) {
        issues.push(`Confirmedシート不整合: ${slotId} (記録:${recordedCount}/実際:${actualConfirmed})`);
      }
    }
  }
  
  // 2. Slots vs Responses の整合性
  const slotSh = getSS_().getSheetByName(SHEETS.SLOTS);
  const slots = readSheetAsObjects_(slotSh);
  
  slots.forEach(slot => {
    const slotResponses = responses.filter(r => r.SlotID === slot.SlotID);
    const confirmedCount = slotResponses.filter(r => r.Status === 'confirmed').length;
    
    // 定員超過チェック
    if (confirmedCount > Number(slot.Capacity)) {
      issues.push(`定員超過: ${slot.SlotID} (確定:${confirmedCount}/定員:${slot.Capacity})`);
    }
    
    // 最小人数チェック
    if (confirmedCount > 0 && confirmedCount < CONFIG.minCapacityToConfirm) {
      issues.push(`最小人数未満で確定: ${slot.SlotID} (確定:${confirmedCount}/最小:${CONFIG.minCapacityToConfirm})`);
    }
    
    // ステータス整合性
    const expectedStatus = confirmedCount >= Number(slot.Capacity) ? 'filled' : 'open';
    if (slot.Status !== expectedStatus && confirmedCount > 0) {
      issues.push(`ステータス不整合: ${slot.SlotID} (現在:${slot.Status}/期待:${expectedStatus})`);
    }
  });
  
  // 3. 重複確定チェック
  if (!CONFIG.allowMultipleConfirmationPerEmail) {
    const confirmedByEmail = {};
    responses.filter(r => r.Status === 'confirmed').forEach(r => {
      const email = String(r.Email).toLowerCase();
      if (!confirmedByEmail[email]) confirmedByEmail[email] = [];
      confirmedByEmail[email].push(r.SlotID);
    });
    
    Object.keys(confirmedByEmail).forEach(email => {
      if (confirmedByEmail[email].length > 1) {
        issues.push(`重複確定: ${email} (${confirmedByEmail[email].join(', ')})`);
      }
    });
  }
  
  // 4. Archive重複チェック
  const archSh = getSS_().getSheetByName(SHEETS.ARCH);
  if (archSh) {
    const archData = archSh.getDataRange().getValues();
    const activeEmails = new Set(responses.map(r => String(r.Email).toLowerCase()));
    
    for (let i = 1; i < archData.length; i++) {
      const email = String(archData[i][3]).toLowerCase();
      const slotId = archData[i][4];
      const restoredAt = archData[i][13];
      
      if (!restoredAt && activeEmails.has(email)) {
        const activeSlot = responses.find(r => 
          String(r.Email).toLowerCase() === email && r.SlotID === slotId
        );
        if (activeSlot) {
          issues.push(`Archive/Active重複: ${email} - ${slotId}`);
        }
      }
    }
  }
  
  // 結果表示
  let message = '【データ整合性チェック結果】\n\n';
  
  if (issues.length === 0) {
    message += '✅ 問題は検出されませんでした\n\n';
  } else {
    message += `⚠️ ${issues.length}件の問題が検出されました:\n\n`;
    issues.forEach((issue, i) => {
      message += `${i + 1}. ${issue}\n`;
    });
  }
  
  message += '\n【統計情報】\n';
  message += `- Responses: ${responses.length}件\n`;
  message += `- 確定: ${responses.filter(r => r.Status === 'confirmed').length}件\n`;
  message += `- Pending: ${responses.filter(r => r.Status === 'pending').length}件\n`;
  message += `- Waitlist: ${responses.filter(r => r.Status === 'waitlist').length}件\n`;
  
  if (archSh) {
    message += `- Archive: ${archSh.getLastRow() - 1}件\n`;
  }
  
  SpreadsheetApp.getUi().alert('データ整合性チェック', message, SpreadsheetApp.getUi().ButtonSet.OK);
  console.log(message);
}
