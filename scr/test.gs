/** ========= テスト専用関数 ========= */
/** 
 * このファイルはテスト用です。
 * 本番環境では削除しても問題ありません。
 */

// テストモード設定（メール送信を無効化）
function enableTestMode() {
  PropertiesService.getScriptProperties().setProperty('TEST_MODE', 'true');
  SpreadsheetApp.getActive().toast('テストモード有効化', 'テスト', 3);
}

function disableTestMode() {
  PropertiesService.getScriptProperties().deleteProperty('TEST_MODE');
  SpreadsheetApp.getActive().toast('テストモード無効化', 'テスト', 3);
}

function isTestMode() {
  return PropertiesService.getScriptProperties().getProperty('TEST_MODE') === 'true';
}

// メール送信のオーバーライド（テスト時のみ）
function sendMailSmart_TEST(opt) {
  if (isTestMode()) {
    console.log('テストモード：メール送信スキップ', {
      to: opt.to,
      subject: opt.subject,
      type: opt.type
    });
    
    // メール送信ログをシートに記録（確認用）
    logMailToSheet_(opt);
    
    return {sent: true, testMode: true};
  }
  // テストモードでない場合は本来の関数を呼び出す
  return sendMailSmart_ORIGINAL(opt);
}

// メール送信ログをシートに記録
function logMailToSheet_(opt) {
  let logSheet = getSS_().getSheetByName('TestMailLog');
  if (!logSheet) {
    logSheet = getSS_().insertSheet('TestMailLog');
    logSheet.appendRow(['Timestamp', 'Type', 'To', 'Subject']);
  }
  logSheet.appendRow([new Date(), opt.type, opt.to, opt.subject]);
}

// 10アカウント分のテストデータ生成
function generateTestData10() {
  // テストモード有効化
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
  
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => s.Status === 'open')
    .slice(0, 10); // 最初の10枠を使用
  
  if (slots.length < 10) {
    SpreadsheetApp.getUi().alert('エラー', 'openステータスの枠が10個未満です。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  let addedCount = 0;
  
  // 各枠に対して申込みを生成
  slots.forEach((slot, slotIndex) => {
    // 各枠に3-5人が申し込む（重複あり）
    const numApplicants = 3 + Math.floor(Math.random() * 3);
    const shuffled = [...testAccounts].sort(() => Math.random() - 0.5);
    const applicants = shuffled.slice(0, numApplicants);
    
    applicants.forEach((account, index) => {
      const timestamp = new Date();
      // 申込み時刻を少しずつずらす（同時申込みのシミュレーション）
      timestamp.setMilliseconds(timestamp.getMilliseconds() + index * 100);
      
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
        'test-data'
      ]);
      addedCount++;
    });
  });
  
  SpreadsheetApp.getActive().toast(
    `テストデータ生成完了：${addedCount}件の申込みを作成しました。30秒後にバッチ処理が実行されます。`, 
    'テスト', 
    10
  );
  
  // バッチ処理をスケジュール
  scheduleDelayedBatch_(30);
}

// 特定の競合状況をテスト
function testConcurrentApplications() {
  enableTestMode();
  
  // 特定の1枠に5人が同時申込み（2名枠の場合、3名がwaitlistになるはず）
  const targetSlot = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .find(s => s.Status === 'open');
  
  if (!targetSlot) {
    SpreadsheetApp.getUi().alert('openな枠がありません');
    return;
  }
  
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const baseTime = new Date();
  
  const applicants = [
    {name: '同時申込A', email: 'concurrent.a@example.com'},
    {name: '同時申込B', email: 'concurrent.b@example.com'},
    {name: '同時申込C', email: 'concurrent.c@example.com'},
    {name: '同時申込D', email: 'concurrent.d@example.com'},
    {name: '同時申込E', email: 'concurrent.e@example.com'}
  ];
  
  applicants.forEach((account, index) => {
    const timestamp = new Date(baseTime.getTime() + index * 50); // 50ミリ秒ずつずらす
    
    respSh.appendRow([
      timestamp,
      account.name,
      account.email,
      targetSlot.SlotID,
      targetSlot.Date,
      targetSlot.Start,
      targetSlot.End,
      'pending',
      false, false, false,
      'concurrent-test'
    ]);
  });
  
  SpreadsheetApp.getActive().toast(
    `${targetSlot.SlotID}に5名が同時申込みしました。30秒後に処理されます。`,
    'テスト',
    10
  );
  
  scheduleDelayedBatch_(30);
}

// キャンセル→補充のテスト
function testCancelAndRefill() {
  enableTestMode();
  
  // 確定済みの最初の人をキャンセル
  const confirmed = getResponses_().filter(r => r.Status === 'confirmed');
  if (confirmed.length === 0) {
    SpreadsheetApp.getUi().alert('確定済みのデータがありません。先にgenerateTestData10()を実行してください。');
    return;
  }
  
  const target = confirmed[0];
  const coSh = ensureCancelOpsSheet_();
  
  // CancelOpsシートに追加
  coSh.appendRow([
    target.Email,
    'confirmed',
    'refill-slot',
    'try-fill',
    'テストキャンセル',
    'pending',
    ''
  ]);
  
  SpreadsheetApp.getActive().toast(
    `${target.Name}（${target.SlotID}）のキャンセル処理を実行します。`,
    'テスト',
    5
  );
  
  // キャンセル処理実行
  applyCancelOps();
}

// テストデータのクリア
function clearAllTestData() {
  const testDomains = ['@example.com'];
  const sheets = [SHEETS.RESP, SHEETS.ARCH];
  let deletedCount = 0;
  
  sheets.forEach(sheetName => {
    const sh = getSS_().getSheetByName(sheetName);
    if (!sh) return;
    
    const data = sh.getDataRange().getValues();
    const emailCol = sheetName === SHEETS.RESP ? 2 : 3; // Email列の位置
    
    for (let i = data.length - 1; i > 0; i--) {
      const email = String(data[i][emailCol] || '').toLowerCase();
      if (testDomains.some(domain => email.includes(domain))) {
        sh.deleteRow(i + 1);
        deletedCount++;
      }
    }
  });
  
  // Confirmedシートのテストデータもクリア
  const confSh = ensureConfirmedSheet_();
  const confData = confSh.getDataRange().getValues();
  if (confData.length > 1) {
    const headers = getConfirmedHeaders();
    for (let i = confData.length - 1; i > 0; i--) {
      let hasTestData = false;
      // 各参加者のメールをチェック
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
  
  // TestMailLogシートも削除
  const logSheet = getSS_().getSheetByName('TestMailLog');
  if (logSheet) {
    getSS_().deleteSheet(logSheet);
  }
  
  SpreadsheetApp.getActive().toast(
    `テストデータを削除しました（${deletedCount}件）`,
    'クリア完了',
    5
  );
  
  // テストモード無効化
  disableTestMode();
}

// テスト実行状況の確認
function showTestStatus() {
  const testEmails = [
    'test.yamada@example.com',
    'test.sato@example.com',
    'test.suzuki@example.com',
    'test.tanaka@example.com',
    'test.takahashi@example.com',
    'test.watanabe@example.com',
    'test.ito@example.com',
    'test.nakamura@example.com',
    'test.kobayashi@example.com',
    'test.kato@example.com',
    'concurrent.a@example.com',
    'concurrent.b@example.com',
    'concurrent.c@example.com',
    'concurrent.d@example.com',
    'concurrent.e@example.com'
  ];
  
  const responses = getResponses_();
  const statusCount = {
    confirmed: 0,
    pending: 0,
    waitlist: 0
  };
  
  testEmails.forEach(email => {
    const userResponses = responses.filter(r => 
      String(r.Email).toLowerCase() === email.toLowerCase()
    );
    userResponses.forEach(r => {
      if (statusCount[r.Status] !== undefined) {
        statusCount[r.Status]++;
      }
    });
  });
  
  const message = `
テストデータ状況:
- confirmed: ${statusCount.confirmed}件
- pending: ${statusCount.pending}件  
- waitlist: ${statusCount.waitlist}件
- テストモード: ${isTestMode() ? '有効' : '無効'}
  `;
  
  SpreadsheetApp.getUi().alert('テスト状況', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// メニューに追加（Code.gsのonOpenから呼び出される）
function addTestMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('テスト機能')
    .addItem('テストモード有効化', 'enableTestMode')
    .addItem('テストモード無効化', 'disableTestMode')
    .addSeparator()
    .addItem('10アカウントのテストデータ生成', 'generateTestData10')
    .addItem('同時申込みテスト（5人→1枠）', 'testConcurrentApplications')
    .addItem('キャンセル→補充テスト', 'testCancelAndRefill')
    .addSeparator()
    .addItem('テスト状況確認', 'showTestStatus')
    .addItem('全テストデータ削除', 'clearAllTestData')
    .addToUi();
}
