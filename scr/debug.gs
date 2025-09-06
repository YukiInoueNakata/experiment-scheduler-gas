// デバッグ用：日付処理の詳細確認
function debugDateIssue() {
  console.log('===== デバッグ開始 =====');
  
  // 1. 現在の日付情報
  const now = new Date();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  
  console.log('現在時刻:', now);
  console.log('今日（0時）:', today);
  console.log('今日の文字列:', normDateStr_(today));
  console.log('明日（0時）:', tomorrow);
  console.log('明日の文字列:', normDateStr_(tomorrow));
  
  // 2. 利用可能なスロットの日付を確認
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => s.Status === 'open');
  
  console.log('\n===== Open状態のスロット =====');
  const slotDates = new Set();
  slots.forEach(s => {
    const dateStr = normDateStr_(s.Date);
    slotDates.add(dateStr);
  });
  
  const sortedDates = Array.from(slotDates).sort();
  sortedDates.forEach(date => {
    const comparison = date < normDateStr_(tomorrow) ? '❌ 過去/今日' : '✅ 明日以降';
    console.log(`${date}: ${comparison}`);
  });
  
  // 3. 現在のResponsesデータを確認
  console.log('\n===== Responsesのテストデータ =====');
  const responses = getResponses_().filter(r => 
    String(r.Email).toLowerCase().includes('@example.com')
  );
  
  const byDate = {};
  responses.forEach(r => {
    const dateStr = normDateStr_(r.Date);
    if (!byDate[dateStr]) byDate[dateStr] = {pending: 0, confirmed: 0, waitlist: 0};
    byDate[dateStr][r.Status]++;
  });
  
  Object.keys(byDate).sort().forEach(date => {
    const comparison = date < normDateStr_(tomorrow) ? '❌ 過去/今日' : '✅ 明日以降';
    const stats = byDate[date];
    console.log(`${date} ${comparison}: pending=${stats.pending}, confirmed=${stats.confirmed}, waitlist=${stats.waitlist}`);
  });
  
  return {
    todayStr: normDateStr_(today),
    tomorrowStr: normDateStr_(tomorrow),
    slotDates: sortedDates,
    responsesByDate: byDate
  };
}

// archivePastDatePending_関数が正しく動作しているか確認
function testArchiveFunction() {
  console.log('===== Archive関数のテスト =====');
  
  // テスト用データを作成
  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  const testData = [
    {date: yesterday, label: '昨日', expectedArchive: true},
    {date: today, label: '今日', expectedArchive: true},
    {date: tomorrow, label: '明日', expectedArchive: false}
  ];
  
  // テストデータを追加
  testData.forEach((test, index) => {
    const dateStr = normDateStr_(test.date);
    respSh.appendRow([
      new Date(),
      `テスト${test.label}`,
      `test.archive.${index}@example.com`,
      `${dateStr}_1100`,
      dateStr,
      '11:00',
      '12:00',
      'pending',
      false, false, false,
      'archive-test'
    ]);
    console.log(`テストデータ追加: ${test.label} (${dateStr})`);
  });
  
  // Archive関数を実行
  console.log('\narchivePastDatePending_()を実行...');
  archivePastDatePending_();
  
  // 結果を確認
  console.log('\n===== 結果 =====');
  testData.forEach((test, index) => {
    const remaining = getResponses_().filter(r => 
      r.Email === `test.archive.${index}@example.com`
    );
    
    if (test.expectedArchive && remaining.length === 0) {
      console.log(`✅ ${test.label}: 正しくArchiveされた`);
    } else if (!test.expectedArchive && remaining.length > 0) {
      console.log(`✅ ${test.label}: 正しく残っている`);
    } else {
      console.log(`❌ ${test.label}: 期待と異なる結果 (残存: ${remaining.length}件)`);
    }
  });
}

// シンプル即時テストで使用されるスロットを確認
function checkSimpleTestSlots() {
  console.log('===== シンプル即時テストのスロット確認 =====');
  
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const tomorrowStr = normDateStr_(tomorrow);
  
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .filter(s => s.Status === 'open')
    .slice(0, 5);  // シンプル即時テストと同じ条件
  
  console.log(`取得された最初の5スロット:`);
  slots.forEach((slot, index) => {
    const dateStr = normDateStr_(slot.Date);
    const isPast = dateStr < tomorrowStr;
    console.log(`${index + 1}. ${slot.SlotID} (${dateStr}) ${isPast ? '❌ 過去/今日' : '✅ 明日以降'}`);
  });
  
  // 明日以降のスロットの数を確認
  const futureSlots = slots.filter(s => normDateStr_(s.Date) >= tomorrowStr);
  console.log(`\n明日以降のスロット: ${futureSlots.length}/5`);
  
  if (futureSlots.length === 0) {
    console.log('⚠️ 警告: テストで使用可能な明日以降のスロットがありません！');
  }
}

// 修正版のarchivePastDatePending_が存在するか確認
function checkArchiveFunctionCode() {
  const funcStr = archivePastDatePending_.toString();
  
  console.log('===== archivePastDatePending_関数の確認 =====');
  
  if (funcStr.includes('tomorrow')) {
    console.log('✅ 修正版（tomorrowを使用）が適用されています');
  } else if (funcStr.includes('today')) {
    console.log('❌ 古いバージョン（todayを使用）のままです');
    console.log('BatchProcess.gsの archivePastDatePending_ 関数を修正版に置き換えてください');
  } else {
    console.log('⚠️ 関数の内容を確認できません');
  }
  
  // 関数の最初の数行を表示
  const lines = funcStr.split('\n').slice(0, 10);
  console.log('\n関数の冒頭部分:');
  lines.forEach(line => console.log(line));
}
