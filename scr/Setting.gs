/** ========= SETTINGS ========= */
/** ★ スプレッドシートID（必須）
 *  URL: https://docs.google.com/spreadsheets/d/【このID】/edit...
 */
const SS_ID = '11mPz0I-Axge7UUtpiaMeDIou_l6PKFtdINwqhwMhHg8';

const CONFIG = {
  title: '実験参加スケジュール',
  tz: 'Asia/Tokyo',

  // 枠生成
  startDate: '2025-09-01',
  endDate:   '2025-09-30',
  timeWindows: ['11:00-12:00','13:20-14:20','15:00-16:00','16:50-17:50'],
  
  // 人数設定（重要）
  capacity: 2,                    // 1枠の最大人数
  minCapacityToConfirm: 2,        // 確定に必要な最小人数
  // 例：
  // capacity: 4, minCapacityToConfirm: 3 → 4名枠で3名以上で確定
  // capacity: 2, minCapacityToConfirm: 1 → 2名枠で1名でも確定可
  // capacity: 3, minCapacityToConfirm: 3 → 3名枠で3名揃わないと確定しない

  // 除外日設定
  excludeWeekends: true,
  excludeDates: ['2025-09-16','2025-09-23'],
  excludeDateTimes: [], //記入例 '2025-09-10 11:00-12:00'
    
 

  // 表示設定
  showOnlyFromTomorrow: true,

  // 確定ポリシー
  allowMultipleConfirmationPerEmail: false,  // falseなら1人1枠まで

  // 総確定人数の上限（例：60名）
  totalConfirmCap: 60,

  // バッチ処理設定
  batchProcessDelaySeconds: 30,   // 申込み後、何秒後に確定処理を実行

  // 表示/通知
  location: '立命館大学 OIC ○号館 ○F 実験室A',
  adminEmails: ['dj.y.nakata@gmail.com'],
  mailFromName: '実験担当（自動送信）',

  // メール送信上限ケア
  mail: {
    reserveForReminders: 50,           // 翌朝リマインド分は確保
    hourlyQueueTriggerMinute: 10       // メールキュー：毎時 xx:10
  }
};
