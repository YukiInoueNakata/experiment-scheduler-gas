/** ========= SETTINGS ========= */
/** ★ スプレッドシートID（必須）
 *  URL: https://docs.google.com/spreadsheets/d/【このID】/edit...
 */
const SS_ID = 'ここにシートID';

const CONFIG = {
  title: '実験参加スケジュール',
  tz: 'Asia/Tokyo',

  // 枠生成
  startDate: '2025-09-01',
  endDate:   '2025-09-30',
  timeWindows: ['11:00-12:00','13:20-14:20','15:00-16:00','16:50-17:50'],
  capacity: 2,

  // 除外
  excludeWeekends: true,
  excludeDates: ['2025-09-16','2025-09-23'],
  excludeDateTimes: [
    // '2025-09-10 11:00-12:00'
  ],

  // 表示
  showOnlyFromTomorrow: true,

  // 確定ポリシー
  requireFullCapacityToConfirm: true,
  allowMultipleConfirmationPerEmail: false,

  // 総確定人数の上限（例：60名）
  totalConfirmCap: 60,

  // 表示/通知
  location: '立命館大学 OIC ○号館 ○F 実験室A',
  adminEmails: ['you@example.com'],
  mailFromName: '実験担当（自動送信）',

  // メール送信上限ケア
  mail: {
    reserveForReminders: 50,           // 翌朝リマインド分は確保
    hourlyQueueTriggerMinute: 10       // メールキュー：毎時 xx:10
  }
};
