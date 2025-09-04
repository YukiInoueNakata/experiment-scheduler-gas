/** ========= 便利関数 ========= */
function _lines(arr){ return arr.join('\n'); }

/** ========= テンプレ ========= */
var TEMPLATES = {
  participant: {
    receiptSubject: '【受付】実験参加のお申込みを受け付けました',
    receiptBody: _lines([
      '{{name}} 様',
      '',
      '以下のご希望を受け付けました。確定可否は別途ご連絡いたします。',
      '',
      '{{lines}}',
      '',
      '— {{fromName}}'
    ]),
    confirmSubject: '【確定】{{when}} 実験参加のご案内',
    confirmBody: _lines([
      '{{name}} 様',
      '',
      '以下の内容でご参加が確定しました。',
      '',
      '日時：{{when}}（{{tz}}）',
      '場所：{{location}}',
      '',
      '※ キャンセル／変更はこのメールにご返信ください。',
      '',
      '— {{fromName}}'
    ]),
    remindSubject: '【前日リマインド】{{when}} 実験参加',
    remindBody: _lines([
      '{{name}} 様',
      '',
      '明日、以下の時間に実験参加のご予約があります。',
      '日時：{{when}}（{{tz}}）',
      '場所：{{location}}',
      '',
      '道中お気をつけてお越しください。',
      '',
      '— {{fromName}}'
    ]),
    cancelSubject: '【受付】キャンセル処理を完了しました',
    cancelBody: _lines([
      '{{name}} 様',
      '',
      '以下の通り、キャンセル処理を完了しました。',
      '',
      '{{lines}}',
      '',
      '— {{fromName}}'
    ]),
    slotCanceledSubject: '【重要】{{when}} の実施中止のご連絡',
    slotCanceledBody: _lines([
      '{{name}} 様',
      '',
      '下記の回は、やむを得ない事情により「実施中止」となりました。',
      '',
      '日時：{{when}}（{{tz}}）',
      '場所：{{location}}',
      '',
      '— {{fromName}}'
    ])
  },
  admin: {
    confirmSubject: '【確定通知/管理者】{{when}} / {{count}}名',
    confirmBody: _lines([
      '{{when}}（{{tz}}）の枠が確定しました。',
      '場所：{{location}}',
      '',
      '参加者:',
      '{{participants}}',
      '',
      '（自動送信）'
    ]),
    dailyDigestSubject: '【日次ダイジェスト】本日以降の確定一覧（{{date}}時点）',
    dailyDigestBodyIntro: '本日（{{date}}）時点で「本日以降」に確定している枠の一覧です。\n\n'
  },
  // 同意文（HTML）
  consentHtml: [
    '<h2>研究のご説明と同意</h2>',
    '<p>この研究は◯◯を目的とし、2名での対話を通じてデータを取得します。所要時間は約◯分、謝礼は◯◯です。</p>',
    '<p>音声・発話内容は記録され、研究目的の範囲で匿名化・統計的に利用します。個人情報は学内規程に基づき適切に管理し、一定期間保管後に廃棄します。</p>',
    '<p>同意はいつでも撤回できます。撤回時は連絡先（◯◯）にご連絡ください。撤回前までに収集されたデータの研究利用については◯◯となります。</p>',
    '<p>問い合わせ先：◯◯研究室（email@example.com）</p>'
  ].join('')
};

/** ========= テンプレ描画＆日時整形 ========= */
function renderTemplate_(tpl, vars) {
  return String(tpl).replace(/\{\{(\w+)\}\}/g, function(_, k){
    return (k in vars) ? String(vars[k]) : '';
  });
}

// '2025-09-05','13:20' → '2025年09月05日(金)13:20'
function fmtJPDateTime_(dateStr, timeStr) {
  function pad2(n){ return ('0'+n).slice(-2); }
  var ymd = String(dateStr).split('-');
  var y = Number(ymd[0]), m = Number(ymd[1]), d = Number(ymd[2]);
  var dowJ = ['日','月','火','水','木','金','土'][ new Date(dateStr+'T00:00:00+09:00').getDay() ];
  var s = String(timeStr);
  var hhmm = (s.length === 5) ? s : (s.slice(0,2)+':'+s.slice(2,4));
  return y+'年'+pad2(m)+'月'+pad2(d)+'日('+dowJ+')'+hhmm;
}
