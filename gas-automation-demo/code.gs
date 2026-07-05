/**
 * お問い合わせフォーム自動化セット（Google Apps Script）
 *
 * できること:
 *  1. Googleフォームの回答を受信したら、スプレッドシートの「回答一覧」に整形して転記
 *  2. 回答者へ自動でお礼メール（自動返信）を送信
 *  3. 管理者へ新着通知メールを送信
 *  4. 毎朝9時に前日の問い合わせ件数をまとめてメールレポート
 *
 * 導入方法は README.md を参照してください。
 */

// ================= 設定（お客様の環境に合わせてここだけ変更） =================

const CONFIG = {
  // 管理者（通知を受け取る人）のメールアドレス
  ADMIN_EMAIL: 'admin@example.com',

  // 自動返信メールの署名に使う店舗・会社名
  SHOP_NAME: '〇〇店',

  // 転記先シート名
  SHEET_NAME: '回答一覧',

  // フォームの質問タイトル（フォーム側の質問名と一致させる）
  Q_NAME: 'お名前',
  Q_EMAIL: 'メールアドレス',
  Q_BODY: 'お問い合わせ内容',
};

// ================= 1. フォーム回答受信時の処理 =================

/**
 * フォーム送信トリガー（onFormSubmit）から呼ばれるメイン処理
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e
 */
function onFormSubmitHandler(e) {
  const answers = parseFormResponse_(e);

  appendToSheet_(answers);
  sendAutoReply_(answers);
  notifyAdmin_(answers);
}

/** フォーム回答をオブジェクトに変換する */
function parseFormResponse_(e) {
  const values = e.namedValues; // { 質問タイトル: [回答], ... }
  const pick = (key) => (values[key] ? values[key][0] : '');

  return {
    timestamp: new Date(),
    name: pick(CONFIG.Q_NAME),
    email: pick(CONFIG.Q_EMAIL),
    body: pick(CONFIG.Q_BODY),
  };
}

/** 回答一覧シートに整形して転記する（ステータス列付き） */
function appendToSheet_(a) {
  const sheet = getOrCreateSheet_();
  sheet.appendRow([
    Utilities.formatDate(a.timestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    a.name,
    a.email,
    a.body,
    '未対応', // ステータス（未対応 / 対応中 / 完了）
  ]);
}

/** 回答者への自動返信メール */
function sendAutoReply_(a) {
  if (!a.email) return; // メールアドレス未入力なら送らない

  const subject = `【${CONFIG.SHOP_NAME}】お問い合わせありがとうございます`;
  const body = [
    `${a.name} 様`,
    '',
    `この度は${CONFIG.SHOP_NAME}へお問い合わせいただき、誠にありがとうございます。`,
    '以下の内容で受け付けました。担当者より2営業日以内にご連絡いたします。',
    '',
    '──────────────────',
    '■ お問い合わせ内容',
    a.body,
    '──────────────────',
    '',
    '※このメールは自動送信です。心当たりのない場合は破棄してください。',
    '',
    CONFIG.SHOP_NAME,
  ].join('\n');

  GmailApp.sendEmail(a.email, subject, body);
}

/** 管理者への新着通知 */
function notifyAdmin_(a) {
  const subject = `【新着】お問い合わせ: ${a.name} 様`;
  const body = [
    '新しいお問い合わせが届きました。',
    '',
    `お名前: ${a.name}`,
    `メール: ${a.email}`,
    `内容:`,
    a.body,
    '',
    `シートを開く: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`,
  ].join('\n');

  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, body);
}

// ================= 2. 日次レポート（毎朝9時） =================

/** 時間主導型トリガーから呼ばれる。前日の問い合わせをまとめて管理者に送る */
function sendDailyReport() {
  const sheet = getOrCreateSheet_();
  const data = sheet.getDataRange().getValues().slice(1); // ヘッダーを除外

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const targetDate = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy/MM/dd');

  const rows = data.filter((r) => String(r[0]).startsWith(targetDate));
  const pending = data.filter((r) => r[4] === '未対応').length;

  const subject = `【日次レポート】${targetDate} の問い合わせ ${rows.length}件`;
  const lines = [
    `${targetDate} の問い合わせ件数: ${rows.length}件`,
    `未対応の合計: ${pending}件`,
    '',
  ];
  rows.forEach((r, i) => lines.push(`${i + 1}. ${r[0]} ${r[1]} 様: ${String(r[3]).slice(0, 40)}...`));

  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, lines.join('\n'));
}

// ================= 共通ユーティリティ =================

/** 回答一覧シートを取得（なければヘッダー付きで作成） */
function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(['受信日時', 'お名前', 'メールアドレス', '内容', 'ステータス']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:E1').setBackground('#2f5496').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sheet;
}
