/**
 * フォーム回答の自動振り分け・通知・日次集計（Google Apps Script）
 *
 * できること
 *   1. Googleフォームの回答を「種別」ごとのシートへ自動で振り分ける
 *   2. 必須項目の抜けやメールアドレスの誤りをチェックし、要確認として記録する
 *   3. 担当者へ通知メールを送る（種別ごとに宛先を変えられる）
 *   4. 毎朝、前日分の件数を種別ごとに集計してレポートシートへ追記・メール送信
 *
 * 使い方は README.md を参照してください。
 * 設定はすべて「設定」シートで変更でき、コードを書き換える必要はありません。
 */

// ---- シート名（変更する場合はここだけ直す） ----
const SHEET = {
  RAW: 'フォームの回答 1', // フォーム連携で自動作成されるシート
  CONFIG: '設定',
  REPORT: '日次レポート',
  LOG: '処理ログ',
  REVIEW: '要確認',
};

// 振り分け先シートに書き出す列
const OUTPUT_HEADERS = ['受付日時', '受付番号', 'お名前', 'メールアドレス', '種別', 'お問い合わせ内容', '担当', 'ステータス'];
const REVIEW_HEADERS = ['受付日時', '受付番号', '理由', '元データ'];

/**
 * フォーム送信時に実行（トリガー：フォーム送信時）
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e
 */
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000); // 同時送信で受付番号が重複しないようにする
  try {
    const answer = normalizeAnswer_(e.namedValues);
    const config = loadConfig_();
    const errors = validate_(answer);
    const receiptNo = issueReceiptNo_();

    if (errors.length > 0) {
      appendRow_(SHEET.REVIEW, REVIEW_HEADERS, [
        answer.timestamp, receiptNo, errors.join(' / '), JSON.stringify(e.namedValues),
      ]);
      notify_(config.fallbackTo, `【要確認】フォーム回答 ${receiptNo}`, `入力内容に不備がありました。\n理由：${errors.join(' / ')}`);
      log_('要確認', receiptNo, errors.join(' / '));
      return;
    }

    // 設定にない種別は「その他」として扱い、元の種別は内容の先頭に残す
    const key = config.routes[answer.category] ? answer.category : 'その他';
    const rule = config.routes[key];
    const message = key === answer.category ? answer.message : `［種別：${answer.category}］${answer.message}`;
    appendRow_(rule.sheet, OUTPUT_HEADERS, [
      answer.timestamp, receiptNo, answer.name, answer.email, key, message, rule.owner, '未対応',
    ]);

    notify_(rule.to, `【${answer.category}】新しいお問い合わせ ${receiptNo}`, buildMailBody_(answer, receiptNo, rule));
    log_('振り分け', receiptNo, `${answer.category} → ${rule.sheet}`);
  } catch (err) {
    log_('エラー', '-', String(err && err.stack || err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 前日分の件数を種別ごとに集計（トリガー：毎日 午前8〜9時）
 */
function dailyReport() {
  const config = loadConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { start, end } = yesterdayRange_(new Date());

  const counts = {};
  let total = 0;
  Object.keys(config.routes).forEach((category) => {
    const sheet = ss.getSheetByName(config.routes[category].sheet);
    if (!sheet || sheet.getLastRow() < 2) { counts[category] = 0; return; }
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, OUTPUT_HEADERS.length).getValues();
    const n = rows.filter((r) => r[0] instanceof Date && r[0] >= start && r[0] < end && r[4] === category).length;
    counts[category] = n;
    total += n;
  });
  const review = countInRange_(ss.getSheetByName(SHEET.REVIEW), start, end);

  const categories = Object.keys(config.routes);
  appendRow_(SHEET.REPORT, ['対象日', ...categories, '要確認', '合計'], [
    Utilities.formatDate(start, 'Asia/Tokyo', 'yyyy/MM/dd'), ...categories.map((c) => counts[c]), review, total + review,
  ]);

  const lines = categories.map((c) => `・${c}：${counts[c]}件`).concat([`・要確認：${review}件`, `合計：${total + review}件`]);
  notify_(config.reportTo, `【日次レポート】${Utilities.formatDate(start, 'Asia/Tokyo', 'M/d')}のお問い合わせ`, lines.join('\n'));
  log_('日次レポート', '-', `合計${total + review}件`);
}

/**
 * 初回だけ実行：必要なシートとトリガーを作成
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureConfigSheet_(ss);
  ensureSheet_(SHEET.LOG, ['日時', '処理', '受付番号', '詳細']);
  ensureSheet_(SHEET.REVIEW, REVIEW_HEADERS);
  // 日次レポートは種別の列が設定で変わるため、初回集計時に見出し付きで作成する

  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t)); // 二重登録を防ぐ
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('dailyReport').timeBased().everyDays(1).atHour(8).inTimezone('Asia/Tokyo').create();
  log_('セットアップ', '-', 'シートとトリガーを作成しました');
}

// ===================================================================
// 以下は内部処理（通常は変更不要）
// ===================================================================

/** フォームの質問名ゆれを吸収して、決まった形に整える */
function normalizeAnswer_(namedValues) {
  const pick = (...keys) => {
    for (const k of keys) {
      if (namedValues[k] && String(namedValues[k][0]).trim() !== '') return String(namedValues[k][0]).trim();
    }
    return '';
  };
  const ts = pick('タイムスタンプ', 'Timestamp');
  return {
    timestamp: ts ? new Date(ts) : new Date(),
    name: pick('お名前', '氏名', 'Name'),
    email: pick('メールアドレス', 'Email', 'Email Address'),
    category: pick('種別', 'お問い合わせ種別', 'カテゴリ') || 'その他',
    message: pick('お問い合わせ内容', '内容', 'Message'),
  };
}

/** 入力チェック。問題があれば理由の配列を返す */
function validate_(a) {
  const errors = [];
  if (!a.name) errors.push('お名前が未入力');
  if (!a.message) errors.push('お問い合わせ内容が未入力');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a.email)) errors.push('メールアドレスの形式が不正');
  if (a.message && a.message.length > 5000) errors.push('内容が長すぎます（5000文字超）');
  return errors;
}

/** 設定シートを読み込む */
function loadConfig_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CONFIG);
  if (!sheet) throw new Error('「設定」シートがありません。setup() を実行してください。');
  const rows = sheet.getDataRange().getValues().slice(1); // 1行目は見出し
  const config = { routes: {}, reportTo: '', fallbackTo: '' };
  rows.forEach(([key, sheetName, owner, to]) => {
    if (!key) return;
    if (key === '日次レポート送信先') config.reportTo = String(to || sheetName);
    else if (key === '要確認の送信先') config.fallbackTo = String(to || sheetName);
    else config.routes[String(key)] = { sheet: String(sheetName), owner: String(owner), to: String(to) };
  });
  if (!config.routes['その他']) throw new Error('設定シートに「その他」の行が必要です');
  return config;
}

/** 設定シートがなければひな形を作る */
function ensureConfigSheet_(ss) {
  if (ss.getSheetByName(SHEET.CONFIG)) return;
  const s = ss.insertSheet(SHEET.CONFIG);
  s.getRange(1, 1, 7, 4).setValues([
    ['種別', '振り分け先シート', '担当', '通知先メール'],
    ['資料請求', '資料請求', '営業部', 'sales@example.com'],
    ['お見積り', 'お見積り', '営業部', 'sales@example.com'],
    ['採用', '採用', '人事部', 'hr@example.com'],
    ['その他', 'その他', '総務部', 'info@example.com'],
    ['日次レポート送信先', '', '', 'manager@example.com'],
    ['要確認の送信先', '', '', 'info@example.com'],
  ]);
  s.setFrozenRows(1);
}

/** 受付番号（例：R-20261113-0007）を日ごとの連番で発行 */
function issueReceiptNo_() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  const key = 'seq_' + today;
  const next = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(next));
  return `R-${today}-${String(next).padStart(4, '0')}`;
}

/** シートがなければ見出し付きで作成。空のシートには見出しを入れる */
function ensureSheet_(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 見出しを保証したうえで1行追加 */
function appendRow_(sheetName, headers, row) {
  ensureSheet_(sheetName, headers).appendRow(row);
}

function buildMailBody_(a, receiptNo, rule) {
  return [
    `新しいお問い合わせを受け付けました（担当：${rule.owner}）。`,
    '',
    `受付番号：${receiptNo}`,
    `お名前　：${a.name}`,
    `メール　：${a.email}`,
    `種別　　：${a.category}`,
    '',
    '―― お問い合わせ内容 ――',
    a.message,
    '',
    `一覧：${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`,
  ].join('\n');
}

function notify_(to, subject, body) {
  if (!to) return;
  MailApp.sendEmail({ to, subject, body });
}

function log_(type, receiptNo, detail) {
  appendRow_(SHEET.LOG, ['日時', '処理', '受付番号', '詳細'], [new Date(), type, receiptNo, detail]);
}

function yesterdayRange_(now) {
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const start = new Date(end.getTime() - 24 * 60 * 60 * 1000);
  return { start, end };
}

function countInRange_(sheet, start, end) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .filter(([d]) => d instanceof Date && d >= start && d < end).length;
}

// Node.js でのテスト用（GAS上では無視される）
if (typeof module !== 'undefined') {
  module.exports = { onFormSubmit, dailyReport, setup, normalizeAnswer_, validate_ };
}
