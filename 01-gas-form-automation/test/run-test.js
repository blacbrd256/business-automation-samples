/**
 * GASのサービス（SpreadsheetApp・MailApp など）をメモリ上で再現し、
 * Code.gs をそのまま動かして動作を確認するテスト。
 * 実行：node test/run-test.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

// ---------- GASサービスのモック ----------
function createMockEnv() {
  const sheets = new Map();
  const mails = [];
  const props = {};

  class Sheet {
    constructor(name) { this.name = name; this.rows = []; }
    appendRow(r) { this.rows.push([...r]); }
    setFrozenRows() {}
    getLastRow() { return this.rows.length; }
    getRange(row, col, numRows = 1, numCols = 1) {
      const self = this;
      return {
        getValues() {
          return Array.from({ length: numRows }, (_, i) =>
            Array.from({ length: numCols }, (_, j) => (self.rows[row - 1 + i] || [])[col - 1 + j] ?? ''));
        },
        setValues(v) { v.forEach((r, i) => { self.rows[row - 1 + i] = [...r]; }); },
      };
    }
    getDataRange() { return this.getRange(1, 1, this.rows.length, Math.max(...this.rows.map((r) => r.length))); }
  }

  const ss = {
    getSheetByName: (n) => sheets.get(n) || null,
    insertSheet: (n) => { const s = new Sheet(n); sheets.set(n, s); return s; },
    getUrl: () => 'https://docs.google.com/spreadsheets/d/xxxx',
  };

  const triggers = [];
  const triggerBuilder = (fn) => {
    const t = { fn };
    const b = {
      forSpreadsheet: () => b, onFormSubmit: () => { t.type = 'フォーム送信時'; return b; },
      timeBased: () => b, everyDays: () => { t.type = '毎日'; return b; }, atHour: (h) => { t.hour = h; return b; },
      inTimezone: () => b, create: () => { triggers.push(t); return t; },
    };
    return b;
  };

  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    MailApp: { sendEmail: (m) => mails.push(m) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
    ScriptApp: { getProjectTriggers: () => [...triggers], deleteTrigger: (t) => triggers.splice(triggers.indexOf(t), 1), newTrigger: triggerBuilder },
    Utilities: {
      formatDate: (d, _tz, fmt) => {
        const p = (n) => String(n).padStart(2, '0');
        return fmt.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate()))
          .replace('M', d.getMonth() + 1).replace('d', d.getDate());
      },
    },
    console, Date, JSON, String, Number, Object, Array, Math, Error,
  };
  return { context, sheets, mails, triggers };
}

function load(env) {
  const code = fs.readFileSync(path.join(__dirname, '../src/Code.gs'), 'utf8');
  vm.createContext(env.context);
  vm.runInContext(code, env.context);
  return env.context;
}

const form = (ts, name, email, category, message) => ({
  namedValues: { タイムスタンプ: [ts], お名前: [name], メールアドレス: [email], 種別: [category], お問い合わせ内容: [message] },
});

// ---------- テスト ----------
const env = createMockEnv();
const gas = load(env);
gas.setup();

assert.ok(env.sheets.get('設定'), '設定シートが作成される');
assert.strictEqual(env.triggers.length, 2, 'トリガーが2つ作成される');
gas.setup();
assert.strictEqual(env.triggers.length, 2, 'setupを2回実行してもトリガーは重複しない');

const now = new Date();
const y = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 10, 0, 0);
const at = (h, m) => new Date(y.getFullYear(), y.getMonth(), y.getDate(), h, m).toISOString();

const submissions = [
  form(at(9, 12), '山田 太郎', 'yamada@example.com', '資料請求', 'サービス資料を送ってください。'),
  form(at(10, 3), '佐藤 花子', 'sato@example.com', 'お見積り', '月500件の処理で見積りをお願いします。'),
  form(at(11, 45), '鈴木 一郎', 'suzuki@example.com', '採用', 'エンジニア職の募集状況を知りたいです。'),
  form(at(13, 20), '高橋 美咲', 'takahashi@example', 'お見積り', '至急お願いします。'),          // メール形式エラー
  form(at(15, 8), '田中 健', 'tanaka@example.com', '資料請求', ''),                           // 内容未入力
  form(at(16, 30), '伊藤 誠', 'ito@example.com', 'その他', '請求書の宛名を変更したいです。'),
  form(at(17, 55), '渡辺 愛', 'watanabe@example.com', '取材依頼', '御社サービスについて取材させてください。'), // 設定にない種別
];
submissions.forEach((s) => gas.onFormSubmit(s));

const rows = (n) => (env.sheets.get(n)?.rows.slice(1) || []);
assert.strictEqual(rows('資料請求').length, 1);
assert.strictEqual(rows('お見積り').length, 1);
assert.strictEqual(rows('採用').length, 1);
assert.strictEqual(rows('その他').length, 2, '未知の種別は「その他」へ');
assert.strictEqual(rows('要確認').length, 2, '不備のある回答は要確認へ');
assert.match(rows('資料請求')[0][1], /^R-\d{8}-0001$/, '受付番号の形式');
assert.ok(env.mails.find((m) => m.to === 'sales@example.com' && m.subject.includes('資料請求')), '担当者へ通知');
assert.ok(env.mails.find((m) => m.subject.startsWith('【要確認】')), '不備は要確認として通知');

gas.dailyReport();
const report = rows('日次レポート')[0];
assert.deepStrictEqual(report.slice(1), [1, 1, 1, 2, 2, 7], '日次集計：種別ごと・要確認・合計');
assert.ok(env.mails.find((m) => m.subject.startsWith('【日次レポート】')), '日次レポートを送信');

console.log('✔ すべてのテストに合格しました');

// スクリーンショット用にシート内容をJSONで書き出す
const dump = {};
for (const [name, s] of env.sheets) dump[name] = s.rows.map((r) => r.map((c) => (c instanceof Date ? c.toISOString() : c)));
fs.writeFileSync(path.join(__dirname, 'result.json'), JSON.stringify({ sheets: dump, mails: env.mails }, null, 2));
