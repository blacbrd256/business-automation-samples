/**
 * Code.gs をそのまま動かし、生成AI・WordPress・Bluesky へのリクエスト内容と
 * シートの状態遷移を確認するテスト。外部APIは呼ばず、応答を差し替えて検証する。
 * 実行：TZ=Asia/Tokyo node test/run-test.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

function createEnv() {
  const sheets = new Map();
  const requests = [];
  const responders = []; // [{match: RegExp, reply: (req)=>({code, body})}]
  const props = {};

  class Sheet {
    constructor() { this.rows = []; }
    appendRow(r) { this.rows.push([...r]); }
    setFrozenRows() {}
    getLastRow() { return this.rows.length; }
    getRange(row, col, numRows = 1, numCols = 1) {
      const s = this;
      const cell = (r, c) => { s.rows[r] = s.rows[r] || []; return s.rows[r]; };
      return {
        getValues: () => Array.from({ length: numRows }, (_, i) =>
          Array.from({ length: numCols }, (_, j) => (s.rows[row - 1 + i] || [])[col - 1 + j] ?? '')),
        setValues: (v) => v.forEach((r, i) => r.forEach((x, j) => { cell(row - 1 + i)[col - 1 + j] = x; })),
        setValue: (x) => { cell(row - 1)[col - 1] = x; },
      };
    }
  }
  const ss = {
    getSheetByName: (n) => sheets.get(n) || null,
    insertSheet: (n) => { const s = new Sheet(); sheets.set(n, s); return s; },
  };

  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] ?? null }) },
    ScriptApp: {
      getProjectTriggers: () => [], deleteTrigger() {},
      newTrigger: () => { const b = { timeBased: () => b, everyHours: () => b, everyMinutes: () => b, create: () => ({}) }; return b; },
    },
    UrlFetchApp: {
      fetch: (url, opt) => {
        const req = { url, ...opt, body: JSON.parse(opt.payload) };
        requests.push(req);
        const r = responders.find((x) => x.match.test(url));
        if (!r) throw new Error('想定外のリクエスト: ' + url);
        const { code, body } = r.reply(req);
        return { getResponseCode: () => code, getContentText: () => (typeof body === 'string' ? body : JSON.stringify(body)) };
      },
    },
    Utilities: {
      base64Encode: (s) => Buffer.from(s).toString('base64'),
      sleep: () => {},
      formatDate: (d) => {
        const p = (n) => String(n).padStart(2, '0');
        return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:00`;
      },
    },
    encodeURIComponent, unescape, console, Date, JSON, String, Math, Error, Object, Array,
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/Code.gs'), 'utf8'), context);
  return { gas: context, sheets, requests, responders, props };
}

const env = createEnv();
const { gas, sheets, requests, responders, props } = env;
Object.assign(props, {
  AI_PROVIDER: 'openai', AI_API_KEY: 'sk-test', AI_MODEL: 'gpt-test',
  WP_SITE_URL: 'https://media.example.com/', WP_USER: 'editor', WP_APP_PASSWORD: 'abcd efgh',
  BSKY_HANDLE: 'nobcelly-demo.bsky.social', BSKY_APP_PASSWORD: 'app-pass',
});
gas.setup();
const q = sheets.get('投稿キュー');
const now = new Date();
const past = new Date(now.getTime() - 3600e3);
const future = new Date(now.getTime() + 86400e3);
q.appendRow(['GAS 自動化', '中小企業の事務担当者向け', '転記作業が月10時間かかっている', past, '公開', '未生成']);
q.appendRow(['フォーム 自動返信', '', '', future, '公開', '未生成']);
q.appendRow(['経費精算 効率化', '', '', past, '下書き', '未生成']);

// ---- 1. 下書き生成 ----
const article = {
  title: 'GASでスプレッドシートの転記作業を自動化する方法',
  meta: '毎月の転記作業に時間を取られていませんか。Google Apps Scriptで自動化する手順を、事務担当者向けにわかりやすく解説します。',
  body: '<h2>なぜ自動化するのか</h2><p onclick="x()">手作業は<strong>ミス</strong>の原因です。</p><script>alert(1)</script><a href="x">リンク</a>',
  sns: 'スプレッドシートの転記に毎月10時間かかっていませんか？GASで自動化する方法をまとめました。 #業務効率化 #GAS',
};
let call = 0;
responders.push({
  match: /api\.openai\.com/,
  reply: () => {
    call++;
    if (call === 3) return { code: 200, body: { choices: [{ message: { content: 'すみません、JSONにできませんでした' } }] } };
    const out = call === 1 ? article : { ...article, title: 'フォームの自動返信メールをGASで設定する手順', sns: 'Googleフォームに自動返信メールを付ける方法を解説しました。 #GAS' };
    return { code: 200, body: { choices: [{ message: { content: '```json\n' + JSON.stringify(out) + '\n```' } }] } };
  },
});
gas.generateDrafts();

const row = (n) => q.rows[n - 1];
assert.strictEqual(row(2)[5], '承認待ち');
assert.strictEqual(row(2)[6], article.title);
assert.ok(!/script|onclick|<a /.test(row(2)[8]), 'script・イベント属性・許可外タグを除去');
assert.ok(row(2)[8].includes('<strong>ミス</strong>'), '許可したタグは残す');
assert.strictEqual(row(4)[5], 'エラー', 'JSONでない出力はエラーにして止める');
assert.match(row(4)[12], /JSON/);
const aiReq = requests[0];
assert.strictEqual(aiReq.headers.Authorization, 'Bearer sk-test');
assert.ok(aiReq.body.messages[0].content.includes('参考メモにない数字・固有名詞・効果を作らない'), 'プロンプトに事実の捏造防止を含む');

// ---- 2. 承認前は投稿しない ----
requests.length = 0;
gas.publishApproved(now);
assert.strictEqual(requests.length, 0, '承認されていない記事は投稿しない');

// ---- 3. 承認後に投稿：Blueskyが一時的に落ちている ----
row(2)[5] = '承認';
row(3)[5] = '承認'; // 投稿日時が未来なので今回は対象外
let wpCount = 0;
let bskyDown = true;
responders.push({ match: /\/wp-json\/wp\/v2\/posts$/, reply: () => { wpCount++; return { code: 201, body: { link: 'https://media.example.com/gas-automation/' } }; } });
responders.push({ match: /createSession/, reply: () => (bskyDown ? { code: 503, body: 'Service Unavailable' } : { code: 200, body: { accessJwt: 'jwt', did: 'did:plc:abc' } }) });
responders.push({ match: /createRecord/, reply: () => ({ code: 200, body: { uri: 'at://did:plc:abc/app.bsky.feed.post/3kxyz' } }) });

gas.publishApproved(now);
assert.strictEqual(row(2)[5], 'エラー');
assert.strictEqual(row(2)[10], 'https://media.example.com/gas-automation/', 'WordPress成功分のURLは保存');
assert.strictEqual(requests.filter((r) => /createSession/.test(r.url)).length, 3, '5xxは3回まで再試行');
const wpReq = requests.find((r) => /wp-json/.test(r.url));
assert.strictEqual(wpReq.url, 'https://media.example.com/wp-json/wp/v2/posts', 'サイトURL末尾のスラッシュを正規化');
assert.strictEqual(wpReq.headers.Authorization, 'Basic ' + Buffer.from('editor:abcd efgh').toString('base64'));
assert.strictEqual(wpReq.body.status, 'publish');
assert.ok(!requests.some((r) => r.body.title === undefined && /wp-json/.test(r.url)));

// ---- 4. 復旧後に再実行：WordPressへは二重投稿しない ----
bskyDown = false;
row(2)[5] = '承認';
requests.length = 0;
gas.publishApproved(now);
assert.strictEqual(wpCount, 1, 'WordPressへの投稿は1回だけ');
assert.strictEqual(row(2)[5], '公開済み');
assert.strictEqual(row(2)[11], 'https://bsky.app/profile/nobcelly-demo.bsky.social/post/3kxyz');

const post = requests.find((r) => /createRecord/.test(r.url)).body.record;
assert.ok(post.text.endsWith('https://media.example.com/gas-automation/'), '投稿文の末尾に記事URL');
const f = post.facets[0].index;
const bytes = Buffer.from(post.text);
assert.strictEqual(bytes.slice(f.byteStart, f.byteEnd).toString(), 'https://media.example.com/gas-automation/', '日本語を含む本文でもリンク位置が正しい');

// ---- 5. 予約投稿と文字数 ----
const later = new Date(future.getTime() + 60e3);
gas.publishApproved(later);
const futureReq = requests.filter((r) => /wp-json/.test(r.url)).pop();
assert.strictEqual(row(3)[5], '公開済み');
assert.strictEqual(futureReq.body.status, 'publish', '投稿日時になったら公開');
const longText = 'あ'.repeat(400);
const fitted = gas.fitForSns_(longText, 'https://example.com/a', 300);
assert.ok([...fitted].length <= 300, 'SNSの文字数上限に収める');
assert.ok(fitted.endsWith('…\nhttps://example.com/a'));

// ---- 6. Claudeへの切り替え ----
props.AI_PROVIDER = 'anthropic';
props.AI_MODEL = 'claude-test';
responders.unshift({ match: /api\.anthropic\.com/, reply: () => ({ code: 200, body: { content: [{ type: 'text', text: JSON.stringify({ ...article, title: '請求書作成をスプレッドシートで自動化する方法', sns: '毎月の請求書作成を自動化する方法をまとめました。 #業務効率化' }) }] } }) });
q.appendRow(['請求書 自動作成', '', '', past, '下書き', '未生成']);
gas.generateDrafts();
const cReq = requests.filter((r) => /anthropic/.test(r.url)).pop();
assert.strictEqual(cReq.headers['x-api-key'], 'sk-test');
assert.strictEqual(cReq.body.model, 'claude-test');
assert.strictEqual(row(5)[5], '承認待ち');
// 「下書き」指定の行はWordPressに下書き保存する
row(5)[5] = '承認';
gas.publishApproved(later);
assert.strictEqual(requests.filter((r) => /wp-json/.test(r.url)).pop().body.status, 'draft', '下書き指定はdraftで保存');

// ---- 7. APIキー未設定ならデモ文で動作確認できる ----
delete props.AI_API_KEY;
q.appendRow(['キー未設定', '', '', past, '下書き', '未生成']);
gas.generateDrafts();
assert.match(row(6)[6], /^【デモ】/);

console.log('✔ すべてのテストに合格しました');

const dump = { sheets: {} };
for (const [n, s] of sheets) dump.sheets[n] = s.rows.map((r) => r.map((c) => (c instanceof Date ? c.toISOString() : c)));
dump.requests = { wordpress: wpReq.body, bluesky: post };
fs.writeFileSync(path.join(__dirname, 'result.json'), JSON.stringify(dump, null, 2));
