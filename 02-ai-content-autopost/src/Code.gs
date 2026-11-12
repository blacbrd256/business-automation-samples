/**
 * 生成AIで記事とSNS投稿文を作り、承認後にWordPress・Blueskyへ自動投稿（Google Apps Script）
 *
 * 流れ
 *   1. 「投稿キュー」シートにキーワード・テーマ・参考メモを入力（ステータス：未生成）
 *   2. generateDrafts() … 生成AIで タイトル／メタディスクリプション／本文／SNS投稿文 を作成 → 「承認待ち」
 *   3. 人が内容を確認し、ステータスを「承認」に変更
 *   4. publishApproved() … 投稿日時を過ぎた「承認」行を WordPress（予約・下書き可）と Bluesky に投稿 → 「公開済み」
 *
 * 生成AIは OpenAI / Anthropic(Claude) / Google(Gemini) を設定で切り替えられます。
 * APIキーやパスワードはシートではなく「スクリプト プロパティ」に保存します。
 */

const SHEET_QUEUE = '投稿キュー';
const SHEET_LOG = '処理ログ';

// 投稿キューの列（1始まり）
const COL = {
  KEYWORD: 1, THEME: 2, NOTES: 3, PUBLISH_AT: 4, WP_STATUS: 5, STATUS: 6,
  TITLE: 7, META: 8, BODY: 9, SNS: 10, WP_URL: 11, SNS_URL: 12, ERROR: 13,
};
const QUEUE_HEADERS = ['キーワード', 'テーマ・切り口', '参考メモ', '投稿日時', 'WP公開方法', 'ステータス',
  'タイトル', 'メタディスクリプション', '本文(HTML)', 'SNS投稿文', 'WordPress URL', 'Bluesky URL', 'エラー'];

const STATUS = { NEW: '未生成', REVIEW: '承認待ち', APPROVED: '承認', DONE: '公開済み', ERROR: 'エラー' };
const MAX_PER_RUN = 5; // 1回の実行で処理する件数（GASの実行時間制限対策）

// ===================================================================
// 1. 下書き生成
// ===================================================================
function generateDrafts() {
  const sheet = getQueue_();
  const rows = readRows_(sheet).filter((r) => r.status === STATUS.NEW && r.keyword).slice(0, MAX_PER_RUN);
  rows.forEach((r) => {
    try {
      const draft = generateArticle_(r);
      sheet.getRange(r.row, COL.TITLE, 1, 4).setValues([[draft.title, draft.meta, draft.body, draft.sns]]);
      setStatus_(sheet, r.row, STATUS.REVIEW, '');
      log_('生成', r.row, draft.title);
    } catch (err) {
      setStatus_(sheet, r.row, STATUS.ERROR, String(err.message || err));
      log_('生成エラー', r.row, String(err.message || err));
    }
  });
}

// ===================================================================
// 2. 承認済みを投稿
// ===================================================================
function publishApproved(now) {
  now = now || new Date();
  const sheet = getQueue_();
  const rows = readRows_(sheet)
    .filter((r) => r.status === STATUS.APPROVED && (!r.publishAt || r.publishAt <= now))
    .slice(0, MAX_PER_RUN);

  rows.forEach((r) => {
    try {
      // 途中で失敗しても二重投稿しないよう、成功した投稿先はURLを先に保存する
      let wpUrl = r.wpUrl;
      if (!wpUrl) {
        wpUrl = postToWordPress_(r, now);
        sheet.getRange(r.row, COL.WP_URL).setValue(wpUrl);
      }
      let snsUrl = r.snsUrl;
      if (!snsUrl && r.sns) {
        snsUrl = postToBluesky_(fitForSns_(r.sns, wpUrl, 300));
        sheet.getRange(r.row, COL.SNS_URL).setValue(snsUrl);
      }
      setStatus_(sheet, r.row, STATUS.DONE, '');
      log_('公開', r.row, `${wpUrl} / ${snsUrl || '-'}`);
    } catch (err) {
      setStatus_(sheet, r.row, STATUS.ERROR, String(err.message || err));
      log_('公開エラー', r.row, String(err.message || err));
    }
  });
}

// ===================================================================
// 初期設定
// ===================================================================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET_QUEUE)) {
    const s = ss.insertSheet(SHEET_QUEUE);
    s.appendRow(QUEUE_HEADERS);
    s.setFrozenRows(1);
  }
  if (!ss.getSheetByName(SHEET_LOG)) ss.insertSheet(SHEET_LOG).appendRow(['日時', '処理', '行', '詳細']);

  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('generateDrafts').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('publishApproved').timeBased().everyMinutes(15).create();
}

// ===================================================================
// 生成AI
// ===================================================================

/** 記事一式を生成して検証済みのオブジェクトで返す */
function generateArticle_(r) {
  const prompt = buildPrompt_(r);
  const text = callLLM_(prompt);
  const draft = parseJson_(text);
  ['title', 'meta', 'body', 'sns'].forEach((k) => {
    if (!draft[k] || typeof draft[k] !== 'string') throw new Error(`AIの出力に「${k}」がありません`);
  });
  draft.title = draft.title.slice(0, 60);
  draft.meta = draft.meta.slice(0, 120);
  draft.body = sanitizeHtml_(draft.body);
  return draft;
}

function buildPrompt_(r) {
  return [
    'あなたは日本語のWebメディア編集者です。次の条件でブログ記事を作成してください。',
    `- 狙うキーワード：${r.keyword}`,
    `- テーマ・切り口：${r.theme || '指定なし'}`,
    `- 参考メモ（事実情報）：${r.notes || 'なし'}`,
    '',
    '守ること：',
    '- 参考メモにない数字・固有名詞・効果を作らない。不確かなことは書かない',
    '- 他サイトの文章を写さず、自分の言葉で書く',
    '- 見出しは h2 / h3、本文は 1,500〜2,500字。読者の疑問に答える順に構成する',
    '- 最後に、問い合わせを促す短いまとめを入れる',
    '',
    '次のJSONだけを出力してください（前後に説明文を付けない）：',
    '{"title":"32字前後のタイトル","meta":"100〜120字のメタディスクリプション","body":"本文のHTML（h2,h3,p,ul,li,strongのみ）","sns":"記事を紹介するSNS投稿文（200字以内・ハッシュタグ2個まで・URLは入れない）"}',
  ].join('\n');
}

/** 設定に応じて OpenAI / Claude / Gemini を呼ぶ。キーが未設定ならデモ用の固定文を返す */
function callLLM_(prompt) {
  const p = PropertiesService.getScriptProperties();
  const provider = (p.getProperty('AI_PROVIDER') || 'openai').toLowerCase();
  const key = p.getProperty('AI_API_KEY');
  const model = p.getProperty('AI_MODEL');
  if (!key) return demoResponse_();

  let req;
  if (provider === 'anthropic') {
    req = {
      url: 'https://api.anthropic.com/v1/messages',
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: { model, max_tokens: 4096, messages: [{ role: 'user', content: prompt }] },
      pick: (j) => j.content[0].text,
    };
  } else if (provider === 'gemini') {
    req = {
      url: `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(key)}`,
      headers: {},
      payload: { contents: [{ parts: [{ text: prompt }] }], generationConfig: { responseMimeType: 'application/json' } },
      pick: (j) => j.candidates[0].content.parts[0].text,
    };
  } else {
    req = {
      url: 'https://api.openai.com/v1/chat/completions',
      headers: { Authorization: `Bearer ${key}` },
      payload: { model, messages: [{ role: 'user', content: prompt }], response_format: { type: 'json_object' } },
      pick: (j) => j.choices[0].message.content,
    };
  }
  const json = fetchJson_(req.url, { method: 'post', headers: req.headers, payload: req.payload });
  return req.pick(json);
}

// ===================================================================
// 投稿先
// ===================================================================

/** WordPress REST API（アプリケーションパスワード認証）で投稿し、URLを返す */
function postToWordPress_(r, now) {
  now = now || new Date();
  const p = PropertiesService.getScriptProperties();
  const site = String(p.getProperty('WP_SITE_URL') || '').replace(/\/$/, '');
  const auth = Utilities.base64Encode(`${p.getProperty('WP_USER')}:${p.getProperty('WP_APP_PASSWORD')}`);
  if (!site) throw new Error('WP_SITE_URL が未設定です');

  const status = r.wpStatus === '下書き' ? 'draft' : (r.publishAt && r.publishAt > now ? 'future' : 'publish');
  const payload = { title: r.title, content: r.body, excerpt: r.meta, status };
  if (status === 'future') payload.date = Utilities.formatDate(r.publishAt, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");

  const json = fetchJson_(`${site}/wp-json/wp/v2/posts`, {
    method: 'post', headers: { Authorization: `Basic ${auth}` }, payload,
  });
  return json.link;
}

/** Bluesky（AT Protocol）に投稿し、投稿URLを返す */
function postToBluesky_(text) {
  const p = PropertiesService.getScriptProperties();
  const handle = p.getProperty('BSKY_HANDLE');
  const session = fetchJson_('https://bsky.social/xrpc/com.atproto.server.createSession', {
    method: 'post', payload: { identifier: handle, password: p.getProperty('BSKY_APP_PASSWORD') },
  });
  const record = { $type: 'app.bsky.feed.post', text, createdAt: new Date().toISOString(), facets: linkFacets_(text) };
  const res = fetchJson_('https://bsky.social/xrpc/com.atproto.repo.createRecord', {
    method: 'post', headers: { Authorization: `Bearer ${session.accessJwt}` },
    payload: { repo: session.did, collection: 'app.bsky.feed.post', record },
  });
  const rkey = res.uri.split('/').pop();
  return `https://bsky.app/profile/${handle}/post/${rkey}`;
}

/** Blueskyで本文中のURLをリンクにするための位置情報（バイト単位） */
function linkFacets_(text) {
  const facets = [];
  const re = /https?:\/\/[^\s]+/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const byteStart = utf8Length_(text.slice(0, m.index));
    facets.push({
      index: { byteStart, byteEnd: byteStart + utf8Length_(m[0]) },
      features: [{ $type: 'app.bsky.richtext.facet#link', uri: m[0] }],
    });
  }
  return facets;
}

/** SNSの文字数上限に収まるよう本文を切り詰め、URLを末尾に付ける */
function fitForSns_(text, url, limit) {
  const suffix = url ? `\n${url}` : '';
  const room = limit - [...suffix].length;
  const chars = [...text.trim()];
  const body = chars.length > room ? chars.slice(0, room - 1).join('') + '…' : chars.join('');
  return body + suffix;
}

// ===================================================================
// 共通処理
// ===================================================================

/** JSONのAPIを呼ぶ。429/5xx は待ってから最大3回まで再試行 */
function fetchJson_(url, { method, headers, payload }) {
  let lastErr;
  for (let attempt = 1; attempt <= 3; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method, contentType: 'application/json', headers: headers || {},
      payload: JSON.stringify(payload), muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) return JSON.parse(res.getContentText());
    lastErr = new Error(`HTTP ${code}: ${res.getContentText().slice(0, 200)}`);
    if (code !== 429 && code < 500) break;
    Utilities.sleep(1000 * Math.pow(2, attempt));
  }
  throw lastErr;
}

/** AIの出力からJSON部分だけを取り出す（```json で囲まれていても可） */
function parseJson_(text) {
  const m = String(text).match(/\{[\s\S]*\}/);
  if (!m) throw new Error('AIの出力がJSON形式ではありません');
  return JSON.parse(m[0]);
}

/** 許可したタグ以外を除去（script等の混入防止） */
function sanitizeHtml_(html) {
  return String(html)
    .replace(/<\s*(script|style|iframe)[\s\S]*?<\s*\/\s*\1\s*>/gi, '')
    .replace(/<(?!\/?(h2|h3|p|ul|ol|li|strong|em|br)\b)[^>]*>/gi, '')
    .replace(/\son\w+="[^"]*"/gi, '');
}

function readRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, QUEUE_HEADERS.length).getValues().map((v, i) => ({
    row: i + 2, keyword: v[COL.KEYWORD - 1], theme: v[COL.THEME - 1], notes: v[COL.NOTES - 1],
    publishAt: v[COL.PUBLISH_AT - 1] instanceof Date ? v[COL.PUBLISH_AT - 1] : null,
    wpStatus: v[COL.WP_STATUS - 1], status: v[COL.STATUS - 1], title: v[COL.TITLE - 1], meta: v[COL.META - 1],
    body: v[COL.BODY - 1], sns: v[COL.SNS - 1], wpUrl: v[COL.WP_URL - 1], snsUrl: v[COL.SNS_URL - 1],
  }));
}

function setStatus_(sheet, row, status, error) {
  sheet.getRange(row, COL.STATUS).setValue(status);
  sheet.getRange(row, COL.ERROR).setValue(error);
}

function getQueue_() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUEUE);
  if (!s) throw new Error('「投稿キュー」シートがありません。setup() を実行してください。');
  return s;
}

function log_(type, row, detail) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (s) s.appendRow([new Date(), type, row, detail]);
}

function utf8Length_(s) {
  return unescape(encodeURIComponent(s)).length;
}

/** APIキー未設定時のデモ出力（動作確認用） */
function demoResponse_() {
  return JSON.stringify({
    title: '【デモ】スプレッドシート業務を自動化する3つの方法',
    meta: 'APIキー未設定のため、デモ用の固定文を表示しています。設定後は指定キーワードに合わせた記事が生成されます。',
    body: '<h2>デモ記事</h2><p>APIキーを設定すると、ここに生成された本文が入ります。</p>',
    sns: '【デモ】スプレッドシートの手作業を減らす方法をまとめました。 #業務効率化',
  });
}

if (typeof module !== 'undefined') module.exports = {};
