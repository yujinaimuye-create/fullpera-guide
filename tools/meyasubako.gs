/**
 * フルペラフェス 目安箱（ガイドページ用）
 *
 * - ページから送られた質問を「質問」シートに1行ずつ追加します。
 * - 「回答」欄に書いた行だけを、ページの一覧に表示します。
 * - 「非表示」にチェックを入れた行は、回答があってもページに出しません。
 * - ページに出すのは「内容・質問・回答・回答日」だけです。拠点とニックネームは出しません。
 */
const SHEET_NAME = '質問';
const HEADERS = ['受付日時', '拠点', '内容', 'ニックネーム', '質問', '回答', '回答日', '非表示'];
const BASES = ['清水', '横浜', '東京', 'その他'];
const CATS = ['参加・費用', '会場・当日', 'プログラム', '有志バンド', 'その他'];
const MAX_PER_10_MIN = 30; // いたずら対策：10分間に受け付ける件数の上限

/** 最初に1回だけ実行して、シートと見出しを作ります。 */
function setup() {
  const sh = getSheet_();
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#E8E9FF');
    sh.setColumnWidth(5, 320);
    sh.setColumnWidth(6, 320);
    sh.getRange('A:A').setNumberFormat('m/d h:mm');
    sh.getRange('G:G').setNumberFormat('m/d');
  }
}

/** ページから質問を受け取ります。 */
function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.hp) return json_({ ok: true }); // 見えない欄に入力があれば機械的な投稿として捨てる
    const text = clean_(d.text, 500);
    if (text.length < 5) return json_({ ok: false, error: 'short' });

    const cache = CacheService.getScriptCache();
    const count = Number(cache.get('count') || 0);
    if (count >= MAX_PER_10_MIN) return json_({ ok: false, error: 'busy' });
    cache.put('count', String(count + 1), 600);

    const base = BASES.indexOf(d.base) >= 0 ? d.base : 'その他';
    const cat = CATS.indexOf(d.cat) >= 0 ? d.cat : 'その他';
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = getSheet_();
      if (sh.getLastRow() === 0) setup();
      const row = sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, 7).setValues([[new Date(), base, cat, clean_(d.name, 20), text, '', '']]);
      sh.getRange(row, 8).insertCheckboxes();
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  }
}

/** 回答済みの質問を、ページに返します。 */
function doGet() {
  const sh = getSheet_();
  if (sh.getLastRow() < 2) return json_({ ok: true, items: [] });
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();
  const items = rows
    .filter(function (r) { return String(r[5]).trim() && r[7] !== true; })
    .map(function (r) {
      return {
        cat: CATS.indexOf(r[2]) >= 0 ? r[2] : 'その他',
        q: String(r[4]),
        a: String(r[5]),
        at: r[6] instanceof Date ? Utilities.formatDate(r[6], 'Asia/Tokyo', 'M/d') : ''
      };
    })
    .reverse();
  return json_({ ok: true, items: items });
}

/** 「回答」欄に書き込んだら、「回答日」に今日の日付を自動で入れます。 */
function onEdit(e) {
  const r = e.range;
  if (r.getSheet().getName() !== SHEET_NAME || r.getRow() < 2 || r.getColumn() !== 6) return;
  const dateCell = r.getSheet().getRange(r.getRow(), 7);
  if (String(r.getValue()).trim() && !dateCell.getValue()) dateCell.setValue(new Date());
}

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // 数式として動かないようにする
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
