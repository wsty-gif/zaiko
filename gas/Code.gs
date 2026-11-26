/**
 * 在庫確認アプリ（iPad）→ このスプレッドシートに在庫数を書き込む Apps Script
 *
 * 1. KEY を好きな合言葉（英数字）に書き換える
 * 2. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *    次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 * 3. 表示されたウェブアプリの URL と合言葉を、アプリの「設定 → スプレッドシート」に入れる
 *
 * 書き込むのは「在庫一覧」シートだけです（無ければ作ります）。ほかのシートには触れません。
 */
const KEY = 'ここを合言葉に書き換える';
const SHEET_NAME = '在庫一覧';

function doGet(e) {
  if (!keyOk_(e)) return json_({ ok: false, error: 'key' });
  return json_({ ok: true, app: 'sweets-zaiko', sheet: SHEET_NAME });
}

function doPost(e) {
  if (!keyOk_(e)) return json_({ ok: false, error: 'key' });
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'json' });
  }
  if (!data || data.app !== 'sweets-zaiko' || !Array.isArray(data.header) || !Array.isArray(data.rows)) {
    return json_({ ok: false, error: 'format' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    write_(data);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true, rows: data.rows.length });
}

function keyOk_(e) {
  return KEY !== 'ここを合言葉に書き換える' && !!e && !!e.parameter && e.parameter.key === KEY;
}

function write_(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const w = d.header.length;
  // 列数をそろえ、「=」などで始まる文字は数式として扱われないようにする
  const rows = d.rows.map(function (r) {
    const a = r.slice(0, w);
    while (a.length < w) a.push('');
    return a.map(function (v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v; });
  });
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['最終更新', d.sentAtText || '', '在庫確認の日付', d.dateText || '']]);
  sh.getRange(3, 1, 1, w).setValues([d.header]).setFontWeight('bold').setBackground('#efe7df');
  if (rows.length) {
    sh.getRange(4, 1, rows.length, w).setValues(rows);
    const f = d.flagCol;
    const bg = rows.map(function (r) {
      const v = String(r[f] || '');
      const c = v.indexOf('購入') === 0 ? '#fff0e3' : v ? '#fff7d6' : '#ffffff';
      return new Array(w).fill(c);
    });
    sh.getRange(4, 1, rows.length, w).setBackgrounds(bg);
  }
  sh.setFrozenRows(3);
  sh.autoResizeColumns(1, w);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
