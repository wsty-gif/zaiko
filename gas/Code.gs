/**
 * 在庫確認アプリ（iPad）→ このスプレッドシートの「在庫管理」シートに在庫数を書き込む Apps Script
 *
 * 1. KEY を好きな合言葉（英数字）に書き換える
 * 2. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *    次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 * 3. 表示されたウェブアプリの URL と合言葉を、アプリの「設定 → スプレッドシート」に入れる
 *
 * 「在庫管理」シートの1行目の見出し（商品名・在庫数・発注目安・単位・保管場所）を見て、
 * 商品名が同じ行の値だけを書き換えます。見出しに「購入必要」「確認日」が無ければ右に足します。
 * シートに無い商品は下に行を足します。ほかの列・ほかのシートには触れません。
 *
 * XLSX_ID を入れた場合は、書き込んだあとにスプレッドシート全体を .xlsx に書き出し、
 * そのファイル（元の「在庫管理.xlsx」）の中身を置き換えます。ファイルの ID・名前・共有・URL は変わりません。
 * 置き換える前の中身は、ドライブの「版を管理」に残ります。.xlsx を直接書き換えても、次の書き込みで上書きされます。
 * このときは、同じフォルダの appsscript.json の oauthScopes（ドライブの編集）もプロジェクトに入れてください。
 * 入れないと、Apps Script がドライブの「閲覧」だけの許可を求め、.xlsx を置き換えられません（upload-403）。
 */
const KEY = 'ここを合言葉に書き換える';
const SHEET_NAME = '在庫管理';
const SPREADSHEET_ID = '';   // スプレッドシートから開いた Apps Script なら空のままで OK
const XLSX_ID = '';          // 同じ内容を入れる .xlsx ファイルの ID（使わないなら空のまま）

const COLS = ['商品名', '在庫数', '発注目安', '単位', '保管場所', '購入必要', '確認日'];
const FIELDS = { '在庫数': 'stock', '発注目安': 'target', '単位': 'unit', '保管場所': 'loc', '購入必要': 'need', '確認日': 'date' };

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
  if (!data || data.app !== 'sweets-zaiko' || !Array.isArray(data.items)) return json_({ ok: false, error: 'format' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let xlsx = '';
  try {
    write_(data);
    if (XLSX_ID) xlsx = copyToXlsx_();
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true, rows: data.items.length, xlsx: xlsx });
}

/** スプレッドシート全体を .xlsx に書き出し、XLSX_ID のファイルの中身を置き換える。失敗してもスプレッドシートへの書き込みは取り消さない */
function copyToXlsx_() {
  try {
    SpreadsheetApp.flush();
    const file = DriveApp.getFileById(XLSX_ID);
    if (file.getMimeType() !== MimeType.MICROSOFT_EXCEL) return 'not-xlsx';
    const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
    const ex = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + book_().getId() + '/export?format=xlsx',
      { headers: auth, muteHttpExceptions: true });
    if (ex.getResponseCode() !== 200) return 'export-' + ex.getResponseCode();
    const up = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + XLSX_ID + '?uploadType=media&supportsAllDrives=true', {
      method: 'patch', contentType: MimeType.MICROSOFT_EXCEL, payload: ex.getBlob().getBytes(),
      headers: auth, muteHttpExceptions: true
    });
    return up.getResponseCode() === 200 ? 'ok' : 'upload-' + up.getResponseCode() + ' ' + up.getContentText().slice(0, 300);
  } catch (err) {
    return 'error-' + err;
  }
}

/** エディタで選んで「実行」する用：今のスプレッドシートの内容で .xlsx を置き換える（初回はドライブを使う許可の確認が出ます） */
function copyToXlsxNow() {
  Logger.log(XLSX_ID ? copyToXlsx_() : 'XLSX_ID が空です');
}

function keyOk_(e) {
  return KEY !== 'ここを合言葉に書き換える' && !!e && !!e.parameter && e.parameter.key === KEY;
}

function book_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function safe_(v) {
  // 「=」などで始まる文字が数式として扱われないようにする
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : (v === null || v === undefined ? '' : v);
}

function write_(d) {
  const book = book_();
  const sh = book.getSheetByName(SHEET_NAME) || book.insertSheet(SHEET_NAME);

  // 見出し（1行目）。足りない見出しは右に足す
  const width = Math.max(sh.getLastColumn(), 1);
  const header = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
  while (header.length && header[header.length - 1] === '') header.pop();
  COLS.forEach(function (h) {
    if (header.indexOf(h) < 0) { header.push(h); sh.getRange(1, header.length).setValue(h); }
  });
  const col = {};
  COLS.forEach(function (h) { col[h] = header.indexOf(h) + 1; });

  // 商品名の列から、行の位置を探す（前の名前でも探し、見つかったら今の名前に直す）
  const lastRow = sh.getLastRow();
  const names = lastRow >= 2 ? sh.getRange(2, col['商品名'], lastRow - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  while (names.length && names[names.length - 1] === '') names.pop();
  const rowOf = {};
  let namesChanged = false;
  d.items.forEach(function (it) {
    let i = names.indexOf(it.name);
    if (i < 0) {
      (it.aliases || []).some(function (a) { const j = names.indexOf(a); if (j >= 0) { i = j; return true; } return false; });
      if (i >= 0) { names[i] = it.name; namesChanged = true; }
    }
    if (i < 0) { names.push(it.name); i = names.length - 1; namesChanged = true; }
    rowOf[it.name] = i;
  });
  const n = names.length;
  if (!n) return;
  if (namesChanged) sh.getRange(2, col['商品名'], n, 1).setValues(names.map(function (x) { return [x]; }));

  // 列ごとにまとめて書く（ほかの列の数式などは触らない）
  Object.keys(FIELDS).forEach(function (h) {
    const range = sh.getRange(2, col[h], n, 1);
    const vals = range.getValues();
    d.items.forEach(function (it) { vals[rowOf[it.name]][0] = safe_(it[FIELDS[h]]); });
    range.setValues(vals);
  });
  // 購入が必要な行の「購入必要」をオレンジにする
  const needRange = sh.getRange(2, col['購入必要'], n, 1);
  const bg = needRange.getValues().map(function (r) { return [String(r[0]).indexOf('購入') === 0 ? '#fff0e3' : '#ffffff']; });
  needRange.setBackgrounds(bg);

  // 最終更新（見出しの右）
  let at = header.indexOf('最終更新') + 1;
  if (!at) { at = header.length + 2; sh.getRange(1, at).setValue('最終更新'); }
  sh.getRange(1, at + 1).setValue(d.sentAtText || '');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
