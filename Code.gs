// 水やり管理アプリ - データAPI(Apps Script)
// 画面(HTML/CSS/JS)はGitHub Pagesに移し、このスクリプトは
// Googleスプレッドシート(DB)とGoogleドライブ(写真置き場)への
// 読み書きだけを行う「JSON API」として動かします。
//
// 使い方:
//  1. このファイルの内容を、今まで通り Extensions > Apps Script の
//     Code.gs に丸ごと貼り付けて保存する(index.html / Script2.html は
//     もう使わないので、内容を空にするか削除してOK)。
//  2. 「デプロイを管理」→ 鉛筆アイコン→ バージョン:「新バージョン」→ デプロイ。
//     (URLはこれまでと同じものを使い続けられます)
//  3. そのURLを GitHub Pages 側の index.html 内、API_URL に設定する。

const KV_SHEET_NAME = 'KV';
const IMAGE_FOLDER_NAME = '水やり管理_写真';

// ==== エントリーポイント ====

// 読み取り: GET {URL}?action=get&key=xxxx
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'get') {
      const value = getValue(e.parameter.key);
      return jsonOut_({ ok: true, value: value });
    }
    if (action === 'ping') {
      return jsonOut_({ ok: true, pong: true });
    }
    return jsonOut_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// 書き込み・画像アップロード: POST {URL}  (body: JSONテキスト, Content-Type: text/plain)
// body例: {"action":"set","key":"...","value":"..."}
//        {"action":"uploadImage","dataUrl":"...","filename":"..."}
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'set') {
      setValue(body.key, body.value);
      return jsonOut_({ ok: true });
    }
    if (action === 'uploadImage') {
      const url = uploadImage(body.dataUrl, body.filename);
      return jsonOut_({ ok: true, url: url });
    }
    return jsonOut_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==== KV(キー・バリュー)シート ====

/**
 * KV(キー・バリュー)シートを取得。なければ自動作成する
 */
function getKvSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(KV_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(KV_SHEET_NAME);
    sheet.appendRow(['key', 'value', 'updatedAt']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * key に一致する行番号(1始まり)を返す。見つからなければ -1
 */
function findRowByKey_(sheet, key) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) return i + 2;
  }
  return -1;
}

/**
 * 値を1件取得
 */
function getValue(key) {
  const sheet = getKvSheet_();
  const row = findRowByKey_(sheet, key);
  if (row === -1) return null;
  const value = sheet.getRange(row, 2).getValue();
  return value === '' ? null : String(value);
}

/**
 * 値を1件保存
 */
function setValue(key, value) {
  const sheet = getKvSheet_();
  const row = findRowByKey_(sheet, key);
  const now = new Date();
  if (row === -1) {
    sheet.appendRow([key, value, now]);
  } else {
    sheet.getRange(row, 2, 1, 2).setValues([[value, now]]);
  }
  return true;
}

// ==== 写真アップロード ====

/**
 * 写真をGoogleドライブに保存し、閲覧用URLを返す
 * base64Data: "data:image/jpeg;base64,xxxxx" 形式の文字列
 */
function uploadImage(base64Data, filename) {
  const folder = getOrCreateImageFolder_();

  const commaIdx = base64Data.indexOf(',');
  const meta = base64Data.substring(0, commaIdx);
  const raw = base64Data.substring(commaIdx + 1);
  const contentTypeMatch = meta.match(/data:(.*);base64/);
  const contentType = contentTypeMatch ? contentTypeMatch[1] : 'image/jpeg';

  const bytes = Utilities.base64Decode(raw);
  const blob = Utilities.newBlob(bytes, contentType, filename || ('photo_' + new Date().getTime() + '.jpg'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}

/**
 * 写真保存用フォルダを取得(なければスプレッドシートと同じ場所に作成)
 */
function getOrCreateImageFolder_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  const existing = parentFolder.getFoldersByName(IMAGE_FOLDER_NAME);
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder(IMAGE_FOLDER_NAME);
}
