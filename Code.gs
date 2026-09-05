// 水やり管理アプリ - データAPI(Apps Script)
// 画面(HTML/CSS/JS)はGitHub Pagesに移し、このスクリプトは
// Googleスプレッドシート(DB)とGoogleドライブ(写真置き場)への
// 読み書きだけを行う「データAPI」として動かします。
//
// 使い方:
//  1. このファイルの内容を、今まで通り Extensions > Apps Script の
//     Code.gs に丸ごと貼り付けて保存する(index.html / Script2.html は
//     もう使わないので、内容を空にするか削除してOK)。
//  2. 「デプロイを管理」→ 鉛筆アイコン→ バージョン:「新バージョン」→ デプロイ。
//     (URLはこれまでと同じものを使い続けられます)
//
// 【重要】Apps Scriptのウェブアプリは、別サイト(GitHub Pages)からの
// 読み書きを許可する印(CORSヘッダー)を技術的に返せません。そのため
// GitHub Pages側は fetch() ではなく、「隠しフォームでこのURLに
// GET/POSTし、ここが返すHTMLページの中の<script>から
// postMessageで結果を送り返す」という方式で通信しています。
// なので、ここのdoGet/doPostは "JSONそのもの" ではなく、
// postMessageするだけの小さなHTMLページを返します。

const KV_SHEET_NAME = 'KV';
const IMAGE_FOLDER_NAME = '水やり管理_写真';

// ==== エントリーポイント ====

// 読み取り: 隠しフォームからの GET (action=get, key=xxxx, requestId=xxxx)
function doGet(e) {
  const requestId = e.parameter.requestId || '';
  try {
    const action = e.parameter.action;
    if (action === 'get') {
      const value = getValue(e.parameter.key);
      return bridgeOut_({ ok: true, value: value, requestId: requestId });
    }
    if (action === 'getMulti') {
      // 起動時など、複数のキーを一度にまとめて読み込むための一括取得。
      // 1件ずつ同時に何本も通信すると、隠しiframeの通信同士が
      // ぶつかってタイムアウトしやすくなるため、1回の通信で
      // まとめて返せるようにしている。
      const keys = (e.parameter.keys || '').split(',').filter(function(k){ return k; });
      const values = {};
      keys.forEach(function(k){ values[k] = getValue(k); });
      return bridgeOut_({ ok: true, values: values, requestId: requestId });
    }
    if (action === 'ping') {
      return bridgeOut_({ ok: true, pong: true, requestId: requestId });
    }
    return bridgeOut_({ ok: false, error: 'unknown action: ' + action, requestId: requestId });
  } catch (err) {
    return bridgeOut_({ ok: false, error: String(err), requestId: requestId });
  }
}

// 書き込み・画像アップロード: 隠しフォームからの POST
// (action=set, key, value, requestId)
// (action=uploadImage, dataUrl, filename, requestId)
function doPost(e) {
  const requestId = (e.parameter && e.parameter.requestId) || '';
  try {
    const action = e.parameter.action;
    if (action === 'set') {
      setValue(e.parameter.key, e.parameter.value);
      return bridgeOut_({ ok: true, requestId: requestId });
    }
    if (action === 'uploadImage') {
      const url = uploadImage(e.parameter.dataUrl, e.parameter.filename);
      return bridgeOut_({ ok: true, url: url, requestId: requestId });
    }
    return bridgeOut_({ ok: false, error: 'unknown action: ' + action, requestId: requestId });
  } catch (err) {
    return bridgeOut_({ ok: false, error: String(err), requestId: requestId });
  }
}

/**
 * postMessageで結果を親ウィンドウ(GitHub Pages側)に送り返すだけの
 * 小さなHTMLページを作る。ALLOWALLにしておかないと、他サイトの
 * iframeの中に表示すること自体をApps Script側に拒否されてしまう。
 */
function bridgeOut_(obj) {
  const json = JSON.stringify(obj).replace(/</g, '\\u003c');
  // Apps Scriptはこのページ自体をさらに内側のiframe(googleusercontent.com)
  // に入れ子で読み込むため、parent(1つ上の階層)ではなくtop(一番外側=
  // GitHub Pagesのページ本体)宛てにpostMessageする必要がある
  const html = '<script>top.postMessage(Object.assign({__gasBridge:true}, ' + json + '), "*");</script>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
