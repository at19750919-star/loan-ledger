/**
 * 借貸帳簿 - Google Apps Script 後端
 *
 * 部署步驟：
 * 1. 開啟一個新的 Google Sheet（檔名隨意，例如「借貸帳簿」）
 * 2. 選單：擴充功能 → Apps Script
 * 3. 把此整份檔案貼進去，覆蓋預設的 Code.gs
 * 4. 儲存 → 上方「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 * 5. 執行身分：自己；誰可以存取：「所有人」
 * 6. 點「部署」→ 複製「網頁應用程式 URL」
 * 7. 回到借貸帳簿頁面，點右上「◆ 雲端設定」→ 貼上 URL → 儲存
 *
 * 注意：此 API 沒有驗證，URL 等同密碼，請勿分享。
 */

const SHEET_NAME = 'State';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange('A1').setValue('');
  }
  return sh;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const sh = getSheet_();
    const raw = sh.getRange('A1').getValue();
    if (!raw) return jsonOut_({ ok: true, data: null });
    const data = JSON.parse(raw);
    return jsonOut_({ ok: true, data: data });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = e.postData && e.postData.contents;
    if (!body) return jsonOut_({ ok: false, error: 'empty body' });
    // 驗證 JSON 合法
    const parsed = JSON.parse(body);
    if (!parsed || typeof parsed !== 'object') {
      return jsonOut_({ ok: false, error: 'invalid json' });
    }
    const sh = getSheet_();
    sh.getRange('A1').setValue(body);
    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
