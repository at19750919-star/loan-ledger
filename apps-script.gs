/**
 * 借貸帳簿 - Google Apps Script 後端（表格版）
 *
 * 資料存三個分頁：
 *   - Debtors：借款人清單
 *   - Payments：付款紀錄（以 debtor_id 關聯）
 *   - Meta：其他 state 欄位（如 seeded）
 *
 * 部署步驟：
 * 1. 開啟 Google Sheet（若沿用之前那份，三個分頁會自動建立）
 * 2. 擴充功能 → Apps Script → 把本檔貼進去覆蓋原本的
 * 3. 儲存 → 部署 → 管理部署作業 → 選取原部署 → 右上「編輯」圖示
 *    → 版本改「新增版本」→ 部署
 *    （重新部署後 URL 不變，網頁端不用改）
 * 4. 如果你想「從頭來」，可以把舊的 State 分頁手動刪掉
 */

const DEBTORS_SHEET = 'Debtors';
const PAYMENTS_SHEET = 'Payments';
const META_SHEET = 'Meta';
const EXTRA_SHEETS = ['股東']; // 程式不會讀寫，僅確保分頁存在供手動記錄

const DEBTOR_COLS = ['id', 'name', 'day', 'amount', 'principal', 'interest', 'phone', 'notes', 'createdAt'];
const DEBTOR_HEADERS = ['編號', '姓名', '月付款日', '月應收金額', '本金', '利息', '電話', '備註', '建立時間'];
const PAYMENT_COLS = ['debtor_id', 'id', 'date', 'principal', 'interest', 'note'];
const PAYMENT_HEADERS = ['借款人編號', '付款編號', '日期', '本金', '利息', '備註'];
const META_COLS = ['key', 'value'];
const META_HEADERS = ['項目', '值'];

function getOrCreateSheet_(name, cols, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    // 每次都把第一列覆寫成最新中文標頭（避免舊部署留下英文 header）
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getNs_(e) {
  const ns = (e && e.parameter && e.parameter.ns) ? String(e.parameter.ns).trim() : '';
  if (!ns) return '';
  if (!/^[a-zA-Z0-9_]+$/.test(ns)) return ''; // 只允許英數底線
  return ns;
}

function sheetName_(base, ns) {
  return ns ? base + '_' + ns : base;
}

function ensureExtraSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const name of EXTRA_SHEETS) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readSheet_(sheet, cols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, cols.length).getValues();
  const rows = [];
  for (const row of data) {
    // 整列空就略過
    if (row.every(v => v === '' || v === null)) continue;
    const obj = {};
    cols.forEach((c, i) => { obj[c] = row[i]; });
    rows.push(obj);
  }
  return rows;
}

function writeSheet_(sheet, cols, rows) {
  // 清除舊資料（保留 header）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, cols.length).clearContent();
  }
  if (!rows.length) return;
  const values = rows.map(r => cols.map(c => {
    const v = r[c];
    return (v === undefined || v === null) ? '' : v;
  }));
  sheet.getRange(2, 1, values.length, cols.length).setValues(values);
}

function numOrZero_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function strOrEmpty_(v) {
  return (v === null || v === undefined) ? '' : String(v);
}

function doGet(e) {
  try {
    const ns = getNs_(e);
    const debtorsSh = getOrCreateSheet_(sheetName_(DEBTORS_SHEET, ns), DEBTOR_COLS, DEBTOR_HEADERS);
    const paymentsSh = getOrCreateSheet_(sheetName_(PAYMENTS_SHEET, ns), PAYMENT_COLS, PAYMENT_HEADERS);
    const metaSh = getOrCreateSheet_(sheetName_(META_SHEET, ns), META_COLS, META_HEADERS);
    ensureExtraSheets_();

    const debtorRows = readSheet_(debtorsSh, DEBTOR_COLS);
    const paymentRows = readSheet_(paymentsSh, PAYMENT_COLS);
    const metaRows = readSheet_(metaSh, META_COLS);

    // 沒有任何借款人時，回傳 data:null 讓前端判斷「雲端為空」
    if (!debtorRows.length) {
      return jsonOut_({ ok: true, data: null });
    }

    // 依 debtor_id 分組付款
    const paymentsById = {};
    for (const p of paymentRows) {
      const did = String(p.debtor_id || '');
      if (!did) continue;
      if (!paymentsById[did]) paymentsById[did] = [];
      paymentsById[did].push({
        id: strOrEmpty_(p.id),
        date: strOrEmpty_(p.date),
        principal: numOrZero_(p.principal),
        interest: numOrZero_(p.interest),
        note: strOrEmpty_(p.note),
      });
    }

    const debtors = debtorRows.map(d => ({
      id: strOrEmpty_(d.id),
      name: strOrEmpty_(d.name),
      day: numOrZero_(d.day),
      amount: numOrZero_(d.amount),
      principal: numOrZero_(d.principal),
      interest: numOrZero_(d.interest),
      phone: strOrEmpty_(d.phone),
      notes: strOrEmpty_(d.notes),
      createdAt: strOrEmpty_(d.createdAt),
      payments: paymentsById[strOrEmpty_(d.id)] || [],
    }));

    const state = { debtors: debtors };
    for (const m of metaRows) {
      const k = strOrEmpty_(m.key);
      if (!k) continue;
      let v = m.value;
      // 嘗試還原 boolean/number
      if (v === 'true') v = true;
      else if (v === 'false') v = false;
      state[k] = v;
    }

    return jsonOut_({ ok: true, data: state });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err), stack: String(err.stack || '') });
  }
}

function doPost(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : '';
    if (action === 'appendPayment') {
      return doAppendPayment_(e);
    }

    const body = e.postData && e.postData.contents;
    if (!body) return jsonOut_({ ok: false, error: 'empty body' });
    const state = JSON.parse(body);
    if (!state || typeof state !== 'object') return jsonOut_({ ok: false, error: 'invalid json' });
    const debtors = Array.isArray(state.debtors) ? state.debtors : [];

    const ns = getNs_(e);
    const debtorsSh = getOrCreateSheet_(sheetName_(DEBTORS_SHEET, ns), DEBTOR_COLS, DEBTOR_HEADERS);
    const paymentsSh = getOrCreateSheet_(sheetName_(PAYMENTS_SHEET, ns), PAYMENT_COLS, PAYMENT_HEADERS);
    const metaSh = getOrCreateSheet_(sheetName_(META_SHEET, ns), META_COLS, META_HEADERS);
    ensureExtraSheets_();

    // Debtors
    const debtorRows = debtors.map(d => ({
      id: strOrEmpty_(d.id),
      name: strOrEmpty_(d.name),
      day: numOrZero_(d.day),
      amount: numOrZero_(d.amount),
      principal: numOrZero_(d.principal),
      interest: numOrZero_(d.interest),
      phone: strOrEmpty_(d.phone),
      notes: strOrEmpty_(d.notes),
      createdAt: strOrEmpty_(d.createdAt),
    }));
    writeSheet_(debtorsSh, DEBTOR_COLS, debtorRows);

    // Payments（展平）
    const paymentRows = [];
    for (const d of debtors) {
      const did = strOrEmpty_(d.id);
      const ps = Array.isArray(d.payments) ? d.payments : [];
      for (const p of ps) {
        paymentRows.push({
          debtor_id: did,
          id: strOrEmpty_(p.id),
          date: strOrEmpty_(p.date),
          principal: numOrZero_(p.principal),
          interest: numOrZero_(p.interest),
          note: strOrEmpty_(p.note),
        });
      }
    }
    writeSheet_(paymentsSh, PAYMENT_COLS, paymentRows);

    // Meta（state 中 debtors 以外的欄位）
    const metaRows = [];
    for (const k of Object.keys(state)) {
      if (k === 'debtors') continue;
      const v = state[k];
      if (v === null || v === undefined) continue;
      if (typeof v === 'object') continue; // 跳過巢狀物件
      metaRows.push({ key: k, value: String(v) });
    }
    writeSheet_(metaSh, META_COLS, metaRows);

    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err), stack: String(err.stack || '') });
  }
}

/**
 * 只 append 一筆付款，不覆蓋整份資料。
 *
 * 呼叫方式：
 *   POST <exec>?action=appendPayment[&ns=xxx]
 *   Content-Type: application/json
 *   Body: {
 *     "debtor_id": "d_xxx",
 *     "payment": {
 *       "id": "tg_<gmail_messageId>",   // 冪等鍵：相同 id 不會重複寫
 *       "date": "2026-04-27",
 *       "principal": 0,
 *       "interest": 60000,
 *       "note": "自動同步（玉山）匯款人：王*哲"
 *     }
 *   }
 *
 * 回應：
 *   { ok: true, appended: true,  debtor_id, payment_id }   // 寫入成功
 *   { ok: true, appended: false, reason: 'duplicate id' }  // 同 id 已存在，跳過
 *   { ok: false, error: '...' }                            // 失敗
 */
function doAppendPayment_(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'lock timeout' });
  }
  try {
    const body = e.postData && e.postData.contents;
    if (!body) return jsonOut_({ ok: false, error: 'empty body' });
    const req = JSON.parse(body);
    if (!req || typeof req !== 'object') return jsonOut_({ ok: false, error: 'invalid json' });

    const debtorId = strOrEmpty_(req.debtor_id);
    const payment = req.payment;
    if (!debtorId) return jsonOut_({ ok: false, error: 'missing debtor_id' });
    if (!payment || typeof payment !== 'object') return jsonOut_({ ok: false, error: 'missing payment' });

    const ns = getNs_(e);
    const debtorsSh = getOrCreateSheet_(sheetName_(DEBTORS_SHEET, ns), DEBTOR_COLS, DEBTOR_HEADERS);
    const paymentsSh = getOrCreateSheet_(sheetName_(PAYMENTS_SHEET, ns), PAYMENT_COLS, PAYMENT_HEADERS);

    // 確認 debtor 存在
    const debtors = readSheet_(debtorsSh, DEBTOR_COLS);
    const exists = debtors.some(d => strOrEmpty_(d.id) === debtorId);
    if (!exists) return jsonOut_({ ok: false, error: 'debtor not found: ' + debtorId });

    // 冪等檢查：若同 debtor 下已有相同 payment.id，直接回 appended:false
    const paymentId = strOrEmpty_(payment.id);
    if (paymentId) {
      const all = readSheet_(paymentsSh, PAYMENT_COLS);
      const dup = all.some(p =>
        strOrEmpty_(p.debtor_id) === debtorId &&
        strOrEmpty_(p.id) === paymentId
      );
      if (dup) {
        return jsonOut_({
          ok: true, appended: false, reason: 'duplicate id',
          debtor_id: debtorId, payment_id: paymentId
        });
      }
    }

    // append 一列（不清空舊資料）
    const row = [
      debtorId,
      paymentId,
      strOrEmpty_(payment.date),
      numOrZero_(payment.principal),
      numOrZero_(payment.interest),
      strOrEmpty_(payment.note),
    ];
    paymentsSh.appendRow(row);

    return jsonOut_({
      ok: true, appended: true,
      debtor_id: debtorId, payment_id: paymentId
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err), stack: String(err.stack || '') });
  } finally {
    lock.releaseLock();
  }
}
