/**
 * Gmail → loan-ledger 每日自動還款同步
 *
 * 架構：本檔是「獨立 Apps Script 專案」，部署在「收款信箱所在的 Google 帳號」下，
 *       透過 UrlFetchApp 呼叫 loan-ledger Web App 的 doGet/doPost 寫回 Sheets。
 *
 * 部署步驟（請用收信的 Google 帳號，例如太太的帳號登入）：
 *   1. 到 https://script.google.com/，新增空白專案（名稱自訂，例如「還款信件自動同步」）
 *   2. 把本檔全部貼入，儲存
 *   3. 填入下方 CONFIG.LOAN_LEDGER_URL（至 loan-ledger Apps Script 編輯器點「部署 → 管理部署作業」複製 /exec 網址）
 *   4. 手動執行一次 syncRepaymentsDaily，同意 GmailApp + UrlFetchApp 權限
 *   5. 觸發條件 → 新增：函式 syncRepaymentsDaily、時間驅動、日計時器、午夜~凌晨 1 點
 *   6. 專案時區確認為 (GMT+08:00) Taipei
 *
 * 日常調整：若某家銀行的信抓不到或姓名抓錯，改下方 BANK_PARSERS 的正則即可。
 */

const CONFIG = {
  LOAN_LEDGER_URL: '',           // 例：https://script.google.com/macros/s/XXXXXX/exec
  LOAN_LEDGER_NS: '',            // loan-ledger 若有用命名空間就填，否則留空
  DONE_LABEL: '已同步還款',       // 成功同步後的 Gmail label
  PENDING_LABEL: '還款待確認',    // 姓名對不上的 Gmail label
  LOOKBACK_HOURS: 24,
  TIMEZONE: 'Asia/Taipei',
};

/**
 * 每家銀行一組解析設定。senderQuery 用於 Gmail 搜尋；amountRegex/payerRegex
 * 用於從主旨＋內文抽資料。請依實際信件樣本調整——尤其 payerRegex，
 * 匯款人／備註的欄位名稱各行不同。
 */
const BANK_PARSERS = {
  cathay: {
    senderQuery: 'from:cathaybk.com.tw',
    subjectKeywords: ['跨行匯入', '活存入帳', '轉入通知', '入帳'],
    amountRegex: /(?:新台幣|NT\$|TWD)\s*([\d,]+)\s*元?/,
    payerRegex: /(?:匯款人|轉帳人|備註)[:：]?\s*([一-龥A-Za-z0-9]+)/,
  },
  esun: {
    senderQuery: 'from:esunbank.com.tw',
    subjectKeywords: ['入帳通知', '轉帳通知', '匯入'],
    amountRegex: /(?:新台幣|NT\$|TWD)\s*([\d,]+)\s*元?/,
    payerRegex: /(?:匯款人|轉帳人|備註)[:：]?\s*([一-龥A-Za-z0-9]+)/,
  },
  ctbc: {
    senderQuery: 'from:ctbcbank.com',
    subjectKeywords: ['交易通知', '入帳通知', '匯入'],
    amountRegex: /(?:新台幣|NT\$|TWD)\s*([\d,]+)\s*元?/,
    payerRegex: /(?:匯款人|轉帳人|備註)[:：]?\s*([一-龥A-Za-z0-9]+)/,
  },
};

function syncRepaymentsDaily() {
  const doneLabel = getOrCreateLabel_(CONFIG.DONE_LABEL);
  const pendingLabel = getOrCreateLabel_(CONFIG.PENDING_LABEL);

  const state = fetchLedgerState_();
  if (!state || !Array.isArray(state.debtors) || state.debtors.length === 0) {
    Logger.log('loan-ledger 目前無借款人，結束');
    return;
  }

  const threads = fetchBankNotices_(CONFIG.LOOKBACK_HOURS);
  Logger.log('找到 ' + threads.length + ' 個符合條件的 thread');

  const threadsDone = [];
  const threadsPending = [];
  let applied = 0, pending = 0, skipped = 0;

  for (const thread of threads) {
    let didApply = false, didDefer = false;
    for (const msg of thread.getMessages()) {
      const notice = parseNotice_(msg);
      if (!notice || !notice.amount) continue;

      const paymentId = 'gmail_' + notice.messageId;
      if (hasPaymentId_(state, paymentId)) {
        // 已經寫過：代表上次 push 成功但 label 失敗；這次只補 label
        didApply = true;
        continue;
      }

      const debtorId = matchDebtor_(state, notice.payerName);
      if (!debtorId) {
        didDefer = true;
        continue;
      }

      appendPayment_(state, debtorId, {
        id: paymentId,
        date: notice.date,
        principal: notice.amount,
        interest: 0,
        note: '自動同步（' + notice.bank + '）匯款人：' + notice.payerName + '；本利未拆分',
      });
      didApply = true;
    }

    if (didApply) { threadsDone.push(thread); applied++; }
    else if (didDefer) { threadsPending.push(thread); pending++; }
    else { skipped++; }
  }

  if (threadsDone.length > 0) {
    pushLedgerState_(state);
    for (const t of threadsDone) t.addLabel(doneLabel);
  }
  for (const t of threadsPending) t.addLabel(pendingLabel);

  Logger.log('完成：已套用 ' + applied + '、待確認 ' + pending + '、略過 ' + skipped);
}

function fetchBankNotices_(lookbackHours) {
  const days = Math.max(1, Math.ceil(lookbackHours / 24));
  const senders = Object.values(BANK_PARSERS).map(p => '(' + p.senderQuery + ')').join(' OR ');
  const q = 'newer_than:' + days + 'd (' + senders + ')'
    + ' -label:"' + CONFIG.DONE_LABEL + '"'
    + ' -label:"' + CONFIG.PENDING_LABEL + '"';
  return GmailApp.search(q);
}

function parseNotice_(message) {
  const from = message.getFrom();
  const subject = message.getSubject();

  let bankKey = null, parser = null;
  for (const key of Object.keys(BANK_PARSERS)) {
    const p = BANK_PARSERS[key];
    const domain = p.senderQuery.replace(/^from:/, '');
    if (from.indexOf(domain) === -1) continue;
    if (!p.subjectKeywords.some(kw => subject.indexOf(kw) !== -1)) continue;
    bankKey = key; parser = p; break;
  }
  if (!parser) return null;

  const text = subject + '\n' + message.getPlainBody();
  const amountMatch = text.match(parser.amountRegex);
  const amount = amountMatch ? Number(amountMatch[1].replace(/,/g, '')) : 0;
  const payerMatch = text.match(parser.payerRegex);
  const payerName = payerMatch ? payerMatch[1].trim() : '';
  const date = Utilities.formatDate(message.getDate(), CONFIG.TIMEZONE, 'yyyy-MM-dd');

  return { bank: bankKey, amount, payerName, date, messageId: message.getId() };
}

function matchDebtor_(state, payerName) {
  if (!payerName) return null;
  for (const d of state.debtors) {
    const name = String(d.name || '').trim();
    if (!name) continue;
    // 雙向 includes：匯款人可能是「王小明」、或「王小明 XX 公司」；反向也允許借款人名字是匯款人的子字串
    if (payerName.indexOf(name) !== -1 || name.indexOf(payerName) !== -1) {
      return String(d.id);
    }
  }
  return null;
}

function appendPayment_(state, debtorId, payment) {
  for (const d of state.debtors) {
    if (String(d.id) !== String(debtorId)) continue;
    if (!Array.isArray(d.payments)) d.payments = [];
    d.payments.push(payment);
    return;
  }
}

function hasPaymentId_(state, paymentId) {
  for (const d of state.debtors) {
    if (!Array.isArray(d.payments)) continue;
    if (d.payments.some(p => String(p.id) === String(paymentId))) return true;
  }
  return false;
}

function fetchLedgerState_() {
  const resp = UrlFetchApp.fetch(buildLedgerUrl_(), { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code !== 200) throw new Error('fetchLedgerState_ HTTP ' + code + ': ' + resp.getContentText());
  const body = JSON.parse(resp.getContentText());
  if (!body || body.ok !== true) throw new Error('fetchLedgerState_ payload error: ' + resp.getContentText());
  return body.data || { debtors: [] };
}

function pushLedgerState_(state) {
  const resp = UrlFetchApp.fetch(buildLedgerUrl_(), {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(state),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  if (code !== 200) throw new Error('pushLedgerState_ HTTP ' + code + ': ' + resp.getContentText());
  const body = JSON.parse(resp.getContentText());
  if (!body || body.ok !== true) throw new Error('pushLedgerState_ payload error: ' + resp.getContentText());
}

function buildLedgerUrl_() {
  if (!CONFIG.LOAN_LEDGER_URL || CONFIG.LOAN_LEDGER_URL.indexOf('script.google.com') === -1) {
    throw new Error('請先在 CONFIG.LOAN_LEDGER_URL 填入 loan-ledger 的 Web App /exec 網址');
  }
  return CONFIG.LOAN_LEDGER_URL + (CONFIG.LOAN_LEDGER_NS ? '?ns=' + encodeURIComponent(CONFIG.LOAN_LEDGER_NS) : '');
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}
