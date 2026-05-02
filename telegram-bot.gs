/**
 * 借貸還款通知 v3.0 — 自動推播 + 自動登記
 *
 * 部署位置：太太 Google 帳號的 Apps Script（因為要讀她的 Gmail）
 * 取代之前的 v1.0 / v2.0。
 *
 * 架構：
 *   ┌─ 太太帳號（這份檔案）─────────────────┐
 *   │  讀玉山入帳信 → 解析 → 比對借款人      │
 *   │       ↓                                │
 *   │  HTTP POST 你後端的 appendPayment      │
 *   │       ↓                                │
 *   │  推 Telegram（已登記 / 待手動）         │
 *   └───────────────────────────────────────┘
 *   ┌─ 你帳號（loan-ledger 後端）────────────┐
 *   │  appendPayment 端點 → 寫進 Sheet      │
 *   └───────────────────────────────────────┘
 *
 *   太太帳號只授權 Gmail + UrlFetch（不需要 Sheet 權限）。
 *   未來要改規則 / 加借款人 → 都在你帳號改，不用再登她帳號。
 *
 * 部署步驟（請用太太帳號登入）：
 *   1. 開啟既有的「借貸還款通知」Apps Script
 *   2. 把 程式碼.gs 整份內容刪掉，把本檔貼進去
 *   3. 確認下方 LOAN_LEDGER_URL 是你後端的 /exec 網址（不含 ?action）
 *   4. 確認 Script Properties 已有 TELEGRAM_TOKEN / TELEGRAM_CHAT_ID（v1.0 已設過就不用再設）
 *   5. 儲存 → 執行 testTelegram() 確認 Telegram 通
 *   6. 執行 testMatching() 確認比對邏輯沒問題
 *   7. 執行 testParseBody() 確認解析最近一封玉山信沒問題
 *   8. 觸發條件：保留原本「每小時跑 checkBankEmails」的 trigger 即可
 *
 * 比對邏輯：
 *   1. 從備註抓「漢字 + 遮罩(O/o/〇/*) + 漢字」這個 3 字模式
 *   2. 取第一字 + 第三字組成 regex「first.third」
 *   3. 對借款人姓名做 regex 比對（從你後端動態讀，不寫死）
 *   4. 多筆候選時用金額分（金額相等 或 金額是月應收的整數倍 → 預收）
 *   5. 全比對不到 → 推 Telegram 等你手動處理
 *
 * 範例：
 *   備註「葉O姐」+ 70000 → 五股葉小姐(長期)
 *   備註「葉O姐」+ 40000 → 五股葉小姐(短期)
 *   備註「王O哲」+ 60000 → 台中王浩哲
 *   備註「台中王O哲」+ 60000 → 台中王浩哲（多餘字 OK）
 *   備註「陳O生」+ 120000 → 三重陳先生（120000 = 60000×2，預收 2 期）
 */

// ============ 你要改的設定 ============
// 你 loan-ledger 後端的 /exec 網址（從「部署 → 管理部署作業」複製）
const LOAN_LEDGER_URL = 'https://script.google.com/macros/s/AKfycbxauwz6lSMqJPb8sqyKjGho_mgZGtcGwcrLsTeiyuhWkQ8pwO1mMtrTXGx_R6Pd37Vr/exec';

// ============ 主程式（每小時跑）============
function checkBankEmails() {
  const props = PropertiesService.getScriptProperties();
  const TOKEN = props.getProperty('TELEGRAM_TOKEN');
  const CHAT_ID = props.getProperty('TELEGRAM_CHAT_ID');
  if (!TOKEN || !CHAT_ID) {
    Logger.log('❌ 尚未設定 TELEGRAM_TOKEN 或 TELEGRAM_CHAT_ID');
    return;
  }

  const query = 'from:Service@info.esunbank.com -label:已通知 newer_than:2d';
  const threads = GmailApp.search(query);
  Logger.log(`找到 ${threads.length} 封玉山通知信`);
  if (threads.length === 0) return;

  // 從你後端動態抓借款人清單（不寫死，未來新增/改名都自動跟上）
  const debtors = loadDebtors();
  if (!debtors) {
    Logger.log('❌ 無法從後端讀取借款人清單，停止本次處理');
    return;
  }
  Logger.log(`讀到 ${debtors.length} 位借款人`);

  let label = GmailApp.getUserLabelByName('已通知');
  if (!label) label = GmailApp.createLabel('已通知');

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      try {
        const messageId = message.getId();
        const body = message.getPlainBody();
        const txns = parseEsunMessages(body);
        if (txns.length === 0) {
          Logger.log(`- 無交易明細：${message.getSubject()}`);
          continue;
        }
        for (let i = 0; i < txns.length; i++) {
          const txn = txns[i];
          if (txn.type !== '轉入') {
            Logger.log(`- 略過轉出：${txn.amount}`);
            continue;
          }
          const paymentId = `tg_${messageId}_${i}`;
          const matched = matchDebtor(debtors, txn);
          if (matched) {
            const result = appendPayment(matched, txn, paymentId);
            if (result === 'appended') {
              sendTelegram(TOKEN, CHAT_ID, formatAutoLogged(matched, txn));
              Logger.log(`✓ 已登記：${matched.name} ${txn.amount}`);
            } else if (result === 'duplicate') {
              Logger.log(`- 重複跳過：${matched.name} ${paymentId}`);
              // 重複時不再推 Telegram，避免騷擾
            } else {
              sendTelegram(TOKEN, CHAT_ID, formatWriteFailed(matched, txn));
              Logger.log(`⚠️ 寫入失敗：${matched.name}`);
            }
          } else {
            sendTelegram(TOKEN, CHAT_ID, formatNeedManual(txn));
            Logger.log(`⚠️ 待手動：金額 ${txn.amount} 備註「${txn.remark}」`);
          }
        }
      } catch (e) {
        Logger.log(`❌ 處理訊息失敗：${e}\n${e.stack || ''}`);
      }
    }
    thread.addLabel(label);
  }
}

// ============ 解析玉山信件（一封多筆）============
// 每行格式：帳務異動 帳務異動-轉入 0842979***334 2026/04/27 交易金額：NTD 60,000 王O哲
function parseEsunMessages(body) {
  const lines = String(body || '').split(/\r?\n/);
  const rowRegex = /^帳務異動\s+帳務異動-(轉入|轉出)\s+(\S+)\s+(\d{4}\/\d{1,2}\/\d{1,2})\s+交易金額[：:]\s*NTD\s+([\d,]+)\s+(.+?)$/;
  const txns = [];
  for (const line of lines) {
    const trimmed = line.trim();
    const m = trimmed.match(rowRegex);
    if (!m) continue;
    txns.push({
      type: m[1],                                      // '轉入' 或 '轉出'
      account: m[2],                                   // 0842979***334
      date: m[3],                                      // 2026/04/27
      amount: parseInt(m[4].replace(/,/g, ''), 10),    // 60000
      remark: m[5].trim(),                             // 王O哲
    });
  }
  return txns;
}

// ============ 比對借款人 ============
// 演算法：
//   1. 從備註找「漢字 + 遮罩 + 漢字」3 字模式 → 取第一字+第三字
//   2. 用 first.third regex 比對借款人姓名
//   3. 比對不到 → 後備：看備註是否含借款人姓名 2 字以上連續片段
//   4. 候選多筆 → 用金額過濾（剛好等於月應收 或 月應收的整數倍 = 預收 N 期）
function matchDebtor(debtors, txn) {
  const remark = String(txn.remark || '');
  const amount = Number(txn.amount) || 0;

  // 第一步：找遮罩模式
  // 遮罩字元：O o 〇 ○ ● * ＊ Ｏ · ・
  const maskedRegex = /([一-龥])[Oo〇○●*＊Ｏ·・]([一-龥])/;
  const m = remark.match(maskedRegex);
  let candidates = [];
  if (m) {
    const first = m[1];
    const third = m[2];
    const nameRegex = new RegExp(first + '.' + third);
    candidates = debtors.filter(d => nameRegex.test(String(d.name || '')));
  }

  // 第二步：遮罩比對沒命中 → 後備（直接看備註是否含姓名 2 字以上片段）
  if (candidates.length === 0) {
    candidates = debtors.filter(d => {
      const name = String(d.name || '').replace(/[（()【】\s（）]/g, '');
      if (name.length < 2) return false;
      for (let i = 0; i + 2 <= name.length; i++) {
        const sub = name.substring(i, i + 2);
        if (remark.indexOf(sub) !== -1) return true;
      }
      return false;
    });
  }

  if (candidates.length === 0) return null;

  // 用金額過濾
  const byAmount = candidates.filter(d => {
    const a = Number(d.amount) || 0;
    if (a === 0) return false;
    return amount === a || (amount > 0 && amount % a === 0);
  });

  if (byAmount.length === 1) return byAmount[0];
  if (byAmount.length === 0) return null; // 名字對但金額對不起來
  return null;                            // 多筆無法分辨
}

// ============ 從你後端抓借款人清單 ============
function loadDebtors() {
  try {
    const res = UrlFetchApp.fetch(LOAN_LEDGER_URL, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      Logger.log('loadDebtors HTTP ' + res.getResponseCode() + ': ' + res.getContentText().substring(0, 300));
      return null;
    }
    const json = JSON.parse(res.getContentText());
    if (!json.ok || !json.data || !Array.isArray(json.data.debtors)) {
      Logger.log('loadDebtors 回應異常：' + res.getContentText().substring(0, 300));
      return null;
    }
    return json.data.debtors;
  } catch (e) {
    Logger.log('loadDebtors 例外：' + e);
    return null;
  }
}

// ============ 呼叫後端 appendPayment ============
// 回傳：'appended' | 'duplicate' | 'failed'
function appendPayment(debtor, txn, paymentId) {
  try {
    const dateIso = toIsoDate(txn.date);
    const payload = {
      debtor_id: String(debtor.id),
      payment: {
        id: paymentId,
        date: dateIso,
        principal: 0,
        interest: Number(txn.amount) || 0,
        note: `🤖自動登記 玉山入帳 備註：${txn.remark}`,
      },
    };
    const res = UrlFetchApp.fetch(LOAN_LEDGER_URL + '?action=appendPayment', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      Logger.log('appendPayment HTTP ' + res.getResponseCode() + ': ' + res.getContentText());
      return 'failed';
    }
    const json = JSON.parse(res.getContentText());
    if (!json.ok) {
      Logger.log('appendPayment 後端回 fail：' + res.getContentText());
      return 'failed';
    }
    return json.appended ? 'appended' : 'duplicate';
  } catch (e) {
    Logger.log('appendPayment 例外：' + e);
    return 'failed';
  }
}

// ============ 工具：YYYY/M/D → YYYY-MM-DD ============
function toIsoDate(yyyySlashMmDd) {
  const m = String(yyyySlashMmDd || '').match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (!m) return '';
  return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
}

// ============ Telegram 訊息格式 ============
function formatAutoLogged(debtor, txn) {
  return `✅ 已自動登記\n\n` +
    `👤 借款人：${debtor.name}\n` +
    `💵 金額：NTD ${fmt(txn.amount)}\n` +
    `📅 日期：${txn.date}\n` +
    `📝 備註：${txn.remark}\n\n` +
    `📒 已寫入借貸帳簿`;
}
function formatNeedManual(txn) {
  return `⚠️ 待手動確認\n\n` +
    `💵 金額：NTD ${fmt(txn.amount)}\n` +
    `📅 日期：${txn.date}\n` +
    `📝 備註：${txn.remark}\n` +
    `🏦 帳號：${txn.account}\n\n` +
    `❓ 找不到對應借款人\n👉 請打開借貸帳簿手動處理`;
}
function formatWriteFailed(debtor, txn) {
  return `⚠️ 寫入帳簿失敗\n\n` +
    `已比對到：${debtor.name}\n` +
    `金額：NTD ${fmt(txn.amount)}\n` +
    `日期：${txn.date}\n\n` +
    `請手動到網頁登記，並檢查 Apps Script 執行記錄`;
}
function fmt(n) {
  if (typeof n !== 'number') return String(n || '');
  return n.toLocaleString('zh-TW');
}

// ============ Telegram 發送 ============
function sendTelegram(token, chatId, text) {
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: text }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    Logger.log('Telegram 發送失敗：' + res.getContentText());
  }
}

// ============ 測試函式 ============

// 手動跑一次完整流程
function testRun() { checkBankEmails(); }

// 測 Telegram 連線（不掃信、不寫帳簿）
function testTelegram() {
  const props = PropertiesService.getScriptProperties();
  const t = props.getProperty('TELEGRAM_TOKEN');
  const c = props.getProperty('TELEGRAM_CHAT_ID');
  if (!t || !c) { Logger.log('❌ TELEGRAM_TOKEN 或 TELEGRAM_CHAT_ID 未設定'); return; }
  sendTelegram(t, c, '🧪 v3.0 連線測試 — ' + new Date().toLocaleString('zh-TW'));
  Logger.log('✓ 已發送測試訊息');
}

// 測比對邏輯（會 GET 後端拿借款人清單，不寫帳簿、不發 Telegram）
function testMatching() {
  const debtors = loadDebtors();
  if (!debtors) { Logger.log('❌ 無法讀取借款人清單'); return; }
  Logger.log(`借款人清單：${debtors.map(d => d.name + '/' + d.amount).join('、')}`);

  const samples = [
    { type: '轉入', amount: 70000,  remark: '葉O姐',    date: '2026/04/24' }, // → 五股葉小姐(長期)
    { type: '轉入', amount: 40000,  remark: '葉O姐',    date: '2026/04/24' }, // → 五股葉小姐(短期)
    { type: '轉入', amount: 60000,  remark: '王O哲',    date: '2026/04/05' }, // → 台中王浩哲
    { type: '轉入', amount: 60000,  remark: '台中王O哲', date: '2026/04/05' }, // → 台中王浩哲
    { type: '轉入', amount: 60000,  remark: '周O生',    date: '2026/04/23' }, // → 板橋周先生
    { type: '轉入', amount: 60000,  remark: '陳O生',    date: '2026/04/28' }, // → 三重陳先生
    { type: '轉入', amount: 120000, remark: '陳O生',    date: '2026/04/29' }, // → 三重陳先生（預收 2 期）
    { type: '轉入', amount: 359,    remark: 'N391747302', date: '2026/04/27' }, // → 無法比對
  ];
  for (const s of samples) {
    const m = matchDebtor(debtors, s);
    Logger.log(`金額 ${s.amount} 備註「${s.remark}」→ ${m ? '✓ ' + m.name + ' [' + m.id + ']' : '✗ 無法比對'}`);
  }
}

// 測解析最近一封玉山信（不寫帳簿、不發 Telegram）
function testParseBody() {
  const threads = GmailApp.search('from:Service@info.esunbank.com newer_than:14d');
  if (!threads.length) { Logger.log('❌ 沒找到玉山通知信'); return; }
  const msg = threads[0].getMessages()[0];
  Logger.log('Subject: ' + msg.getSubject());
  const body = msg.getPlainBody();
  const txns = parseEsunMessages(body);
  Logger.log(`抓到 ${txns.length} 筆交易`);
  for (const t of txns) Logger.log(JSON.stringify(t));
}
