/**
 * RASTA HOUSE — Google Apps Script backend
 * Demo-credit simulation only. This does not verify or process real payments.
 *
 * Required Script Properties:
 *   TELEGRAM_BOT_TOKEN       Telegram bot token
 *   TELEGRAM_CHAT_ID         Admin chat ID (defaults to CFG.CHAT_ID)
 *   TELEGRAM_ADMIN_USER_ID   Admin Telegram user ID; for a private bot chat
 *                            it may be left blank and the chat ID is used.
 *   SPREADSHEET_ID           Saved automatically by setup() when run
 *
 * Deploy as Web App: Execute as Me; Who has access: Anyone.
 */
const CFG = {
  CHAT_ID: '9268792685',
  USERS: 'Users',
  BALANCES: 'DemoBalances',
  REQUESTS: 'Requests',
  LEDGER: 'BalanceLedger',
  MIN_DEPOSIT: 150,
  MAX_DEPOSIT: 10000,
  MIN_WITHDRAW: 300,
  MAX_WITHDRAW: 10000
};

const HEADERS = {
  Users: ['PlayerID','FullName','Phone','PasswordHash','PasswordSalt','JoinedAt'],
  DemoBalances: ['Phone','FullName','Balance','UpdatedAt'],
  Requests: ['RequestID','CreatedAt','Type','Phone','FullName','Amount','Details','TransactionID','Status','ProcessedAt','AdminNote'],
  BalanceLedger: ['RequestID','Phone','Decision','Type','Amount','BeforeBalance','AfterBalance','LedgerStatus','CreatedAt','AppliedAt']
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('RASTA ADMIN')
    .addItem('Setup / Repair Sheets', 'setup')
    .addItem('Configure Telegram Bot', 'configureTelegram')
    .addItem('Connect Telegram Approval Buttons', 'configureTelegramWebhook')
    .addItem('Check Telegram Webhook', 'checkTelegramWebhook')
    .addItem('Test Telegram', 'testTelegram')
    .addSeparator()
    .addItem('Approve Selected Request', 'approveSelectedRequest')
    .addItem('Reject Selected Request', 'rejectSelectedRequest')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the bound Rasta House spreadsheet and run Setup / Repair Sheets there.');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  Object.keys(HEADERS).forEach(function(name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = HEADERS[name];
    headers.forEach(function(h, i) {
      if (sh.getRange(1, i + 1).getValue() !== h) sh.getRange(1, i + 1).setValue(h);
    });
    sh.setFrozenRows(1);
  });
  SpreadsheetApp.getUi().alert('RASTA HOUSE setup complete. Spreadsheet connection saved. Users, DemoBalances, Requests and BalanceLedger are ready.');
}

function configureTelegram() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const result = ui.prompt('Configure Telegram', 'Paste your Telegram bot token. Leave blank to keep the saved token.', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const token = result.getResponseText().trim();
  if (token) props.setProperty('TELEGRAM_BOT_TOKEN', token);
  if (!getBotToken_()) {
    ui.alert('No Telegram token saved. Run Configure Telegram again.');
    return;
  }
  if (!props.getProperty('TELEGRAM_CHAT_ID')) props.setProperty('TELEGRAM_CHAT_ID', CFG.CHAT_ID);
  ui.alert('Telegram token saved. The configured chat ID is ' + (props.getProperty('TELEGRAM_CHAT_ID') || CFG.CHAT_ID) + '.');
}

function testTelegram() {
  const result = sendTelegram_('RASTA HOUSE TEST\nTelegram connection is working.');
  SpreadsheetApp.getUi().alert(result ? 'Telegram test sent.' : 'Telegram test failed. Check the bot token, chat ID and Apps Script execution log.');
}

function checkTelegramWebhook() {
  const token = getBotToken_();
  if (!token) throw new Error('Telegram bot token is missing. Use RASTA ADMIN > Configure Telegram Bot.');
  const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getWebhookInfo', { muteHttpExceptions: true });
  const result = JSON.parse(response.getContentText());
  if (!result.ok) {
    SpreadsheetApp.getUi().alert('Telegram webhook check failed: ' + (result.description || 'Unknown error'));
    return result;
  }
  const info = result.result || {};
  const message = 'Webhook URL: ' + (info.url || '(not set)') +
    '\nPending updates: ' + (info.pending_update_count || 0) +
    '\nLast error: ' + (info.last_error_message || 'None') +
    '\nLast error date: ' + (info.last_error_date ? new Date(info.last_error_date * 1000).toLocaleString() : 'None');
  SpreadsheetApp.getUi().alert(message);
  return info;
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = String(p.action || '').trim();
    if (action === 'register') return json_(registerPlayer_(p));
    if (action === 'login') return json_(loginPlayer_(p));
    if (action === 'getBalance') return json_(getBalance_(p.phone));
    if (action === 'getHistory') return json_(getHistory_(p.phone));
    if (action === 'ping') return json_({ status: 'success', message: 'RASTA backend online' });
    return json_({ status: 'error', message: 'Unknown action.' });
  } catch (err) {
    console.error('doGet: ' + String(err && err.stack || err));
    return json_({ status: 'error', message: String(err.message || err) });
  }
}

function doPost(e) {
  let body = {};
  try {
    if (e && e.postData && e.postData.contents &&
        String(e.postData.type || '').toLowerCase().indexOf('application/json') >= 0) {
      body = JSON.parse(e.postData.contents);
    }
  } catch (parseError) {
    console.error('Could not parse JSON POST: ' + parseError);
  }
  if (body && body.callback_query) return handleTelegramCallback_(body.callback_query);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = (e && e.parameter) || {};
    const action = String(p.action || '').trim().toLowerCase();
    if (action !== 'deposit' && action !== 'withdraw') {
      return json_({ status: 'error', message: 'Unknown transaction action.' });
    }
    const phone = cleanPhone_(p.phone);
    const user = findUserByPhone_(phone);
    if (!user) return json_({ status: 'error', message: 'Player is not registered. Please register or log in again.' });

    const amount = Number(p.amount);
    const min = action === 'deposit' ? CFG.MIN_DEPOSIT : CFG.MIN_WITHDRAW;
    if (!Number.isFinite(amount) || Math.floor(amount) !== amount || amount < min || amount > CFG.MAX_DEPOSIT) {
      return json_({ status: 'error', message: 'Amount is outside the allowed range.' });
    }

    const details = String(p.details || '').trim().slice(0, 1000);
    const transactionId = String(p.transactionId || extractTransactionId_(details) || '').trim().slice(0, 160);
    if (action === 'deposit' && !transactionId) {
      return json_({ status: 'error', message: 'Payment transaction ID is required.' });
    }
    if (action === 'deposit' && transactionIdExists_(transactionId)) {
      return json_({ status: 'error', message: 'This payment transaction ID has already been submitted.' });
    }

    const balanceRow = ensureBalanceRow_(user);
    const balanceSheet = sheet_(CFG.BALANCES);
    const currentBalance = Number(balanceSheet.getRange(balanceRow, 3).getValue()) || 0;
    if (action === 'withdraw' && currentBalance < amount) {
      return json_({ status: 'error', message: 'Insufficient demo-credit balance.' });
    }

    const reqSheet = sheet_(CFG.REQUESTS);
    const requestId = 'RH-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const now = new Date();
    reqSheet.appendRow([requestId, now, action.toUpperCase(), phone, user.name, amount, details, transactionId, 'PENDING', '', '']);

    if (action === 'withdraw') {
      balanceSheet.getRange(balanceRow, 3).setValue(currentBalance - amount);
      balanceSheet.getRange(balanceRow, 4).setValue(now);
    }

    const message = (action === 'deposit' ? '🟡 NEW DEPOSIT REQUEST' : '🟠 NEW WITHDRAWAL REQUEST') +
      '\nRequest: ' + requestId + '\nPlayer: ' + user.name + '\nPhone: ' + phone +
      '\nAmount: ' + amount + ' demo credits\nDetails: ' + details +
      (transactionId ? '\nTransaction ID: ' + transactionId : '');
    const notified = sendTelegramRequest_(message, requestId);
    return json_({
      status: 'success',
      message: notified ? 'Request recorded and sent to admin for review.' : 'Request recorded, but Telegram notification failed. Admin must review the Requests sheet.',
      requestId: requestId,
      notificationSent: notified
    });
  } catch (err) {
    console.error('doPost transaction: ' + String(err && err.stack || err));
    return json_({ status: 'error', message: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function registerPlayer_(p) {
  const name = String(p.name || '').trim();
  const phone = cleanPhone_(p.phone);
  const password = String(p.password || '');
  if (!name || !/^251[79]\d{8}$/.test(phone) || password.length < 8) {
    return { status: 'error', message: 'Enter your full name, valid Ethiopian phone and password (8+ characters).' };
  }
  if (findUserByPhone_(phone)) return { status: 'error', message: 'This phone number is already registered. Please log in.' };
  const salt = Utilities.getUuid() + Utilities.getUuid();
  const hash = passwordHash_(password, salt);
  const id = 'RH-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  sheet_(CFG.USERS).appendRow([id, name, phone, hash, salt, new Date()]);
  sheet_(CFG.BALANCES).appendRow([phone, name, 0, new Date()]);
  return { status: 'success', id: id, name: name, phone: phone, balance: 0 };
}

function loginPlayer_(p) {
  const phone = cleanPhone_(p.phone);
  const password = String(p.password || '');
  const user = findUserByPhone_(phone);
  if (!user || passwordHash_(password, user.salt) !== user.hash) {
    return { status: 'error', message: 'Invalid phone number or password.' };
  }
  const row = ensureBalanceRow_(user);
  const balance = Number(sheet_(CFG.BALANCES).getRange(row, 3).getValue()) || 0;
  return { status: 'success', id: user.id, name: user.name, phone: user.phone, balance: balance };
}

function getBalance_(rawPhone) {
  const phone = cleanPhone_(rawPhone);
  const user = findUserByPhone_(phone);
  if (!user) return { status: 'error', message: 'Player not found.' };
  const row = ensureBalanceRow_(user);
  return { status: 'success', balance: Number(sheet_(CFG.BALANCES).getRange(row, 3).getValue()) || 0 };
}

function getHistory_(rawPhone) {
  const phone = cleanPhone_(rawPhone);
  if (!findUserByPhone_(phone)) return { status: 'error', message: 'Player not found.' };
  const sh = sheet_(CFG.REQUESTS);
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (cleanPhone_(r[3]) !== phone) continue;
    out.push({
      id: String(r[0] || ''),
      date: r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '',
      type: String(r[2] || ''),
      amount: Number(r[5]) || 0,
      details: String(r[6] || ''),
      status: String(r[8] || 'PENDING').toUpperCase()
    });
  }
  out.reverse();
  return { status: 'success', transactions: out };
}

function approveSelectedRequest() { processSelectedRequest_('APPROVED'); }
function rejectSelectedRequest() { processSelectedRequest_('REJECTED'); }

function processSelectedRequest_(decision) {
  const sh = sheet_(CFG.REQUESTS);
  const row = sh.getActiveRange() ? sh.getActiveRange().getRow() : 0;
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Select a request row first.');
    return;
  }
  const requestId = String(sh.getRange(row, 1).getValue() || '');
  const result = processRequestById_(requestId, decision);
  SpreadsheetApp.getUi().alert(result.message);
}

function configureTelegramWebhook() {
  const token = getBotToken_();
  const url = ScriptApp.getService().getUrl();
  if (!token || !url) throw new Error('Missing bot token or deployed Web App URL. Deploy as Execute as Me / Anyone first.');
  const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/setWebhook', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ url: url, allowed_updates: ['message', 'callback_query'] }),
    muteHttpExceptions: true
  });
  const result = JSON.parse(response.getContentText());
  SpreadsheetApp.getUi().alert(result.ok
    ? 'Telegram approval webhook connected to this Apps Script deployment.'
    : 'Webhook setup failed: ' + (result.description || 'Unknown error'));
  return result;
}

function sendTelegramRequest_(message, requestId) {
  const token = getBotToken_();
  if (!token) {
    console.error('Telegram bot token is missing.');
    return false;
  }
  const chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID') || CFG.CHAT_ID;
  try {
    const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: chatId,
        text: message + '\n\nChoose an action:',
        reply_markup: {
          inline_keyboard: [[
            { text: '✅ APPROVE', callback_data: 'RH|APPROVED|' + requestId },
            { text: '❌ REJECT', callback_data: 'RH|REJECTED|' + requestId }
          ]]
        }
      }),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (!result.ok) console.error('Telegram sendMessage failed: ' + (result.description || response.getContentText()));
    return !!result.ok;
  } catch (err) {
    console.error('Telegram request notification: ' + err);
    return false;
  }
}

function handleTelegramCallback_(q) {
  const token = getBotToken_();
  const props = PropertiesService.getScriptProperties();
  const allowedChat = String(props.getProperty('TELEGRAM_CHAT_ID') || CFG.CHAT_ID);
  const messageChat = String(q.message && q.message.chat ? q.message.chat.id : '');
  const senderId = String(q.from && q.from.id ? q.from.id : '');
  const chatType = String(q.message && q.message.chat ? q.message.chat.type : '');
  const allowedAdmin = String(props.getProperty('TELEGRAM_ADMIN_USER_ID') || (chatType === 'private' ? allowedChat : ''));
  const callbackId = String(q.id || '');
  const rawData = String(q.data || '');

  let decision = '';
  let requestId = '';
  const parts = rawData.split('|');
  if (parts.length === 3 && parts[0] === 'RH') {
    decision = parts[1].toUpperCase();
    requestId = parts[2].toUpperCase();
  } else {
    const actionMatch = rawData.match(/APPROVE(?:D)?|REJECT(?:ED)?/i);
    const idMatch = rawData.match(/RH-[A-Z0-9]{8}/i);
    if (actionMatch && idMatch) {
      decision = /^APPROVE/i.test(actionMatch[0]) ? 'APPROVED' : 'REJECTED';
      requestId = idMatch[0].toUpperCase();
    }
  }

  if (messageChat !== allowedChat || !allowedAdmin || senderId !== allowedAdmin ||
      !/^RH-[A-Z0-9]{8}$/.test(requestId) ||
      !['APPROVED', 'REJECTED'].includes(decision)) {
    answerTelegramCallback_(callbackId, 'Not authorized or unsupported button.', true);
    return json_({ status: 'error', message: 'Unauthorized or unsupported Telegram callback.' });
  }

  try {
    const result = processRequestById_(requestId, decision);
    answerTelegramCallback_(callbackId, result.message, result.status !== 'success');
    if (q.message && q.message.message_id && result.status === 'success') {
      const suffix = '\n\n' + (decision === 'APPROVED' ? '✅ APPROVED' : '❌ REJECTED') + ' by admin';
      UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/editMessageText', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          chat_id: messageChat,
          message_id: q.message.message_id,
          text: String(q.message.text || '') + suffix,
          reply_markup: { inline_keyboard: [] }
        }),
        muteHttpExceptions: true
      });
    }
    return json_({ status: result.status, message: result.message });
  } catch (err) {
    console.error('Telegram callback: ' + String(err && err.stack || err));
    answerTelegramCallback_(callbackId, 'Approval failed: ' + String(err.message || err), true);
    return json_({ status: 'error', message: String(err.message || err) });
  }
}

function answerTelegramCallback_(callbackId, message, alert) {
  const token = getBotToken_();
  if (!token || !callbackId) return;
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/answerCallbackQuery', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      callback_query_id: callbackId,
      text: String(message || '').slice(0, 180),
      show_alert: !!alert
    }),
    muteHttpExceptions: true
  });
}

/**
 * Idempotent request processing:
 * 1. Create a PREPARED ledger row with before/after balance.
 * 2. Set the balance to the recorded after-balance (never add twice).
 * 3. Mark ledger APPLIED and request APPROVED/REJECTED.
 * If Apps Script stops between steps, pressing the same button again resumes
 * the same ledger operation without crediting the player twice.
 */
function processRequestById_(requestId, decision) {
  if (!['APPROVED', 'REJECTED'].includes(decision)) throw new Error('Invalid decision.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const requestSheet = sheet_(CFG.REQUESTS);
    const ledgerSheet = sheet_(CFG.LEDGER);
    if (requestSheet.getLastRow() < 2) throw new Error('No requests found.');

    const rows = requestSheet.getRange(2, 1, requestSheet.getLastRow() - 1, HEADERS.Requests.length).getValues();
    let row = -1;
    let request = null;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(requestId)) {
        row = i + 2;
        request = rows[i];
        break;
      }
    }
    if (row < 0) throw new Error('Request ID not found: ' + requestId);

    const currentStatus = String(request[8] || '').toUpperCase();
    if (currentStatus === 'APPROVED' || currentStatus === 'REJECTED') {
      return { status: 'error', message: 'Already ' + currentStatus + '. No balance change made.' };
    }
    if (currentStatus !== 'PENDING' && currentStatus !== 'PROCESSING') {
      throw new Error('Request is not pending. Current status: ' + currentStatus);
    }

    const type = String(request[2] || '').toUpperCase();
    if (!['DEPOSIT', 'WITHDRAW'].includes(type)) throw new Error('Invalid request type.');
    const phone = cleanPhone_(request[3]);
    const amount = Number(request[5]);
    if (!Number.isFinite(amount) || amount <= 0) throw new Error('Invalid request amount.');
    const user = findUserByPhone_(phone);
    if (!user) throw new Error('Player record not found for request ' + requestId + '.');

    let ledgerRow = findLedgerRow_(requestId);
    if (!ledgerRow) {
      const balanceRow = ensureBalanceRow_(user);
      const balanceSheet = sheet_(CFG.BALANCES);
      const before = Number(balanceSheet.getRange(balanceRow, 3).getValue()) || 0;
      let after = before;
      if (type === 'DEPOSIT' && decision === 'APPROVED') after = before + amount;
      if (type === 'WITHDRAW' && decision === 'REJECTED') after = before + amount;
      ledgerSheet.appendRow([requestId, phone, decision, type, amount, before, after, 'PREPARED', new Date(), '']);
      ledgerRow = ledgerSheet.getLastRow();
    }

    const ledger = ledgerSheet.getRange(ledgerRow, 1, 1, HEADERS.BalanceLedger.length).getValues()[0];
    const savedDecision = String(ledger[2] || '').toUpperCase();
    if (savedDecision !== decision) throw new Error('This request already has a ledger operation for ' + savedDecision + '. Contact admin; do not press the other button.');
    const balanceRow = ensureBalanceRow_(user);
    const balanceSheet = sheet_(CFG.BALANCES);
    const afterBalance = Number(ledger[6]);
    if (!Number.isFinite(afterBalance)) throw new Error('Ledger after-balance is invalid.');

    // Setting a recorded absolute balance is safe to repeat after a timeout.
    balanceSheet.getRange(balanceRow, 3).setValue(afterBalance);
    balanceSheet.getRange(balanceRow, 4).setValue(new Date());
    ledgerSheet.getRange(ledgerRow, 8, 1, 2).setValues([['APPLIED', new Date()]]);
    requestSheet.getRange(row, 9, 1, 3).setValues([[
      decision,
      new Date(),
      decision === 'REJECTED' && type === 'WITHDRAW' ? 'Withdrawal refunded' : 'Processed through Telegram / RASTA ADMIN'
    ]]);

    const message = type + ' ' + decision.toLowerCase() + '. Player balance: ' + afterBalance;
    sendTelegram_((decision === 'APPROVED' ? '✅ APPROVED' : '❌ REJECTED') +
      '\nRequest: ' + requestId + '\nType: ' + type + '\nPlayer: ' + request[4] +
      '\nPhone: ' + phone + '\nAmount: ' + amount +
      '\nNew balance: ' + afterBalance +
      (type === 'WITHDRAW' && decision === 'REJECTED' ? '\nReserved amount refunded.' : ''));
    return { status: 'success', message: message, balance: afterBalance };
  } finally {
    lock.releaseLock();
  }
}

function findLedgerRow_(requestId) {
  const sh = sheet_(CFG.LEDGER);
  if (sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(requestId)) return i + 2;
  }
  return 0;
}

function transactionIdExists_(id) {
  const wanted = String(id).trim().toLowerCase();
  const sh = sheet_(CFG.REQUESTS);
  if (sh.getLastRow() < 2) return false;
  const vals = sh.getRange(2, 8, sh.getLastRow() - 1, 1).getDisplayValues();
  return vals.some(function(r) { return String(r[0]).trim().toLowerCase() === wanted; });
}

function extractTransactionId_(details) {
  const m = String(details || '').match(/transaction\s*id\s*:\s*([^|]+)/i);
  return m ? m[1].trim() : '';
}

function findUserByPhone_(phone) {
  const sh = sheet_(CFG.USERS);
  if (sh.getLastRow() < 2) return null;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.Users.length).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (cleanPhone_(vals[i][2]) === phone) {
      return {
        id: String(vals[i][0]),
        name: String(vals[i][1]),
        phone: phone,
        hash: String(vals[i][3]),
        salt: String(vals[i][4]),
        row: i + 2
      };
    }
  }
  return null;
}

function ensureBalanceRow_(user) {
  const sh = sheet_(CFG.BALANCES);
  if (sh.getLastRow() >= 2) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
    for (let i = 0; i < vals.length; i++) {
      if (cleanPhone_(vals[i][0]) === user.phone) return i + 2;
    }
  }
  sh.appendRow([user.phone, user.name, 0, new Date()]);
  return sh.getLastRow();
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('Spreadsheet ID is not configured. Open the bound spreadsheet and run RASTA ADMIN > Setup / Repair Sheets once.');
  props.setProperty('SPREADSHEET_ID', active.getId());
  return active;
}

function sheet_(name) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const headers = HEADERS[name];
  if (headers) {
    headers.forEach(function(h, i) {
      if (sh.getRange(1, i + 1).getValue() !== h) sh.getRange(1, i + 1).setValue(h);
    });
  }
  return sh;
}

function cleanPhone_(value) {
  let s = String(value || '').replace(/\D/g, '');
  if (/^0[79]\d{8}$/.test(s)) s = '251' + s.slice(1);
  else if (/^[79]\d{8}$/.test(s)) s = '251' + s;
  return s;
}

function passwordHash_(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + '|' + String(password),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function getBotToken_() {
  const p = PropertiesService.getScriptProperties();
  return p.getProperty('TELEGRAM_BOT_TOKEN') || p.getProperty('RASTA_BOT_TOKEN') ||
    p.getProperty('BOT_TOKEN') || p.getProperty('TELEGRAM_TOKEN') || '';
}

function sendTelegram_(message) {
  const token = getBotToken_();
  if (!token) return false;
  const chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID') || CFG.CHAT_ID;
  try {
    const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: message }),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (!result.ok) console.error('Telegram send failed: ' + (result.description || response.getContentText()));
    return !!result.ok;
  } catch (err) {
    console.error('Telegram: ' + err);
    return false;
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
