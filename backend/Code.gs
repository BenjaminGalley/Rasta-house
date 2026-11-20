/**
 * RASTA HOUSE — Google Apps Script backend
 * Demo-credit simulation only. No real payment processing.
 *
 * Required Script Properties:
 *   TELEGRAM_BOT_TOKEN = existing Telegram bot token
 *   TELEGRAM_CHAT_ID   = 9268792685 (default below)
 *
 * Run setup() once, then deploy as Web App:
 * Execute as: Me | Who has access: Anyone
 */
const CFG = {
  CHAT_ID: '9268792685',
  USERS: 'Users',
  BALANCES: 'DemoBalances',
  REQUESTS: 'Requests',
  MIN_DEPOSIT: 150,
  MAX_DEPOSIT: 10000,
  MIN_WITHDRAW: 300,
  MAX_WITHDRAW: 10000
};

const HEADERS = {
  Users: ['PlayerID','FullName','Phone','PasswordHash','PasswordSalt','JoinedAt'],
  DemoBalances: ['Phone','FullName','Balance','UpdatedAt'],
  Requests: ['RequestID','CreatedAt','Type','Phone','FullName','Amount','Details','TransactionID','Status','ProcessedAt','AdminNote']
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('RASTA ADMIN')
    .addItem('Setup / Repair Sheets', 'setup')
    .addItem('Configure Telegram Bot', 'configureTelegram')
    .addItem('Connect Telegram Approval Buttons', 'configureTelegramWebhook')
    .addItem('Test Telegram', 'testTelegram')
    .addSeparator()
    .addItem('Approve Selected Request', 'approveSelectedRequest')
    .addItem('Reject Selected Request', 'rejectSelectedRequest')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = HEADERS[name];
    const existing = sh.getRange(1,1,1,Math.max(headers.length, sh.getLastColumn() || 1)).getValues()[0];
    headers.forEach((h,i) => { if (existing[i] !== h) sh.getRange(1,i+1).setValue(h); });
    sh.setFrozenRows(1);
  });
  SpreadsheetApp.getUi().alert('RASTA HOUSE setup complete. Users, DemoBalances and Requests are ready.');
}

function configureTelegram() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const current = getBotToken_();
  const result = ui.prompt('Configure Telegram', 'Paste your bot token. Leave blank to keep the existing token.', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const token = result.getResponseText().trim();
  if (token) props.setProperty('TELEGRAM_BOT_TOKEN', token);
  if (!getBotToken_()) {
    ui.alert('No Telegram token saved. Run Configure Telegram again.');
    return;
  }
  props.setProperty('TELEGRAM_CHAT_ID', CFG.CHAT_ID);
  ui.alert('Telegram settings saved. Use Test Telegram to verify.');
}

function testTelegram() {
  const ok = sendTelegram_('🌿 RASTA HOUSE TEST\nTelegram notifications are connected.');
  SpreadsheetApp.getUi().alert(ok ? 'Telegram test sent.' : 'Telegram test failed. Check the bot token and chat ID.');
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = String(p.action || '').trim();
    if (action === 'register') return json_(registerPlayer_(p));
    if (action === 'login') return json_(loginPlayer_(p));
    if (action === 'getBalance') return json_(getBalance_(p.phone));
    if (action === 'getHistory') return json_(getHistory_(p.phone));
    if (action === 'ping') return json_({status:'success',message:'RASTA backend online'});
    return json_({status:'error',message:'Unknown action.'});
  } catch (err) {
    return json_({status:'error',message:String(err.message || err)});
  }
}

function doPost(e) {
  // Telegram sends JSON updates; website forms send URL-encoded parameters.
  let body = {};
  try {
    if (e && e.postData && e.postData.contents) {
      const raw = e.postData.contents;
      if (String(e.postData.type || '').toLowerCase().indexOf('application/json') >= 0) body = JSON.parse(raw);
    }
  } catch (_) {}
  if (body && body.callback_query) return handleTelegramCallback_(body.callback_query);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = (e && e.parameter) || {};
    const action = String(p.action || '').trim();
    if (action !== 'deposit' && action !== 'withdraw') {
      return json_({status:'error',message:'Unknown transaction action.'});
    }
    const phone = cleanPhone_(p.phone);
    const user = findUserByPhone_(phone);
    if (!user) return json_({status:'error',message:'Player is not registered. Please register or log in again.'});
    const amount = Number(p.amount);
    const min = action === 'deposit' ? CFG.MIN_DEPOSIT : CFG.MIN_WITHDRAW;
    if (!Number.isFinite(amount) || Math.floor(amount) !== amount || amount < min || amount > CFG.MAX_WITHDRAW) {
      return json_({status:'error',message:'Amount is outside the allowed range.'});
    }
    const details = String(p.details || '').trim();
    const transactionId = String(p.transactionId || extractTransactionId_(details) || '').trim();
    if (action === 'deposit' && !transactionId) return json_({status:'error',message:'Payment transaction ID is required.'});
    if (action === 'deposit' && transactionIdExists_(transactionId)) {
      return json_({status:'error',message:'This payment transaction ID has already been submitted.'});
    }
    const balanceRow = ensureBalanceRow_(user);
    const balanceSheet = sheet_(CFG.BALANCES);
    const currentBalance = Number(balanceSheet.getRange(balanceRow,3).getValue()) || 0;
    if (action === 'withdraw' && currentBalance < amount) {
      return json_({status:'error',message:'Insufficient demo-credit balance.'});
    }
    const reqSheet = sheet_(CFG.REQUESTS);
    const requestId = 'RH-' + Utilities.getUuid().slice(0,8).toUpperCase();
    const now = new Date();
    reqSheet.appendRow([requestId,now,action.toUpperCase(),phone,user.name,amount,details,transactionId,'PENDING','', '']);
    if (action === 'withdraw') {
      balanceSheet.getRange(balanceRow,3).setValue(currentBalance - amount);
      balanceSheet.getRange(balanceRow,4).setValue(now);
    }
    const message = (action === 'deposit' ? '🟡 NEW DEPOSIT REQUEST' : '🟠 NEW WITHDRAWAL REQUEST') +
      '\nID: ' + requestId + '\nPlayer: ' + user.name + '\nPhone: ' + phone +
      '\nAmount: ' + amount + ' demo credits\nDetails: ' + details +
      (transactionId ? '\nTransaction ID: ' + transactionId : '');
    sendTelegramRequest_(message, requestId);
    return json_({status:'success',message:'Request recorded and awaiting review.',requestId:requestId});
  } catch (err) {
    return json_({status:'error',message:String(err.message || err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function registerPlayer_(p) {
  const name = String(p.name || '').trim();
  const phone = cleanPhone_(p.phone);
  const password = String(p.password || '');
  if (!name || !/^251[79]\d{8}$/.test(phone) || password.length < 8) {
    return {status:'error',message:'Enter your full name, valid Ethiopian phone and password (8+ characters).'};
  }
  if (findUserByPhone_(phone)) return {status:'error',message:'This phone number is already registered. Please log in.'};
  const salt = Utilities.getUuid() + Utilities.getUuid();
  const hash = passwordHash_(password,salt);
  const id = 'RH-' + Utilities.getUuid().slice(0,8).toUpperCase();
  sheet_(CFG.USERS).appendRow([id,name,phone,hash,salt,new Date()]);
  const bal = sheet_(CFG.BALANCES);
  bal.appendRow([phone,name,0,new Date()]);
  return {status:'success',id:id,name:name,phone:phone,balance:0};
}

function loginPlayer_(p) {
  const phone = cleanPhone_(p.phone);
  const password = String(p.password || '');
  const user = findUserByPhone_(phone);
  if (!user || passwordHash_(password,user.salt) !== user.hash) {
    return {status:'error',message:'Invalid phone number or password.'};
  }
  const row = ensureBalanceRow_(user);
  const balance = Number(sheet_(CFG.BALANCES).getRange(row,3).getValue()) || 0;
  return {status:'success',id:user.id,name:user.name,phone:user.phone,balance:balance};
}

function getBalance_(rawPhone) {
  const phone = cleanPhone_(rawPhone);
  const user = findUserByPhone_(phone);
  if (!user) return {status:'error',message:'Player not found.'};
  const row = ensureBalanceRow_(user);
  return {status:'success',balance:Number(sheet_(CFG.BALANCES).getRange(row,3).getValue()) || 0};
}

function getHistory_(rawPhone) {
  const phone = cleanPhone_(rawPhone);
  if (!findUserByPhone_(phone)) return {status:'error',message:'Player not found.'};
  const sh = sheet_(CFG.REQUESTS);
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i=1;i<values.length;i++) {
    const r=values[i];
    if (cleanPhone_(r[3]) !== phone) continue;
    out.push({
      id:String(r[0] || ''),
      date:r[1] ? Utilities.formatDate(new Date(r[1]),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm') : '',
      type:String(r[2] || ''),
      amount:Number(r[5]) || 0,
      details:String(r[6] || ''),
      status:String(r[8] || 'PENDING').toUpperCase()
    });
  }
  out.reverse();
  return {status:'success',transactions:out};
}

function approveSelectedRequest() { processSelectedRequest_('APPROVED'); }
function rejectSelectedRequest() { processSelectedRequest_('REJECTED'); }

function processSelectedRequest_(decision) {
  const sh = sheet_(CFG.REQUESTS);
  const row = sh.getActiveRange() ? sh.getActiveRange().getRow() : 0;
  if (row < 2) { SpreadsheetApp.getUi().alert('Select a request row first.'); return; }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r=sh.getRange(row,1,1,HEADERS.Requests.length).getValues()[0];
    const status=String(r[8] || '').toUpperCase();
    if (status !== 'PENDING') { SpreadsheetApp.getUi().alert('This request is already '+status+'.'); return; }
    const type=String(r[2] || '').toUpperCase();
    const phone=cleanPhone_(r[3]);
    const amount=Number(r[5]) || 0;
    const user=findUserByPhone_(phone);
    if (!user) throw new Error('Player record not found for '+phone);
    const balRow=ensureBalanceRow_(user);
    const balSh=sheet_(CFG.BALANCES);
    let balance=Number(balSh.getRange(balRow,3).getValue()) || 0;
    if (type==='DEPOSIT' && decision==='APPROVED') balance+=amount;
    if (type==='WITHDRAW' && decision==='REJECTED') balance+=amount;
    balSh.getRange(balRow,3).setValue(balance);
    balSh.getRange(balRow,4).setValue(new Date());
    sh.getRange(row,9).setValue(decision);
    sh.getRange(row,10).setValue(new Date());
    sh.getRange(row,11).setValue(decision==='REJECTED' && type==='WITHDRAW' ? 'Withdrawal refunded' : '');
    sendTelegram_((decision==='APPROVED'?'✅ APPROVED':'❌ REJECTED')+
      '\nRequest: '+r[0]+'\nType: '+type+'\nPlayer: '+r[4]+'\nPhone: '+phone+
      '\nAmount: '+amount+'\nNew balance: '+balance+
      (type==='WITHDRAW' && decision==='REJECTED' ? '\nReserved amount refunded.' : ''));
    SpreadsheetApp.getUi().alert(type+' request '+decision.toLowerCase()+'.');
  } finally { lock.releaseLock(); }
}

// Run once after deploying the updated Web App. This registers Telegram's
// button callbacks with this same Apps Script deployment.
function configureTelegramWebhook() {
  const token = getBotToken_();
  const url = ScriptApp.getService().getUrl();
  if (!token || !url) throw new Error('Missing bot token or deployed Web App URL. Deploy as Execute as Me / Anyone first.');
  const response = UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/setWebhook', {
    method:'post', contentType:'application/json',
    payload:JSON.stringify({url:url,allowed_updates:['message','callback_query']}),
    muteHttpExceptions:true
  });
  const result = JSON.parse(response.getContentText());
  SpreadsheetApp.getUi().alert(result.ok ? 'Telegram approval buttons connected.' : 'Webhook setup failed: '+(result.description || 'Unknown error'));
  return result;
}

function sendTelegramRequest_(message, requestId) {
  const token = getBotToken_();
  if (!token) return false;
  const chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID') || CFG.CHAT_ID;
  try {
    const response = UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/sendMessage', {
      method:'post', contentType:'application/json',
      payload:JSON.stringify({
        chat_id:chatId, text:message,
        reply_markup:{inline_keyboard:[[
          {text:'Approve',callback_data:'RH|APPROVED|'+requestId},
          {text:'Reject',callback_data:'RH|REJECTED|'+requestId}
        ]]}
      }),
      muteHttpExceptions:true
    });
    return !!JSON.parse(response.getContentText()).ok;
  } catch(err) { console.error('Telegram request notification: '+err); return false; }
}

function handleTelegramCallback_(q) {
  const token = getBotToken_();
  const props = PropertiesService.getScriptProperties();
  const allowedChat = String(props.getProperty('TELEGRAM_CHAT_ID') || CFG.CHAT_ID);
  const callbackId = String(q.id || '');
  const messageChat = String(q.message && q.message.chat ? q.message.chat.id : '');
  const rawData = String(q.data || '');
  // Accept current RH|APPROVED|ID callbacks and common legacy button payloads.
  let decision = '', requestId = '';
  const parts = rawData.split('|');
  if (parts.length === 3 && parts[0] === 'RH') {
    decision = parts[1].toUpperCase();
    requestId = parts[2];
  } else {
    const actionMatch = rawData.match(/APPROVE(?:D)?|REJECT(?:ED)?/i);
    const idMatch = rawData.match(/RH-[A-Z0-9]{8}/i);
    if (actionMatch && idMatch) {
      decision = /^APPROVE/i.test(actionMatch[0]) ? 'APPROVED' : 'REJECTED';
      requestId = idMatch[0].toUpperCase();
    }
  }
  if (messageChat !== allowedChat || !requestId ||
      !['APPROVED','REJECTED'].includes(decision)) {
    answerTelegramCallback_(callbackId, 'Not authorized or unsupported button. Update the Apps Script deployment.', true);
    return json_({status:'error',message:'Unauthorized or unsupported Telegram callback.'});
  }
  try {
    const result = processRequestById_(requestId, decision);
    answerTelegramCallback_(callbackId, result.message, false);
    if (q.message && q.message.message_id) {
      const suffix = result.status === 'success' ? '\n\n'+(decision === 'APPROVED' ? '✅ APPROVED' : '❌ REJECTED')+' by admin' : '';
      UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/editMessageText', {
        method:'post', contentType:'application/json',
        payload:JSON.stringify({chat_id:messageChat,message_id:q.message.message_id,
          text:String(q.message.text || '')+suffix,reply_markup:{inline_keyboard:[]}}),
        muteHttpExceptions:true
      });
    }
    return json_({status:result.status,message:result.message});
  } catch(err) {
    answerTelegramCallback_(callbackId, 'Approval failed: '+String(err.message || err), true);
    return json_({status:'error',message:String(err.message || err)});
  }
}

function answerTelegramCallback_(callbackId, message, alert) {
  const token = getBotToken_();
  if (!token || !callbackId) return;
  UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/answerCallbackQuery', {
    method:'post', contentType:'application/json',
    payload:JSON.stringify({callback_query_id:callbackId,text:String(message || '').slice(0,180),show_alert:!!alert}),
    muteHttpExceptions:true
  });
}

function processRequestById_(requestId, decision) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(CFG.REQUESTS);
    if (sh.getLastRow() < 2) throw new Error('No requests found.');
    const rows = sh.getRange(2,1,sh.getLastRow()-1,HEADERS.Requests.length).getValues();
    let row = -1, r = null;
    for (let i=0;i<rows.length;i++) if (String(rows[i][0]) === String(requestId)) { row=i+2; r=rows[i]; break; }
    if (row < 0) throw new Error('Request ID not found: '+requestId);
    const status = String(r[8] || '').toUpperCase();
    if (status !== 'PENDING') return {status:'error',message:'Already '+status+'. No balance change made.'};
    const type = String(r[2] || '').toUpperCase();
    if (!['DEPOSIT','WITHDRAW'].includes(type)) throw new Error('Invalid request type.');
    const phone = cleanPhone_(r[3]), amount = Number(r[5]) || 0;
    const user = findUserByPhone_(phone);
    if (!user) throw new Error('Player record not found.');
    const balRow = ensureBalanceRow_(user), balSh = sheet_(CFG.BALANCES);
    let balance = Number(balSh.getRange(balRow,3).getValue()) || 0;
    if (type === 'DEPOSIT' && decision === 'APPROVED') balance += amount;
    if (type === 'WITHDRAW' && decision === 'REJECTED') balance += amount;
    balSh.getRange(balRow,3).setValue(balance);
    balSh.getRange(balRow,4).setValue(new Date());
    sh.getRange(row,9,1,3).setValues([[decision,new Date(),decision==='REJECTED' && type==='WITHDRAW' ? 'Withdrawal refunded' : '']]);
    return {status:'success',message:type+' '+decision.toLowerCase()+'. Balance: '+balance};
  } finally { lock.releaseLock(); }
}

function transactionIdExists_(id) {
  const wanted=String(id).trim().toLowerCase();
  const sh=sheet_(CFG.REQUESTS);
  if (sh.getLastRow()<2) return false;
  const vals=sh.getRange(2,8,sh.getLastRow()-1,1).getDisplayValues();
  return vals.some(r=>String(r[0]).trim().toLowerCase()===wanted);
}

function extractTransactionId_(details) {
  const m=String(details||'').match(/transaction\s*id\s*:\s*([^|]+)/i);
  return m ? m[1].trim() : '';
}

function findUserByPhone_(phone) {
  const sh=sheet_(CFG.USERS);
  if (sh.getLastRow()<2) return null;
  const vals=sh.getRange(2,1,sh.getLastRow()-1,HEADERS.Users.length).getValues();
  for (let i=0;i<vals.length;i++) {
    if (cleanPhone_(vals[i][2])===phone) return {
      id:String(vals[i][0]),name:String(vals[i][1]),phone:phone,
      hash:String(vals[i][3]),salt:String(vals[i][4]),row:i+2
    };
  }
  return null;
}

function ensureBalanceRow_(user) {
  const sh=sheet_(CFG.BALANCES);
  if (sh.getLastRow()>=2) {
    const vals=sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues();
    for(let i=0;i<vals.length;i++) if(cleanPhone_(vals[i][0])===user.phone) return i+2;
  }
  sh.appendRow([user.phone,user.name,0,new Date()]);
  return sh.getLastRow();
}

function sheet_(name) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  const headers=HEADERS[name];
  if(headers) {
    headers.forEach((h,i)=>{ if(sh.getRange(1,i+1).getValue()!==h) sh.getRange(1,i+1).setValue(h); });
  }
  return sh;
}

function cleanPhone_(value) {
  let s=String(value||'').replace(/\D/g,'');
  if(/^0[79]\d{8}$/.test(s)) s='251'+s.slice(1);
  else if(/^[79]\d{8}$/.test(s)) s='251'+s;
  return s;
}

function passwordHash_(password,salt) {
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(salt)+'|'+String(password),Utilities.Charset.UTF_8);
  return bytes.map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
}

function getBotToken_() {
  const p=PropertiesService.getScriptProperties();
  return p.getProperty('TELEGRAM_BOT_TOKEN') || p.getProperty('RASTA_BOT_TOKEN') ||
    p.getProperty('BOT_TOKEN') || p.getProperty('TELEGRAM_TOKEN') || '';
}

function sendTelegram_(message) {
  const token=getBotToken_();
  if(!token) return false;
  const chatId=PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID') || CFG.CHAT_ID;
  try {
    const response=UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/sendMessage',{
      method:'post',contentType:'application/json',
      payload:JSON.stringify({chat_id:chatId,text:message}),
      muteHttpExceptions:true
    });
    const result=JSON.parse(response.getContentText());
    return !!result.ok;
  } catch(err) { console.error('Telegram: '+err); return false; }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
