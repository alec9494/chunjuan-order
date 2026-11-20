/******** 設定 ********/
const DEFAULT_PRICE = 150;      // 預設單價（之後可在後台修改）
const ITEM_NAME = '招牌商品';    // 品項名稱（顯示用）
const STORE_NAME = '我的小店';   // 店名
const SHEET_NAME = '訂單';
const PENDING_SHEET = 'LINE待付款';
const TZ = 'Asia/Taipei';
const DEFAULT_PIN = '1234';     // 請到「專案設定 > 指令碼屬性」新增 PIN 改掉
const PAY_MINUTES = 15;         // LINE 點餐後幾分鐘內要到櫃台付款，逾時自動取消
const DEFAULT_MAX_PUSH = 200;   // 每月最多發幾則 LINE 通知（輕用量方案免費 200 則）

/*  指令碼屬性（專案設定 > 指令碼屬性）要新增：
 *    PIN                    店員登入密碼
 *    LINE_LOGIN_CHANNEL_ID  LINE Login channel 的 Channel ID（驗證客人身分用）
 *    LINE_TOKEN             Messaging API channel 的 Channel access token（發通知用）
 *    LIFF_ID                LIFF ID（通知訊息裡附上取餐券連結，可不填）
 *    MAX_PUSH               每月通知上限（可不填，預設 200）
 */

/******** 網頁入口 ********/
// 店員：.../exec?p=staff    客人查詢：.../exec  或  .../exec?n=12    看板：.../exec?p=api
function doGet(e) {
  if (e.parameter.p === 'api') {
    const s = read_();
    return ContentService.createTextOutput(s.called + ',' + s.last + ',' + s.ring)
      .setMimeType(ContentService.MimeType.TEXT);
  }
  const isStaff = e.parameter.p === 'staff';
  const t = HtmlService.createTemplateFromFile(isStaff ? 'Staff' : 'Customer');
  t.url = ScriptApp.getService().getUrl();
  t.n = e.parameter.n || '';
  t.price = price_();
  t.itemName = ITEM_NAME;
  t.storeName = STORE_NAME;
  return t.evaluate()
    .setTitle(isStaff ? 'POS 叫號' : '叫號查詢')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// LINE 取餐頁（liff.html）呼叫的 API
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    out = { ok: true, data: lineApi_(req) };
  } catch (err) {
    out = { ok: false, error: err.message, relogin: err.relogin === true };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/******** 工具 ********/
function props_() { return PropertiesService.getScriptProperties(); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'HH:mm:ss'); }
function hm_(ms) { return Utilities.formatDate(new Date(ms), TZ, 'HH:mm'); }
function month_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM'); }
function pickCode_() { return String(10 + Math.floor(Math.random() * 90)); }

const HEADER = ['日期', '號碼', '數量', '金額', '收款', '找零', '下單時間', '狀態', '叫號時間', '領取時間',
                '來源', 'LINE ID', 'LINE 名稱', '取餐碼'];
const COLS = HEADER.length;
const P_HEADER = ['ID', '付款碼', '建立(ms)', '日期', '建立時間', 'LINE ID', 'LINE 名稱', '數量', '單價', '狀態', '號碼'];
const P_COLS = P_HEADER.length;

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, COLS).getValue() !== HEADER[COLS - 1]) {
    sh.getRange(1, 1, 1, COLS).setValues([HEADER]);   // 舊表補上新欄位
  }
  return sh;
}

function pendingSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(PENDING_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PENDING_SHEET);
    sh.appendRow(P_HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 今天所有訂單
function todayOrders_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const d = today_();
  return sh.getRange(2, 1, last - 1, COLS).getDisplayValues()
    .map((r, i) => ({ row: i + 2, date: r[0], n: +r[1], qty: +r[2], amount: +r[3], paid: +r[4], time: r[6],
                      status: r[7], calledAt: r[8], pickedAt: r[9], source: r[10], uid: r[11],
                      name: r[12], pickCode: r[13] }))
    .filter(o => o.date === d);
}

function findOrder_(sh, n) {
  const o = todayOrders_(sh).reverse().find(x => x.n === +n);
  if (!o) throw new Error('找不到今天的 ' + n + ' 號');
  return o;
}

// 今天的 LINE 待付款單（順便把逾時的標成「逾時」）
function todayPending_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const d = today_(), limit = Date.now() - PAY_MINUTES * 60000;
  const rows = sh.getRange(2, 1, last - 1, P_COLS).getDisplayValues();
  const list = [];
  rows.forEach((r, i) => {
    if (r[3] !== d) return;
    const p = { row: i + 2, id: r[0], code: r[1], created: +r[2], uid: r[5], name: r[6],
                qty: +r[7], price: +r[8], status: r[9], n: +r[10] || 0 };
    if (p.status === '待付款' && p.created < limit) {
      sh.getRange(p.row, 10).setValue('逾時');
      p.status = '逾時';
    }
    list.push(p);
  });
  return list;
}

// 叫號時記下時間；第一次叫到 LINE 客人時發通知
function markCalled_(n) {
  if (!(n > 0)) return;
  const sh = sheet_();
  try {
    const o = findOrder_(sh, n);
    if (o.calledAt) return;
    sh.getRange(o.row, 9).setValue(now_());
    if (o.uid && o.status !== '作廢') {
      const liff = props_().getProperty('LIFF_ID');
      pushLine_(o.uid, `🔔 ${n} 號好了！\n請到櫃台出示 LINE 取餐券取餐（取餐碼 ${o.pickCode}）` +
        (liff ? `\n\n取餐券：https://liff.line.me/${liff}` : ''));
    }
  } catch (e) {}
}

function price_() {
  return +props_().getProperty('PRICE') || DEFAULT_PRICE;
}

function checkPin_(pin) {
  const real = props_().getProperty('PIN') || DEFAULT_PIN;
  if (String(pin) !== String(real)) throw new Error('PIN 錯誤');
}

// 讀狀態（不寫入）；換日自動視為從 0 開始
function read_() {
  const p = props_().getProperties();
  if (p.DATE !== today_()) return { last: 0, called: 0, ring: 0, lineSeq: 0 };
  return { last: +p.LAST || 0, called: +p.CALLED || 0, ring: +p.RING || 0, lineSeq: +p.LINE_SEQ || 0 };
}

function write_(s) {
  props_().setProperties({
    DATE: today_(), LAST: String(s.last), CALLED: String(s.called),
    RING: String(s.ring || 0), LINE_SEQ: String(s.lineSeq || 0)
  });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function lineOpen_() { return props_().getProperty('LINE_OPEN') !== '0'; }

/******** LINE：推播 ********/
function pushCount_() { return +props_().getProperty('PUSH_' + month_()) || 0; }
function maxPush_() { return +props_().getProperty('MAX_PUSH') || DEFAULT_MAX_PUSH; }

function pushLine_(uid, text) {
  const token = props_().getProperty('LINE_TOKEN');
  if (!token || !uid) return false;
  if (pushCount_() >= maxPush_()) return false;          // 本月額度用完就不發，避免超額收費
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: uid, messages: [{ type: 'text', text }] })
  });
  if (res.getResponseCode() === 200) {
    props_().setProperty('PUSH_' + month_(), String(pushCount_() + 1));
    return true;
  }
  console.warn('LINE push 失敗', res.getResponseCode(), res.getContentText());
  return false;
}

/******** LINE：客人身分 ********/
function verifyIdToken_(idToken) {
  const cid = props_().getProperty('LINE_LOGIN_CHANNEL_ID');
  if (!cid) throw new Error('店家尚未設定 LINE_LOGIN_CHANNEL_ID');
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', muteHttpExceptions: true,
    payload: { id_token: String(idToken || ''), client_id: cid }
  });
  const j = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200 || !j.sub) {
    const err = new Error('LINE 登入已過期，請重新開啟頁面'); err.relogin = true; throw err;
  }
  return { uid: j.sub, name: j.name || 'LINE 用戶', pic: j.picture || '' };
}

function session_(sess) {
  const v = sess && CacheService.getScriptCache().get('s_' + sess);
  if (!v) { const err = new Error('請重新登入'); err.relogin = true; throw err; }
  return JSON.parse(v);
}

// 客人目前的取餐券
function ticketFor_(uid) {
  const s = read_();
  const orders = todayOrders_(sheet_()).filter(o => o.uid === uid && o.status !== '作廢');
  const pend = todayPending_(pendingSheet_()).filter(p => p.uid === uid);
  const lastPend = pend[pend.length - 1];
  const lastOrder = orders[orders.length - 1];

  if (lastPend && lastPend.status === '待付款') {
    return { stage: 'pay', code: lastPend.code, qty: lastPend.qty, amount: lastPend.qty * lastPend.price,
             expiresAt: lastPend.created + PAY_MINUTES * 60000, expiresHm: hm_(lastPend.created + PAY_MINUTES * 60000) };
  }
  if (lastOrder) {
    const base = { n: lastOrder.n, qty: lastOrder.qty, amount: lastOrder.amount, pickCode: lastOrder.pickCode };
    if (lastOrder.status === '已領') return Object.assign(base, { stage: 'done', pickedAt: lastOrder.pickedAt.slice(0, 5) });
    if (lastOrder.n <= s.called) return Object.assign(base, { stage: 'call', ring: s.ring });
    return Object.assign(base, { stage: 'wait', ahead: lastOrder.n - s.called - 1, called: s.called });
  }
  if (lastPend && (lastPend.status === '逾時' || lastPend.status === '取消')) {
    return { stage: 'expired', code: lastPend.code, reason: lastPend.status };
  }
  return null;
}

function lineApi_(req) {
  const base = () => ({ open: lineOpen_(), price: price_(), item: ITEM_NAME, store: STORE_NAME,
                        payMinutes: PAY_MINUTES, called: read_().called });

  if (req.action === 'login') {
    const u = verifyIdToken_(req.idToken);
    const sess = Utilities.getUuid();
    CacheService.getScriptCache().put('s_' + sess, JSON.stringify(u), 21600);   // 6 小時
    return Object.assign(base(), { sess, name: u.name, pic: u.pic, ticket: ticketFor_(u.uid) });
  }

  const u = session_(req.sess);

  if (req.action === 'state') {
    return Object.assign(base(), { ticket: ticketFor_(u.uid) });
  }

  if (req.action === 'order') {
    if (!lineOpen_()) throw new Error('目前暫停 LINE 點餐，請直接到櫃台');
    const PRICE = price_();
    if (+req.price !== PRICE) throw new Error('價格已更新為 $' + PRICE + '，請重新確認');
    const qty = Math.floor(+req.qty);
    if (!(qty >= 1 && qty <= 50)) throw new Error('數量不正確');
    return withLock_(() => {
      const t = ticketFor_(u.uid);
      if (t && (t.stage === 'pay' || t.stage === 'wait' || t.stage === 'call'))
        throw new Error('您還有一筆訂單尚未完成');
      const s = read_();
      s.lineSeq += 1; write_(s);
      const code = 'L' + String(s.lineSeq).padStart(2, '0');
      pendingSheet_().appendRow([Utilities.getUuid().slice(0, 8), code, Date.now(), today_(), now_(),
                                 u.uid, u.name, qty, PRICE, '待付款', '']);
      return Object.assign(base(), { ticket: ticketFor_(u.uid) });
    });
  }

  if (req.action === 'cancel') {
    return withLock_(() => {
      const sh = pendingSheet_();
      todayPending_(sh).filter(p => p.uid === u.uid && p.status === '待付款')
        .forEach(p => sh.getRange(p.row, 10).setValue('取消'));
      return Object.assign(base(), { ticket: ticketFor_(u.uid) });
    });
  }

  throw new Error('未知的動作');
}

/******** 客人查詢頁（公開） ********/
function getStatus() {
  const s = read_();
  return { called: s.called, last: s.last, ring: s.ring };
}

/******** 店員端（需 PIN） ********/
function staffState(pin) {
  checkPin_(pin);
  const s = read_();
  let orders = 0, qty = 0, revenue = 0, lineRevenue = 0;
  const recent = [], waitingPickup = [];
  todayOrders_(sheet_()).forEach(o => {
    if (o.status === '作廢') return;
    orders++; qty += o.qty; revenue += o.amount;
    if (o.source === 'LINE') lineRevenue += o.amount;
    const picked = o.status === '已領';
    recent.push({ n: o.n, qty: o.qty, amount: o.amount, time: o.time, picked, pickedAt: o.pickedAt,
                  name: o.name, line: o.source === 'LINE', pickCode: o.pickCode });
    if (!picked && o.n <= s.called)
      waitingPickup.push({ n: o.n, qty: o.qty, calledAt: o.calledAt, name: o.name, pickCode: o.pickCode });
  });
  const pendingLine = todayPending_(pendingSheet_())
    .filter(p => p.status === '待付款')
    .map(p => ({ id: p.id, code: p.code, name: p.name, qty: p.qty, amount: p.qty * p.price,
                 left: Math.max(0, Math.ceil((p.created + PAY_MINUTES * 60000 - Date.now()) / 60000)) }));
  return {
    called: s.called, last: s.last, price: price_(),
    orders, qty, revenue, lineRevenue,
    recent: recent.slice(-15).reverse(),
    waitingPickup, pendingLine,
    lineOpen: lineOpen_(), pushCount: pushCount_(), maxPush: maxPush_(),
    lineReady: !!(props_().getProperty('LINE_TOKEN') && props_().getProperty('LINE_LOGIN_CHANNEL_ID'))
  };
}

// 收款出號；pendingId 有值代表收的是 LINE 點的單
function newOrder(pin, qty, paid, expectPrice, pendingId) {
  checkPin_(pin);
  return withLock_(() => {
    let PRICE = price_(), line = null, psh = null;
    if (pendingId) {
      psh = pendingSheet_();
      line = todayPending_(psh).find(p => p.id === pendingId);
      if (!line) throw new Error('找不到這筆 LINE 訂單');
      if (line.status !== '待付款') throw new Error(`${line.code} 目前狀態：${line.status}，請重新整理`);
      qty = line.qty; PRICE = line.price;
    } else if (expectPrice && +expectPrice !== PRICE) {
      throw new Error('單價已改為 $' + PRICE + '，請重新結帳');
    }
    qty = Math.max(1, Math.floor(+qty || 1));
    const amount = qty * PRICE;
    paid = +paid || amount;
    if (paid < amount) throw new Error('收款不足');

    const s = read_();
    s.last += 1;
    write_(s);
    const code = pickCode_();
    sheet_().appendRow([today_(), s.last, qty, amount, paid, paid - amount, now_(), '等待', '', '',
                        line ? 'LINE' : '現場', line ? line.uid : '', line ? line.name : '', code]);
    if (line) {
      psh.getRange(line.row, 10).setValue('已付款');
      psh.getRange(line.row, 11).setValue(s.last);
    }
    return { n: s.last, qty, amount, paid, change: paid - amount, pickCode: code,
             line: !!line, name: line ? line.name : '' };
  });
}

// 店員取消 LINE 待付款單
function cancelPending(pin, id) {
  checkPin_(pin);
  return withLock_(() => {
    const sh = pendingSheet_();
    const p = todayPending_(sh).find(x => x.id === id);
    if (p && p.status === '待付款') sh.getRange(p.row, 10).setValue('取消');
    return staffState(pin);
  });
}

function setLineOpen(pin, open) {
  checkPin_(pin);
  props_().setProperty('LINE_OPEN', open ? '1' : '0');
  return staffState(pin);
}

function callNext(pin) {
  checkPin_(pin);
  return withLock_(() => {
    const s = read_();
    if (s.called >= s.last) throw new Error('沒有等待中的號碼');
    s.called += 1; s.ring += 1;
    write_(s);
    markCalled_(s.called);
    return staffState(pin);
  });
}

function callNumber(pin, n) {
  checkPin_(pin);
  return withLock_(() => {
    const s = read_();
    n = Math.floor(+n);
    if (!(n >= 0 && n <= s.last)) throw new Error('號碼不在範圍內（0～' + s.last + '）');
    s.called = n; s.ring += 1;
    write_(s);
    markCalled_(n);
    return staffState(pin);
  });
}

// 再提醒一次（客人頁面會再震動；不另發 LINE，節省額度）
function reRing(pin) {
  checkPin_(pin);
  return withLock_(() => {
    const s = read_(); s.ring += 1; write_(s);
    return staffState(pin);
  });
}

function setPrice(pin, price) {
  checkPin_(pin);
  price = Math.floor(+price);
  if (!(price > 0 && price < 100000)) throw new Error('單價不正確');
  props_().setProperty('PRICE', String(price));
  return staffState(pin);
}

function voidOrder(pin, n) {
  checkPin_(pin);
  return withLock_(() => {
    const sh = sheet_();
    sh.getRange(findOrder_(sh, n).row, 8).setValue('作廢');
    return staffState(pin);
  });
}

// 標記已領取 / 改回未領（已領過的不能再領）
function setPicked(pin, n, picked) {
  checkPin_(pin);
  return withLock_(() => {
    const sh = sheet_();
    const o = findOrder_(sh, n);
    if (o.status === '作廢') throw new Error(n + ' 號已作廢');
    if (picked && o.status === '已領') throw new Error(`⚠️ ${n} 號已於 ${o.pickedAt} 領取，請勿重複交貨`);
    sh.getRange(o.row, 8).setValue(picked ? '已領' : '等待');
    sh.getRange(o.row, 10).setValue(picked ? now_() : '');
    return staffState(pin);
  });
}

// 在編輯器裡手動執行一次，確認 LINE 設定是否正確
function testLineSetup() {
  const p = props_().getProperties();
  console.log('LINE_LOGIN_CHANNEL_ID：', p.LINE_LOGIN_CHANNEL_ID ? '已設定' : '❌ 未設定');
  console.log('LINE_TOKEN：', p.LINE_TOKEN ? '已設定' : '❌ 未設定');
  console.log('LIFF_ID：', p.LIFF_ID ? p.LIFF_ID : '（未設定，可略過）');
  if (p.LINE_TOKEN) {
    const r = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
      headers: { Authorization: 'Bearer ' + p.LINE_TOKEN }, muteHttpExceptions: true });
    console.log('官方帳號：', r.getResponseCode() === 200 ? JSON.parse(r.getContentText()).displayName : '❌ Token 不正確');
  }
  console.log('本月已發通知：', pushCount_(), '/', maxPush_());
}
