// ══════════════════════════════════════════════════════════
// 英文練習點數機器人（Google Apps Script）
// 1. 把下面三個設定改成你自己的
// 2. 部署 → 新增部署作業 → 網頁應用程式（執行身分：我、存取權：所有人）
// ══════════════════════════════════════════════════════════

const LINE_TOKEN  = '貼上 LINE 的 Channel access token';
const APP_KEY     = '自己想一個連線密碼，例如 apple-2026-xyz';   // 網頁那邊也要填一樣的
const PARENT_CODE = '自己想一個家長暗號，例如 爸爸最帥9527';      // 只有你知道，用來認證家長

// ── 雲端帳本（Google 試算表的「點數」工作表）──
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('點數');
  if (!sh) {
    sh = ss.insertSheet('點數');
    sh.appendRow(['時間', '名字', '點數', '原因', '類型', '來源', '編號']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function props_() { return PropertiesService.getScriptProperties(); }
function r1_(n) { return Math.round((Number(n) || 0) * 10) / 10; } // 點數最小單位 0.1
function rows_() { const v = sheet_().getDataRange().getValues(); v.shift(); return v.filter(r => r[1] !== ''); }
function balances_() {
  const b = {};
  rows_().forEach(r => { const n = String(r[1]); b[n] = r1_((b[n] || 0) + Number(r[2] || 0)); });
  return b;
}
// 經驗值＝讀書練習賺到的點數（跟讀、單字、對話、作業、學科測驗），花掉不會減少
// 不算：簽到(ck/ckm)、寶箱(chest/gift)、任務(wk/mo)、家事(chore)、升級(lv)、家長加的點(give)
function xpRef_(ref) { return !/^(lv|ck|ckm|chest|gift|wk|mo|chore|chorex):/.test(String(ref || '')); }
function earned_() {
  const b = {};
  rows_().forEach(r => { if (String(r[4]) === 'earn' && xpRef_(r[6])) { const n = String(r[1]); b[n] = r1_((b[n] || 0) + Number(r[2] || 0)); } });
  return b;
}
// 每個人簽到過的日期（給每台 iPad 畫簽到印章）
function ckdays_() {
  const b = {};
  rows_().forEach(r => { const ref = String(r[6] || ''); if (/^ck:/.test(ref)) { const n = String(r[1]); (b[n] = b[n] || []).push(ref.slice(ref.lastIndexOf(':') + 1)); } });
  return b;
}
function findName_(name) {
  const known = Object.keys(balances_());
  return known.find(k => k.toLowerCase() === String(name || '').toLowerCase()) || String(name || '');
}
function add_(name, amt, why, type, src, ref) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (ref && rows_().some(r => String(r[6]) === ref && String(r[1]) === name)) return false; // 同一件事不重複給點
    sheet_().appendRow([new Date(), name, amt, why, type, src, ref || '']);
    return true;
  } finally { lock.releaseLock(); }
}
// 家長設定（給點規則、等級公式、家長密碼）：只留一份，最後改的為準
function settings_() { try { return JSON.parse(props_().getProperty('SETTINGS') || 'null'); } catch (_) { return null; } }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ── 網頁來查點數 ──
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== APP_KEY) return json_({ ok: false, error: '連線密碼不對' });
  if (p.sync) return json_(syncGet_(Number(p.since || 0)));
  if (p.img) return json_({ ok: true, data: driveRead_('img_' + p.img) });
  if (p.choreImg) return json_({ ok: true, data: driveRead_('chore_' + p.choreImg) });
  const rows = rows_().slice(-150).reverse().map(r => ({ t: new Date(r[0]).getTime(), name: String(r[1]), amt: Number(r[2]), why: String(r[3]), type: String(r[4]), src: String(r[5]), ref: String(r[6] || '') }));
  return json_({ ok: true, balances: balances_(), earned: earned_(), xpv: 2, ckdays: ckdays_(), rows: rows, group: !!props_().getProperty('GROUP_ID'), settings: settings_(), chores: choreList_() });
}

// ── 網頁送點數過來，或 LINE 傳訊息過來 ──
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (_) {}
  if (body.events) { body.events.forEach(ev => { try { handleEvent_(ev); } catch (err) { console.error(err); } }); return json_({ ok: true }); }
  if (body.key !== APP_KEY) return json_({ ok: false, error: '連線密碼不對' });
  const name = findName_(body.name);
  if (body.action === 'earn') {
    const added = add_(name, r1_(body.amt), body.why || '', 'earn', '練習', body.ref || '');
    return json_({ ok: true, added: added, balances: balances_(), earned: earned_(), xpv: 2 });
  }
  if (body.action === 'adjust') {
    const amt = r1_(body.amt);
    add_(name, amt, body.why || '', amt < 0 ? 'spend' : 'give', '網頁', body.ref || '');
    if (body.notify) push_(String(body.notify).replace('{點數}', balances_()[name] || 0));
    return json_({ ok: true, balances: balances_() });
  }
  if (body.action === 'sync') return json_(syncPut_(body.items || []));
  if (body.action === 'chore') { choreAdd_(body, name); return json_({ ok: true, chores: choreList_() }); }
  if (body.action === 'choreDecide') { choreDecide_(String(body.id), !!body.ok, body.note || ''); return json_({ ok: true, chores: choreList_(), balances: balances_(), earned: earned_(), xpv: 2 }); }
  if (body.action === 'img') {
    driveWrite_('img_' + body.id, String(body.data || ''));
    syncPut_([{ k: 'img', id: String(body.id), u: Date.now(), d: 1 }]);
    return json_({ ok: true });
  }
  if (body.action === 'settings') {
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      const cur = settings_();
      if (body.pinOnly) { const c = cur || { rules: null, at: 0 }; c.pin = body.pin || ''; props_().setProperty('SETTINGS', JSON.stringify(c)); }
      else if (!cur || !cur.rules || Number(body.at) >= Number(cur.at || 0)) props_().setProperty('SETTINGS', JSON.stringify({ rules: body.rules || {}, pin: body.pin || '', at: Number(body.at) || Date.now() }));
    } finally { lock.releaseLock(); }
    return json_({ ok: true, settings: settings_() });
  }
  if (body.action === 'notify') {
    const b = balances_();
    push_(String(body.text || '').replace('{點數}', b[name] || 0));
    return json_({ ok: true, balances: b });
  }
  return json_({ ok: false, error: '不認得的動作' });
}

// ── LINE 指令 ──
const HELP = '🎁 英文練習點數小幫手\n' +
  '・家事 Yuna 洗碗 → 幫 Yuna 登記家事加點（家長）\n' +
  '・家事 → 看家事清單和編號\n' +
  '・取消家事 Yuna → 撤銷 Yuna 最後一筆家事\n' +
  '・點數 → 看每個人有幾點\n' +
  '・總結 → 看今天大家練了多少\n' +
  '・紀錄 Yuna → 看 Yuna 最近的紀錄\n' +
  '・扣 100 Yuna 換Roblox → 扣點（家長）\n' +
  '・加 20 Yuna 幫忙洗碗 → 加點（家長）\n' +
  '・綁定 → 練習通知傳到這個群組（家長）';

function handleEvent_(ev) {
  const src = ev.source || {};
  const uid = src.userId || '';
  const where = src.groupId || src.roomId || src.userId;
  if (ev.type === 'join') {
    if (!props_().getProperty('GROUP_ID')) props_().setProperty('GROUP_ID', where);
    return reply_(ev.replyToken, '大家好！我是英文練習點數小幫手 🎁\n家長請先私訊我「我是家長 暗號」，再到群組打「綁定」。\n\n' + HELP);
  }
  if (ev.type === 'follow') return reply_(ev.replyToken, '嗨！家長請傳「我是家長 暗號」完成設定。\n\n' + HELP);
  if (ev.type !== 'message' || !ev.message || ev.message.type !== 'text') return;

  const text = ev.message.text.trim();
  const admins = JSON.parse(props_().getProperty('ADMINS') || '[]');
  const isAdmin = uid && admins.indexOf(uid) >= 0;
  let m;

  if ((m = text.match(/^我是家長\s*(.+)$/))) {
    if (m[1].trim() !== PARENT_CODE) return reply_(ev.replyToken, '暗號不對喔 🤔');
    if (!uid) return reply_(ev.replyToken, '讀不到你的 LINE 帳號，請先加我為好友，再私訊我一次。');
    if (!isAdmin) { admins.push(uid); props_().setProperty('ADMINS', JSON.stringify(admins)); }
    return reply_(ev.replyToken, '✅ 已設定你是家長，可以加點、扣點了。\n到家庭群組打「綁定」，練習通知就會傳到那裡。');
  }
  if (text === '綁定') {
    if (!isAdmin) return reply_(ev.replyToken, '只有家長可以綁定喔');
    props_().setProperty('GROUP_ID', where);
    return reply_(ev.replyToken, '✅ 綁定好了！之後小朋友完成練習的通知都會傳到這裡。');
  }
  if (/^(點數|查點數|查詢點數)$/.test(text)) {
    const b = balances_(), names = Object.keys(b);
    return reply_(ev.replyToken, names.length ? '🎁 目前點數\n' + names.map(n => '・' + n + '：' + b[n] + ' 點').join('\n') : '還沒有點數紀錄');
  }
  if ((m = text.match(/^紀錄\s*(\S+)?$/))) {
    const who = m[1] ? findName_(m[1]) : '';
    const rs = rows_().filter(r => !who || String(r[1]) === who).slice(-10).reverse();
    if (!rs.length) return reply_(ev.replyToken, '還沒有紀錄');
    return reply_(ev.replyToken, (who ? who + ' ' : '') + '最近 10 筆：\n' + rs.map(r => {
      const d = new Date(r[0]);
      return (d.getMonth() + 1) + '/' + d.getDate() + ' ' + (who ? '' : r[1] + ' ') + (Number(r[2]) > 0 ? '+' : '') + r[2] + ' ' + r[3];
    }).join('\n'));
  }
  if ((m = text.match(/^(扣點|加點|扣|加)\s*(.+)$/))) {
    if (!isAdmin) return reply_(ev.replyToken, '只有家長可以加點、扣點喔 😉');
    const rest = m[2].trim(), num = rest.match(/\d+(?:\.\d+)?/);
    if (!num) return reply_(ev.replyToken, '請寫點數，例如：扣 100 Yuna 換Roblox');
    const toks = rest.replace(num[0], ' ').trim().split(/\s+/).filter(Boolean);
    const known = Object.keys(balances_());
    let name = toks.find(t => known.some(k => k.toLowerCase() === t.toLowerCase())) || toks[0];
    if (!name) return reply_(ev.replyToken, '請寫名字，例如：扣 100 Yuna 換Roblox');
    const why = toks.filter(t => t !== name).join(' ');
    name = findName_(name);
    const sign = m[1].charAt(0) === '扣' ? -1 : 1, amt = r1_(sign * Number(num[0]));
    add_(name, amt, why || (sign < 0 ? '換獎勵' : '額外獎勵'), sign < 0 ? 'spend' : 'give', 'LINE');
    return reply_(ev.replyToken, (sign < 0 ? '🎮 ' : '🎁 ') + name + (sign < 0 ? ' 用了 ' : ' 加了 ') + Math.abs(amt) + ' 點' + (why ? '（' + why + '）' : '') + '\n' + (sign < 0 ? '目前剩下 ' : '目前共 ') + balances_()[name] + ' 點');
  }
  if (/家事/.test(text) && !/^(待確認|待確認家事)$/.test(text)) return reply_(ev.replyToken, isAdmin ? choreCmd_(text) : '只有家長可以登記家事喔 😉');
  if (/^(待確認|待確認家事)$/.test(text)) {
    if (!isAdmin) return reply_(ev.replyToken, '只有家長可以看待確認的家事喔');
    return reply_(ev.replyToken, choreText_());
  }
  if ((m = text.match(/^(通過|退回)\s*(.*)$/))) {
    if (!isAdmin) return reply_(ev.replyToken, '只有家長可以確認家事喔 😉');
    const pend = chorePending_(), ok = m[1] === '通過';
    if (!pend.length) return reply_(ev.replyToken, '目前沒有等待確認的家事 👍');
    const nums = (m[2].match(/^[\d\s,，、]+/) || [''])[0].split(/[\s,，、]+/).filter(Boolean).map(Number);
    const note = m[2].slice(((m[2].match(/^[\d\s,，、]+/) || [''])[0]).length).trim();
    if (!ok && !nums.length) return reply_(ev.replyToken, '要退回哪一項？例如：退回 2 碗還有油');
    const pick = nums.length ? nums.map(n => pend[n - 1]).filter(Boolean) : pend;
    if (!pick.length) return reply_(ev.replyToken, '找不到那個編號，先打「待確認」看看清單');
    pick.forEach(c => choreDecide_(c.id, ok, note));
    const b = balances_();
    return reply_(ev.replyToken, (ok ? '✅ 已通過：\n' : '🔁 已退回：\n') + pick.map(c => '・' + c.name + ' ' + c.chore + (ok ? ' +' + c.pts + ' 點（共 ' + (b[c.name] || 0) + ' 點）' : '')).join('\n') + (note && !ok ? '\n留言：' + note : ''));
  }
  if (/^(總結|今日總結|今天)$/.test(text)) return reply_(ev.replyToken, summaryText_());
  if (/^(說明|help|\?|？)$/i.test(text)) return reply_(ev.replyToken, HELP);
  // 其他聊天內容不理會，免得在群組裡吵
}

function reply_(token, text) {
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + LINE_TOKEN },
    payload: JSON.stringify({ replyToken: token, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] })
  });
}
function push_(text) {
  const to = props_().getProperty('GROUP_ID');
  if (!to || !text) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + LINE_TOKEN },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] })
  });
}

// ── 📅 每天一則總結（iPad 家長專區選「每天一則總結」才會傳）──
const TZ = 'Asia/Taipei';
function dayOf_(d) { return Utilities.formatDate(new Date(d), TZ, 'yyyy-MM-dd'); }
function summaryText_() {
  const now = new Date(), day = dayOf_(now), bal = balances_(), today = {};
  rows_().forEach(r => {
    const n = String(r[1]); if (!(n in today)) today[n] = { earn: 0, spend: 0, notes: [] };
    if (dayOf_(r[0]) !== day) return;
    const t = today[n], a = Number(r[2] || 0), why = String(r[3] || '');
    if (a > 0) t.earn = r1_(t.earn + a); else t.spend = r1_(t.spend - a);
    let m;
    if ((m = why.match(/每日簽到 連續第 (\d+) 天/))) t.notes.push('簽到第 ' + m[1] + ' 天🔥');
    if ((m = why.match(/升到 (Lv\d+ \S+)/))) t.notes.push('⬆️ 升到 ' + m[1]);
    if (/完成「/.test(why)) t.notes.push('🏁 ' + why.replace(/🏁/g, '').replace(/（.*）/, '').trim());
  });
  const names = Object.keys(today), pn = chorePending_().length;
  const tail = pn ? '\n🧹 還有 ' + pn + ' 項家事等你確認（打「待確認」）' : '';
  if (!names.length) return '📅 ' + Utilities.formatDate(now, TZ, 'M/d') + ' 還沒有任何點數紀錄';
  return '📅 ' + Utilities.formatDate(now, TZ, 'M/d') + ' 英文練習總結\n' + names.map(n => {
    const t = today[n];
    if (!t.earn && !t.spend) return '・' + n + '：今天還沒練習喔！（共 ' + (bal[n] || 0) + ' 點）';
    return '・' + n + '：今天 +' + t.earn + (t.spend ? '、用了 ' + t.spend : '') + (t.notes.length ? '（' + t.notes.slice(0, 3).join('、') + '）' : '') + '\n　目前共 ' + (bal[n] || 0) + ' 點';
  }).join('\n') + tail;
}
// 每小時檢查一次：到了設定的時間、今天還沒傳過，就傳總結
function dailySummary() {
  const st = settings_(), r = (st && st.rules) || {};
  if (r.notify !== 'daily') return;
  const now = new Date(), hour = Number(r.notifyHour || 20), day = dayOf_(now);
  if (Number(Utilities.formatDate(now, TZ, 'H')) < hour) return;
  if (props_().getProperty('LAST_SUMMARY') === day) return;
  props_().setProperty('LAST_SUMMARY', day);
  push_(summaryText_());
}
// 在編輯器選「setupDaily」按「執行」一次（只要一次），之後 Google 每小時會自動檢查
function setupDaily() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailySummary').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailySummary').timeBased().everyHours(1).create();
  Logger.log('好了！每天總結已經設定好。');
}

// ── ☁️ 全部資料同步（課文、單字本、成績、學習紀錄、頭像、單字卡圖片）──
// 「同步」工作表一列一筆：類型、編號、更新時間、資料(JSON)、刪除、伺服器時間
// 太大的資料和圖片放在雲端硬碟的「英文練習同步資料」資料夾
function syncSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('同步');
  if (!sh) { sh = ss.insertSheet('同步'); sh.appendRow(['類型', '編號', '更新時間', '資料', '刪除', '伺服器時間']); sh.setFrozenRows(1); }
  return sh;
}
function folder_() {
  const id = props_().getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (_) {} }
  const f = DriveApp.createFolder('英文練習同步資料');
  props_().setProperty('FOLDER_ID', f.getId());
  return f;
}
function driveWrite_(name, text) {
  const f = folder_(), it = f.getFilesByName(name);
  if (it.hasNext()) { it.next().setContent(text); while (it.hasNext()) it.next().setTrashed(true); }
  else f.createFile(name, text, MimeType.PLAIN_TEXT);
}
function driveRead_(name) {
  const it = folder_().getFilesByName(name);
  return it.hasNext() ? it.next().getBlob().getDataAsString() : '';
}
function syncGet_(since) {
  const now = Date.now(), v = syncSheet_().getDataRange().getValues(); v.shift();
  const items = [];
  v.forEach(r => {
    if (Number(r[5]) <= since) return;
    let d = null, raw = String(r[3] || '');
    if (raw.indexOf('@drive:') === 0) raw = driveRead_(raw.slice(7));
    try { d = raw ? JSON.parse(raw) : null; } catch (_) {}
    items.push({ k: String(r[0]), id: String(r[1]), u: Number(r[2]), d: d, del: !!r[4], su: Number(r[5]) });
  });
  return { ok: true, items: items, now: now };
}
function syncPut_(items) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = syncSheet_(), v = sh.getDataRange().getValues(), at = {}, now = Date.now(), add = [], addAt = {};
    for (let i = 1; i < v.length; i++) at[v[i][0] + '|' + v[i][1]] = i + 1;
    items.forEach(it => {
      const key = it.k + '|' + it.id, row = at[key];
      if (row && Number(v[row - 1][2]) > Number(it.u)) return; // 雲端的比較新，不蓋掉
      let raw = it.del ? '' : JSON.stringify(it.d);
      if (raw.length > 45000) { const fn = 'doc_' + it.k + '_' + it.id; driveWrite_(fn, raw); raw = '@drive:' + fn; }
      const vals = [it.k, it.id, Number(it.u) || now, raw, it.del ? 1 : '', now];
      if (row) sh.getRange(row, 1, 1, 6).setValues([vals]);
      else if (key in addAt) add[addAt[key]] = vals;
      else { addAt[key] = add.length; add.push(vals); }
    });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 6).setValues(add);
    return { ok: true, now: now };
  } finally { lock.releaseLock(); }
}
// 在編輯器選「setupSync」按「執行」一次：建立「同步」工作表和雲端硬碟資料夾，並授權使用雲端硬碟
function setupSync() { syncSheet_(); folder_(); Logger.log('好了！全部資料同步已經準備好。'); }

// 在編輯器裡按「執行 → setup」一次，讓 Google 要求授權並建立「點數」工作表
function setup() { sheet_(); Logger.log('好了！「點數」工作表已建立。'); }

// ── 🧹 家事：小朋友在 App 回報，家長在 App 或 LINE 確認後才加點 ──
// 「家事」工作表一列一筆：編號、時間、名字、家事、點數、狀態(wait/ok/back)、留言、有照片、確認時間
function choreSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('家事');
  if (!sh) { sh = ss.insertSheet('家事'); sh.appendRow(['編號', '時間', '名字', '家事', '點數', '狀態', '留言', '照片', '確認時間']); sh.setFrozenRows(1); }
  return sh;
}
function choreRows_() { const v = choreSheet_().getDataRange().getValues(); v.shift(); return v.filter(r => r[0] !== ''); }
function choreObj_(r) { return { id: String(r[0]), t: new Date(r[1]).getTime(), name: String(r[2]), chore: String(r[3]), pts: r1_(r[4]), st: String(r[5] || 'wait'), note: String(r[6] || ''), ph: r[7] ? 1 : 0, dt: r[8] ? new Date(r[8]).getTime() : 0 }; }
function choreList_() { const since = Date.now() - 45 * 864e5; return choreRows_().map(choreObj_).filter(c => c.t >= since || c.st === 'wait').slice(-200); }
function chorePending_() { return choreRows_().map(choreObj_).filter(c => c.st === 'wait').sort((a, b) => a.t - b.t); }
function choreAdd_(b, name) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const id = String(b.id || Date.now());
    if (choreRows_().some(r => String(r[0]) === id)) return;
    if (b.photo) driveWrite_('chore_' + id, String(b.photo));
    choreSheet_().appendRow([id, new Date(Number(b.t) || Date.now()), name, String(b.chore || ''), r1_(b.pts), 'wait', '', b.photo ? 1 : '', '']);
  } finally { lock.releaseLock(); }
}
function choreDecide_(id, ok, note) {
  const sh = choreSheet_(), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) !== id) continue;
    if (String(v[i][5]) === 'ok') return; // 已經通過過了
    sh.getRange(i + 1, 6, 1, 4).setValues([[ok ? 'ok' : 'back', note || '', v[i][7], new Date()]]);
    if (ok) add_(String(v[i][2]), r1_(v[i][4]), '🧹 家事：' + v[i][3], 'earn', '家事', 'chore:' + id);
    return;
  }
}
function choreText_() {
  const p = chorePending_();
  if (!p.length) return '🧹 目前沒有等待確認的家事 👍';
  return '🧹 等你確認的家事\n' + p.map((c, i) => (i + 1) + '. ' + c.name + ' ' + c.chore + ' +' + c.pts + '（' + Utilities.formatDate(new Date(c.t), TZ, 'M/d H:mm') + '）' + (c.ph ? '📷' : '')).join('\n') +
    '\n\n打「通過」全部通過\n打「通過 1 3」只通過第 1、3 項\n打「退回 2 原因」請他重做' + (p.some(c => c.ph) ? '\n📷 有照片的可以到 App 家長專區看' : '');
}

// ── 🧹 家長在 LINE 直接登記家事：「家事 Yuna 洗碗」「Yuna 完成家事 洗碗 倒垃圾」「家事 Yuna 1」「家事 Yuna 搬東西 2」 ──
const CHORE_DEF = '🍽️ 洗碗, 2, 每天, 1\n🗑️ 倒垃圾, 1, 每天, 1\n🧹 整理房間, 1, 每天, 1\n🧺 洗衣服, 1, 每天, 1\n👚 收衣服, 1, 每天, 1\n✨ 其他, 1, 每天, 0';
function choreDefs_() {
  const st = settings_(), src = (st && st.rules && st.rules.chores) || CHORE_DEF;
  return String(src).split(/\n+/).map(l => l.split(/[,，]/).map(x => x.trim())).filter(a => a[0] && isFinite(Number(a[1])) && Number(a[1]) > 0).map(a => {
    const m = a[0].match(/^(\S+)\s+(.+)$/), emo = m && !/[一-鿿A-Za-z0-9]/.test(m[1]);
    return { ic: emo ? m[1] : '🧽', name: emo ? m[2] : a[0], pts: r1_(a[1]), max: Math.max(0, Math.round(Number(a[3]) || 0)) };
  });
}
function choreListText_() {
  return '🧹 家事清單\n' + choreDefs_().map((d, i) => (i + 1) + '. ' + d.ic + ' ' + d.name + ' +' + d.pts).join('\n') +
    '\n\n登記方式：\n家事 Yuna 洗碗\n家事 Yuna 1 3（第 1、3 項）\n家事 Yuna 搬東西 2（清單沒有的，自己寫點數）\n取消家事 Yuna（撤銷最後一筆）';
}
function weekDays_() { const now = new Date(), dow = Number(Utilities.formatDate(now, TZ, 'u')) || 1, out = {}; for (let i = 0; i < dow; i++) out[dayOf_(now.getTime() - i * 864e5)] = 1; return out; }
function choreCmd_(text) {
  let s = text.replace(/完成|做完|做了/g, ' ').trim();
  const cancel = /^取消/.test(s); s = s.replace(/^取消/, '').replace(/家事/g, ' ').trim();
  const toks = s.split(/\s+/).filter(Boolean);
  if (!toks.length) return choreListText_();
  const known = Object.keys(balances_());
  let ni = toks.findIndex(t => known.some(k => k.toLowerCase() === t.toLowerCase()));
  if (ni < 0) ni = toks.findIndex(t => /^[A-Za-z]/.test(t));
  if (ni < 0) ni = 0;
  const name = findName_(toks[ni]), rest = toks.filter((_, i) => i !== ni);
  if (cancel) {
    const sh = choreSheet_(), v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) {
      if (String(v[i][2]) !== name || String(v[i][5]) !== 'ok') continue;
      sh.getRange(i + 1, 6, 1, 4).setValues([['cancel', '家長取消', v[i][7], new Date()]]);
      add_(name, -r1_(v[i][4]), '🧹 取消家事：' + v[i][3], 'earn', '家事', 'chorex:' + v[i][0]);
      return '↩️ 已取消 ' + name + ' 的「' + v[i][3] + '」-' + r1_(v[i][4]) + ' 點\n目前共 ' + (balances_()[name] || 0) + ' 點';
    }
    return name + ' 沒有可以取消的家事';
  }
  if (!rest.length) return '要登記哪一項家事？例如：家事 ' + name + ' 洗碗\n\n' + choreListText_();
  const defs = choreDefs_(), picks = [];
  for (let i = 0; i < rest.length; i++) {
    const t = rest[i];
    if (/^\d+$/.test(t)) { const d = defs[Number(t) - 1]; if (d) picks.push({ name: d.name, pts: d.pts, max: d.max }); else return '清單沒有第 ' + t + ' 項\n\n' + choreListText_(); continue; }
    const d = defs.find(x => x.name === t) || defs.find(x => x.name.indexOf(t) >= 0 || t.indexOf(x.name) >= 0);
    if (d) { picks.push({ name: d.name, pts: d.pts, max: d.max }); continue; }
    const other = defs.find(x => x.name === '其他') || { pts: 1 };
    let pts = other.pts;
    if (/^\d+(\.\d+)?$/.test(rest[i + 1] || '')) { pts = r1_(rest[i + 1]); i++; }
    picks.push({ name: t, pts: pts, max: 0 });
  }
  const today = dayOf_(new Date()), rows = choreRows_().map(choreObj_).filter(c => c.name === name && c.st === 'ok'), lines = [];
  picks.forEach((p, k) => {
    const n = rows.filter(c => c.chore === p.name && dayOf_(c.t) === today).length + picks.slice(0, k).filter(q => q.name === p.name).length;
    const id = 'L' + Date.now() + k;
    choreSheet_().appendRow([id, new Date(), name, p.name, p.pts, 'ok', 'LINE 登記', '', new Date()]);
    add_(name, p.pts, '🧹 家事：' + p.name, 'earn', '家事', 'chore:' + id);
    lines.push(p.name + ' +' + p.pts + (p.max && n >= p.max ? '（今天第 ' + (n + 1) + ' 次）' : ''));
  });
  const wk = weekDays_(), wn = choreRows_().map(choreObj_).filter(c => c.name === name && c.st === 'ok' && wk[dayOf_(c.t)]).length;
  const st = settings_(), tw = st && st.rules && st.rules.tasks && st.rules.tasks.wk && st.rules.tasks.wk.chore, goal = tw ? Number(tw[0]) : 3;
  return '✅ ' + name + '：' + lines.join('、') + '\n目前共 ' + (balances_()[name] || 0) + ' 點' + (goal > 0 ? '\n🗓️ 本週家事 ' + wn + ' / ' + goal + (wn >= goal ? ' 達成！可以到 App 打卡領獎勵' : '') : '');
}
