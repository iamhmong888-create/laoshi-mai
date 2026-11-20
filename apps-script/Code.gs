/**
 * เหล่าซือไหม : ระบบรับการบ้านวิดีโอ (Google Apps Script)
 * - เก็บวิดีโอไว้ใน Google Drive ของครู (โฟลเดอร์ "เหล่าซือไหม - การบ้านวิดีโอ")
 * - เก็บรายชื่อผู้ส่งไว้ใน Google Sheet ในโฟลเดอร์เดียวกัน
 * - รองรับวิดีโอไฟล์ใหญ่ โดยส่งเป็นช่วง ๆ (resumable upload)
 */

// ===== เปลี่ยนรหัสผ่านหน้าครูตรงนี้ =====
const ADMIN_PASSWORD = 'laoshi1234';
const FOLDER_NAME = 'เหล่าซือไหม - การบ้านวิดีโอ';
const SHEET_NAME = 'เหล่าซือไหม - รายชื่อส่งการบ้าน';
const HEADERS = ['เวลาส่ง', 'ชื่อ-สกุล', 'ชื่อไฟล์', 'fileId', 'ขนาด (MB)', 'ตรวจแล้ว'];

/** กดรันฟังก์ชันนี้ 1 ครั้งก่อน Deploy เพื่ออนุญาตสิทธิ์และสร้างโฟลเดอร์/ชีต */
function setup() {
  const folder = getFolder_();
  const ss = getSpreadsheet_();
  Logger.log('โฟลเดอร์: ' + folder.getUrl());
  Logger.log('ชีต: ' + ss.getUrl());
}

function doGet() {
  return json_({ ok: true, app: 'laoshi-mai' });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad request' }); }
  try {
    switch (d.action) {
      case 'init':   return json_(init_(d));
      case 'chunk':  return json_(chunk_(d));
      case 'list':   auth_(d); return json_(list_());
      case 'mark':   auth_(d); return json_(mark_(d));
      case 'delete': auth_(d); return json_(remove_(d));
      default:       return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- student upload ---------- */
function init_(d) {
  const name = clean_(d.name, 80);
  if (name.length < 3) throw new Error('กรุณาใส่ชื่อ-สกุล');
  const size = Number(d.size);
  if (!(size > 0) || size > 300 * 1048576) throw new Error('ขนาดไฟล์ไม่ถูกต้อง');
  const mime = /^video\/[\w.+-]+$/.test(d.mime || '') ? d.mime : 'video/mp4';
  const ext = (String(d.filename || '').match(/\.(\w{2,5})$/) || [, 'mp4'])[1];
  const stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd_HHmm');
  const fileName = stamp + '_' + name.replace(/[\\/:*?"<>|]/g, '') + '.' + ext;

  const res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ name: fileName, mimeType: mime, parents: [getFolder_().getId()], description: name }),
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        'X-Upload-Content-Type': mime,
        'X-Upload-Content-Length': String(size)
      },
      muteHttpExceptions: true
    });
  const h = res.getHeaders();
  const url = h.Location || h.location;
  if (res.getResponseCode() !== 200 || !url) throw new Error('เริ่มอัปโหลดไม่ได้ (' + res.getResponseCode() + ')');
  return { ok: true, uploadUrl: url };
}

function chunk_(d) {
  if (!/^https:\/\/www\.googleapis\.com\/upload\/drive\//.test(d.uploadUrl || '')) throw new Error('bad upload url');
  const bytes = Utilities.base64Decode(d.data);
  const res = UrlFetchApp.fetch(d.uploadUrl, {
    method: 'put',
    contentType: d.mime || 'application/octet-stream',
    payload: bytes,
    headers: { 'Content-Range': 'bytes ' + d.start + '-' + d.end + '/' + d.total },
    muteHttpExceptions: true,
    followRedirects: false
  });
  const code = res.getResponseCode();
  if (code === 308) return { ok: true, done: false };
  if (code === 200 || code === 201) {
    const id = JSON.parse(res.getContentText()).id;
    const file = DriveApp.getFileById(id);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* โดเมนบางแห่งไม่อนุญาต ไม่เป็นไร */ }
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      getSheet_().appendRow([new Date(), clean_(d.name, 80), file.getName(), id, +(Number(d.total) / 1048576).toFixed(1), false]);
    } finally { lock.releaseLock(); }
    return { ok: true, done: true };
  }
  throw new Error('อัปโหลดผิดพลาด (' + code + ')');
}

/* ---------- teacher ---------- */
function auth_(d) {
  if (String(d.pw || '') !== ADMIN_PASSWORD) throw new Error('รหัสผ่านไม่ถูกต้อง');
}

function list_() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getValues() : [];
  const items = rows.filter(r => r[3]).map(r => ({
    time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    name: String(r[1]), file: String(r[2]), id: String(r[3]), mb: Number(r[4]) || 0, checked: r[5] === true
  })).reverse();
  return { ok: true, items: items, sheetUrl: getSpreadsheet_().getUrl(), folderUrl: getFolder_().getUrl() };
}

function mark_(d) {
  const r = findRow_(d.id);
  getSheet_().getRange(r, 6).setValue(!!d.checked);
  return { ok: true };
}

function remove_(d) {
  const r = findRow_(d.id);
  try { DriveApp.getFileById(d.id).setTrashed(true); } catch (e) {}
  getSheet_().deleteRow(r);
  return { ok: true };
}

/* ---------- helpers ---------- */
function findRow_(id) {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) throw new Error('ไม่พบรายการ');
  const ids = sh.getRange(2, 4, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  throw new Error('ไม่พบรายการ');
}

function getFolder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}

function getSpreadsheet_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create(SHEET_NAME);
  const sh = ss.getSheets()[0];
  sh.appendRow(HEADERS);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#b3161b').setFontColor('#ffffff');
  try { DriveApp.getFileById(ss.getId()).moveTo(getFolder_()); } catch (e) {}
  p.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getSheet_() { return getSpreadsheet_().getSheets()[0]; }

function clean_(s, max) { return String(s || '').replace(/[<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, max); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
