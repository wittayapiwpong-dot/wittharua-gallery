/**
 * ระบบหลังบ้าน แกลเลอรีพระเครื่อง วิทย์ ท่าเรือ
 * วางโค้ดนี้ใน Google Sheet: ส่วนขยาย → Apps Script
 * ตั้งรหัสผ่านที่: การตั้งค่าโปรเจ็กต์ → พร็อพเพอร์ตี้ของสคริปต์ → ADMIN_PASSWORD
 */

const SHEET_ITEMS = 'ข้อมูลพระ';
const KEYS = ['name', 'cat', 'era', 'material', 'size', 'source', 'story', 'front', 'back', 'show', 'status', 'extra'];
const FOLDER_NAME = 'วิทย์ ท่าเรือ – รูปพระ';

function doGet() {
  return out_({ ok: true, msg: 'ระบบพร้อมใช้งาน' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'คำขอไม่ถูกต้อง' }); }

  const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!pw) return out_({ ok: false, error: 'ยังไม่ได้ตั้งรหัสผ่าน ADMIN_PASSWORD ใน Apps Script' });
  if (req.password !== pw) { Utilities.sleep(1500); return out_({ ok: false, error: 'wrong_password' }); }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (req.action) {
      case 'login': return out_({ ok: true });
      case 'list': return out_({ ok: true, items: list_() });
      case 'add': return out_({ ok: true, row: save_(null, req.item) });
      case 'update': return out_({ ok: true, row: save_(req.row, req.item, req.expectName) });
      case 'delete': del_(req.row, req.expectName); return out_({ ok: true });
      default: return out_({ ok: false, error: 'ไม่รู้จักคำสั่งนี้' });
    }
  } catch (err) {
    console.error('doPost ' + req.action + ' failed: ' + (err && err.stack || err));
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ITEMS);
  if (!sh) throw new Error('ไม่พบแท็บ ' + SHEET_ITEMS);
  return sh;
}

function list_() {
  const sh = sheet_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, KEYS.length).getDisplayValues()
    .map(function (r, i) { const o = { row: i + 2 }; KEYS.forEach(function (k, j) { o[k] = r[j]; }); return o; })
    .filter(function (o) { return o.name; });
}

function check_(sh, row, expectName) {
  if (!row || row < 2 || row > sh.getLastRow()) throw new Error('ไม่พบรายการนี้ในชีต');
  if (expectName !== undefined && sh.getRange(row, 1).getDisplayValue() !== expectName) {
    throw new Error('ข้อมูลในชีตเปลี่ยนไปแล้ว กดโหลดรายการใหม่แล้วลองอีกครั้ง');
  }
}

function save_(row, item, expectName) {
  const sh = sheet_();
  if (!item || !String(item.name || '').trim()) throw new Error('ต้องใส่ชื่อพระ');
  const it = Object.assign({}, item);
  ['front', 'back', 'extra'].forEach(function (s) {
    if (it[s + 'Data']) it[s] = upload_(it[s + 'Data'], String(it.name).trim() + ' (' + (s === 'front' ? 'หน้า' : s === 'back' ? 'หลัง' : 'เพิ่มเติม') + ')');
    delete it[s + 'Data'];
  });
  const values = [KEYS.map(function (k) { return String(it[k] == null ? '' : it[k]).trim(); })];
  let r = row;
  if (r) check_(sh, r, expectName); else r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, KEYS.length).setNumberFormat('@').setValues(values);
  return r;
}

function del_(row, expectName) {
  const sh = sheet_();
  check_(sh, row, expectName);
  sh.deleteRow(row);
}

function folder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}

function upload_(dataUrl, name) {
  const m = String(dataUrl).match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!m) throw new Error('ไฟล์รูปไม่ถูกต้อง');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name + '.jpg');
  const file = folder_().createFile(blob);
  // โฟลเดอร์แชร์แบบทุกคนที่มีลิงก์อยู่แล้ว ถ้าตั้งสิทธิ์รายไฟล์ไม่ได้ก็ข้ามไป
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { console.warn('setSharing skipped: ' + e.message); }
  return 'https://drive.google.com/file/d/' + file.getId() + '/view';
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// กดเรียกใช้ฟังก์ชันนี้หนึ่งครั้งเพื่อให้สิทธิ์ Google Drive (ใช้เก็บรูปพระ)
function authorizeDrive() {
  Logger.log(folder_().getName());
}
