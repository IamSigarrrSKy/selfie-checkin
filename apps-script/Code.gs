/**
 * รับข้อมูลลงเวลาจากแอปเซลฟี่ → เขียนลง Google Sheet + เก็บรูปใน Google Drive
 * ติดตั้ง: ดู README.md ในโฟลเดอร์นี้
 */
const SHEET_ID   = 'ใส่_SHEET_ID_ตรงนี้';   // ส่วนที่อยู่ระหว่าง /d/ กับ /edit ใน URL ของ Sheet
const SHEET_NAME = 'ลงเวลา';
const FOLDER_NAME = 'รูปเช็คอินเซลฟี่';
const TOKEN = '';   // ตั้งรหัสลับได้ ต้องกรอกให้ตรงกันในหน้าตั้งค่าของแอป

const HEADERS = ['ID', 'วันที่', 'เวลา', 'ประเภท', 'ชื่อ', 'รหัสพนักงาน', 'จุดทำงาน', 'ระยะ (ม.)', 'ในพื้นที่',
  'ละติจูด', 'ลองจิจูด', 'ความแม่นยำ (ม.)', 'หมายเหตุ', 'รูป', 'ทดสอบ', 'บันทึกเมื่อ', 'แก้ไขเมื่อ', 'เวลาเดิมก่อนแก้'];
const COL_PHOTO = 14, COL_CREATED = 16;

function doGet() {
  return ContentService.createTextOutput('Selfie check-in: พร้อมใช้งาน');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (TOKEN && req.token !== TOKEN) return out({ ok: false, error: 'รหัสลับไม่ตรงกัน' });
    const sh = getSheet();

    if (req.action === 'ping') return out({ ok: true, sheet: sh.getParent().getName() + ' / ' + sh.getName() });

    if (req.action === 'add') {
      const r = req.record;
      if (findRow(sh, r.id)) return out({ ok: true, duplicate: true });
      let photoUrl = '';
      if (req.photo) {
        const name = r.date + '_' + r.time.replace(/:/g, '') + '_' + (r.type === 'เข้างาน' ? 'in' : 'out') + '.jpg';
        const blob = Utilities.newBlob(Utilities.base64Decode(req.photo.split(',')[1]), 'image/jpeg', name);
        photoUrl = getFolder().createFile(blob).getUrl();
      }
      sh.appendRow(toRow(r, photoUrl, new Date()));
      return out({ ok: true, photoUrl: photoUrl });
    }

    if (req.action === 'update') {
      const r = req.record, row = findRow(sh, r.id);
      if (!row) return out({ ok: false, error: 'ไม่พบรายการใน Sheet' });
      const cur = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      sh.getRange(row, 1, 1, HEADERS.length).setValues([toRow(r, cur[COL_PHOTO - 1], cur[COL_CREATED - 1])]);
      return out({ ok: true });
    }

    if (req.action === 'delete') {
      const row = findRow(sh, req.id);
      if (row) sh.deleteRow(row);
      return out({ ok: true });
    }

    return out({ ok: false, error: 'ไม่รู้จักคำสั่ง ' + req.action });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function toRow(r, photoUrl, created) {
  return [r.id, r.date, r.time, r.type, r.name, r.empId, r.site, r.dist, r.inRange,
    r.lat, r.lng, r.acc, r.note, photoUrl, r.mock, created, r.editedAt, r.origTime];
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.getRange('B:C').setNumberFormat('@');   // เก็บวันที่/เวลาเป็นข้อความ ไม่ให้ Sheet แปลงเอง
  }
  return sh;
}

function getFolder() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function findRow(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** กด Run ฟังก์ชันนี้ครั้งแรกเพื่ออนุญาตสิทธิ์ และสร้างแท็บ "ลงเวลา" */
function setup() {
  getSheet();
  getFolder();
}
