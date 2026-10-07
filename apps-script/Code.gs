/**
 * รับข้อมูลเช็คอินจากแอปเซลฟี่ → เขียนลง Google Sheet + เก็บรูปใน Google Drive
 * ติดตั้ง: ดู README.md ในโฟลเดอร์นี้ แล้วกด Run ฟังก์ชัน setup() หนึ่งครั้งเพื่อสร้างชีตทั้งหมด
 */
const SHEET_ID    = '';   // เว้นว่างได้ถ้าเปิด Apps Script จากเมนู ส่วนขยาย ของ Sheet นั้นเอง
const TOKEN       = '';   // ตั้งรหัสลับได้ ต้องกรอกให้ตรงกันในหน้าตั้งค่าของแอป
const FOLDER_NAME = 'รูปเช็คอินเซลฟี่';

const LOG = 'เช็คอิน', LEGACY_LOG = 'ลงเวลา', DAILY = 'สรุปรายวัน', MONTHLY = 'สรุปรายเดือน', SITES = 'จุดทำงาน',
  LEAVE = 'การลา', LEAVE_SUM = 'สรุปการลา';

const HEADERS = ['ID', 'วันที่', 'เวลา', 'ประเภท', 'ชื่อ', 'รหัสพนักงาน', 'สถานที่', 'ระยะจากจุดใกล้สุด (ม.)', 'อยู่ในจุดที่บันทึก',
  'ละติจูด', 'ลองจิจูด', 'ความแม่นยำ (ม.)', 'หมายเหตุ', 'รูป', 'ทดสอบ', 'บันทึกเมื่อ', 'แก้ไขเมื่อ', 'เวลาเดิมก่อนแก้',
  'ภาค', 'ลิงก์รูป', 'ลิงก์'];
const COL_CREATED = 16, COL_REGION = 19, COL_PHOTO_URL = 20;

// แต่ละภาคมีแท็บของตัวเอง ระบบจัดรายการเข้าภาคตามจุดทำงาน (หรือชื่อจังหวัดในชื่อจุด ถ้าเพิ่มจุดใหม่ในแอป)
const REGIONS = [
  { name: 'ส่วนกลาง', provinces: ['กรุงเทพ', 'Bangkok', 'นนทบุรี', 'ปทุมธานี', 'สมุทรปราการ'] },
  { name: 'ภาคเหนือ', provinces: ['เชียงใหม่', 'เชียงราย', 'ลำพูน', 'ลำปาง', 'แม่ฮ่องสอน', 'พะเยา', 'แพร่', 'น่าน', 'อุตรดิตถ์', 'พิษณุโลก', 'สุโขทัย', 'ตาก', 'มช'] },
  { name: 'ภาคตะวันออกเฉียงเหนือ', provinces: ['ขอนแก่น', 'อุดรธานี', 'นครราชสีมา', 'อุบลราชธานี', 'มหาสารคาม', 'ร้อยเอ็ด', 'กาฬสินธุ์', 'สกลนคร', 'นครพนม', 'หนองคาย', 'เลย', 'ชัยภูมิ', 'บุรีรัมย์', 'สุรินทร์', 'ศรีสะเกษ', 'มข'] },
  { name: 'ภาคใต้', provinces: ['สงขลา', 'หาดใหญ่', 'ปัตตานี', 'ภูเก็ต', 'สุราษฎร์ธานี', 'นครศรีธรรมราช', 'ตรัง', 'พัทลุง', 'สตูล', 'ยะลา', 'นราธิวาส', 'กระบี่', 'ชุมพร', 'ม.อ.'] },
  { name: 'ภาคตะวันออก', provinces: ['ระยอง', 'ชลบุรี', 'จันทบุรี', 'ตราด', 'ฉะเชิงเทรา', 'ปราจีนบุรี', 'สระแก้ว'] }
];
const OTHER_REGION = 'อื่นๆ';

const SITE_LIST = [
  ['อุทยานวิทยาศาสตร์และเทคโนโลยี ม.เชียงใหม่', 'เชียงใหม่', 'ภาคเหนือ', 18.76467, 98.93700, 300],
  ['อุทยานวิทยาศาสตร์ ม.ขอนแก่น', 'ขอนแก่น', 'ภาคตะวันออกเฉียงเหนือ', 16.45588, 102.81942, 300],
  ['อุทยานวิทยาศาสตร์ ม.สงขลานครินทร์', 'สงขลา', 'ภาคใต้', 7.02194, 100.55542, 300],
  ['ปส. ภาคตะวันออก (ศาลากลาง จ.ระยอง)', 'ระยอง', 'ภาคตะวันออก', 12.70726, 101.18380, 300],
  ['สำนักงานปรมาณูเพื่อสันติ (ส่วนกลาง กรุงเทพฯ)', 'กรุงเทพฯ', 'ส่วนกลาง', 13.85435, 100.56617, 300]
];

function regionOf(site) {
  site = String(site || '');
  const known = SITE_LIST.find(s => s[0] === site);
  if (known) return known[2];
  const hit = REGIONS.find(r => r.provinces.some(p => site.indexOf(p) >= 0));
  return hit ? hit.name : OTHER_REGION;
}

const C = { head: '#0e5a52', headFg: '#ffffff', band: '#f3f6f4', warn: '#fbecd3', edit: '#e3edfb', muted: '#8a9a96', title: '#14211e' };

/* ===================== Web app ===================== */

function doGet() {
  return ContentService.createTextOutput('Selfie check-in: พร้อมใช้งาน');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (TOKEN && req.token !== TOKEN) return out({ ok: false, error: 'รหัสลับไม่ตรงกัน' });
    const sh = logSheet();

    if (req.action === 'ping') return out({ ok: true, sheet: sh.getParent().getName() + ' / ' + sh.getName() });
    if (req.action === 'leave_add') return out(addLeave(req.leave));
    if (req.action === 'leave_delete') return out(deleteLeave(req.id));

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
      if (!row) return out({ ok: true, missing: true });   // ถูกลบออกจาก Sheet ไปแล้ว: ไม่ต้องทำอะไร
      // แอปแก้ได้แค่เวลาเช็คเอาท์: เขียนเฉพาะวันที่/เวลา ช่องอื่นที่หัวหน้าแก้ใน Sheet ไม่ถูกทับ
      if (r.type !== 'เช็คเอาท์' && r.type !== 'ออกงาน') return out({ ok: true, skipped: true });
      sh.getRange(row, 2, 1, 2).setValues([[dateSerial(r.date), timeSerial(r.time)]]);
      return out({ ok: true });
    }

    if (req.action === 'delete') {
      const row = findRow(sh, req.id);
      if (row) {
        // ย้ายรูปของรายการนี้ไปถังขยะใน Drive (กู้คืนได้ 30 วัน)
        const url = String(sh.getRange(row, COL_PHOTO_URL).getValue() || '');
        const fid = (url.match(/\/d\/([^/]+)/) || [])[1];
        if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) { /* ไฟล์ถูกลบไปแล้ว */ } }
        sh.deleteRow(row);
      }
      return out({ ok: true });
    }

    return out({ ok: false, error: 'ไม่รู้จักคำสั่ง ' + req.action });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// วันที่ "2026-10-01" และเวลา "08:30:00" แปลงเป็นตัวเลขวันที่/เวลาของ Sheet เอง ไม่พึ่งการแปลงอัตโนมัติ
function dateSerial(s) {
  const m = String(s).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  return m ? (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000 : s;
}
function timeSerial(s) {
  const m = String(s).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  return m ? (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400 : s;
}

function toRow(r, photoUrl, created) {
  const photo = photoUrl ? '=HYPERLINK("' + photoUrl + '","ดูรูป")' : '';
  return [r.id, dateSerial(r.date), timeSerial(r.time), r.type, r.name, '', r.site, r.dist, r.inRange,
    r.lat, r.lng, '', r.note, photo, r.mock, created, "", "",   // ไม่บันทึกประวัติการแก้ไขลง Sheet
    regionOf(r.site), photoUrl || '', r.link || 'A'];
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

function book() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function logSheet() {
  return book().getSheetByName(LOG) || buildLog(book());
}

function getFolder() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

/* ===================== เมนูลบข้อมูลใน Sheet ===================== */

/** เมนู "ระบบเช็คอิน" ขึ้นเองทุกครั้งที่เปิด Sheet */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบเช็คอิน')
    .addItem('ลบแถวที่เลือก (แท็บเช็คอิน / การลา)', 'deleteSelectedRows')
    .addItem('ลบข้อมูลทั้งเดือน…', 'deleteMonth')
    .addSeparator()
    .addItem('เก็บกวาดแถวที่ลบไม่หมด', 'cleanupBlankRows')
    .addToUi();
}

// สคริปต์ที่ไม่ได้สร้างจากเมนู ส่วนขยาย ของ Sheet: onOpen ปกติไม่ทำงาน ต้องติดตั้ง trigger เปิดไฟล์ให้แทน
function installMenuTrigger(ss) {
  if (SpreadsheetApp.getActiveSpreadsheet()) return;   // สคริปต์ผูกกับ Sheet อยู่แล้ว เมนูขึ้นเอง
  const handlers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  if (handlers.indexOf('onOpen') < 0) ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create();
  if (handlers.indexOf('onSheetEdit') < 0) ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(ss).onEdit().create();
}

// หัวหน้าแก้หรือเพิ่มแถวเองในแท็บเช็คอิน/การลา: เติม ID ภาค ประเภท และแปลงวันที่/เวลาที่พิมพ์ให้เป็นค่าวันที่จริง
function onSheetEdit(e) {
  const sh = e && e.range && e.range.getSheet();
  if (!sh || [LOG, LEAVE].indexOf(sh.getName()) < 0) return;
  const first = Math.max(2, e.range.getRow()), last = e.range.getLastRow();
  if (last < first) return;
  const n = last - first + 1, width = sh.getName() === LOG ? HEADERS.length : LEAVE_HEADERS.length;
  const vals = sh.getRange(first, 1, n, width).getValues();
  vals.forEach((v, i) => {
    if (v[1] === '' || v[1] === null) return;   // ไม่มีวันที่ = แถวว่าง
    const row = first + i;
    if (!v[0]) sh.getRange(row, 1).setValue('manual-' + Date.now().toString(36) + i);
    if (typeof v[1] === 'string') sh.getRange(row, 2).setValue(dateSerial(v[1].trim()));
    if (sh.getName() === LOG) {
      if (typeof v[2] === 'string' && v[2]) sh.getRange(row, 3).setValue(timeSerial(v[2].trim()));
      if (!v[3]) sh.getRange(row, 4).setValue('เช็คอิน');
      if (!v[COL_REGION - 1] || (e.range.getColumn() <= 7 && e.range.getLastColumn() >= 7)) {
        sh.getRange(row, COL_REGION).setValue(regionOf(v[6]));
      }
    }
  });
}

// แท็บที่เป็นสูตร: ขึ้นคำเตือนเมื่อมีคนพิมพ์ทับ (ให้ไปแก้ที่แท็บเช็คอิน/การลาแทน)
function protectFormulaTabs(ss, names) {
  names.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    sh.protect().setDescription('แท็บนี้คำนวณอัตโนมัติ แก้ข้อมูลที่แท็บ "' + LOG + '" หรือ "' + LEAVE + '" แทน')
      .setWarningOnly(true);
  });
}

// ลบทั้งแถว (รวมคอลัมน์ที่ซ่อนอยู่) ของแถวที่เลือกในแท็บเช็คอินหรือการลา
function deleteSelectedRows() {
  const ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActiveSheet();
  if ([LOG, LEAVE].indexOf(sh.getName()) < 0) {
    ui.alert('เปิดแท็บ "' + LOG + '" หรือ "' + LEAVE + '" แล้วเลือกแถวที่ต้องการลบก่อน');
    return;
  }
  const rows = {};
  sh.getActiveRangeList().getRanges().forEach(r => {
    for (let i = r.getRow(); i < r.getRow() + r.getNumRows(); i++) if (i > 1 && i <= sh.getLastRow()) rows[i] = true;
  });
  const list = Object.keys(rows).map(Number).sort((a, b) => b - a);
  if (!list.length) { ui.alert('ยังไม่ได้เลือกแถวข้อมูล'); return; }
  if (ui.alert('ลบ ' + list.length + ' แถวจากแท็บ "' + sh.getName() + '"?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  list.forEach(i => sh.deleteRow(i));
  ui.alert('ลบแล้ว ' + list.length + ' แถว');
}

// ลบเช็คอินและการลาทั้งเดือน รูปของเดือนนั้นย้ายไปถังขยะใน Drive (กู้คืนได้ 30 วัน)
function deleteMonth() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('ลบข้อมูลทั้งเดือน', 'พิมพ์เดือน/ปี เช่น 10/2569 หรือ 10/2026', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const m = res.getResponseText().trim().match(/^(\d{1,2})\s*\/\s*(\d{4})$/);
  if (!m || +m[1] < 1 || +m[1] > 12) { ui.alert('รูปแบบไม่ถูกต้อง พิมพ์เป็น เดือน/ปี เช่น 10/2569'); return; }
  const month = +m[1], year = +m[2] > 2400 ? +m[2] - 543 : +m[2];
  const inMonth = v => v instanceof Date && v.getFullYear() === year && v.getMonth() + 1 === month;

  const ss = book(), log = ss.getSheetByName(LOG), lv = ss.getSheetByName(LEAVE);
  const pick = (sh) => {
    if (!sh || sh.getLastRow() < 2) return [];
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    return vals.map((v, i) => ({ row: i + 2, v })).filter(x => inMonth(x.v[1]));
  };
  const a = pick(log), b = pick(lv);
  if (!a.length && !b.length) { ui.alert('ไม่มีข้อมูลของเดือน ' + month + '/' + (year + 543)); return; }
  const msg = 'เดือน ' + month + '/' + (year + 543) + ': เช็คอิน ' + a.length + ' รายการ, วันลา ' + b.length + ' วัน\n' +
    'รูปของเดือนนี้จะถูกย้ายไปถังขยะใน Google Drive (กู้คืนได้ภายใน 30 วัน)\nลบเลยไหม?';
  if (ui.alert('ยืนยันการลบ', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  a.forEach(x => {   // ย้ายรูปไปถังขยะ
    const url = String(x.v[COL_PHOTO_URL - 1] || ''), id = (url.match(/\/d\/([^/]+)/) || [])[1];
    if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* ไฟล์ถูกลบไปแล้ว */ } }
  });
  a.map(x => x.row).sort((p, q) => q - p).forEach(r => log.deleteRow(r));
  b.map(x => x.row).sort((p, q) => q - p).forEach(r => lv.deleteRow(r));
  ui.alert('ลบข้อมูลเดือน ' + month + '/' + (year + 543) + ' แล้ว');
}

// แถวที่ถูกลบด้วยปุ่ม Delete (ช่องวันที่ว่าง แต่คอลัมน์ที่ซ่อนยังมีค่า) ลบทั้งแถวทิ้ง
function cleanupBlankRows(silent) {
  const ss = book();
  let n = 0;
  [LOG, LEAVE].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const dates = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
    for (let i = dates.length - 1; i >= 0; i--) if (dates[i][0] === '' || dates[i][0] === null) { sh.deleteRow(i + 2); n++; }
  });
  if (silent !== true) SpreadsheetApp.getUi().alert(n ? 'ลบแถวที่ค้างอยู่ ' + n + ' แถว' : 'ไม่มีแถวค้าง');
  return n;
}

/* ===================== ออกแบบชีต ===================== */

/** กด Run ฟังก์ชันนี้หนึ่งครั้ง: สร้าง/จัดรูปแบบทุกชีต (รันซ้ำได้ ข้อมูลเดิมไม่หาย) */
function setup() {
  const ss = book();
  cleanupBlankRows(true);
  installMenuTrigger(ss);
  ss.setSpreadsheetLocale('th_TH');
  ss.setSpreadsheetTimeZone('Asia/Bangkok');
  buildLog(ss);
  buildDaily(ss);
  buildMonthly(ss);
  buildSites(ss);
  buildLeave(ss);
  buildLeaveSummary(ss);
  const regionNames = REGIONS.map(r => r.name).concat([OTHER_REGION]);
  regionNames.forEach(name => buildRegion(ss, name));
  getFolder();

  // เรียงแท็บ: สรุปรายเดือน → สรุปรายวัน → สรุปการลา → แต่ละภาค → เช็คอิน → การลา → จุดทำงาน
  const order = [MONTHLY, DAILY, LEAVE_SUM].concat(regionNames, [LOG, LEAVE, SITES]);
  order.forEach((name, i) => { ss.setActiveSheet(ss.getSheetByName(name)); ss.moveActiveSheet(i + 1); });
  // ลบชีตว่างเริ่มต้น (Sheet1 / แผ่น1)
  ss.getSheets().forEach(s => {
    if (order.indexOf(s.getName()) < 0 && s.getLastRow() === 0 && s.getLastColumn() === 0) ss.deleteSheet(s);
  });
  protectFormulaTabs(ss, [MONTHLY, DAILY, LEAVE_SUM].concat(regionNames, [SITES]));
  ss.setActiveSheet(ss.getSheetByName(MONTHLY));
}

function sheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader(sh, row, n) {
  sh.getRange(row, 1, 1, n).setBackground(C.head).setFontColor(C.headFg).setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
  sh.setRowHeight(row, 36);
  sh.setFrozenRows(row);
}

function trimColumns(sh, n) {
  const extra = sh.getMaxColumns() - n;
  if (extra > 0) sh.deleteColumns(n + 1, extra);
  if (extra < 0) sh.insertColumnsAfter(sh.getMaxColumns(), -extra);
}

function band(sh, range) {
  sh.getBandings().forEach(b => b.remove());
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor('#ffffff').setSecondRowColor(C.band);
}

/* ---------- เช็คอิน (ข้อมูลดิบจากแอป) ---------- */
function buildLog(ss) {
  const legacy = ss.getSheetByName(LEGACY_LOG);   // แท็บชื่อเดิม: เปลี่ยนชื่อ ข้อมูลอยู่ครบ
  if (legacy && !ss.getSheetByName(LOG)) legacy.setName(LOG);
  const sh = sheet(ss, LOG), n = HEADERS.length;
  trimColumns(sh, n);
  sh.getRange(1, 1, 1, n).setValues([HEADERS]);
  styleHeader(sh, 1, n);
  sh.setFrozenColumns(3);

  const widths = [70, 95, 75, 75, 140, 95, 260, 70, 70, 90, 90, 90, 220, 70, 60, 140, 140, 140, 170, 120, 70];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.hideColumns(1);               // ID ใช้อ้างอิงตอนแก้ไขจากแอป ไม่ต้องเห็น
  sh.hideColumns(COL_PHOTO_URL);   // URL รูปแบบเต็ม ใช้ทำลิงก์ในแท็บภาค
  sh.hideColumns(17, 2);           // คอลัมน์ประวัติการแก้ไข (ไม่ใช้แล้ว)
  sh.hideColumns(6);               // รหัสพนักงาน (ไม่ใช้)
  if (sh.getLastRow() > 1) sh.getRange(2, 6, sh.getLastRow() - 1, 1).clearContent();
  sh.hideColumns(12);              // ความแม่นยำ GPS (ไม่ใช้)
  if (sh.getLastRow() > 1) sh.getRange(2, 12, sh.getLastRow() - 1, 1).clearContent();
  if (sh.getLastRow() > 1) sh.getRange(2, 17, sh.getLastRow() - 1, 2).clearContent();

  // เติมภาคให้แถวเก่าที่ยังไม่มี
  const last = sh.getLastRow();
  if (last > 1) {
    const sitesCol = sh.getRange(2, 7, last - 1, 1).getValues();
    const regionCol = sh.getRange(2, COL_REGION, last - 1, 1).getValues();
    sh.getRange(2, COL_REGION, last - 1, 1).setValues(regionCol.map((v, i) => [(v[0] && v[0] !== OTHER_REGION) ? v[0] : regionOf(sitesCol[i][0])]));
    // แถวที่บันทึกก่อนมีคอลัมน์ลิงก์รูป: ดึง URL จากสูตร HYPERLINK ในคอลัมน์รูป
    const photoF = sh.getRange(2, 14, last - 1, 1).getFormulas();
    const urlCol = sh.getRange(2, COL_PHOTO_URL, last - 1, 1).getValues();
    sh.getRange(2, COL_PHOTO_URL, last - 1, 1).setValues(urlCol.map((v, i) => {
      const m = String(photoF[i][0]).match(/HYPERLINK\("([^"]+)"/i);
      return [v[0] || (m ? m[1] : '')];
    }));
    // วันที่/เวลาที่ค้างเป็นข้อความ: แปลงเป็นตัวเลข
    const dt = sh.getRange(2, 2, last - 1, 2).getValues();
    dt.forEach((v, i) => {
      if (typeof v[0] === 'string' && v[0]) sh.getRange(i + 2, 2).setValue(dateSerial(v[0].trim()));
      if (typeof v[1] === 'string' && v[1]) sh.getRange(i + 2, 3).setValue(timeSerial(v[1].trim()));
    });
  }

  const rows = sh.getMaxRows() - 1;
  sh.getRange(2, 2, rows, 1).setNumberFormat('ddd d mmm yyyy');
  sh.getRange(2, 3, rows, 1).setNumberFormat('HH:mm:ss');
  sh.getRange(2, 10, rows, 2).setNumberFormat('0.000000');
  sh.getRange(2, 16, rows, 1).setNumberFormat('d/m/yyyy HH:mm');
  sh.getRange(2, 2, rows, 4).setHorizontalAlignment('center');
  sh.getRange(2, 8, rows, 2).setHorizontalAlignment('center');
  sh.getRange(2, 14, rows, 2).setHorizontalAlignment('center');
  sh.getRange(2, 13, rows, 1).setWrap(true);
  band(sh, sh.getRange(2, 1, rows, n));

  const all = sh.getRange(2, 1, rows, n);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$O2="ใช่"')
      .setFontColor(C.muted).setItalic(true).setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('เช็คเอาท์')
      .setFontColor('#a15c00').setBold(true).setRanges([sh.getRange(2, 4, rows, 1)]).build()
  ]);
  return sh;
}

/* ---------- สรุปรายวัน (สูตรคำนวณเอง ไม่ต้องแก้) ---------- */
function buildDaily(ss) {
  const sh = sheet(ss, DAILY);
  sh.clear();
  const L = "'" + LOG + "'!";
  const head = ['วันที่', 'สถานที่', 'เช็คอิน', 'เช็คเอาท์', 'จำนวนครั้ง', 'หมายเหตุ'];
  trimColumns(sh, head.length);
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 1, head.length);

  const real = L + 'O2:O,"<>ใช่"';   // ไม่นับรายการโหมดทดสอบ
  sh.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + L + 'B2:B,' + L + 'B2:B<>"",' + L + 'O2:O<>"ใช่")),1,FALSE),)');
  sh.getRange('B2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER(' + L + 'G2:G,' + L + 'B2:B=d,' + L + 'O2:O<>"ใช่"))),))))');
  sh.getRange('C2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,LET(v,MINIFS(' + L + 'C2:C,' + L + 'B2:B,d,' + L + 'D2:D,"<>เช็คเอาท์",' + L + 'D2:D,"<>ออกงาน",' + real + '),IF(v=0,,v)))))');
  sh.getRange('D2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,LET(v,MAX(MAXIFS(' + L + 'C2:C,' + L + 'B2:B,d,' + L + 'D2:D,"เช็คเอาท์",' + real + '),MAXIFS(' + L + 'C2:C,' + L + 'B2:B,d,' + L + 'D2:D,"ออกงาน",' + real + ')),IF(v=0,,v)))))');
  sh.getRange('E2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,COUNTIFS(' + L + 'B2:B,d,' + real + '))))');
  sh.getRange('F2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,IFERROR(TEXTJOIN(" / ",TRUE,FILTER(' + L + 'M2:M,' + L + 'B2:B=d,' + L + 'M2:M<>"",' + L + 'O2:O<>"ใช่")),))))');

  const rows = sh.getMaxRows() - 1;
  sh.getRange(2, 1, rows, 1).setNumberFormat('ddd d mmm yyyy').setHorizontalAlignment('left');
  sh.getRange(2, 3, rows, 2).setNumberFormat('HH:mm');
  sh.getRange(2, 5, rows, 1).setNumberFormat('0 "ครั้ง"');
  sh.getRange(2, 3, rows, 3).setHorizontalAlignment('center');
  sh.getRange(2, 2, rows, 1).setWrap(true);
  sh.getRange(2, 6, rows, 1).setWrap(true);
  [130, 320, 110, 80, 90, 280].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  band(sh, sh.getRange(2, 1, rows, head.length));
  sh.setConditionalFormatRules([]);
  return sh;
}

/* ---------- สรุปรายเดือน ---------- */
function buildMonthly(ss) {
  const sh = sheet(ss, MONTHLY);
  sh.clear();
  const D = "'" + DAILY + "'!";
  sh.getRange('A1').setValue('สรุปการเช็คอิน').setFontSize(16).setFontWeight('bold').setFontColor(C.title);
  sh.getRange('A2').setFormula('="ผู้เช็คอิน: "&IFERROR(INDEX(FILTER(\'' + LOG + '\'!E2:E,\'' + LOG + '\'!E2:E<>""),1),"–")&"   ·   อัปเดตอัตโนมัติจากแอป"')
    .setFontColor(C.muted);
  sh.setRowHeight(1, 34);

  const head = ['เดือน', 'วันที่เช็คอิน', 'จำนวนครั้ง', 'เช็คอินเฉลี่ย', 'สถานที่'];
  trimColumns(sh, head.length);
  sh.getRange(4, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 4, head.length);

  const range = (m) => D + 'A2:A,">="&' + m + ',' + D + 'A2:A,"<="&EOMONTH(' + m + ',0)';
  sh.getRange('A5').setFormula('=IFERROR(SORT(UNIQUE(FILTER(ARRAYFORMULA(EOMONTH(' + D + 'A2:A,-1)+1),' + D + 'A2:A<>"")),1,FALSE),)');
  sh.getRange('B5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,COUNTIFS(' + range('m') + '))))');
  sh.getRange('C5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,SUMIFS(' + D + 'E2:E,' + range('m') + '))))');
  sh.getRange('D5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,IFERROR(AVERAGEIFS(' + D + 'C2:C,' + range('m') + ',' + D + 'C2:C,"<>"),))))');
  sh.getRange('E5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER(' + D + 'B2:B,' + D + 'A2:A>=m,' + D + 'A2:A<=EOMONTH(m,0)))),))))');

  const rows = sh.getMaxRows() - 4;
  sh.getRange(5, 1, rows, 1).setNumberFormat('mmmm yyyy').setHorizontalAlignment('left');
  sh.getRange(5, 2, rows, 1).setNumberFormat('0 "วัน"');
  sh.getRange(5, 3, rows, 1).setNumberFormat('0 "ครั้ง"');
  sh.getRange(5, 4, rows, 1).setNumberFormat('HH:mm');
  sh.getRange(5, 2, rows, 3).setHorizontalAlignment('center');
  sh.getRange(5, 5, rows, 1).setWrap(true);
  [140, 110, 110, 110, 420].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  band(sh, sh.getRange(5, 1, rows, head.length));
  return sh;
}

/* ---------- แท็บแยกภาค (ดึงจากแท็บเช็คอินอัตโนมัติ) ---------- */
const REGION_TAB_COLORS = { 'ส่วนกลาง': '#0e5a52', 'ภาคเหนือ': '#2e7d5b', 'ภาคตะวันออกเฉียงเหนือ': '#b8862b', 'ภาคใต้': '#2a6cb0', 'ภาคตะวันออก': '#8a4fb0' };

function buildRegion(ss, name) {
  const sh = sheet(ss, name);
  sh.clear();
  sh.getBandings().forEach(b => b.remove());
  const L = "'" + LOG + "'!", R = '$A$1';
  const real = L + 'S2:S=' + R + ',' + L + 'O2:O<>"ใช่",' + L + 'B2:B<>""';   // รายการของภาคนี้ ไม่รวมโหมดทดสอบ
  const head = ['วันที่', 'เวลา', 'ประเภท', 'สถานที่', 'หมายเหตุ', 'รูป'];
  trimColumns(sh, Math.max(head.length, 6));
  sh.setTabColor(REGION_TAB_COLORS[name] || C.muted);

  // หัวแท็บ: ชื่อภาค (สูตรอ้างอิงเซลล์ A1) + จุดที่บันทึกไว้ในภาค
  sh.getRange('A1').setValue(name).setFontSize(16).setFontWeight('bold').setFontColor(C.title);
  sh.setRowHeight(1, 34);
  const siteNames = SITE_LIST.filter(s => s[2] === name).map(s => s[0]);
  sh.getRange('A2').setValue(siteNames.length ? 'จุดที่บันทึกไว้: ' + siteNames.join(', ') : 'สถานที่ที่จัดเข้าภาคใดไม่ได้')
    .setFontColor(C.muted);

  // ตัวเลขสรุปของภาค
  sh.getRange('A3:F3').setValues([['วันที่เช็คอิน', '', 'จำนวนครั้ง', '', 'ล่าสุด', '']]);
  sh.getRange('B3').setFormula('=IFERROR(ROWS(UNIQUE(FILTER(' + L + 'B2:B,' + real + '))),0)').setNumberFormat('0 "วัน"');
  sh.getRange('D3').setFormula('=COUNTIFS(' + L + 'S2:S,' + R + ',' + L + 'O2:O,"<>ใช่",' + L + 'B2:B,"<>")').setNumberFormat('0 "ครั้ง"');
  sh.getRange('F3').setFormula('=IFERROR(MAX(FILTER(' + L + 'B2:B,' + real + ')),"–")').setNumberFormat('d mmm yyyy');
  sh.getRange('A3:F3').setFontColor(C.muted);
  ['B3', 'D3', 'F3'].forEach(a => sh.getRange(a).setFontWeight('bold').setFontColor(C.title).setHorizontalAlignment('left'));

  // ตารางรายการ: เรียงใหม่สุดก่อน
  sh.getRange(5, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 5, head.length);
  const cols = ['B', 'C', 'D', 'G', 'M', 'T'].map(c => L + c + '2:' + c).join(',');
  sh.getRange('A6').setFormula('=IFERROR(LET(f,FILTER({' + cols + '},' + real + '),s,SORT(f,1,FALSE,2,FALSE),' +
    'HSTACK(CHOOSECOLS(s,1,2,3,4,5),' +
    'MAP(CHOOSECOLS(s,6),LAMBDA(u,IF(u="","",HYPERLINK(u,"ดูรูป")))))),"ยังไม่มีรายการ")');

  const rows = sh.getMaxRows() - 5;
  sh.getRange(6, 1, rows, 1).setNumberFormat('ddd d mmm yyyy').setHorizontalAlignment('left');
  sh.getRange(6, 2, rows, 1).setNumberFormat('HH:mm').setHorizontalAlignment('center');
  sh.getRange(6, 3, rows, 1).setHorizontalAlignment('center');
  sh.getRange(6, 6, rows, 1).setHorizontalAlignment('center');
  sh.getRange(6, 4, rows, 2).setWrap(true);
  [130, 70, 90, 300, 240, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  band(sh, sh.getRange(6, 1, rows, head.length));
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('เช็คเอาท์')
      .setFontColor('#a15c00').setBold(true).setRanges([sh.getRange(6, 3, rows, 1)]).build()
  ]);
  return sh;
}

/* ---------- การลา: หนึ่งแถวต่อหนึ่งวันลา (ID เดียวกันสำหรับการลาครั้งเดียวกัน) ---------- */
const LEAVE_HEADERS = ['ID', 'วันที่ลา', 'ประเภท', 'ชื่อ', 'รหัสพนักงาน', 'หมายเหตุ', 'แจ้งเมื่อ'];

function leaveSheet() {
  return book().getSheetByName(LEAVE) || buildLeave(book());
}

function addLeave(l) {
  const sh = leaveSheet();
  if (findRow(sh, l.id)) return { ok: true, duplicate: true };
  const name = String(l.name || '').trim() || 'ไม่ระบุชื่อ';
  const rows = l.days.map(d => [l.id, dateSerial(d), l.type, name, '', l.note, l.created]);
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, LEAVE_HEADERS.length).setValues(rows);
  return { ok: true };
}

function deleteLeave(id) {
  const sh = leaveSheet(), n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true };
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 2);
  return { ok: true };
}

function buildLeave(ss) {
  const sh = sheet(ss, LEAVE), n = LEAVE_HEADERS.length;
  trimColumns(sh, n);
  sh.getRange(1, 1, 1, n).setValues([LEAVE_HEADERS]);
  styleHeader(sh, 1, n);
  sh.hideColumns(1);
  sh.hideColumns(5);   // รหัสพนักงาน (ไม่ใช้)
  if (sh.getLastRow() > 1) sh.getRange(2, 5, sh.getLastRow() - 1, 1).clearContent();
  const rows = sh.getMaxRows() - 1;
  sh.getRange(2, 2, rows, 1).setNumberFormat('ddd d mmm yyyy').setHorizontalAlignment('left');
  sh.getRange(2, 3, rows, 1).setHorizontalAlignment('center');
  sh.getRange(2, 6, rows, 1).setWrap(true);
  [70, 140, 100, 160, 110, 260, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  band(sh, sh.getRange(2, 1, rows, n));
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ลาป่วย')
      .setFontColor('#b4282f').setBold(true).setRanges([sh.getRange(2, 3, rows, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ลาพักผ่อน')
      .setFontColor('#0e5a52').setBold(true).setRanges([sh.getRange(2, 3, rows, 1)]).build()
  ]);
  return sh;
}

/* ---------- สรุปการลารายเดือน (แยกตามผู้ลา) ---------- */
function buildLeaveSummary(ss) {
  const sh = sheet(ss, LEAVE_SUM);
  sh.clear();
  sh.getBandings().forEach(b => b.remove());
  const V = "'" + LEAVE + "'!";
  const head = ['เดือน', 'ผู้ลา', 'ลาป่วย', 'ลาพักผ่อน', 'รวม', 'วันที่ลา'];
  trimColumns(sh, head.length);
  sh.setTabColor('#b4282f');

  sh.getRange('A1').setValue('สรุปการลารายเดือน').setFontSize(16).setFontWeight('bold').setFontColor(C.title);
  sh.setRowHeight(1, 34);
  const yr = (type) => 'COUNTIFS(' + V + 'B2:B,">="&DATE(YEAR(TODAY()),1,1),' + V + 'B2:B,"<="&DATE(YEAR(TODAY()),12,31),' + V + 'C2:C,"' + type + '")';
  sh.getRange('A2').setFormula('="ปีนี้ทุกคน · ลาป่วย "&' + yr('ลาป่วย') + '&" วัน · ลาพักผ่อน "&' + yr('ลาพักผ่อน') + '&" วัน"')
    .setFontColor(C.muted);

  sh.getRange(4, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 4, head.length);

  // แถวละหนึ่งคนต่อหนึ่งเดือน เรียงเดือนล่าสุดก่อน แล้วเรียงตามชื่อ
  sh.getRange('A5').setFormula('=IFERROR(SORT(UNIQUE(FILTER({ARRAYFORMULA(EOMONTH(' + V + 'B2:B,-1)+1),' + V + 'D2:D},' + V + 'B2:B<>"")),1,FALSE,2,TRUE),)');
  const cnt = (type) => 'COUNTIFS(' + V + 'B2:B,">="&m,' + V + 'B2:B,"<="&EOMONTH(m,0),' + V + 'D2:D,n,' + V + 'C2:C,"' + type + '")';
  const days = (type) => 'IFERROR("' + type.replace('ลา', '') + ' "&TEXTJOIN(", ",TRUE,ARRAYFORMULA(TEXT(SORT(FILTER(' + V + 'B2:B,' +
    V + 'B2:B>=m,' + V + 'B2:B<=EOMONTH(m,0),' + V + 'D2:D=n,' + V + 'C2:C="' + type + '")),"d"))),"")';
  sh.getRange('C5').setFormula('=MAP(A5:A,B5:B,LAMBDA(m,n,IF(m="",,' + cnt('ลาป่วย') + ')))');
  sh.getRange('D5').setFormula('=MAP(A5:A,B5:B,LAMBDA(m,n,IF(m="",,' + cnt('ลาพักผ่อน') + ')))');
  sh.getRange('E5').setFormula('=MAP(C5:C,D5:D,LAMBDA(s,v,IF(AND(s="",v=""),,s+v)))');
  sh.getRange('F5').setFormula('=MAP(A5:A,B5:B,LAMBDA(m,n,IF(m="",,TRIM(' + days('ลาป่วย') + '&"   "&' + days('ลาพักผ่อน') + '))))');

  const rows = sh.getMaxRows() - 4;
  sh.getRange(5, 1, rows, 1).setNumberFormat('mmmm yyyy').setHorizontalAlignment('left');
  sh.getRange(5, 2, rows, 1).setFontWeight('bold');
  sh.getRange(5, 3, rows, 3).setNumberFormat('0 "วัน"').setHorizontalAlignment('center');
  sh.getRange(5, 6, rows, 1).setWrap(true);
  [130, 200, 90, 100, 80, 340].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  band(sh, sh.getRange(5, 1, rows, head.length));
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setFontColor('#b4282f').setBold(true).setRanges([sh.getRange(5, 3, rows, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setFontColor('#0e5a52').setBold(true).setRanges([sh.getRange(5, 4, rows, 1)]).build()
  ]);
  return sh;
}

/* ---------- จุดทำงาน (อ้างอิง) ---------- */
function buildSites(ss) {
  const sh = sheet(ss, SITES);
  sh.clear();
  const head = ['จุดทำงาน', 'จังหวัด', 'ภาค', 'ละติจูด', 'ลองจิจูด', 'รัศมี (ม.)', 'แผนที่'];
  trimColumns(sh, head.length);
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 1, head.length);
  const rows = SITE_LIST.map(s => s.concat(['=HYPERLINK("https://www.google.com/maps?q=' + s[3] + ',' + s[4] + '","เปิดแผนที่")']));
  sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.getRange(2, 4, rows.length, 2).setNumberFormat('0.00000');
  sh.getRange(2, 2, rows.length, 6).setHorizontalAlignment('center');
  [300, 100, 170, 100, 100, 90, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(rows.length + 3, 1).setValue('รายการนี้ใช้อ้างอิงเท่านั้น การเพิ่มหรือแก้จุดทำงานให้ทำในแท็บ "จุดทำงาน" ของแอป')
    .setFontColor(C.muted).setFontStyle('italic');
  band(sh, sh.getRange(2, 1, rows.length, head.length));
  return sh;
}

/* ===================== ใบลงชื่อปฏิบัติงาน (ไฟล์แยก หนึ่งแท็บต่อหนึ่งคน ดาวน์โหลดเป็น Excel) ===================== */
const SIGN_SHEET_ID = '';   // ID ของไฟล์ใบลงชื่อ (ไฟล์ Google Sheet อีกไฟล์)
const SIGN_CFG = 'ตั้งค่า', SIGN_NAMES = 'รายชื่อ', SIGN_HOLIDAYS = 'วันหยุด', SIGN_TAG = 'ci.signsheet';
const SIGN_FONT = 'TH SarabunIT๙';
const TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม',
  'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const TH_DAYS = ['วันอาทิตย์', 'วันจันทร์', 'วันอังคาร', 'วันพุธ', 'วันพฤหัสบดี', 'วันศุกร์', 'วันเสาร์'];
// หัวกระดาษบรรทัดที่ 2 ของแต่ละภาค (แก้ได้ในแท็บตั้งค่าของไฟล์ใบลงชื่อ)
const SIGN_CENTERS = [
  ['ส่วนกลาง', 'สำนักงานปรมาณูเพื่อสันติ กรุงเทพมหานคร'],
  ['ภาคเหนือ', 'ศูนย์ปรมาณูเพื่อสันติภูมิภาค ภาคเหนือ จังหวัดเชียงใหม่'],
  ['ภาคตะวันออกเฉียงเหนือ', 'ศูนย์ปรมาณูเพื่อสันติภูมิภาค ภาคตะวันออกเฉียงเหนือ จังหวัดขอนแก่น'],
  ['ภาคใต้', 'ศูนย์ปรมาณูเพื่อสันติภูมิภาค ภาคใต้ จังหวัดสงขลา'],
  ['ภาคตะวันออก', 'ศูนย์ปรมาณูเพื่อสันติภูมิภาค ภาคตะวันออก จังหวัดระยอง'],
  [OTHER_REGION, 'สำนักงานปรมาณูเพื่อสันติ']
];
const SIGN_SIGNER = ['ลงชื่อ.................................................', '(นางสาวธนวรรณ  แจ่มสุวรรณ)',
  'ผชช.เฉพาะด้านพัฒนาระบบบริหารจัดการด้านพลังงานปรมาณู', 'ปฏิบัติหน้าที่ หปสภ.', '......./......../............'];
// วันหยุดราชการที่วันที่ตายตัว วันหยุดทางพุทธศาสนาและวันหยุดชดเชยต้องเพิ่มเองในแท็บวันหยุด
const SIGN_FIXED_HOLIDAYS = [['01-01', 'วันขึ้นปีใหม่'], ['04-06', 'วันจักรี'], ['04-13', 'วันสงกรานต์'], ['04-14', 'วันสงกรานต์'],
  ['04-15', 'วันสงกรานต์'], ['05-04', 'วันฉัตรมงคล'], ['06-03', 'วันเฉลิมพระชนมพรรษาสมเด็จพระราชินี'],
  ['07-28', 'วันเฉลิมพระชนมพรรษาพระบาทสมเด็จพระเจ้าอยู่หัว'], ['08-12', 'วันแม่แห่งชาติ'], ['10-13', 'วันนวมินทรมหาราช'],
  ['10-23', 'วันปิยมหาราช'], ['12-05', 'วันพ่อแห่งชาติ'], ['12-10', 'วันรัฐธรรมนูญ'], ['12-31', 'วันสิ้นปี']];

function signBook() {
  if (!SIGN_SHEET_ID) throw new Error('ยังไม่ได้ใส่ SIGN_SHEET_ID');
  return SpreadsheetApp.openById(SIGN_SHEET_ID);
}

function onSignOpen() {
  SpreadsheetApp.getUi().createMenu('ใบลงชื่อ')
    .addItem('สร้าง/อัปเดตใบลงชื่อ (เดือนในแท็บตั้งค่า)', 'makeSignSheets')
    .addSeparator()
    .addItem('ดาวน์โหลดเป็น Excel…', 'signExcelHelp')
    .addToUi();
}

function signExcelHelp() {
  SpreadsheetApp.getUi().alert('ดาวน์โหลดเป็น Excel',
    'เมนู ไฟล์ → ดาวน์โหลด → Microsoft Excel (.xlsx)\nได้ไฟล์เดียว หนึ่งแท็บต่อหนึ่งคน ของเดือนที่สร้างไว้ล่าสุด',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** กด Run หนึ่งครั้ง: สร้างแท็บตั้งค่า/รายชื่อ/วันหยุด ในไฟล์ใบลงชื่อ และติดตั้งเมนู (รันซ้ำได้ ค่าที่แก้ไว้ไม่หาย) */
function setupSignSheet() {
  const ss = signBook();
  ss.setSpreadsheetLocale('th_TH');
  ss.setSpreadsheetTimeZone('Asia/Bangkok');
  if (ScriptApp.getProjectTriggers().every(t => t.getHandlerFunction() !== 'onSignOpen')) {
    ScriptApp.newTrigger('onSignOpen').forSpreadsheet(ss).onOpen().create();
  }

  // ตั้งค่า: เดือน ปี หัวกระดาษแต่ละภาค ผู้ลงนาม
  const cfg = sheet(ss, SIGN_CFG);
  trimColumns(cfg, 3);
  cfg.getRange('A1').setValue('ตั้งค่าใบลงชื่อ').setFontSize(16).setFontWeight('bold').setFontColor(C.title);
  cfg.getRange('A3:A4').setValues([['เดือน'], ['ปี (พ.ศ.)']]).setFontWeight('bold');
  const now = new Date();
  if (!cfg.getRange('B3').getValue()) cfg.getRange('B3').setValue(TH_MONTHS[now.getMonth()]);
  if (!cfg.getRange('B4').getValue()) cfg.getRange('B4').setValue(now.getFullYear() + 543);
  cfg.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TH_MONTHS, true).build());
  cfg.getRange('B4').setNumberFormat('0');
  cfg.getRange('B3:B4').setBackground(C.warn).setFontWeight('bold').setHorizontalAlignment('left');
  cfg.getRange('C3').setValue('← เลือกเดือนและปี แล้วกดเมนู ใบลงชื่อ → สร้าง/อัปเดตใบลงชื่อ').setFontColor(C.muted);

  cfg.getRange('A6:B6').setValues([['ภาค', 'หัวกระดาษบรรทัดที่ 2 (ต่อท้ายด้วย "ประจำเดือน …" ให้อัตโนมัติ)']]);
  styleHeader(cfg, 6, 2);
  cfg.setFrozenRows(0);
  const have = cfg.getRange(7, 1, SIGN_CENTERS.length, 2).getValues();
  cfg.getRange(7, 1, SIGN_CENTERS.length, 2).setValues(SIGN_CENTERS.map((c, i) => [c[0], have[i][1] || c[1]]));

  const sRow = 8 + SIGN_CENTERS.length;
  cfg.getRange(sRow, 1, 1, 2).setValues([['บรรทัด', 'ผู้ลงนามรับรอง (ท้ายใบลงชื่อทุกคน)']]);
  styleHeader(cfg, sRow, 2);
  cfg.setFrozenRows(0);
  const haveS = cfg.getRange(sRow + 1, 2, SIGN_SIGNER.length, 1).getValues();
  cfg.getRange(sRow + 1, 1, SIGN_SIGNER.length, 2).setValues(SIGN_SIGNER.map((s, i) => [i + 1, haveS[i][0] || s]));
  cfg.getRange(sRow + 1, 1, SIGN_SIGNER.length, 1).setHorizontalAlignment('center');
  [150, 520, 420].forEach((w, i) => cfg.setColumnWidth(i + 1, w));
  cfg.setTabColor(C.head);

  // รายชื่อ: เพิ่มชื่อใหม่ให้อัตโนมัติ เอาเครื่องหมายออกถ้าไม่ต้องทำใบลงชื่อให้คนนั้น
  const nm = sheet(ss, SIGN_NAMES);
  trimColumns(nm, 3);
  nm.getRange(1, 1, 1, 3).setValues([['ชื่อ-สกุล (ตรงกับที่พิมพ์ในแอป)', 'ภาค (ว่าง = ตามจุดที่เช็คอิน)', 'ทำใบลงชื่อ']]);
  styleHeader(nm, 1, 3);
  nm.getRange(2, 2, nm.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(SIGN_CENTERS.map(c => c[0]), true).setAllowInvalid(false).build());
  nm.getRange(2, 3, nm.getMaxRows() - 1, 1).setHorizontalAlignment('center');
  [280, 220, 110].forEach((w, i) => nm.setColumnWidth(i + 1, w));
  nm.setTabColor(C.head);

  // วันหยุดราชการ: ใส่วันที่ตายตัวของปีที่ตั้งไว้ให้ก่อน
  const hd = sheet(ss, SIGN_HOLIDAYS);
  trimColumns(hd, 4);
  hd.getRange(1, 1, 1, 2).setValues([['วันที่', 'ชื่อวันหยุด']]);
  styleHeader(hd, 1, 2);
  if (hd.getLastRow() < 2) {
    const y = Number(cfg.getRange('B4').getValue()) - 543;
    hd.getRange(2, 1, SIGN_FIXED_HOLIDAYS.length, 2)
      .setValues(SIGN_FIXED_HOLIDAYS.map(h => [dateSerial(y + '-' + h[0]), h[1]]));
  }
  hd.getRange(2, 1, hd.getMaxRows() - 1, 1).setNumberFormat('ddd d mmm yyyy').setHorizontalAlignment('left');
  [160, 360, 20, 340].forEach((w, i) => hd.setColumnWidth(i + 1, w));
  hd.getRange('D1').setValue('เพิ่มวันหยุดทางพุทธศาสนา วันหยุดชดเชย และวันหยุดพิเศษตามประกาศ ครม. เอง ' +
    '(พิมพ์วันที่แบบ 2026-03-03) ปีใหม่ให้เพิ่มแถวของปีนั้นต่อท้าย').setFontColor(C.muted).setWrap(true);
  hd.setTabColor(C.head);

  ss.getSheets().forEach(s => {   // ลบชีตว่างเริ่มต้น (Sheet1 / แผ่น1)
    if ([SIGN_CFG, SIGN_NAMES, SIGN_HOLIDAYS].indexOf(s.getName()) < 0 && s.getLastRow() === 0 && !isSignTab(s)) ss.deleteSheet(s);
  });
  orderSignTabs(ss);
}

function isSignTab(s) {
  return s.getDeveloperMetadata().some(m => m.getKey() === SIGN_TAG);
}

// แท็บใบลงชื่อเรียงตามชื่อ แท็บตั้งค่าไว้ท้ายสุด
function orderSignTabs(ss) {
  const people = ss.getSheets().filter(isSignTab).map(s => s.getName()).sort((a, b) => a.localeCompare(b, 'th'));
  people.concat([SIGN_CFG, SIGN_NAMES, SIGN_HOLIDAYS]).forEach((name, i) => {
    const s = ss.getSheetByName(name);
    if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); }
  });
  ss.setActiveSheet(ss.getSheets()[0]);
}

function normName(s) {
  return String(s || '').replace(/\s+/g, ' ').trim();
}

/** สร้างใบลงชื่อของเดือนที่ตั้งไว้ คนละแท็บ (ดึงจากแท็บเช็คอินและการลาของไฟล์หลัก) */
function makeSignSheets() {
  const ss = signBook(), tz = 'Asia/Bangkok';
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { /* รันจากหน้าแก้สคริปต์ */ }
  const say = msg => ui ? ui.alert(msg) : Logger.log(msg);
  const cfg = ss.getSheetByName(SIGN_CFG);
  if (!cfg) { say('ยังไม่ได้ตั้งค่า: รัน setupSignSheet ก่อน'); return; }
  const month = TH_MONTHS.indexOf(String(cfg.getRange('B3').getValue()).trim()) + 1;
  let yearBE = Number(cfg.getRange('B4').getValue());
  if (yearBE && yearBE < 2400) yearBE += 543;
  if (!month || !yearBE) { say('เลือกเดือนและปี (พ.ศ.) ในแท็บ "' + SIGN_CFG + '" ก่อน'); return; }
  const year = yearBE - 543, days = new Date(year, month, 0).getDate();
  const ym = year + '-' + ('0' + month).slice(-2);
  const keyOf = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');

  const centers = {};
  cfg.getRange(7, 1, SIGN_CENTERS.length, 2).getValues().forEach(r => { if (r[0]) centers[r[0]] = r[1]; });
  const signer = cfg.getRange(9 + SIGN_CENTERS.length, 2, SIGN_SIGNER.length, 1).getValues().map(r => r[0]);

  // เช็คอิน: เวลาเข้า = เช็คอินแรกของวัน, เวลาออก = เช็คเอาท์สุดท้ายของวัน (ไม่นับรายการทดสอบ)
  const people = {};
  const person = n => people[n] || (people[n] = { regions: {}, monthRegions: {}, days: {}, leaves: {} });
  const log = book().getSheetByName(LOG);
  if (log && log.getLastRow() > 1) {
    const n = log.getLastRow() - 1;
    const vals = log.getRange(2, 1, n, HEADERS.length).getValues();
    const times = log.getRange(2, 3, n, 1).getDisplayValues();
    vals.forEach((v, i) => {
      const name = normName(v[4]);
      if (!name || !(v[1] instanceof Date) || v[14] === 'ใช่') return;
      const p = person(name), key = keyOf(v[1]), region = v[COL_REGION - 1] || regionOf(v[6]);
      p.regions[region] = (p.regions[region] || 0) + 1;
      if (key.slice(0, 7) !== ym) return;
      const m = String(times[i][0]).match(/^(\d{1,2}):(\d{2})/);
      if (!m) return;
      p.monthRegions[region] = (p.monthRegions[region] || 0) + 1;
      const t = ('0' + m[1]).slice(-2) + ':' + m[2], url = String(v[COL_PHOTO_URL - 1] || '');
      const d = p.days[key] || (p.days[key] = {});
      if (v[3] === 'เช็คเอาท์' || v[3] === 'ออกงาน') {
        if (!d.out || t > d.out.t) d.out = { t: t, url: url };
      } else if (!d.in || t < d.in.t) d.in = { t: t, url: url };
    });
  }
  const lv = book().getSheetByName(LEAVE);
  if (lv && lv.getLastRow() > 1) {
    lv.getRange(2, 1, lv.getLastRow() - 1, LEAVE_HEADERS.length).getValues().forEach(v => {
      const name = normName(v[3]);
      if (!name || !(v[1] instanceof Date)) return;
      const key = keyOf(v[1]);
      if (key.slice(0, 7) === ym) person(name).leaves[key] = v[2];
    });
  }
  const holidays = {};
  const hd = ss.getSheetByName(SIGN_HOLIDAYS);
  if (hd && hd.getLastRow() > 1) {
    hd.getRange(2, 1, hd.getLastRow() - 1, 2).getValues().forEach(v => {
      if (v[0] instanceof Date) holidays[keyOf(v[0])] = String(v[1] || 'วันหยุดราชการ');
    });
  }

  // รายชื่อ: เพิ่มคนที่มีข้อมูลเดือนนี้แต่ยังไม่อยู่ในรายชื่อ
  const nm = ss.getSheetByName(SIGN_NAMES);
  const listed = {};
  if (nm.getLastRow() > 1) {
    nm.getRange(2, 1, nm.getLastRow() - 1, 3).getValues().forEach(r => { if (normName(r[0])) listed[normName(r[0])] = r; });
  }
  const fresh = Object.keys(people).filter(n => !listed[n] &&
    (Object.keys(people[n].days).length || Object.keys(people[n].leaves).length)).sort((a, b) => a.localeCompare(b, 'th'));
  if (fresh.length) {
    nm.getRange(nm.getLastRow() + 1, 1, fresh.length, 3).setValues(fresh.map(n => [n, '', true]));
    fresh.forEach(n => { listed[n] = [n, '', true]; });
  }
  if (nm.getLastRow() > 1) nm.getRange(2, 3, nm.getLastRow() - 1, 1).insertCheckboxes();

  const names = Object.keys(listed).filter(n => listed[n][2] === true).sort((a, b) => a.localeCompare(b, 'th'));
  const top = o => Object.keys(o).sort((a, b) => o[b] - o[a])[0];
  const made = {};
  names.forEach(name => {
    const p = people[name] || { regions: {}, monthRegions: {}, days: {}, leaves: {} };
    const region = listed[name][1] || top(p.monthRegions) || top(p.regions) || OTHER_REGION;
    const head = (centers[region] || centers[OTHER_REGION] || '') + ' ประจำเดือน ' + TH_MONTHS[month - 1] + ' ' + yearBE;
    made[writeSignTab(ss, name, head, year, month, days, p, holidays, signer).getName()] = true;
  });

  // ลบแท็บใบลงชื่อของคนที่เอาออกจากรายชื่อแล้ว
  ss.getSheets().forEach(s => { if (!made[s.getName()] && isSignTab(s)) ss.deleteSheet(s); });
  orderSignTabs(ss);
  say('สร้างใบลงชื่อเดือน ' + TH_MONTHS[month - 1] + ' ' + yearBE + ' แล้ว ' + names.length + ' คน' +
    (fresh.length ? '\nเพิ่มชื่อใหม่ในแท็บรายชื่อ: ' + fresh.join(', ') : '') +
    '\n\nดาวน์โหลดเป็น Excel: ไฟล์ → ดาวน์โหลด → Microsoft Excel (.xlsx)');
}

// หนึ่งแท็บต่อหนึ่งคน หน้าตาเหมือนใบลงชื่อเดิม (ฟอนต์ TH SarabunIT๙ ตอนเปิดใน Excel)
function writeSignTab(ss, name, head, year, month, days, p, holidays, signer) {
  const tab = name.replace(/[\[\]*?\/\\:]/g, ' ').slice(0, 99);
  let sh = ss.getSheetByName(tab);
  if (!sh) {
    sh = ss.insertSheet(tab);
    sh.addDeveloperMetadata(SIGN_TAG);
  }
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  trimColumns(sh, 8);
  const need = 4 + days + 1 + signer.length;
  if (sh.getMaxRows() > need) sh.deleteRows(need + 1, sh.getMaxRows() - need);
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.setHiddenGridlines(true);

  sh.getRange(1, 1, need, 8).setFontFamily(SIGN_FONT).setFontSize(18).setVerticalAlignment('middle');
  sh.setRowHeights(1, need, 31);
  [116, 68, 116, 130, 108, 130, 112, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange('A1:H1').merge().setValue('ใบลงชื่อปฏิบัติงาน').setHorizontalAlignment('center');
  sh.getRange('A2:H2').merge().setValue(head).setHorizontalAlignment('center');
  sh.getRange('A4:H4').setValues([['วันที่', 'ชื่อ-สกุล', '', 'ลงชื่อเข้างาน', 'เวลาเข้างาน', 'ลงชื่อออกงาน', 'เวลาออกงาน', 'หมายเหตุ']])
    .setHorizontalAlignment('center');

  const mm = ('0' + month).slice(-2), rows = [], gray = [];
  const link = (x, text) => x.url ? '=HYPERLINK("' + x.url + '","' + text + '")' : text;
  for (let d = 1; d <= days; d++) {
    const key = year + '-' + mm + '-' + ('0' + d).slice(-2);
    const dow = new Date(year, month - 1, d).getDay(), day = p.days[key] || {};
    const off = holidays[key] || (dow === 0 || dow === 6 ? TH_DAYS[dow] : '');
    if (off) gray.push(d);
    rows.push([d + '/' + mm + '/' + (year + 543), name, '',
      day.in ? link(day.in, 'เช็คอินผ่านระบบ') : '', day.in ? timeSerial(day.in.t) : '',
      day.out ? link(day.out, 'เช็คเอาท์ผ่านระบบ') : '', day.out ? timeSerial(day.out.t) : '',
      p.leaves[key] || off]);
  }
  sh.getRange(5, 1, days, 1).setNumberFormat('@');   // วันที่แบบ พ.ศ. เป็นข้อความ แสดงเหมือนกันทั้ง Sheet และ Excel
  sh.getRange(5, 1, days, 8).setValues(rows);
  sh.getRange(5, 5, days, 1).setNumberFormat('HH:mm');
  sh.getRange(5, 7, days, 1).setNumberFormat('HH:mm');
  sh.getRange(4, 2, days + 1, 2).mergeAcross();
  sh.getRange(5, 1, days, 7).setHorizontalAlignment('center');
  sh.getRange(5, 8, days, 1).setHorizontalAlignment('left');
  sh.getRange(5, 4, days, 1).setFontSize(16);
  sh.getRange(5, 6, days, 1).setFontSize(16);
  sh.getRange(4, 1, days + 1, 8).setBorder(true, true, true, true, true, true);
  gray.forEach(d => sh.getRange(4 + d, 1, 1, 8).setBackground('#e8e8e8'));

  signer.forEach((line, i) => sh.getRange(6 + days + i, 5, 1, 4).merge().setValue(line).setHorizontalAlignment('center'));
  return sh;
}
