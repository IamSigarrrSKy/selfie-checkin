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
  ['ปส. ภาคตะวันออก (ศาลากลาง จ.ระยอง)', 'ระยอง', 'ภาคตะวันออก', 12.70726, 101.18380, 300]
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
      if (!row) return out({ ok: false, error: 'ไม่พบรายการใน Sheet' });
      const cur = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      sh.getRange(row, 1, 1, HEADERS.length).setValues([toRow(r, cur[COL_PHOTO_URL - 1], cur[COL_CREATED - 1])]);
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
  return [r.id, dateSerial(r.date), timeSerial(r.time), r.type, r.name, r.empId, r.site, r.dist, r.inRange,
    r.lat, r.lng, r.acc, r.note, photo, r.mock, created, "", "",   // ไม่บันทึกประวัติการแก้ไขลง Sheet
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

/* ===================== ออกแบบชีต ===================== */

/** กด Run ฟังก์ชันนี้หนึ่งครั้ง: สร้าง/จัดรูปแบบทุกชีต (รันซ้ำได้ ข้อมูลเดิมไม่หาย) */
function setup() {
  const ss = book();
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
  if (sh.getLastRow() > 1) sh.getRange(2, 17, sh.getLastRow() - 1, 2).clearContent();

  // เติมภาคให้แถวเก่าที่ยังไม่มี
  const last = sh.getLastRow();
  if (last > 1) {
    const sitesCol = sh.getRange(2, 7, last - 1, 1).getValues();
    const regionCol = sh.getRange(2, COL_REGION, last - 1, 1).getValues();
    sh.getRange(2, COL_REGION, last - 1, 1).setValues(regionCol.map((v, i) => [v[0] || regionOf(sitesCol[i][0])]));
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
      .setFontColor(C.muted).setItalic(true).setRanges([all]).build()
  ]);
  return sh;
}

/* ---------- สรุปรายวัน (สูตรคำนวณเอง ไม่ต้องแก้) ---------- */
function buildDaily(ss) {
  const sh = sheet(ss, DAILY);
  sh.clear();
  const L = "'" + LOG + "'!";
  const head = ['วันที่', 'สถานที่', 'เช็คอินครั้งแรก', 'ล่าสุด', 'จำนวนครั้ง', 'หมายเหตุ'];
  trimColumns(sh, head.length);
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 1, head.length);

  const real = L + 'O2:O,"<>ใช่"';   // ไม่นับรายการโหมดทดสอบ
  sh.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + L + 'B2:B,' + L + 'B2:B<>"",' + L + 'O2:O<>"ใช่")),1,FALSE),)');
  sh.getRange('B2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER(' + L + 'G2:G,' + L + 'B2:B=d,' + L + 'O2:O<>"ใช่"))),))))');
  sh.getRange('C2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,LET(v,MINIFS(' + L + 'C2:C,' + L + 'B2:B,d,' + real + '),IF(v=0,,v)))))');
  sh.getRange('D2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,LET(v,MAXIFS(' + L + 'C2:C,' + L + 'B2:B,d,' + real + '),IF(v=0,,v)))))');
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

  const head = ['เดือน', 'วันที่เช็คอิน', 'จำนวนครั้ง', 'ครั้งแรกเฉลี่ย', 'สถานที่'];
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
const REGION_TAB_COLORS = { 'ภาคเหนือ': '#2e7d5b', 'ภาคตะวันออกเฉียงเหนือ': '#b8862b', 'ภาคใต้': '#2a6cb0', 'ภาคตะวันออก': '#8a4fb0' };

function buildRegion(ss, name) {
  const sh = sheet(ss, name);
  sh.clear();
  sh.getBandings().forEach(b => b.remove());
  const L = "'" + LOG + "'!", R = '$A$1';
  const real = L + 'S2:S=' + R + ',' + L + 'O2:O<>"ใช่"';   // รายการของภาคนี้ ไม่รวมโหมดทดสอบ
  const head = ['วันที่', 'เวลา', 'สถานที่', 'หมายเหตุ', 'รูป'];
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
  sh.getRange('B3').setFormula('=IFERROR(COUNTA(UNIQUE(FILTER(' + L + 'B2:B,' + real + '))),0)').setNumberFormat('0 "วัน"');
  sh.getRange('D3').setFormula('=COUNTIFS(' + L + 'S2:S,' + R + ',' + L + 'O2:O,"<>ใช่")').setNumberFormat('0 "ครั้ง"');
  sh.getRange('F3').setFormula('=IFERROR(MAX(FILTER(' + L + 'B2:B,' + real + ')),"–")').setNumberFormat('d mmm yyyy');
  sh.getRange('A3:F3').setFontColor(C.muted);
  ['B3', 'D3', 'F3'].forEach(a => sh.getRange(a).setFontWeight('bold').setFontColor(C.title).setHorizontalAlignment('left'));

  // ตารางรายการ: เรียงใหม่สุดก่อน
  sh.getRange(5, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 5, head.length);
  const cols = ['B', 'C', 'G', 'M', 'T'].map(c => L + c + '2:' + c).join(',');
  sh.getRange('A6').setFormula('=IFERROR(LET(f,FILTER({' + cols + '},' + real + '),s,SORT(f,1,FALSE,2,FALSE),' +
    'HSTACK(CHOOSECOLS(s,1,2,3,4),' +
    'MAP(CHOOSECOLS(s,5),LAMBDA(u,IF(u="","",HYPERLINK(u,"ดูรูป")))))),"ยังไม่มีรายการ")');

  const rows = sh.getMaxRows() - 5;
  sh.getRange(6, 1, rows, 1).setNumberFormat('ddd d mmm yyyy').setHorizontalAlignment('left');
  sh.getRange(6, 2, rows, 1).setNumberFormat('HH:mm').setHorizontalAlignment('center');
  sh.getRange(6, 5, rows, 1).setHorizontalAlignment('center');
  sh.getRange(6, 3, rows, 2).setWrap(true);
  [130, 70, 320, 260, 70, 90].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  band(sh, sh.getRange(6, 1, rows, head.length));
  sh.setConditionalFormatRules([]);
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
  const rows = l.days.map(d => [l.id, dateSerial(d), l.type, l.name, l.empId, l.note, l.created]);
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

/* ---------- สรุปการลารายเดือน ---------- */
function buildLeaveSummary(ss) {
  const sh = sheet(ss, LEAVE_SUM);
  sh.clear();
  sh.getBandings().forEach(b => b.remove());
  const V = "'" + LEAVE + "'!";
  const head = ['เดือน', 'ลาป่วย', 'ลาพักผ่อน', 'รวม', 'วันที่ลา'];
  trimColumns(sh, head.length);
  sh.setTabColor('#b4282f');

  sh.getRange('A1').setValue('สรุปการลารายเดือน').setFontSize(16).setFontWeight('bold').setFontColor(C.title);
  sh.setRowHeight(1, 34);
  const yr = (type) => 'COUNTIFS(' + V + 'B2:B,">="&DATE(YEAR(TODAY()),1,1),' + V + 'B2:B,"<="&DATE(YEAR(TODAY()),12,31),' + V + 'C2:C,"' + type + '")';
  sh.getRange('A2').setFormula('="ปีนี้ · ลาป่วย "&' + yr('ลาป่วย') + '&" วัน · ลาพักผ่อน "&' + yr('ลาพักผ่อน') + '&" วัน"')
    .setFontColor(C.muted);

  sh.getRange(4, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 4, head.length);

  const inM = 'B2:B,">="&m,' + V + 'B2:B,"<="&EOMONTH(m,0)';
  const days = (type) => 'IFERROR("' + type.replace('ลา', '') + ' "&TEXTJOIN(", ",TRUE,ARRAYFORMULA(TEXT(SORT(FILTER(' + V + 'B2:B,' +
    V + 'B2:B>=m,' + V + 'B2:B<=EOMONTH(m,0),' + V + 'C2:C="' + type + '")),"d"))),"")';
  sh.getRange('A5').setFormula('=IFERROR(SORT(UNIQUE(FILTER(ARRAYFORMULA(EOMONTH(' + V + 'B2:B,-1)+1),' + V + 'B2:B<>"")),1,FALSE),)');
  sh.getRange('B5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,COUNTIFS(' + V + inM + ',' + V + 'C2:C,"ลาป่วย"))))');
  sh.getRange('C5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,COUNTIFS(' + V + inM + ',' + V + 'C2:C,"ลาพักผ่อน"))))');
  sh.getRange('D5').setFormula('=MAP(B5:B,C5:C,LAMBDA(s,v,IF(AND(s="",v=""),,s+v)))');
  sh.getRange('E5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,TRIM(' + days('ลาป่วย') + '&"   "&' + days('ลาพักผ่อน') + '))))');

  const rows = sh.getMaxRows() - 4;
  sh.getRange(5, 1, rows, 1).setNumberFormat('mmmm yyyy').setHorizontalAlignment('left');
  sh.getRange(5, 2, rows, 3).setNumberFormat('0 "วัน"').setHorizontalAlignment('center');
  sh.getRange(5, 5, rows, 1).setWrap(true);
  [140, 100, 110, 90, 380].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  band(sh, sh.getRange(5, 1, rows, head.length));
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setFontColor('#b4282f').setBold(true).setRanges([sh.getRange(5, 2, rows, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setFontColor('#0e5a52').setBold(true).setRanges([sh.getRange(5, 3, rows, 1)]).build()
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
