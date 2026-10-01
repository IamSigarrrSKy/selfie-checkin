/**
 * รับข้อมูลลงเวลาจากแอปเซลฟี่ → เขียนลง Google Sheet + เก็บรูปใน Google Drive
 * ติดตั้ง: ดู README.md ในโฟลเดอร์นี้ แล้วกด Run ฟังก์ชัน setup() หนึ่งครั้งเพื่อสร้างชีตทั้งหมด
 */
const SHEET_ID    = '';   // เว้นว่างได้ถ้าเปิด Apps Script จากเมนู ส่วนขยาย ของ Sheet นั้นเอง
const TOKEN       = '';   // ตั้งรหัสลับได้ ต้องกรอกให้ตรงกันในหน้าตั้งค่าของแอป
const FOLDER_NAME = 'รูปเช็คอินเซลฟี่';

const LOG = 'ลงเวลา', DAILY = 'สรุปรายวัน', MONTHLY = 'สรุปรายเดือน', SITES = 'จุดทำงาน';

const HEADERS = ['ID', 'วันที่', 'เวลา', 'ประเภท', 'ชื่อ', 'รหัสพนักงาน', 'จุดทำงาน', 'ระยะ (ม.)', 'ในพื้นที่',
  'ละติจูด', 'ลองจิจูด', 'ความแม่นยำ (ม.)', 'หมายเหตุ', 'รูป', 'ทดสอบ', 'บันทึกเมื่อ', 'แก้ไขเมื่อ', 'เวลาเดิมก่อนแก้'];
const COL_PHOTO = 14, COL_CREATED = 16;

const SITE_LIST = [
  ['อุทยานวิทยาศาสตร์และเทคโนโลยี ม.เชียงใหม่', 'เชียงใหม่', 18.76467, 98.93700, 300],
  ['อุทยานวิทยาศาสตร์ ม.ขอนแก่น', 'ขอนแก่น', 16.45588, 102.81942, 300],
  ['อุทยานวิทยาศาสตร์ ม.สงขลานครินทร์', 'สงขลา', 7.02194, 100.55542, 300],
  ['ปส. ภาคตะวันออก (ศาลากลาง จ.ระยอง)', 'ระยอง', 12.70726, 101.18380, 300]
];

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

// วันที่ "2026-10-01" และเวลา "08:30:00" ถูกเขียนเป็นข้อความ แล้ว Sheet แปลงเป็นวันที่/เวลาจริงให้เอง
function toRow(r, photoUrl, created) {
  const photo = photoUrl ? '=HYPERLINK("' + photoUrl + '","ดูรูป")' : '';
  return [r.id, r.date, r.time, r.type, r.name, r.empId, r.site, r.dist, r.inRange,
    r.lat, r.lng, r.acc, r.note, photo, r.mock, created, r.editedAt, r.origTime];
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
  getFolder();

  // เรียงแท็บ: สรุปรายเดือน → สรุปรายวัน → ลงเวลา → จุดทำงาน
  [MONTHLY, DAILY, LOG, SITES].forEach((name, i) => { ss.setActiveSheet(ss.getSheetByName(name)); ss.moveActiveSheet(i + 1); });
  // ลบชีตว่างเริ่มต้น (Sheet1 / แผ่น1)
  ss.getSheets().forEach(s => {
    if ([MONTHLY, DAILY, LOG, SITES].indexOf(s.getName()) < 0 && s.getLastRow() === 0 && s.getLastColumn() === 0) ss.deleteSheet(s);
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
}

function band(sh, range) {
  sh.getBandings().forEach(b => b.remove());
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor('#ffffff').setSecondRowColor(C.band);
}

/* ---------- ลงเวลา (ข้อมูลดิบจากแอป) ---------- */
function buildLog(ss) {
  const sh = sheet(ss, LOG), n = HEADERS.length;
  sh.getRange(1, 1, 1, n).setValues([HEADERS]);
  styleHeader(sh, 1, n);
  trimColumns(sh, n);
  sh.setFrozenColumns(3);

  const widths = [70, 95, 75, 75, 140, 95, 260, 70, 70, 90, 90, 90, 220, 70, 60, 140, 140, 140];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.hideColumns(1);   // ID ใช้อ้างอิงตอนแก้ไขจากแอป ไม่ต้องเห็น

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
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I2="ไม่"')
      .setBackground(C.warn).setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$Q2<>""')
      .setBackground(C.edit).setRanges([sh.getRange(2, 2, rows, 3), sh.getRange(2, 17, rows, 2)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('เข้างาน')
      .setFontColor('#1d7a46').setBold(true).setRanges([sh.getRange(2, 4, rows, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ออกงาน')
      .setFontColor('#a15c00').setBold(true).setRanges([sh.getRange(2, 4, rows, 1)]).build()
  ]);
  return sh;
}

/* ---------- สรุปรายวัน (สูตรคำนวณเอง ไม่ต้องแก้) ---------- */
function buildDaily(ss) {
  const sh = sheet(ss, DAILY);
  sh.clear();
  const L = "'" + LOG + "'!";
  const head = ['วันที่', 'จุดทำงาน', 'เข้างาน', 'ออกงาน', 'ชั่วโมงทำงาน', 'นอกพื้นที่', 'แก้ไข', 'หมายเหตุ'];
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 1, head.length);
  trimColumns(sh, head.length);

  const real = L + 'O2:O,"<>ใช่"';   // ไม่นับรายการโหมดทดสอบ
  sh.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + L + 'B2:B,' + L + 'B2:B<>"",' + L + 'O2:O<>"ใช่")),1,FALSE),)');
  sh.getRange('B2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER(' + L + 'G2:G,' + L + 'B2:B=d,' + L + 'O2:O<>"ใช่"))),))))');
  sh.getRange('C2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,LET(v,MINIFS(' + L + 'C2:C,' + L + 'B2:B,d,' + L + 'D2:D,"เข้างาน",' + real + '),IF(v=0,,v)))))');
  sh.getRange('D2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,LET(v,MAXIFS(' + L + 'C2:C,' + L + 'B2:B,d,' + L + 'D2:D,"ออกงาน",' + real + '),IF(v=0,,v)))))');
  sh.getRange('E2').setFormula('=MAP(C2:C,D2:D,LAMBDA(i,o,IF(OR(i="",o=""),,IF(o>i,o-i,))))');
  sh.getRange('F2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,IF(COUNTIFS(' + L + 'B2:B,d,' + L + 'I2:I,"ไม่",' + real + ')>0,"นอกพื้นที่",))))');
  sh.getRange('G2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,IF(COUNTIFS(' + L + 'B2:B,d,' + L + 'Q2:Q,"<>",' + real + ')>0,"แก้ไข",))))');
  sh.getRange('H2').setFormula('=MAP(A2:A,LAMBDA(d,IF(d="",,IFERROR(TEXTJOIN(" / ",TRUE,FILTER(' + L + 'M2:M,' + L + 'B2:B=d,' + L + 'M2:M<>"",' + L + 'O2:O<>"ใช่")),))))');

  const rows = sh.getMaxRows() - 1;
  sh.getRange(2, 1, rows, 1).setNumberFormat('ddd d mmm yyyy');
  sh.getRange(2, 3, rows, 2).setNumberFormat('HH:mm');
  sh.getRange(2, 5, rows, 1).setNumberFormat('[h]:mm');
  sh.getRange(2, 1, rows, 1).setHorizontalAlignment('left');
  sh.getRange(2, 3, rows, 5).setHorizontalAlignment('center');
  sh.getRange(2, 8, rows, 1).setWrap(true);
  [130, 260, 80, 80, 100, 90, 70, 280].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  band(sh, sh.getRange(2, 1, rows, head.length));

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('นอกพื้นที่')
      .setBackground(C.warn).setFontColor('#a15c00').setRanges([sh.getRange(2, 6, rows, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('แก้ไข')
      .setBackground(C.edit).setFontColor('#2a5db0').setRanges([sh.getRange(2, 7, rows, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",OR($C2="",$D2=""))')
      .setFontColor(C.muted).setRanges([sh.getRange(2, 5, rows, 1)]).build()
  ]);
  return sh;
}

/* ---------- สรุปรายเดือน ---------- */
function buildMonthly(ss) {
  const sh = sheet(ss, MONTHLY);
  sh.clear();
  const D = "'" + DAILY + "'!";
  sh.getRange('A1').setValue('สรุปการลงเวลาปฏิบัติงานต่างจังหวัด').setFontSize(16).setFontWeight('bold').setFontColor(C.title);
  sh.getRange('A2').setFormula('="ผู้ลงเวลา: "&IFERROR(INDEX(FILTER(\'' + LOG + '\'!E2:E,\'' + LOG + '\'!E2:E<>""),1),"–")&"   ·   อัปเดตอัตโนมัติจากแอป"')
    .setFontColor(C.muted);
  sh.setRowHeight(1, 34);

  const head = ['เดือน', 'วันที่มาทำงาน', 'ชั่วโมงรวม', 'เฉลี่ย/วัน', 'เข้างานเฉลี่ย', 'วันนอกพื้นที่', 'จุดทำงานที่ไป'];
  sh.getRange(4, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 4, head.length);
  trimColumns(sh, head.length);

  const range = (m) => D + 'A2:A,">="&' + m + ',' + D + 'A2:A,"<="&EOMONTH(' + m + ',0)';
  sh.getRange('A5').setFormula('=IFERROR(SORT(UNIQUE(FILTER(ARRAYFORMULA(EOMONTH(' + D + 'A2:A,-1)+1),' + D + 'A2:A<>"")),1,FALSE),)');
  sh.getRange('B5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,COUNTIFS(' + range('m') + '))))');
  sh.getRange('C5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,SUMIFS(' + D + 'E2:E,' + range('m') + '))))');
  sh.getRange('D5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,IFERROR(AVERAGEIFS(' + D + 'E2:E,' + range('m') + ',' + D + 'E2:E,">0"),))))');
  sh.getRange('E5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,IFERROR(AVERAGEIFS(' + D + 'C2:C,' + range('m') + ',' + D + 'C2:C,"<>"),))))');
  sh.getRange('F5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,COUNTIFS(' + range('m') + ',' + D + 'F2:F,"นอกพื้นที่"))))');
  sh.getRange('G5').setFormula('=MAP(A5:A,LAMBDA(m,IF(m="",,IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER(' + D + 'B2:B,' + D + 'A2:A>=m,' + D + 'A2:A<=EOMONTH(m,0)))),))))');

  const rows = sh.getMaxRows() - 4;
  sh.getRange(5, 1, rows, 1).setNumberFormat('mmmm yyyy').setHorizontalAlignment('left');
  sh.getRange(5, 2, rows, 1).setNumberFormat('0 "วัน"');
  sh.getRange(5, 3, rows, 2).setNumberFormat('[h]:mm "ชม."');
  sh.getRange(5, 5, rows, 1).setNumberFormat('HH:mm');
  sh.getRange(5, 6, rows, 1).setNumberFormat('0 "วัน";;"–"');
  sh.getRange(5, 2, rows, 5).setHorizontalAlignment('center');
  sh.getRange(5, 7, rows, 1).setWrap(true);
  [140, 110, 110, 100, 110, 110, 320].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  band(sh, sh.getRange(5, 1, rows, head.length));
  return sh;
}

/* ---------- จุดทำงาน (อ้างอิง) ---------- */
function buildSites(ss) {
  const sh = sheet(ss, SITES);
  sh.clear();
  const head = ['จุดทำงาน', 'จังหวัด', 'ละติจูด', 'ลองจิจูด', 'รัศมี (ม.)', 'แผนที่'];
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  styleHeader(sh, 1, head.length);
  trimColumns(sh, head.length);
  const rows = SITE_LIST.map(s => s.concat(['=HYPERLINK("https://www.google.com/maps?q=' + s[2] + ',' + s[3] + '","เปิดแผนที่")']));
  sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.getRange(2, 3, rows.length, 2).setNumberFormat('0.00000');
  sh.getRange(2, 2, rows.length, 5).setHorizontalAlignment('center');
  [300, 100, 100, 100, 90, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(rows.length + 3, 1).setValue('รายการนี้ใช้อ้างอิงเท่านั้น การเพิ่มหรือแก้จุดทำงานให้ทำในแท็บ "จุดทำงาน" ของแอป')
    .setFontColor(C.muted).setFontStyle('italic');
  band(sh, sh.getRange(2, 1, rows.length, head.length));
  return sh;
}
