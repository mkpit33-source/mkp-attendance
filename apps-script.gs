// ============================================================
//  ระบบเช็คชื่อรายวันตอนเช้า — ทั้งโรงเรียนเมืองแกพิทยาสรรค์
//  (ระบบแยกอิสระ ไม่เกี่ยวกับเว็บห้องเรียนส่วนตัวของครูม่อน)
//
//  ⚠️ ระบบนี้เปิดสาธารณะ ไม่มีรหัสผ่าน/PIN ป้องกันใดๆ ทั้งสิ้น
//     เป็นการตัดสินใจของครูที่รับทราบความเสี่ยงเรื่องข้อมูลนักเรียน
//     (ชื่อนักเรียนที่ขาด/ลากิจ/ลาป่วย/ไปกิจกรรม) แล้วว่าใครก็ตามที่มีลิงก์ระบบ
//     จะเข้าถึงข้อมูลนี้ได้ทันที — ไม่ใช่ความผิดพลาดที่ลืมใส่การป้องกัน
//
//  เดิมโครงสร้างตรงกับฟอร์มกระดาษ "แบบรายงานสถิติการมาเรียนประจำวัน" ของครู (แยกนับชาย/หญิง/รวม
//  ทุกหมวด) — ตัดการแยกเพศออกแล้ว 2569 (ครูยอมรับว่าไม่ตรงกับฟอร์มกระดาษนั้นอีกต่อไป) ตอนนี้นับ
//  แค่ยอดรวมต่อหมวด (มา/ขาด/ลากิจ/ลาป่วย/ไปกิจกรรม) + มี "% มาเรียน" ต่อห้องในแดชบอร์ดแทน
//
//  ขั้นตอนติดตั้ง (ทำครั้งเดียว):
//  1. สร้าง Google Sheet ใหม่เปล่าๆ 1 ไฟล์ (ชื่ออะไรก็ได้ เช่น "เช็คชื่อรายวัน MKP")
//  2. คัดลอก ID ของ Sheet จาก URL (ส่วนตัวยาวๆ ระหว่าง /d/ กับ /edit)
//     แล้วนำมาแทนที่ข้อความ 'ใส่ ID ของ Google Sheet ที่สร้างใหม่ตรงนี้' ด้านล่าง
//  3. เปิดเมนู Extensions > Apps Script ของ Sheet นั้น แล้ววางโค้ดทั้งไฟล์นี้ทับ
//  4. รันฟังก์ชัน setupSheets() หนึ่งครั้ง (เลือกฟังก์ชันนี้ที่ dropdown ด้านบน แล้วกด Run)
//     — จะสร้าง 8 sheets ให้อัตโนมัติ: "ห้องเรียน", "นักเรียน", "เช็คชื่อรายวัน", "วันหยุดเพิ่มเติม",
//       "เช็คชื่อเย็น", "คะแนนความประพฤติ", "ตั้งค่ากิจการนักเรียน", "เกณฑ์คะแนนความประพฤติ"
//       (5 อันหลังเพิ่มเข้ามาปี 2569)
//     — Sheet "ห้องเรียน" จะใส่ห้องเริ่มต้นให้ 7 ห้องอัตโนมัติ (ม.1/1, ม.2/1, ม.3/1, ม.3/2, ม.4/1, ม.5/1, ม.6/1)
//     ⚠️ ถ้าระบบเดิมใช้งานอยู่แล้ว (เคยรัน setupSheets() รอบก่อนหน้านี้) ให้วางโค้ดใหม่ทับแล้วรัน
//        setupSheets() ซ้ำอีกครั้ง — โค้ดจะสร้างเฉพาะ sheet ใหม่ที่ยังไม่มีให้อัตโนมัติเท่านั้น
//        ไม่กระทบ/ไม่ลบข้อมูลเดิมใน sheet เก่าแต่อย่างใด (เช็คด้วย `if (!s) s = ss.insertSheet(name)`)
//       แก้ไข/เพิ่ม/ลบเองทีหลังได้ตามจริง
//  4.5 ⚠️ สำคัญ: หลัง setupSheets() ครั้งแรก ให้เปิดหน้า settings.html (PIN เริ่มต้น 8 หลัก
//        คือ 12345678 — ครูกำหนดเองแล้ว) แล้วเปลี่ยน PIN อีก 2 ระดับที่เหลือทันที (เช็คชื่อ 4 หลัก
//        ค่าเริ่มต้น 0000 / คะแนนความประพฤติ 6 หลัก ค่าเริ่มต้น 000000 — ยังเป็นค่าเดาง่าย ห้ามปล่อย
//        ไว้แบบนั้นตอนใช้งานจริง) พร้อมทั้งเปลี่ยน PIN 8 หลักนี้เป็นเลขที่เดายากกว่านี้ในภายหลังด้วย
//  5. เปิด Sheet "นักเรียน" → ใส่รายชื่อนักเรียนจริงของทุกห้อง 4 คอลัมน์: ห้อง, เลขที่, ชื่อ-นามสกุล, เพศ
//     ⚠️ คอลัมน์ "เพศ" ต้องพิมพ์ตรงเป๊ะว่า "ชาย" หรือ "หญิง" เท่านั้น (ไม่งั้นระบบจะนับแยกชาย/หญิงผิด
//     แต่ยอดรวมทั้งหมดยังถูกอยู่ เพราะนับจากจำนวนคนไม่ได้อิงเพศ)
//     — ขั้นตอนนี้ครูกรอกเอง ไม่ต้องให้ใครเห็นข้อมูลนี้ผ่านช่องทางอื่น
//     ⚠️ ถ้าเคยรัน setupSheets() ไปแล้วรอบก่อน (Sheet "เช็คชื่อรายวัน" มีข้อมูลอยู่แล้ว) ระบบจะไม่เพิ่ม
//        หัวคอลัมน์ใหม่ให้อัตโนมัติ — ไปเพิ่มคำว่า "ผู้รายงาน" เองในคอลัมน์ถัดจากคอลัมน์สุดท้ายได้เลย
//        (แค่พิมพ์หัวตาราง ไม่มีคอลัมน์นี้ก็ไม่กระทบการทำงานของระบบ เป็นแค่ความสะดวกตอนเปิดดูชีต)
//  6. Deploy > New deployment > เลือกประเภท "Web app"
//     - Execute as: Me
//     - Who has access: Anyone
//     กด Deploy แล้วคัดลอก URL ที่ได้
//  7. นำ URL ไปแทนที่ SCRIPT_URL ในไฟล์ report.html, report-evening.html, behavior.html และ dashboard.html
//
//  (เดิมมีระบบแจ้งเตือนอัตโนมัติเข้า LINE Official Account 5 แบบ — ตัดออกถาวรแล้ว 2569
//   ครูใช้ปุ่ม "คัดลอกส่งต่อ" ในหน้า dashboard.html ส่งเข้า LINE/OpenChat เองแทน)
// ============================================================

const SPREADSHEET_ID = '1CkgncHeiyJJv5Mixbe39DEphABpMPmEYRGAM2OsBKlI';
const TZ = 'Asia/Bangkok'; // ใช้เวลาไทยเป็นหลักเสมอ กันปัญหาเรื่องเขตเวลาของเซิร์ฟเวอร์ Google

// ------------------------------------------------------------
//  สร้าง 4 Sheets พร้อม header (รันครั้งเดียวตอนตั้งระบบ) + ใส่ห้องเริ่มต้น 7 ห้อง
// ------------------------------------------------------------
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheets = [
    { name: 'ห้องเรียน',        headers: ['ห้อง'] },
    { name: 'นักเรียน',          headers: ['ห้อง', 'เลขที่', 'ชื่อ-นามสกุล', 'เพศ'] },
    // เดิมแยกนับชาย/หญิง ตรงกับฟอร์มกระดาษราชการ "แบบรายงานสถิติการมาเรียนประจำวัน" — ครูตัดสินใจ
    // เอาการแยกเพศออกแล้ว (2569) ใช้ "% มาเรียน" ต่อห้องแทน (คำนวณสดฝั่ง dashboard ไม่ได้เก็บในชีท)
    // โครงสร้างตอนนี้เหมือน Sheet "เช็คชื่อเย็น" เป๊ะ (ไม่มีเพศ)
    { name: 'เช็คชื่อรายวัน',    headers: [
        'วันที่', 'ห้อง', 'เวลาที่ส่ง',
        'มา-รวม',
        'ขาด-รวม', 'รายชื่อขาด',
        'ลากิจ-รวม', 'รายชื่อลากิจ',
        'ลาป่วย-รวม', 'รายชื่อลาป่วย',
        'ไปกิจกรรม-รวม', 'รายชื่อไปกิจกรรม',
        'ผู้รายงาน',
      ] },
    { name: 'วันหยุดเพิ่มเติม',  headers: ['วันที่ (yyyy-mm-dd)', 'หมายเหตุ'] },
    // ---- งานกิจการนักเรียน: เช็คกลับ 15:30 (เพิ่ม 2569) ----
    // เช็คอิสระเต็มรูปแบบเหมือนเช้าทุกประการ (ไม่กรองจากข้อมูลเช้า) หมวดหมู่เดียวกับเช้าเป๊ะ
    // (มา/ขาด/ลากิจ/ลาป่วย/ไปกิจกรรม) แค่ไม่แยกชาย/หญิง เพราะไม่มีฟอร์มกระดาษราชการบังคับ
    { name: 'เช็คชื่อเย็น', headers: [
        'วันที่', 'ห้อง', 'เวลาที่ส่ง',
        'มา-รวม',
        'ขาด-รวม', 'รายชื่อขาด',
        'ลากิจ-รวม', 'รายชื่อลากิจ',
        'ลาป่วย-รวม', 'รายชื่อลาป่วย',
        'ไปกิจกรรม-รวม', 'รายชื่อไปกิจกรรม',
        'ผู้รายงาน',
      ] },
    // ---- งานกิจการนักเรียน: คะแนนความประพฤติ (เพิ่ม 2569) ----
    // ledger แบบ append-only ทุกเหตุการณ์เป็นคนละแถว ไม่ใช่ยอดสะสมที่แก้ทับ —
    // คะแนนคงเหลือคำนวณสดจากผลรวมทุกครั้งที่อ่าน (ดู handleBehaviorSummary)
    { name: 'คะแนนความประพฤติ', headers: [
        'วันที่', 'เวลา', 'ภาคเรียน', 'ห้อง', 'ชื่อนักเรียน',
        'ประเภท', 'เหตุผล', 'คะแนน', 'ผู้รายงาน',
      ] },
    { name: 'ตั้งค่ากิจการนักเรียน', headers: ['รายการ', 'ค่า'] },
    // ---- ระบบตั้งค่าเบื้องหลัง (เพิ่ม 2569) — ให้แก้ผ่านหน้าเว็บได้ ไม่ต้องพึ่งครูภูริณัฐคนเดียว ----
    { name: 'เกณฑ์คะแนนความประพฤติ', headers: ['เหตุผล', 'ประเภท', 'คะแนน'] },
  ];

  sheets.forEach(({ name, headers }) => {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (s.getLastRow() === 0) s.appendRow(headers);
  });

  // ตั้งค่าเริ่มต้นภาคเรียนปัจจุบัน (ถ้ายังไม่มีใครกรอก) — ครูกิจการนักเรียนแก้ค่านี้เองตอนขึ้นภาคเรียนใหม่
  // คะแนนความประพฤติจะคำนวณกรองด้วยค่านี้เท่านั้น แถวเก่าภาคเรียนก่อนหน้ายังอยู่ในชีทเป็นประวัติ ไม่ลบทิ้ง
  const cfgSheet = ss.getSheetByName('ตั้งค่ากิจการนักเรียน');
  if (cfgSheet.getLastRow() <= 1) {
    cfgSheet.appendRow(['ภาคเรียนปัจจุบัน', '1/2569']);
  }

  // ⚠️ PIN เริ่มต้น 3 ระดับ (เพิ่ม 2569) — เป็นค่า placeholder เดาง่ายโดยตั้งใจ ต้องเข้าหน้า
  // settings.html แล้วเปลี่ยนทันทีหลัง deploy ครั้งแรก (แก้ผ่านหน้าเว็บได้เลย ไม่ต้องมาแก้ในชีทตรงๆ)
  const pinDefaults = [
    ['PIN เช็คชื่อ', '0000'],
    ['PIN คะแนนความประพฤติ', '000000'],
    ['PIN ตั้งค่าเบื้องหลัง', '12345678'], // ครูกำหนดค่าเริ่มต้นนี้เอง 2026-08-03 — ควรเปลี่ยนเป็นเลขที่เดายากกว่านี้ทีหลังผ่านหน้า settings.html
  ];
  const cfgRows = cfgSheet.getDataRange().getValues();
  pinDefaults.forEach(([key, def]) => {
    const exists = cfgRows.some(r => String(r[0]).trim() === key);
    if (!exists) cfgSheet.appendRow([key, def]);
  });

  // ใส่เกณฑ์คะแนนความประพฤติเริ่มต้น (ถ้ายังไม่มีใครกรอก) — ย้ายมาจากที่เคย hardcode ใน behavior.html
  const critSheet = ss.getSheetByName('เกณฑ์คะแนนความประพฤติ');
  if (critSheet.getLastRow() <= 1) {
    const defaultCriteria = [
      ['ทรงผม', 'หัก', 5],
      ['การแต่งกาย (เสื้อลอยชาย)', 'หัก', 5],
      ['รองเท้า', 'หัก', 5],
      ['มาสาย หลังเวลา 08.05 น.', 'หัก', 5],
      ['หนีเรียน', 'หัก', 10],
      ['หนีกิจกรรม', 'หัก', 10],
      ['ไม่สวมหมวกนิรภัย', 'หัก', 10],
      ['การพนันและสิ่งเสพติด', 'หัก', 20],
      ['ขับขี่ยานพาหนะสร้างความเดือดร้อนรำคาญ', 'หัก', 20],
    ];
    defaultCriteria.forEach(row => critSheet.appendRow(row));
  }

  // ใส่ห้องเริ่มต้นให้ทั้ง 7 ห้อง ถ้ายังไม่มีใครกรอกอะไรเลย (แก้ไข/เพิ่ม/ลบเองทีหลังได้)
  const roomSheet = ss.getSheetByName('ห้องเรียน');
  if (roomSheet.getLastRow() <= 1) {
    const defaultRooms = ['ม.1/1', 'ม.2/1', 'ม.3/1', 'ม.3/2', 'ม.4/1', 'ม.5/1', 'ม.6/1'];
    defaultRooms.forEach(r => roomSheet.appendRow([r]));
  }

  SpreadsheetApp.flush();
  Logger.log('setupSheets เสร็จแล้ว — ห้องเรียนใส่ค่าเริ่มต้นให้ 7 ห้องแล้ว ไปกรอกรายชื่อนักเรียน+เพศใน Sheet "นักเรียน" ต่อได้เลย');
}

// ------------------------------------------------------------
//  GET — ดึงรายชื่อห้อง / รายชื่อนักเรียน / สรุปยอดวันนี้ / ยอดย้อนหลัง
// ------------------------------------------------------------
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'rooms')    return handleRooms();
    if (action === 'students') return handleStudents(e.parameter);
    if (action === 'mystatus') return handleMyStatus(e.parameter);
    if (action === 'today')    return handleByDate(todayStr());
    if (action === 'bydate')   return handleByDate(e.parameter.date);
    if (action === 'rangeSummary') return handleRangeSummary(e.parameter);
    // ---- งานกิจการนักเรียน: เช็คกลับ 15:30 ----
    if (action === 'eveningByDate')    return handleEveningByDate(e.parameter.date || todayStr());
    if (action === 'dailyComparison')  return handleDailyComparison(e.parameter);
    if (action === 'attendanceRate')   return handleAttendanceRate(e.parameter);
    // ---- งานกิจการนักเรียน: คะแนนความประพฤติ ----
    if (action === 'behaviorSummary') return handleBehaviorSummary(e.parameter);
    if (action === 'currentTerm')     return handleCurrentTerm();
    // ---- งานกิจการนักเรียน: ตั้งค่าเบื้องหลัง ----
    if (action === 'settingsData')   return handleSettingsData();
    return respond({ status: 'error', message: 'unknown action' });
  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

// ------------------------------------------------------------
//  POST — รับรายงานเช็คชื่อจากหัวหน้าห้อง
// ------------------------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'submit')         return handleSubmit(data);
    if (data.action === 'submitEvening')  return handleSubmitEvening(data);
    if (data.action === 'submitBehavior') return handleSubmitBehavior(data);
    // ---- งานกิจการนักเรียน: ตั้งค่าเบื้องหลัง ----
    if (data.action === 'verifyPin')      return handleVerifyPin(data);
    if (data.action === 'updatePin')      return handleUpdatePin(data);
    if (data.action === 'updateTerm')     return handleUpdateTerm(data);
    if (data.action === 'addRoom')        return handleAddRoom(data);
    if (data.action === 'deleteRoom')     return handleDeleteRoom(data);
    if (data.action === 'renameRoom')     return handleRenameRoom(data);
    if (data.action === 'addCriteria')    return handleAddCriteria(data);
    if (data.action === 'updateCriteria') return handleUpdateCriteria(data);
    if (data.action === 'deleteCriteria') return handleDeleteCriteria(data);
    return respond({ status: 'error', message: 'unknown action' });
  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

// ------------------------------------------------------------
//  คืนรายชื่อห้องทั้งหมด
// ------------------------------------------------------------
function handleRooms() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('ห้องเรียน');
  if (!sheet) return respond({ status: 'error', message: 'ไม่พบ Sheet "ห้องเรียน" — รัน setupSheets() ก่อน' });

  const rows  = sheet.getDataRange().getValues();
  const rooms = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) rooms.push(String(rows[i][0]).trim());
  }
  return respond({ status: 'ok', rooms });
}

// ------------------------------------------------------------
//  คืนรายชื่อนักเรียนของห้องที่ระบุ (พร้อมเพศ) เรียงตามเลขที่
// ------------------------------------------------------------
function handleStudents(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheet = ss.getSheetByName('นักเรียน');
  if (!sheet) return respond({ status: 'error', message: 'ไม่พบ Sheet "นักเรียน" — รัน setupSheets() ก่อน' });

  const room = params.room;
  const rows = sheet.getDataRange().getValues();
  const students = [];

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === room && rows[i][2]) {
      students.push({ no: rows[i][1], name: String(rows[i][2]).trim(), gender: String(rows[i][3] || '').trim() });
    }
  }
  students.sort((a, b) => Number(a.no) - Number(b.no));

  return respond({ status: 'ok', students });
}

// ------------------------------------------------------------
//  บันทึกรายงานเช็คชื่อ — นับยอดรวมทุกหมวด (เดิมแยกชาย/หญิงตามฟอร์มกระดาษ ตัดออกแล้ว 2569)
//  ถ้าห้องนี้ส่งของ "วันนี้" มาแล้ว → แก้ไขแถวเดิมแทน ไม่สร้างซ้ำ
// ------------------------------------------------------------
// เดิมนับแยกชาย/หญิงตามฟอร์มกระดาษ ตัดออกแล้ว 2569 — pattern เดียวกับ handleSubmitEvening เป๊ะ
function handleSubmit(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const room = normalizeRoom(data.room);

  // 1) นับจำนวนเต็มของห้องจาก Sheet "นักเรียน" (ไม่สนใจเพศแล้ว)
  const stuSheet = ss.getSheetByName('นักเรียน');
  const stuRows  = stuSheet.getDataRange().getValues();
  let total = 0;
  for (let i = 1; i < stuRows.length; i++) {
    if (normalizeRoom(stuRows[i][0]) === room && stuRows[i][2]) total++;
  }

  // 2) ตัดชื่อซ้ำในลิสต์เดียวกันออก (กันหัวหน้าห้องติ๊กชื่อเดิมซ้ำโดยไม่ตั้งใจ)
  const absentNames        = uniqueList(data.absentNames);
  const personalLeaveNames = uniqueList(data.personalLeaveNames);
  const sickLeaveNames     = uniqueList(data.sickLeaveNames);
  const activityNames      = uniqueList(data.activityNames);

  // 3) มา = จำนวนเต็ม - (ขาด+ลากิจ+ลาป่วย+ไปกิจกรรม)
  const presentTotal = Math.max(0, total - absentNames.length - personalLeaveNames.length - sickLeaveNames.length - activityNames.length);

  const dateStr       = todayStr();
  const nowDate       = new Date();
  const timeStr       = Utilities.formatDate(nowDate, TZ, 'HH:mm:ss');
  const reporterName  = String(data.reporterName || '').trim();

  const rowValues = [
    dateStr, room, timeStr,
    presentTotal,
    absentNames.length, absentNames.join(', '),
    personalLeaveNames.length, personalLeaveNames.join(', '),
    sickLeaveNames.length, sickLeaveNames.join(', '),
    activityNames.length, activityNames.join(', '),
    reporterName,
  ];

  // 4) หาว่าห้องนี้ + วันนี้ เคยส่งแล้วหรือยัง
  //    หมายเหตุสำคัญ: Google Sheets มักแปลงค่าที่หน้าตาเหมือนวันที่ (เช่น "2026-08-03") ให้เป็น
  //    เซลล์ชนิด Date เองอัตโนมัติตอนเขียน แม้โค้ดจะส่งเป็น string ก็ตาม — ต้องใช้ cellDateStr()
  //    แปลงกลับเป็น string ก่อนเทียบเสมอ ห้ามเทียบ logRows[i][0] === dateStr ตรงๆ เด็ดขาด
  const logSheet = ss.getSheetByName('เช็คชื่อรายวัน');
  const logRows  = logSheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < logRows.length; i++) {
    if (cellDateStr(logRows[i][0]) === dateStr && normalizeRoom(logRows[i][1]) === room) { foundRow = i + 1; } // +1 เพราะ sheet เริ่มแถว 1; ใช้แถวล่าสุดถ้าเคยมีข้อมูลซ้ำ
  }

  if (foundRow > 0) {
    logSheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]); // แก้ไขแถวเดิมแทนที่
  } else {
    logSheet.appendRow(rowValues);
  }

  // report.html (ไม่ได้แก้ไฟล์นี้) อ่าน .total จากทุกฟิลด์ที่นี่ — คงรูปแบบ object ไว้แม้ไม่มีเพศแล้ว
  return respond({
    status: 'ok',
    present:       { total: presentTotal },
    absent:        { total: absentNames.length },
    personalLeave: { total: personalLeaveNames.length },
    sickLeave:     { total: sickLeaveNames.length },
    activity:      { total: activityNames.length },
  });
}

// ------------------------------------------------------------
//  เช็คว่า "ห้องนี้" ส่งรายงานวันนี้ไปแล้วหรือยัง
//  ใช้ตอนหัวหน้าห้องเปิดหน้ารายงาน กันกดส่งซ้ำโดยไม่รู้ตัวว่าเขียนทับของเดิม
// ------------------------------------------------------------
function handleMyStatus(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const room = normalizeRoom(params.room);

  const logSheet = ss.getSheetByName('เช็คชื่อรายวัน');
  const logRows  = logSheet.getDataRange().getValues();
  const dateStr  = todayStr();

  let latest = null;
  for (let i = 1; i < logRows.length; i++) {
    const r = logRows[i];
    if (cellDateStr(r[0]) === dateStr && normalizeRoom(r[1]) === room) latest = r;
  }
  if (latest) {
      const r = latest;
      return respond({
        status: 'ok', submitted: true, time: cellTimeStr(r[2]),
        present:       { total: r[3]  },
        absent:        { total: r[4]  },
        personalLeave: { total: r[6]  },
        sickLeave:     { total: r[8]  },
        activity:      { total: r[10] },
      });
  }
  return respond({ status: 'ok', submitted: false });
}

// ------------------------------------------------------------
//  สรุปยอดของวันที่ระบุ ทุกห้อง (ห้องไหนยังไม่ส่ง = submitted:false)
//  "จำนวนเต็ม" ของแต่ละห้องนับจาก Sheet "นักเรียน" เสมอ ไม่ว่าจะส่งรายงานวันนี้หรือยัง
// ------------------------------------------------------------
function handleByDate(dateStr) {
  dateStr = normalizeDateParam(dateStr);
  return respond(getDailySummaryData(dateStr));
}

// เนื้อหาการคำนวณจริงของ handleByDate แยกออกมาเป็นฟังก์ชันของตัวเอง
// (เดิมฟังก์ชันแจ้งเตือน LINE ก็เรียกใช้ตัวนี้ร่วมกันด้วย — ตัดออกไปแล้ว 2569)
function getDailySummaryData(dateStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const roomSheet = ss.getSheetByName('ห้องเรียน');
  const roomRows  = roomSheet.getDataRange().getValues();
  const allRooms  = [];
  for (let i = 1; i < roomRows.length; i++) {
    if (roomRows[i][0]) allRooms.push(String(roomRows[i][0]).trim());
  }

  // นับจำนวนเต็มต่อห้องจาก Sheet "นักเรียน" (เดิมแยกชาย/หญิง ตัดออกแล้ว 2569 — เหลือแค่ยอดรวม)
  const stuSheet = ss.getSheetByName('นักเรียน');
  const stuRows  = stuSheet.getDataRange().getValues();
  const rosterTotals = {}; // room -> total (number)
  for (let i = 1; i < stuRows.length; i++) {
    const room = normalizeRoom(stuRows[i][0]);
    if (!room || !stuRows[i][2]) continue;
    rosterTotals[room] = (rosterTotals[room] || 0) + 1;
  }

  const logSheet = ss.getSheetByName('เช็คชื่อรายวัน');
  const logRows  = logSheet.getDataRange().getValues();
  const submittedMap = {}; // room -> row data
  for (let i = 1; i < logRows.length; i++) {
    const r = logRows[i];
    if (cellDateStr(r[0]) === dateStr) {
      const room = normalizeRoom(r[1]);
      if (!room) continue;
      submittedMap[room] = {
        time:          cellTimeStr(r[2]),
        present:       { total: Number(r[3])  || 0 },
        absent:        { total: Number(r[4])  || 0, names: r[5]  || '' },
        personalLeave: { total: Number(r[6])  || 0, names: r[7]  || '' },
        sickLeave:     { total: Number(r[8])  || 0, names: r[9]  || '' },
        activity:      { total: Number(r[10]) || 0, names: r[11] || '' },
      };
    }
  }

  let submittedCount = 0;
  const grand = {
    roster:        { total: 0 },
    present:       { total: 0 },
    absent:        { total: 0 },
    personalLeave: { total: 0 },
    sickLeave:     { total: 0 },
    activity:      { total: 0 },
  };
  function addCat(target, src) {
    target.total += src.total;
  }

  const rooms = allRooms.map(room => {
    const roster = { total: rosterTotals[room] || 0 };
    addCat(grand.roster, roster);

    const rec = submittedMap[room];
    if (rec) {
      submittedCount++;
      addCat(grand.present, rec.present);
      addCat(grand.absent, rec.absent);
      addCat(grand.personalLeave, rec.personalLeave);
      addCat(grand.sickLeave, rec.sickLeave);
      addCat(grand.activity, rec.activity);
      // % มาเรียน = มา/ทั้งหมด — คำนวณสดตรงนี้เลย (เดิมโมดูลนี้ไม่มี % เพราะแยกชาย/หญิงแทน)
      const rate = roster.total > 0 ? Math.round((rec.present.total / roster.total) * 1000) / 10 : null;
      return {
        room, roster, submitted: true, time: rec.time, rate,
        present: rec.present, absent: rec.absent,
        personalLeave: rec.personalLeave, sickLeave: rec.sickLeave, activity: rec.activity,
      };
    }
    return { room, roster, submitted: false, rate: null };
  });

  return {
    status: 'ok',
    date: dateStr,
    rooms,
    totals: { ...grand, submittedCount, totalRooms: allRooms.length },
  };
}

function handleRangeSummary(params) {
  const startDate = normalizeDateParam(params.startDate || '');
  const endDate = normalizeDateParam(params.endDate || '');
  if (!startDate || !endDate) return respond({ status: 'error', message: 'กรุณาระบุช่วงวันที่' });
  if (startDate > endDate) return respond({ status: 'error', message: 'ช่วงวันที่ไม่ถูกต้อง' });
  return respond(getRangeSummaryData(startDate, endDate));
}

function getRangeSummaryData(startDate, endDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const roomRows = ss.getSheetByName('ห้องเรียน').getDataRange().getValues();
  const allRooms = [];
  for (let i = 1; i < roomRows.length; i++) {
    const room = normalizeRoom(roomRows[i][0]);
    if (room) allRooms.push(room);
  }

  const stuRows = ss.getSheetByName('นักเรียน').getDataRange().getValues();
  const rosterTotals = {};
  for (let i = 1; i < stuRows.length; i++) {
    const room = normalizeRoom(stuRows[i][0]);
    if (!room || !stuRows[i][2]) continue;
    rosterTotals[room] = (rosterTotals[room] || 0) + 1;
  }

  const latestByDateRoom = {};
  const logRows = ss.getSheetByName('เช็คชื่อรายวัน').getDataRange().getValues();
  for (let i = 1; i < logRows.length; i++) {
    const r = logRows[i];
    const date = cellDateStr(r[0]);
    if (date < startDate || date > endDate) continue;
    const room = normalizeRoom(r[1]);
    if (!room) continue;
    latestByDateRoom[`${date}||${room}`] = r;
  }

  const roomMap = {};
  allRooms.forEach(room => {
    roomMap[room] = {
      room,
      roster: { total: rosterTotals[room] || 0 },
      submittedDays: 0,
      present: { total: 0 },
      absent: { total: 0 },
      personalLeave: { total: 0 },
      sickLeave: { total: 0 },
      activity: { total: 0 },
    };
  });

  Object.keys(latestByDateRoom).forEach(key => {
    const r = latestByDateRoom[key];
    const room = normalizeRoom(r[1]);
    if (!roomMap[room]) {
      roomMap[room] = {
        room,
        roster: { total: rosterTotals[room] || 0 },
        submittedDays: 0,
        present: { total: 0 },
        absent: { total: 0 },
        personalLeave: { total: 0 },
        sickLeave: { total: 0 },
        activity: { total: 0 },
      };
    }
    roomMap[room].submittedDays++;
    roomMap[room].present.total += Number(r[3]) || 0;
    roomMap[room].absent.total += Number(r[4]) || 0;
    roomMap[room].personalLeave.total += Number(r[6]) || 0;
    roomMap[room].sickLeave.total += Number(r[8]) || 0;
    roomMap[room].activity.total += Number(r[10]) || 0;
  });

  const totals = {
    submittedRoomDays: 0,
    totalRooms: allRooms.length,
    present: { total: 0 },
    absent: { total: 0 },
    personalLeave: { total: 0 },
    sickLeave: { total: 0 },
    activity: { total: 0 },
  };

  const rooms = allRooms.map(room => roomMap[room]).map(r => {
    totals.submittedRoomDays += r.submittedDays;
    totals.present.total += r.present.total;
    totals.absent.total += r.absent.total;
    totals.personalLeave.total += r.personalLeave.total;
    totals.sickLeave.total += r.sickLeave.total;
    totals.activity.total += r.activity.total;
    const denominator = r.roster.total * r.submittedDays;
    r.rate = denominator > 0 ? Math.round((r.present.total / denominator) * 1000) / 10 : null;
    return r;
  });

  const totalPossible = rooms.reduce((sum, r) => sum + (r.roster.total * r.submittedDays), 0);
  totals.rate = totalPossible > 0 ? Math.round((totals.present.total / totalPossible) * 1000) / 10 : null;

  return { status: 'ok', startDate, endDate, rooms, totals };
}

// ==============================================================
//  งานกิจการนักเรียน — โมดูล A: เช็คกลับ 15:30 (รื้อดีไซน์ 2026-08-03 — เช็คอิสระเต็มรูปแบบ)
//  หัวหน้าห้องเช็คชื่อรอบเย็นอิสระจากเช้าโดยสิ้นเชิง (roster เต็มห้อง 5 หมวดเหมือนเช้าทุกประการ:
//  มา/ขาด/ลากิจ/ลาป่วย/ไปกิจกรรม) แล้วระบบเปรียบเทียบผลเช้า-เย็นต่อคนภายหลังแทน — ไม่ใช่การกรอง
//  roster จากเช้ามาให้เย็นเหมือนดีไซน์เดิม (ถูกแทนที่แล้ว ดู git history ถ้าต้องการเทียบ)
// ==============================================================

// ------------------------------------------------------------
//  บันทึกรายงานเช็คกลับ — pattern เดียวกับ handleSubmit (เช้า) เป๊ะ แค่ไม่แยกชาย/หญิง
//  ถ้าห้องนี้ส่งของ "วันนี้" มาแล้ว → แก้ไขแถวเดิมแทน ไม่สร้างซ้ำ (เหมือนเช้า)
// ------------------------------------------------------------
function handleSubmitEvening(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const room = normalizeRoom(data.room);

  const stuSheet = ss.getSheetByName('นักเรียน');
  const stuRows  = stuSheet.getDataRange().getValues();
  let total = 0;
  for (let i = 1; i < stuRows.length; i++) {
    if (normalizeRoom(stuRows[i][0]) === room && stuRows[i][2]) total++;
  }

  const absentNames        = uniqueList(data.absentNames);
  const personalLeaveNames = uniqueList(data.personalLeaveNames);
  const sickLeaveNames     = uniqueList(data.sickLeaveNames);
  const activityNames      = uniqueList(data.activityNames);

  const presentTotal = Math.max(0, total - absentNames.length - personalLeaveNames.length - sickLeaveNames.length - activityNames.length);

  const dateStr      = todayStr();
  const nowDate      = new Date();
  const timeStr      = Utilities.formatDate(nowDate, TZ, 'HH:mm:ss');
  const reporterName = String(data.reporterName || '').trim();

  const rowValues = [
    dateStr, room, timeStr,
    presentTotal,
    absentNames.length, absentNames.join(', '),
    personalLeaveNames.length, personalLeaveNames.join(', '),
    sickLeaveNames.length, sickLeaveNames.join(', '),
    activityNames.length, activityNames.join(', '),
    reporterName,
  ];

  const logSheet = ss.getSheetByName('เช็คชื่อเย็น');
  const logRows  = logSheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < logRows.length; i++) {
    if (cellDateStr(logRows[i][0]) === dateStr && normalizeRoom(logRows[i][1]) === room) { foundRow = i + 1; }
  }

  if (foundRow > 0) {
    logSheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    logSheet.appendRow(rowValues);
  }

  return respond({
    status: 'ok',
    present:       presentTotal,
    absent:        absentNames.length,
    personalLeave: personalLeaveNames.length,
    sickLeave:     sickLeaveNames.length,
    activity:      activityNames.length,
  });
}

// ------------------------------------------------------------
//  สรุปเช็คกลับของทุกห้อง ณ วันที่ระบุ — ใช้ในแดชบอร์ด (ไม่มีการแจ้งเตือน LINE ใดๆ
//  สำหรับโมดูลนี้ตามที่ครูยืนยัน — ครูต้องเปิดแดชบอร์ดดูเอง)
// ------------------------------------------------------------
function handleEveningByDate(dateStr) {
  dateStr = normalizeDateParam(dateStr);
  return respond(getEveningSummaryData(dateStr));
}

function getEveningSummaryData(dateStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const roomSheet = ss.getSheetByName('ห้องเรียน');
  const roomRows  = roomSheet.getDataRange().getValues();
  const allRooms  = [];
  for (let i = 1; i < roomRows.length; i++) {
    if (roomRows[i][0]) allRooms.push(String(roomRows[i][0]).trim());
  }

  const rosterTotals = {};
  const stuRows = ss.getSheetByName('นักเรียน').getDataRange().getValues();
  for (let i = 1; i < stuRows.length; i++) {
    const room = normalizeRoom(stuRows[i][0]);
    if (!room || !stuRows[i][2]) continue;
    rosterTotals[room] = (rosterTotals[room] || 0) + 1;
  }

  const logRows = ss.getSheetByName('เช็คชื่อเย็น').getDataRange().getValues();
  const submittedMap = {};
  for (let i = 1; i < logRows.length; i++) {
    const r = logRows[i];
    if (cellDateStr(r[0]) === dateStr) {
      const room = normalizeRoom(r[1]);
      if (!room) continue;
      submittedMap[room] = {
        time:          cellTimeStr(r[2]),
        present:       Number(r[3]) || 0,
        absent:        { total: Number(r[4])  || 0, names: r[5]  || '' },
        personalLeave: { total: Number(r[6])  || 0, names: r[7]  || '' },
        sickLeave:     { total: Number(r[8])  || 0, names: r[9]  || '' },
        activity:      { total: Number(r[10]) || 0, names: r[11] || '' },
      };
    }
  }

  const rooms = allRooms.map(room => {
    const roster = rosterTotals[room] || 0;
    const rec = submittedMap[room];
    if (rec) return Object.assign({ room, roster, submitted: true }, rec);
    return { room, roster, submitted: false };
  });

  return { status: 'ok', date: dateStr, rooms };
}

// ------------------------------------------------------------
//  โหลด 3 sheets ที่ใช้ร่วมกันในการเปรียบเทียบเช้า-เย็นครั้งเดียว — ใช้ตอนต้องวนหลายห้อง
//  (เช่น attendanceRate) กันอ่านชีทซ้ำห้องละ 3 ครั้ง
// ------------------------------------------------------------
function loadComparisonContext() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    stuRows:    ss.getSheetByName('นักเรียน').getDataRange().getValues(),
    morningLog: ss.getSheetByName('เช็คชื่อรายวัน').getDataRange().getValues(),
    eveningLog: ss.getSheetByName('เช็คชื่อเย็น').getDataRange().getValues(),
  };
}

// จัดกลุ่มสถานะ 5 หมวด (มา/ขาด/ลากิจ/ลาป่วย/ไปกิจกรรม) ให้เหลือ 3 กลุ่มสำหรับ derivation:
// PRESENT = มา, ไปกิจกรรม (ถือว่าอยู่ในความดูแลของโรงเรียนทั้งคู่)
// ABSENT  = ขาด
// LEAVE   = ลากิจ, ลาป่วย (ชนะทุกกรณี — ถ้าลาก็คือลา ไม่ว่าอีกฝั่งจะเป็นอะไร)
function statusBucket(status) {
  if (status === 'ขาด') return 'ABSENT';
  if (status === 'ลากิจ' || status === 'ลาป่วย') return 'LEAVE';
  return 'PRESENT'; // มา, ไปกิจกรรม
}

// เทียบสถานะเช้า-เย็นต่อคน คืนผลลัพธ์ ปกติ/สาย/หนี/ขาด/ลา ตามกฎ (เช็คจากบนลงล่าง ข้อไหนตรงก่อนใช้ข้อนั้น)
// 1) ฝั่งใดฝั่งหนึ่งเป็น LEAVE → "ลา"   2) PRESENT ทั้งคู่ → "ปกติ"
// 3) เช้า ABSENT + เย็น PRESENT → "สาย" (มาสาย)   4) เช้า PRESENT + เย็น ABSENT → "หนี" (หนีกลางวัน)
// 5) ABSENT ทั้งคู่ → "ขาด"
// ครูยืนยันแล้ว 2026-08-03 ว่าเช้า "ไปกิจกรรม" + เย็น "ขาด" ก็นับเป็น "หนี" เหมือนเช้ามา+เย็นขาด
function deriveDailyResult(morningStatus, eveningStatus) {
  const m = statusBucket(morningStatus);
  const e = statusBucket(eveningStatus);
  if (m === 'LEAVE' || e === 'LEAVE') return 'ลา';
  if (m === 'PRESENT' && e === 'PRESENT') return 'ปกติ';
  if (m === 'ABSENT'  && e === 'PRESENT') return 'สาย';
  if (m === 'PRESENT' && e === 'ABSENT')  return 'หนี';
  return 'ขาด'; // ABSENT + ABSENT
}

// แปลงแถวเช้า/เย็น 1 แถว เป็น map ชื่อ -> หมวด (มา/ขาด/ลากิจ/ลาป่วย/ไปกิจกรรม) ครบทุกคนใน roster
// cols ระบุตำแหน่งคอลัมน์ "รายชื่อ" ของ 4 หมวดที่ไม่ใช่ "มา" (ต่างกันระหว่างเช้า/เย็นเพราะเช้ามีเพศ)
function statusMapFromRow(roster, row, cols) {
  const map = {};
  roster.forEach(name => { map[name] = 'มา'; }); // default ทุกคน = มา ถ้าไม่อยู่ใน 4 หมวดด้านล่าง
  if (!row) return null; // แถวนี้ยังไม่มี = ยังไม่ส่งรายงานของช่วงเวลานี้เลย
  const setStatus = (col, label) => {
    String(row[col] || '').split(',').map(s => s.trim()).filter(Boolean).forEach(n => { if (n in map) map[n] = label; });
  };
  setStatus(cols.absent,   'ขาด');
  setStatus(cols.personal, 'ลากิจ');
  setStatus(cols.sick,     'ลาป่วย');
  setStatus(cols.activity, 'ไปกิจกรรม');
  return map;
}

// คอลัมน์ "รายชื่อ" 4 หมวดของแต่ละชีท (ดู header ใน setupSheets())
// เดิมเช้ามีคอลัมน์เพศ (กว้างกว่า) ตัดออกแล้ว 2569 — ตอนนี้ schema เช้า/เย็นเหมือนกันเป๊ะ ใช้ค่าเดียวกัน
const EVENING_NAME_COLS = { absent: 5,  personal: 7,  sick: 9,  activity: 11 }; // เช็คชื่อเย็น (ไม่มีเพศ)
const MORNING_NAME_COLS = EVENING_NAME_COLS; // เช็คชื่อรายวัน (ไม่มีเพศแล้ว — โครงสร้างเหมือนเย็นทุกประการ)

// เปรียบเทียบเช้า-เย็นของห้อง/วันที่ระบุ 1 ห้อง — คืนสถานะรายคน + สรุปนับแต่ละผลลัพธ์
// รับ ctx (จาก loadComparisonContext) เป็น optional เพื่อไม่ต้องอ่านชีทซ้ำเวลาวนหลายห้อง
function getDailyComparisonData(dateStr, room, ctx) {
  ctx = ctx || loadComparisonContext();
  room = normalizeRoom(room);

  const roster = [];
  for (let i = 1; i < ctx.stuRows.length; i++) {
    if (normalizeRoom(ctx.stuRows[i][0]) === room && ctx.stuRows[i][2]) roster.push(String(ctx.stuRows[i][2]).trim());
  }

  let morningRow = null;
  for (let i = 1; i < ctx.morningLog.length; i++) {
    if (cellDateStr(ctx.morningLog[i][0]) === dateStr && normalizeRoom(ctx.morningLog[i][1]) === room) { morningRow = ctx.morningLog[i]; }
  }
  let eveningRow = null;
  for (let i = 1; i < ctx.eveningLog.length; i++) {
    if (cellDateStr(ctx.eveningLog[i][0]) === dateStr && normalizeRoom(ctx.eveningLog[i][1]) === room) { eveningRow = ctx.eveningLog[i]; }
  }

  const morningDone = !!morningRow;
  const eveningDone = !!eveningRow;
  const morningMap  = statusMapFromRow(roster, morningRow, MORNING_NAME_COLS);
  const eveningMap  = statusMapFromRow(roster, eveningRow, EVENING_NAME_COLS);

  const students = roster.map(name => {
    if (!morningDone || !eveningDone) {
      return {
        name,
        morningStatus: morningDone ? morningMap[name] : null,
        eveningStatus: eveningDone ? eveningMap[name] : null,
        result: 'ไม่ทราบ', // รอข้อมูลอีกฝั่ง ไม่เดาผล
      };
    }
    const morningStatus = morningMap[name];
    const eveningStatus = eveningMap[name];
    return { name, morningStatus, eveningStatus, result: deriveDailyResult(morningStatus, eveningStatus) };
  });

  const summary = { ปกติ: 0, สาย: 0, หนี: 0, ขาด: 0, ลา: 0, ไม่ทราบ: 0 };
  students.forEach(s => { summary[s.result]++; });

  return { room, morningDone, eveningDone, total: roster.length, students, summary };
}

// doGet action=dailyComparison — ระบุ room คืนห้องเดียว, ไม่ระบุ room คืนทุกห้อง (ใช้โดยแดชบอร์ด)
function handleDailyComparison(params) {
  const dateStr = normalizeDateParam(params.date || todayStr());

  if (params.room) {
    return respond(Object.assign({ status: 'ok', date: dateStr }, getDailyComparisonData(dateStr, params.room)));
  }

  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const roomRows = ss.getSheetByName('ห้องเรียน').getDataRange().getValues();
  const allRooms = [];
  for (let i = 1; i < roomRows.length; i++) if (roomRows[i][0]) allRooms.push(String(roomRows[i][0]).trim());

  const ctx   = loadComparisonContext();
  const rooms = allRooms.map(room => getDailyComparisonData(dateStr, room, ctx));
  return respond({ status: 'ok', date: dateStr, rooms });
}

// ------------------------------------------------------------
//  % การมาเรียนรวม ต่อห้อง + รวมทั้งโรงเรียน ของวันที่ระบุ
//  ⚠️ สมมติฐานที่ Cody ตั้งเอง (ยังไม่ได้ยืนยัน 100% กับครู): นับ "สาย" รวมเป็นมาเรียนด้วย
//  (attended = ปกติ + สาย) ถ้าครูต้องการไม่นับสายรวม ให้แก้บรรทัด `const attended = ...` ด้านล่าง
//  เป็น `const attended = s.summary.ปกติ;` แทน — ห้องที่เช้า/เย็นยังไม่ครบ (ไม่ทราบ) ไม่ถูกนับ
//  ในตัวหารรวมทั้งโรงเรียน กันเปอร์เซ็นต์เพี้ยนจากข้อมูลไม่ครบ
// ------------------------------------------------------------
function getAttendanceRateData(dateStr) {
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const roomRows = ss.getSheetByName('ห้องเรียน').getDataRange().getValues();
  const allRooms = [];
  for (let i = 1; i < roomRows.length; i++) if (roomRows[i][0]) allRooms.push(String(roomRows[i][0]).trim());

  const ctx = loadComparisonContext();
  let schoolAttended = 0, schoolTotal = 0;

  const rooms = allRooms.map(room => {
    const data     = getDailyComparisonData(dateStr, room, ctx);
    const attended = data.summary.ปกติ + data.summary.สาย;
    const complete = data.morningDone && data.eveningDone;
    // ต้อง complete ด้วยเสมอ ไม่ใช่แค่ total > 0 — ห้องที่เช้า/เย็นยังไม่ครบ ทุกคนจะถูกนับเป็น
    // "ไม่ทราบ" (ไม่ใช่ปกติ/สาย) ทำให้ attended=0 เสมอ ถ้าไม่กันจุดนี้จะโชว์ 0.0% หลอกว่าห้องนั้น
    // ไม่มีใครมาเลย ทั้งที่จริงแค่ยังไม่ถึงเวลาส่งเช็คกลับ (พบระหว่าง QC 2026-08-03)
    const rate     = complete && data.total > 0 ? Math.round((attended / data.total) * 1000) / 10 : null;
    if (complete) { schoolAttended += attended; schoolTotal += data.total; }
    return { room, rate, attended, total: data.total, complete };
  });

  const schoolRate = schoolTotal > 0 ? Math.round((schoolAttended / schoolTotal) * 1000) / 10 : null;
  return { rooms, schoolRate, schoolAttended, schoolTotal };
}

function handleAttendanceRate(params) {
  const dateStr = normalizeDateParam(params.date || todayStr());
  return respond(Object.assign({ status: 'ok', date: dateStr }, getAttendanceRateData(dateStr)));
}

// ==============================================================
//  งานกิจการนักเรียน — โมดูล B: คะแนนความประพฤติ
//  ⚠️ ครูทุกคนกรอกได้ (ไม่ใช่แค่หัวหน้าห้อง) กันการเข้าถึงด้วย PIN ฝั่งหน้าเว็บเท่านั้น
//  (ดูคอมเมนต์ที่หัว behavior.html) — endpoint พวกนี้ "ไม่มี" การตรวจ PIN ฝั่ง backend เลย
//  เพราะ Apps Script deploy แบบ "Anyone" ไม่มี concept ของ session/login ให้ตรวจ
// ==============================================================

// ------------------------------------------------------------
//  คืนค่าภาคเรียนปัจจุบันจาก Sheet "ตั้งค่ากิจการนักเรียน" (ให้หน้าเว็บโชว์/แนบไปกับข้อมูล)
// ------------------------------------------------------------
function handleCurrentTerm() {
  return respond({ status: 'ok', term: getCurrentTermValue() });
}

function getCurrentTermValue() {
  return getSettingValue('ภาคเรียนปัจจุบัน');
}

// ==============================================================
//  งานกิจการนักเรียน — โมดูล C: ตั้งค่าเบื้องหลัง (เพิ่ม 2569)
//  เป้าหมาย: ให้ระบบดูแลต่อได้โดยไม่ต้องพึ่งครูภูริณัฐคนเดียว — ตั้งค่าทุกอย่างแก้ผ่านหน้าเว็บได้
//  ไม่ต้องแก้โค้ด/เปิด Google Sheet ตรงๆ
//  ⚠️ PIN 8 หลักของหน้านี้เป็นแค่ตัวกันคนไม่ตั้งใจเปิดผิดเหมือนโมดูลอื่น ไม่ใช่การป้องกันจริงจัง
//  (endpoint พวกนี้เรียกตรงได้เสมอโดยไม่ผ่าน PIN เลยถ้ารู้ SCRIPT_URL — ดูคอมเมนต์ที่หัว behavior.html)
// ==============================================================

// อ่าน/เขียนค่าคู่ รายการ-ค่า ใน Sheet "ตั้งค่ากิจการนักเรียน" — ใช้ร่วมกันทั้งภาคเรียนและ PIN ทุกระดับ
function getSettingValue(key) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('ตั้งค่ากิจการนักเรียน');
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) return String(rows[i][1]).trim();
  }
  return '';
}

function setSettingValue(key, value) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('ตั้งค่ากิจการนักเรียน');
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]); // แถวนี้ยังไม่เคยมี (เช่น ระบบเก่าก่อนอัปเดต) สร้างใหม่ให้เอง
}

// PIN 3 ระดับ ผูกกับ key ใน Sheet "ตั้งค่ากิจการนักเรียน"
const PIN_KEYS = {
  checkin:  'PIN เช็คชื่อ',              // เช็คชื่อเช้า+เย็น ใช้ PIN เดียวกัน (4 หลัก)
  behavior: 'PIN คะแนนความประพฤติ',       // 6 หลัก
  settings: 'PIN ตั้งค่าเบื้องหลัง',       // 8 หลัก
};

// ตรวจ PIN ฝั่ง server — ไม่ส่งค่า PIN จริงกลับไปให้ client เห็นเด็ดขาด (คืนแค่ true/false)
function handleVerifyPin(data) {
  const key = PIN_KEYS[data.tier];
  if (!key) return respond({ status: 'error', message: 'ไม่รู้จัก tier ของ PIN' });
  const valid = String(data.pin || '').trim() === getSettingValue(key);
  return respond({ status: 'ok', valid });
}

function handleUpdatePin(data) {
  const key = PIN_KEYS[data.tier];
  if (!key) return respond({ status: 'error', message: 'ไม่รู้จัก tier ของ PIN' });
  const newPin = String(data.newPin || '').trim();
  if (!newPin) return respond({ status: 'error', message: 'กรุณาระบุ PIN ใหม่' });
  setSettingValue(key, newPin);
  return respond({ status: 'ok' });
}

function handleUpdateTerm(data) {
  const term = String(data.term || '').trim();
  if (!term) return respond({ status: 'error', message: 'กรุณาระบุภาคเรียน' });
  setSettingValue('ภาคเรียนปัจจุบัน', term);
  return respond({ status: 'ok' });
}

// คืนค่าตั้งค่าทั้งหมดที่ไม่อ่อนไหว (ภาคเรียน, รายชื่อห้อง, เกณฑ์คะแนน) — ไม่คืนค่า PIN จริงเด็ดขาด
function handleSettingsData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const roomRows = ss.getSheetByName('ห้องเรียน').getDataRange().getValues();
  const rooms = [];
  for (let i = 1; i < roomRows.length; i++) if (roomRows[i][0]) rooms.push(String(roomRows[i][0]).trim());

  const critRows = ss.getSheetByName('เกณฑ์คะแนนความประพฤติ').getDataRange().getValues();
  const criteria = [];
  for (let i = 1; i < critRows.length; i++) {
    if (critRows[i][0]) {
      criteria.push({
        reason: String(critRows[i][0]).trim(),
        type: String(critRows[i][1]).trim() === 'เพิ่ม' ? 'เพิ่ม' : 'หัก',
        points: Number(critRows[i][2]) || 0,
      });
    }
  }

  return respond({ status: 'ok', term: getCurrentTermValue(), rooms, criteria });
}

// ------------------------------------------------------------
//  จัดการรายชื่อห้องเรียน — เพิ่ม/ลบ/แก้ชื่อ
//  ⚠️ ชื่อห้องถูกอ้างอิงเป็น string key ในหลาย Sheet (นักเรียน, เช็คชื่อรายวัน, เช็คชื่อเย็น,
//  คะแนนความประพฤติ) — handleRenameRoom ต้องอัปเดตทุก Sheet พร้อมกัน ไม่งั้นข้อมูลเก่าของห้องนั้น
//  จะหลุดหาไม่เจอทันที (จุดเสี่ยงที่สุดของโมดูลนี้)
// ------------------------------------------------------------
function handleAddRoom(data) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('ห้องเรียน');
  const room  = String(data.room || '').trim();
  if (!room) return respond({ status: 'error', message: 'กรุณาระบุชื่อห้อง' });

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === room) return respond({ status: 'error', message: 'มีห้องนี้อยู่แล้ว' });
  }
  sheet.appendRow([room]);
  return respond({ status: 'ok' });
}

// ลบแค่จาก Sheet "ห้องเรียน" เท่านั้น — ไม่ลบข้อมูลนักเรียน/ประวัติเช็คชื่อ/คะแนนเก่าของห้องนั้น
// (ป้องกันข้อมูลหายโดยไม่ตั้งใจ — ถ้าจะลบข้อมูลจริงต้องไปลบเองใน Sheet ตรงๆ)
function handleDeleteRoom(data) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('ห้องเรียน');
  const room  = String(data.room || '').trim();

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === room) { sheet.deleteRow(i + 1); return respond({ status: 'ok' }); }
  }
  return respond({ status: 'error', message: 'ไม่พบห้องนี้' });
}

function handleRenameRoom(data) {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const oldName = String(data.oldRoom || '').trim();
  const newName = String(data.newRoom || '').trim();
  if (!oldName || !newName) return respond({ status: 'error', message: 'กรุณาระบุชื่อห้องเดิมและชื่อใหม่' });
  if (oldName === newName) return respond({ status: 'ok', changedCount: 0 });

  const roomSheet = ss.getSheetByName('ห้องเรียน');
  const roomRows  = roomSheet.getDataRange().getValues();
  for (let i = 1; i < roomRows.length; i++) {
    if (String(roomRows[i][0]).trim() === newName) return respond({ status: 'error', message: 'มีชื่อห้องนี้อยู่แล้ว เลือกชื่ออื่น' });
  }

  // คอลัมน์ "ห้อง" อยู่ index ต่างกันในแต่ละ Sheet (ดู header ใน setupSheets())
  const targets = [
    { name: 'ห้องเรียน', col: 0 },
    { name: 'นักเรียน', col: 0 },
    { name: 'เช็คชื่อรายวัน', col: 1 },
    { name: 'เช็คชื่อเย็น', col: 1 },
    { name: 'คะแนนความประพฤติ', col: 3 },
  ];
  let changedCount = 0;
  targets.forEach(({ name, col }) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][col]).trim() === oldName) {
        sheet.getRange(i + 1, col + 1).setValue(newName);
        changedCount++;
      }
    }
  });

  return respond({ status: 'ok', changedCount });
}

// ------------------------------------------------------------
//  จัดการเกณฑ์หัก/เพิ่มคะแนนความประพฤติ — ย้ายมาจากที่เคย hardcode ใน behavior.html
// ------------------------------------------------------------
function handleAddCriteria(data) {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('เกณฑ์คะแนนความประพฤติ');
  const reason = String(data.reason || '').trim();
  const type   = data.type === 'เพิ่ม' ? 'เพิ่ม' : 'หัก';
  const points = Math.abs(Number(data.points) || 0);
  if (!reason || !points) return respond({ status: 'error', message: 'กรุณาระบุเหตุผลและคะแนน' });

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === reason) return respond({ status: 'error', message: 'มีเหตุผลนี้อยู่แล้ว' });
  }
  sheet.appendRow([reason, type, points]);
  return respond({ status: 'ok' });
}

function handleUpdateCriteria(data) {
  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet     = ss.getSheetByName('เกณฑ์คะแนนความประพฤติ');
  const oldReason = String(data.oldReason || '').trim();
  const rows      = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === oldReason) {
      sheet.getRange(i + 1, 1, 1, 3).setValues([[
        String(data.reason || '').trim(),
        data.type === 'เพิ่ม' ? 'เพิ่ม' : 'หัก',
        Math.abs(Number(data.points) || 0),
      ]]);
      return respond({ status: 'ok' });
    }
  }
  return respond({ status: 'error', message: 'ไม่พบเหตุผลนี้' });
}

function handleDeleteCriteria(data) {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('เกณฑ์คะแนนความประพฤติ');
  const reason = String(data.reason || '').trim();
  const rows   = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === reason) { sheet.deleteRow(i + 1); return respond({ status: 'ok' }); }
  }
  return respond({ status: 'error', message: 'ไม่พบเหตุผลนี้' });
}

// ------------------------------------------------------------
//  บันทึกเหตุการณ์หัก/เพิ่มคะแนน — append แถวใหม่เสมอ (ไม่ upsert เหมือนเช็คชื่อ)
//  เพราะแต่ละเหตุการณ์เป็นคนละเรื่องกัน ต้องเก็บเป็นประวัติทุกรายการ ไม่ทับกัน
// ------------------------------------------------------------
function handleSubmitBehavior(data) {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const term = getCurrentTermValue();

  const nowDate = new Date();
  const dateStr = todayStr();
  const timeStr = Utilities.formatDate(nowDate, TZ, 'HH:mm:ss');

  const type   = data.type === 'เพิ่ม' ? 'เพิ่ม' : 'หัก'; // กันค่าแปลกปลอมหลุดเข้ามาจากภายนอก
  const points = Math.abs(Number(data.points) || 0);

  const rowValues = [
    dateStr, timeStr, term, data.room, String(data.studentName || '').trim(),
    type, String(data.reason || '').trim(), points, String(data.reporterName || '').trim(),
  ];

  ss.getSheetByName('คะแนนความประพฤติ').appendRow(rowValues);
  return respond({ status: 'ok' });
}

// ------------------------------------------------------------
//  สรุปคะแนนคงเหลือของทุกคนในห้อง (100 + เพิ่ม - หัก เฉพาะภาคเรียนปัจจุบัน) + ประวัติรายคน
//  คำนวณสดจาก ledger ทุกครั้งที่เรียก ไม่มีการเก็บยอดสะสมแยกไว้ที่ไหน (กันข้อมูลไม่ตรงกัน)
// ------------------------------------------------------------
function handleBehaviorSummary(params) {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const room = params.room;
  const term = getCurrentTermValue();

  const stuSheet = ss.getSheetByName('นักเรียน');
  const stuRows  = stuSheet.getDataRange().getValues();
  const roster = [];
  for (let i = 1; i < stuRows.length; i++) {
    if (stuRows[i][0] === room && stuRows[i][2]) roster.push(String(stuRows[i][2]).trim());
  }

  const rows = ss.getSheetByName('คะแนนความประพฤติ').getDataRange().getValues();
  const scores = {}, history = {};
  roster.forEach(name => { scores[name] = 100; history[name] = []; });

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[2]).trim() !== term) continue; // เฉพาะภาคเรียนปัจจุบันเท่านั้น
    if (r[3] !== room) continue;
    const name = String(r[4]).trim();
    if (!(name in scores)) continue; // นักเรียนที่ไม่อยู่ใน roster ปัจจุบันของห้องนี้แล้ว (เช่น ย้ายห้อง) ข้าม
    const pts   = Number(r[7]) || 0;
    const delta = r[5] === 'เพิ่ม' ? pts : -pts;
    scores[name] += delta;
    history[name].push({ date: cellDateStr(r[0]), time: cellTimeStr(r[1]), type: r[5], reason: r[6], points: pts, reporter: r[8] });
  }

  const students = roster.map(name => ({ name, score: scores[name], history: history[name] }));
  return respond({ status: 'ok', term, students });
}

// ------------------------------------------------------------
//  แปลงวันเป็นข้อความไทยเต็มรูปแบบ "วันจันทร์ที่ 3 สิงหาคม 2569" (ปี พ.ศ.)
//  (เดิมใช้เฉพาะในฟังก์ชันแจ้งเตือน LINE ที่ตัดออกไปแล้ว 2569 — เก็บ helper นี้ไว้เผื่อมีจุดอื่น
//  ต้องการฟอร์แมตวันที่ไทยฝั่ง Apps Script ในอนาคต ไม่ได้ถูกเรียกใช้จริงที่ไหนแล้วตอนนี้)
//  ไม่ใช้พารามิเตอร์ locale ของ Utilities.formatDate (ของจริงมีแค่ 3 พารามิเตอร์ ไม่มี locale)
//  ใช้ตาราง lookup แปลงจากชื่อวัน/เดือนภาษาอังกฤษที่ได้มาแทน กันปัญหาไม่ได้ชื่อไทยแบบเงียบๆ
// ------------------------------------------------------------
const THAI_DAY_NAMES = { Sunday: 'อาทิตย์', Monday: 'จันทร์', Tuesday: 'อังคาร', Wednesday: 'พุธ', Thursday: 'พฤหัสบดี', Friday: 'ศุกร์', Saturday: 'เสาร์' };
const THAI_MONTH_NAMES = {
  January: 'มกราคม', February: 'กุมภาพันธ์', March: 'มีนาคม', April: 'เมษายน',
  May: 'พฤษภาคม', June: 'มิถุนายน', July: 'กรกฎาคม', August: 'สิงหาคม',
  September: 'กันยายน', October: 'ตุลาคม', November: 'พฤศจิกายน', December: 'ธันวาคม',
};

function thaiFullDateStr(date) {
  const dayName   = THAI_DAY_NAMES[Utilities.formatDate(date, TZ, 'EEEE')] || '';
  const day       = Utilities.formatDate(date, TZ, 'd');
  const monthName = THAI_MONTH_NAMES[Utilities.formatDate(date, TZ, 'MMMM')] || '';
  const yearBE    = Number(Utilities.formatDate(date, TZ, 'yyyy')) + 543;
  return `วัน${dayName}ที่ ${day} ${monthName} ${yearBE}`;
}

// แปลงเวลาเป็นข้อความสไตล์ไทย "08.21 น." (จุดคั่นชั่วโมง-นาที ไม่ใช่ทวิภาค)
function thaiTimeStr(date) {
  return Utilities.formatDate(date, TZ, 'HH.mm') + ' น.';
}

// ------------------------------------------------------------
//  Helper
// ------------------------------------------------------------
function uniqueList(arr) {
  return Array.from(new Set((arr || []).map(String).map(s => s.trim()).filter(Boolean)));
}

function todayStr() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function normalizeRoom(room) {
  return String(room || '').trim();
}

function normalizeDateParam(value) {
  const raw = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const m = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const day = m[1].padStart(2, '0');
    const month = m[2].padStart(2, '0');
    const yearNum = Number(m[3]);
    const year = String(yearNum > 2400 ? yearNum - 543 : yearNum);
    return `${year}-${month}-${day}`;
  }
  return raw;
}

// แปลงค่าเซลล์ (อาจเป็น Date object หรือ string) ให้เป็น string รูปแบบ yyyy-MM-dd เสมอ
// จำเป็นเพราะ Google Sheets มักแปลงข้อความที่หน้าตาเหมือนวันที่ให้เป็นเซลล์ชนิด Date เองอัตโนมัติ
function cellDateStr(cell) {
  if (cell instanceof Date) {
    const shifted = new Date(cell.getTime() + 12 * 60 * 60 * 1000);
    return Utilities.formatDate(shifted, TZ, 'yyyy-MM-dd');
  }
  return normalizeDateParam(String(cell).trim());
}

// เหตุผลเดียวกับ cellDateStr — Google Sheets แปลงข้อความที่หน้าตาเหมือนเวลา (เช่น "14:03:55")
// ให้เป็นเซลล์ชนิดเวลาเองอัตโนมัติ อ่านกลับมาตรงๆ จะได้ Date แปลกๆ (ปี 1899) ต้องแปลงกลับก่อนใช้
function cellTimeStr(cell) {
  return cell instanceof Date ? Utilities.formatDate(cell, TZ, 'HH:mm:ss') : String(cell).trim();
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function debugDailyRows() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('Spreadsheet ID: ' + SPREADSHEET_ID);
  Logger.log('ชื่อไฟล์ชีต: ' + ss.getName());
  Logger.log('ชื่อชีตทั้งหมด: ' + ss.getSheets().map(s => s.getName()).join(', '));

  const roomSheet = ss.getSheetByName('ห้องเรียน');
  Logger.log('ห้องเรียน จำนวนแถว: ' + (roomSheet ? roomSheet.getLastRow() : 'ไม่พบชีต'));
  if (roomSheet) Logger.log('ห้องเรียน: ' + JSON.stringify(roomSheet.getDataRange().getValues()));

  const dailySheet = ss.getSheetByName('เช็คชื่อรายวัน');
  Logger.log('เช็คชื่อรายวัน จำนวนแถว: ' + (dailySheet ? dailySheet.getLastRow() : 'ไม่พบชีต'));
  if (dailySheet) Logger.log('เช็คชื่อรายวัน: ' + JSON.stringify(dailySheet.getDataRange().getValues()));

  const settingsSheet = ss.getSheetByName('ตั้งค่ากิจการนักเรียน');
  Logger.log('ตั้งค่ากิจการนักเรียน จำนวนแถว: ' + (settingsSheet ? settingsSheet.getLastRow() : 'ไม่พบชีต'));
  if (settingsSheet) Logger.log('ตั้งค่ากิจการนักเรียน: ' + JSON.stringify(settingsSheet.getDataRange().getValues()));
}
