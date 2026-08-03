// ============================================================
//  ระบบเช็คชื่อรายวันตอนเช้า — ทั้งโรงเรียนเมืองแกพิทยาสรรค์
//  (ระบบแยกอิสระ ไม่เกี่ยวกับเว็บห้องเรียนส่วนตัวของครูม่อน)
//
//  ⚠️ ระบบนี้เปิดสาธารณะ ไม่มีรหัสผ่าน/PIN ป้องกันใดๆ ทั้งสิ้น
//     เป็นการตัดสินใจของครูที่รับทราบความเสี่ยงเรื่องข้อมูลนักเรียน
//     (ชื่อนักเรียนที่ขาด/ลากิจ/ลาป่วย/ไปกิจกรรม) แล้วว่าใครก็ตามที่มีลิงก์ระบบ
//     จะเข้าถึงข้อมูลนี้ได้ทันที — ไม่ใช่ความผิดพลาดที่ลืมใส่การป้องกัน
//
//  โครงสร้างตรงกับฟอร์มกระดาษ "แบบรายงานสถิติการมาเรียนประจำวัน" ของครู —
//  แยกนับ ชาย/หญิง/รวม ทุกหมวด (มา/ขาด/ลากิจ/ลาป่วย/ไปกิจกรรม)
//
//  ขั้นตอนติดตั้ง (ทำครั้งเดียว):
//  1. สร้าง Google Sheet ใหม่เปล่าๆ 1 ไฟล์ (ชื่ออะไรก็ได้ เช่น "เช็คชื่อรายวัน MKP")
//  2. คัดลอก ID ของ Sheet จาก URL (ส่วนตัวยาวๆ ระหว่าง /d/ กับ /edit)
//     แล้วนำมาแทนที่ข้อความ 'ใส่ ID ของ Google Sheet ที่สร้างใหม่ตรงนี้' ด้านล่าง
//  3. เปิดเมนู Extensions > Apps Script ของ Sheet นั้น แล้ววางโค้ดทั้งไฟล์นี้ทับ
//  4. รันฟังก์ชัน setupSheets() หนึ่งครั้ง (เลือกฟังก์ชันนี้ที่ dropdown ด้านบน แล้วกด Run)
//     — จะสร้าง 4 sheets ให้อัตโนมัติ: "ห้องเรียน", "นักเรียน", "เช็คชื่อรายวัน", "วันหยุดเพิ่มเติม"
//     — Sheet "ห้องเรียน" จะใส่ห้องเริ่มต้นให้ 7 ห้องอัตโนมัติ (ม.1/1, ม.2/1, ม.3/1, ม.3/2, ม.4/1, ม.5/1, ม.6/1)
//       แก้ไข/เพิ่ม/ลบเองทีหลังได้ตามจริง
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
//  7. นำ URL ไปแทนที่ SCRIPT_URL ในไฟล์ report.html และ dashboard.html ทั้งสองไฟล์
//
//  ขั้นตอนเพิ่ม (ไม่บังคับ) — แจ้งเตือน/รายงานอัตโนมัติเข้ากลุ่ม LINE มี 5 แบบ:
//     ทันทีที่ห้องส่ง      → รายงานของห้องนั้นห้องเดียว (ยอด + ผู้รายงาน)
//     ทันทีที่ครบ 7/7 ห้อง → สรุปทั้งโรงเรียนแบบรายห้อง (ห้องละ 1 บรรทัด)
//     ~08:40              → เตือนให้เริ่มส่งรายงาน (sendDailyReminder)
//     ~08:45              → เตือนเฉพาะห้องที่ยังไม่ส่ง ระบุชื่อห้อง (sendMissingRoomsReminder,
//                            ถ้าครบแล้วจะไม่ส่งซ้ำ)
//     ~09:00              → สรุปทั้งโรงเรียนอีกรอบเผื่อสำรอง (sendDailySummary)
//     (ถ้าอยากเปลี่ยนเวลาไหน แก้ที่ฟังก์ชัน createDailyTrigger() ในโค้ดนี้ได้เลย
//      หาคำว่า .atHour(...).nearMinute(...) ของ trigger นั้นแล้วแก้เลขชั่วโมง/นาทีตามต้องการ)
//  8. (LINE Notify ปิดให้บริการไปแล้วตั้งแต่ 31 มี.ค. 2568 — ใช้ LINE Official Account + Messaging API แทน)
//     สร้าง LINE Official Account ใหม่แยกต่างหาก (ไม่ใช้บัญชีเดิมที่มีคนแอดไว้เพื่อเรื่องอื่น) ที่
//     https://manager.line.biz → เปิดเมนู "ตั้งค่า" > "Messaging API" > เปิดใช้งาน
//     → ไปที่ LINE Developers Console ของ Channel นั้น > แท็บ "Messaging API"
//     → หัวข้อ "Channel access token (long-lived)" กด "Issue" จะได้ Token ยาวๆ
//     คัดลอกมาใส่แทนที่ LINE_CHANNEL_ACCESS_TOKEN ด้านล่าง
//     จากนั้นให้ครูทุกคนสแกน QR Code ของบัญชีนี้ (ดูได้ในหน้า LINE Official Account Manager)
//     เพื่อแอดเป็นเพื่อน — ระบบจะส่งข้อความแบบ "Broadcast" กระจายให้ทุกคนที่แอดพร้อมกันโดยอัตโนมัติ
//  9. (ไม่บังคับ) นำลิงก์หน้า report.html จริง มาใส่แทนที่ REPORT_PAGE_URL ด้านล่าง
//     เพื่อให้ข้อความแจ้งเตือนมีลิงก์ให้กดตรงได้เลย
//  10. ก่อนรันข้อถัดไป เช็คเขตเวลาของโปรเจกต์ก่อน: เมนูฟันเฟือง (Project Settings) ทางซ้าย
//     → เลื่อนลงหา "Time zone" ต้องเป็น "(GMT+07:00) Bangkok" ถ้าไม่ใช่ให้เปลี่ยนก่อน
//     (ถ้าเขตเวลาผิด เวลาที่ตั้งไว้จะเพี้ยนไปเป็นเวลาอื่น)
//  11. รันฟังก์ชัน createDailyTrigger() หนึ่งครั้ง (เลือกที่ dropdown ด้านบน แล้วกด Run)
//     — จะตั้งเวลาให้ระบบส่งข้อความตามเวลาทั้งหมดเข้ากลุ่ม LINE เองทุกวัน
//     (เวลาอาจคลาดเคลื่อนได้ 1-2 นาที ตามธรรมชาติของ Google — ไม่ใช่เป๊ะเวลาวินาที)
//     ระบบจะข้ามการแจ้งเตือนอัตโนมัติในวันเสาร์-อาทิตย์ให้เอง
//  12. ถ้ามีวันหยุดพิเศษ (นักขัตฤกษ์/ปิดเทอม) ที่ไม่ใช่เสาร์-อาทิตย์ → เปิด Sheet
//     "วันหยุดเพิ่มเติม" แล้วพิมพ์วันที่นั้นเพิ่ม (รูปแบบ yyyy-mm-dd เช่น 2026-08-12)
//     ระบบจะข้ามวันนั้นให้อัตโนมัติ ไม่ต้องแก้โค้ด
// ============================================================

const SPREADSHEET_ID = '1CkgncHeiyJJv5Mixbe39DEphABpMPmEYRGAM2OsBKlI';
const TZ = 'Asia/Bangkok'; // ใช้เวลาไทยเป็นหลักเสมอ กันปัญหาเรื่องเขตเวลาของเซิร์ฟเวอร์ Google

// ⚠️ Token นี้คือรหัสจริงที่ให้สิทธิ์ส่งข้อความแทนบัญชี LINE — เป็นความลับ ห้ามเผยแพร่เด็ดขาด
//    (คนละเรื่องกับ PIN/รหัสผ่านที่ตัดออกไปแล้ว — อันนั้นคือรหัสให้คนกรอก อันนี้คือกุญแจ API)
//    ไฟล์นี้เก็บอยู่ใน GitHub แบบสาธารณะ จึงต้องปล่อยให้เป็นข้อความ placeholder ตรงนี้เสมอ —
//    ให้นำ Token จริงไปแทนที่ตรงนี้เฉพาะในหน้า Apps Script editor เท่านั้น (ไม่ใช่ไฟล์ที่จะ push ขึ้น GitHub)
const LINE_CHANNEL_ACCESS_TOKEN = 'ใส่ Channel access token จาก LINE Developers Console ตรงนี้ (เฉพาะใน Apps Script editor เท่านั้น ห้ามใส่ในไฟล์ที่จะขึ้น GitHub)';
const REPORT_PAGE_URL = 'ใส่ลิงก์หน้า report.html ตรงนี้ (ไม่บังคับ เว้นว่างไว้ก็ได้)';

// ------------------------------------------------------------
//  สร้าง 4 Sheets พร้อม header (รันครั้งเดียวตอนตั้งระบบ) + ใส่ห้องเริ่มต้น 7 ห้อง
// ------------------------------------------------------------
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheets = [
    { name: 'ห้องเรียน',        headers: ['ห้อง'] },
    { name: 'นักเรียน',          headers: ['ห้อง', 'เลขที่', 'ชื่อ-นามสกุล', 'เพศ'] },
    { name: 'เช็คชื่อรายวัน',    headers: [
        'วันที่', 'ห้อง', 'เวลาที่ส่ง',
        'มา-ชาย', 'มา-หญิง', 'มา-รวม',
        'ขาด-ชาย', 'ขาด-หญิง', 'ขาด-รวม', 'รายชื่อขาด',
        'ลากิจ-ชาย', 'ลากิจ-หญิง', 'ลากิจ-รวม', 'รายชื่อลากิจ',
        'ลาป่วย-ชาย', 'ลาป่วย-หญิง', 'ลาป่วย-รวม', 'รายชื่อลาป่วย',
        'ไปกิจกรรม-ชาย', 'ไปกิจกรรม-หญิง', 'ไปกิจกรรม-รวม', 'รายชื่อไปกิจกรรม',
        'ผู้รายงาน',
      ] },
    { name: 'วันหยุดเพิ่มเติม',  headers: ['วันที่ (yyyy-mm-dd)', 'หมายเหตุ'] },
  ];

  sheets.forEach(({ name, headers }) => {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (s.getLastRow() === 0) s.appendRow(headers);
  });

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
    if (data.action === 'submit') return handleSubmit(data);
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
//  นับจำนวนชาย/หญิง/รวม ของรายชื่อในลิสต์หนึ่ง โดยเทียบจาก genderMap (ชื่อ -> เพศ)
//  ชื่อที่หาเพศไม่เจอ (พิมพ์ไม่ตรงกับ Sheet "นักเรียน") จะไม่ถูกนับในชาย/หญิง แต่ยังถูกนับใน total
// ------------------------------------------------------------
function genderCountsForList(names, genderMap) {
  let male = 0, female = 0;
  names.forEach(n => {
    const g = genderMap[n];
    if (g === 'ชาย') male++;
    else if (g === 'หญิง') female++;
  });
  return { male, female, total: names.length };
}

// ------------------------------------------------------------
//  บันทึกรายงานเช็คชื่อ — นับแยกชาย/หญิง/รวม ทุกหมวดตามฟอร์มกระดาษ
//  ถ้าห้องนี้ส่งของ "วันนี้" มาแล้ว → แก้ไขแถวเดิมแทน ไม่สร้างซ้ำ
// ------------------------------------------------------------
function handleSubmit(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 1) โหลดรายชื่อนักเรียนของห้องนี้ + สร้าง map ชื่อ -> เพศ + นับจำนวนเต็มของห้อง
  const stuSheet = ss.getSheetByName('นักเรียน');
  const stuRows  = stuSheet.getDataRange().getValues();
  const genderMap = {};
  let maleTotal = 0, femaleTotal = 0;
  for (let i = 1; i < stuRows.length; i++) {
    if (stuRows[i][0] === data.room && stuRows[i][2]) {
      const name   = String(stuRows[i][2]).trim();
      const gender = String(stuRows[i][3] || '').trim();
      genderMap[name] = gender;
      if (gender === 'ชาย') maleTotal++;
      else if (gender === 'หญิง') femaleTotal++;
    }
  }
  const total = maleTotal + femaleTotal;

  // 2) ตัดชื่อซ้ำในลิสต์เดียวกันออก (กันหัวหน้าห้องติ๊กชื่อเดิมซ้ำโดยไม่ตั้งใจ)
  const absentNames        = uniqueList(data.absentNames);
  const personalLeaveNames = uniqueList(data.personalLeaveNames);
  const sickLeaveNames     = uniqueList(data.sickLeaveNames);
  const activityNames      = uniqueList(data.activityNames);

  const absentC   = genderCountsForList(absentNames, genderMap);
  const personalC = genderCountsForList(personalLeaveNames, genderMap);
  const sickC     = genderCountsForList(sickLeaveNames, genderMap);
  const activityC = genderCountsForList(activityNames, genderMap);

  // 3) มา = จำนวนเต็ม - (ขาด+ลากิจ+ลาป่วย+ไปกิจกรรม) คำนวณแยกชาย/หญิง/รวม ไม่ปนกัน
  const presentMale   = Math.max(0, maleTotal   - absentC.male   - personalC.male   - sickC.male   - activityC.male);
  const presentFemale = Math.max(0, femaleTotal - absentC.female - personalC.female - sickC.female - activityC.female);
  const presentTotal  = Math.max(0, total - absentC.total - personalC.total - sickC.total - activityC.total);

  const dateStr       = todayStr();
  const nowDate       = new Date();
  const timeStr       = Utilities.formatDate(nowDate, TZ, 'HH:mm:ss');
  const reporterName  = String(data.reporterName || '').trim();

  const rowValues = [
    dateStr, data.room, timeStr,
    presentMale, presentFemale, presentTotal,
    absentC.male, absentC.female, absentC.total, absentNames.join(', '),
    personalC.male, personalC.female, personalC.total, personalLeaveNames.join(', '),
    sickC.male, sickC.female, sickC.total, sickLeaveNames.join(', '),
    activityC.male, activityC.female, activityC.total, activityNames.join(', '),
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
    if (cellDateStr(logRows[i][0]) === dateStr && logRows[i][1] === data.room) { foundRow = i + 1; break; } // +1 เพราะ sheet เริ่มแถว 1
  }

  if (foundRow > 0) {
    logSheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]); // แก้ไขแถวเดิมแทนที่
  } else {
    logSheet.appendRow(rowValues);
  }

  // 5) ส่งรายงานห้องนี้เข้า LINE ทันที (ไม่รอเวลาไหน) แล้วเช็คว่าครบทุกห้องหรือยัง
  //    ถ้าครบแล้ว ส่งสรุปทั้งโรงเรียนตามด้วยอีกข้อความ
  sendRoomReportToLine(data.room, total, presentTotal, absentC.total, personalC.total, sickC.total, activityC.total, reporterName, nowDate);

  const freshSummary = getDailySummaryData(dateStr);
  if (freshSummary.totals.submittedCount === freshSummary.totals.totalRooms) {
    sendLineNotify(buildCompleteSummaryMessage(freshSummary));
  }

  return respond({
    status: 'ok',
    present:       { male: presentMale, female: presentFemale, total: presentTotal },
    absent:        absentC,
    personalLeave: personalC,
    sickLeave:     sickC,
    activity:      activityC,
  });
}

// ------------------------------------------------------------
//  เช็คว่า "ห้องนี้" ส่งรายงานวันนี้ไปแล้วหรือยัง
//  ใช้ตอนหัวหน้าห้องเปิดหน้ารายงาน กันกดส่งซ้ำโดยไม่รู้ตัวว่าเขียนทับของเดิม
// ------------------------------------------------------------
function handleMyStatus(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const logSheet = ss.getSheetByName('เช็คชื่อรายวัน');
  const logRows  = logSheet.getDataRange().getValues();
  const dateStr  = todayStr();

  for (let i = 1; i < logRows.length; i++) {
    const r = logRows[i];
    if (cellDateStr(r[0]) === dateStr && r[1] === params.room) {
      return respond({
        status: 'ok', submitted: true, time: cellTimeStr(r[2]),
        present:       { male: r[3],  female: r[4],  total: r[5]  },
        absent:        { male: r[6],  female: r[7],  total: r[8]  },
        personalLeave: { male: r[10], female: r[11], total: r[12] },
        sickLeave:     { male: r[14], female: r[15], total: r[16] },
        activity:      { male: r[18], female: r[19], total: r[20] },
      });
    }
  }
  return respond({ status: 'ok', submitted: false });
}

// ------------------------------------------------------------
//  สรุปยอดของวันที่ระบุ ทุกห้อง (ห้องไหนยังไม่ส่ง = submitted:false)
//  "จำนวนเต็ม" ของแต่ละห้องนับจาก Sheet "นักเรียน" เสมอ ไม่ว่าจะส่งรายงานวันนี้หรือยัง
// ------------------------------------------------------------
function handleByDate(dateStr) {
  return respond(getDailySummaryData(dateStr));
}

// เนื้อหาการคำนวณจริงของ handleByDate แยกออกมาเป็นฟังก์ชันของตัวเอง
// เพื่อให้ sendDailySummary() (แจ้งเตือน LINE) เอาไปใช้ซ้ำได้ ไม่ต้องคำนวณเองซ้ำอีกชุด
function getDailySummaryData(dateStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const roomSheet = ss.getSheetByName('ห้องเรียน');
  const roomRows  = roomSheet.getDataRange().getValues();
  const allRooms  = [];
  for (let i = 1; i < roomRows.length; i++) {
    if (roomRows[i][0]) allRooms.push(String(roomRows[i][0]).trim());
  }

  // นับจำนวนเต็มต่อห้อง (ชาย/หญิง/รวม) จาก Sheet "นักเรียน"
  const stuSheet = ss.getSheetByName('นักเรียน');
  const stuRows  = stuSheet.getDataRange().getValues();
  const rosterTotals = {}; // room -> {male, female, total}
  for (let i = 1; i < stuRows.length; i++) {
    const room = stuRows[i][0];
    if (!room || !stuRows[i][2]) continue;
    const gender = String(stuRows[i][3] || '').trim();
    if (!rosterTotals[room]) rosterTotals[room] = { male: 0, female: 0, total: 0 };
    rosterTotals[room].total++;
    if (gender === 'ชาย') rosterTotals[room].male++;
    else if (gender === 'หญิง') rosterTotals[room].female++;
  }

  const logSheet = ss.getSheetByName('เช็คชื่อรายวัน');
  const logRows  = logSheet.getDataRange().getValues();
  const submittedMap = {}; // room -> row data
  for (let i = 1; i < logRows.length; i++) {
    const r = logRows[i];
    if (cellDateStr(r[0]) === dateStr) {
      submittedMap[r[1]] = {
        time: cellTimeStr(r[2]),
        present:       { male: Number(r[3])  || 0, female: Number(r[4])  || 0, total: Number(r[5])  || 0 },
        absent:        { male: Number(r[6])  || 0, female: Number(r[7])  || 0, total: Number(r[8])  || 0, names: r[9]  || '' },
        personalLeave: { male: Number(r[10]) || 0, female: Number(r[11]) || 0, total: Number(r[12]) || 0, names: r[13] || '' },
        sickLeave:     { male: Number(r[14]) || 0, female: Number(r[15]) || 0, total: Number(r[16]) || 0, names: r[17] || '' },
        activity:      { male: Number(r[18]) || 0, female: Number(r[19]) || 0, total: Number(r[20]) || 0, names: r[21] || '' },
      };
    }
  }

  let submittedCount = 0;
  const grand = {
    roster:        { male: 0, female: 0, total: 0 },
    present:       { male: 0, female: 0, total: 0 },
    absent:        { male: 0, female: 0, total: 0 },
    personalLeave: { male: 0, female: 0, total: 0 },
    sickLeave:     { male: 0, female: 0, total: 0 },
    activity:      { male: 0, female: 0, total: 0 },
  };
  function addCat(target, src) {
    target.male += src.male; target.female += src.female; target.total += src.total;
  }

  const rooms = allRooms.map(room => {
    const roster = rosterTotals[room] || { male: 0, female: 0, total: 0 };
    addCat(grand.roster, roster);

    const rec = submittedMap[room];
    if (rec) {
      submittedCount++;
      addCat(grand.present, rec.present);
      addCat(grand.absent, rec.absent);
      addCat(grand.personalLeave, rec.personalLeave);
      addCat(grand.sickLeave, rec.sickLeave);
      addCat(grand.activity, rec.activity);
      return {
        room, roster, submitted: true, time: rec.time,
        present: rec.present, absent: rec.absent,
        personalLeave: rec.personalLeave, sickLeave: rec.sickLeave, activity: rec.activity,
      };
    }
    return { room, roster, submitted: false };
  });

  return {
    status: 'ok',
    date: dateStr,
    rooms,
    totals: { ...grand, submittedCount, totalRooms: allRooms.length },
  };
}

// ------------------------------------------------------------
//  ตั้งเวลาให้ระบบส่งข้อความแจ้งเตือนเข้ากลุ่ม LINE เองทุกวันตอน ~08:40 (รันครั้งเดียวตอนตั้งระบบ)
//  ลบ trigger เดิมของฟังก์ชันนี้ก่อนเสมอ กันสร้างซ้ำถ้ารันคำสั่งนี้มากกว่า 1 ครั้ง
// ------------------------------------------------------------
function createDailyTrigger() {
  const handledByThis = ['sendDailyReminder', 'sendDailySummary', 'sendMissingRoomsReminder'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handledByThis.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('sendDailyReminder')
    .timeBased()
    .atHour(8)
    .nearMinute(40)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('sendMissingRoomsReminder')
    .timeBased()
    .atHour(8)
    .nearMinute(45)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .atHour(9)
    .nearMinute(0)
    .everyDays(1)
    .create();

  Logger.log('ตั้งเวลาอัตโนมัติทุกวันเรียบร้อยแล้ว: เตือนให้ส่งรายงาน ~08:40 / เตือนห้องที่ยังไม่ส่ง ~08:45 / สรุปยอดทั้งโรงเรียนเข้า LINE ~09:00 (นอกจากนี้ระบบส่ง "รายงานทันที" ทุกครั้งที่มีห้องส่งเข้ามา และส่ง "สรุปครบทุกห้อง" อัตโนมัติทันทีที่ห้องสุดท้ายส่งครบ — ข้ามเสาร์-อาทิตย์ และวันที่อยู่ใน Sheet "วันหยุดเพิ่มเติม" ให้อัตโนมัติ)');
}

// ------------------------------------------------------------
//  เช็คว่าวันนี้เป็นวันเรียนหรือไม่ (ไม่ใช่เสาร์-อาทิตย์ และไม่อยู่ใน Sheet "วันหยุดเพิ่มเติม")
// ------------------------------------------------------------
function isSchoolDay(date) {
  const day = date.getDay(); // 0 = อาทิตย์, 6 = เสาร์
  if (day === 0 || day === 6) return false;

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('วันหยุดเพิ่มเติม');
  if (!sheet) return true;

  const dateStr = Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (cellDateStr(rows[i][0]) === dateStr) return false;
  }
  return true;
}

// ------------------------------------------------------------
//  ส่งรายงานสถานะปัจจุบัน (รายห้อง) เข้า LINE — ใช้ร่วมกันโดย sendDailyReminder (~08:40)
//  และ sendDailySummary (~09:00 fallback) เนื้อหาเหมือนกัน ต่างแค่เวลาที่ trigger เรียก
//  ห้องที่ยังไม่ส่งจะโชว์ "⏳ ยังไม่ส่งข้อมูล" ในตัวอยู่แล้ว (จาก buildRoomLine) จึงทำหน้าที่
//  เป็นทั้งรายงานและตัวเตือนในข้อความเดียว — ถ้าครบ 7/7 ห้องไปแล้วก่อนหน้านี้ ข้อความนี้ก็ยัง
//  ส่งซ้ำได้ตามเวลาปกติ ไม่ได้ตัดออก (ครูต้องการให้เห็นสถานะ ณ เวลานั้นเสมอ)
// ------------------------------------------------------------
function sendCurrentReportToLine(now) {
  const data = getDailySummaryData(todayStr());
  const lines = [`📊 สรุปเช็คชื่อ${thaiFullDateStr(now)}`, ''];
  data.rooms.forEach(r => lines.push(buildRoomLine(r)));
  sendLineNotify(lines.join('\n'));
}

// ------------------------------------------------------------
//  ทำงานอัตโนมัติจาก trigger ที่ตั้งไว้ ~08:40 — ส่งรายงานสถานะปัจจุบันเข้า LINE ทันที
//  (ถ้าห้องไหนส่งไปแล้วก่อน 08:40 ก็จะเห็นข้อมูลจริงของห้องนั้นเลย ไม่ใช่แค่ข้อความเตือนเฉยๆ)
//  ข้ามการส่งถ้าวันนี้ไม่ใช่วันเรียน
// ------------------------------------------------------------
function sendDailyReminder() {
  const now = new Date();
  if (!isSchoolDay(now)) return;
  sendCurrentReportToLine(now);
}

// ------------------------------------------------------------
//  ส่งรายงานสถานะปัจจุบันเข้า LINE อีกรอบ — ทำงานอัตโนมัติจาก trigger ที่ตั้งไว้ ~09:00
//  ทำหน้าที่เป็น fallback เผื่อครูอยากดูสรุปช่วงเช้ารวด — ข้ามการส่งถ้าวันนี้ไม่ใช่วันเรียน
// ------------------------------------------------------------
function sendDailySummary() {
  const now = new Date();
  if (!isSchoolDay(now)) return;
  sendCurrentReportToLine(now);
}

// ------------------------------------------------------------
//  แจ้งเตือนห้องที่ยังไม่ส่งรายงาน — ทำงานอัตโนมัติจาก trigger ที่ตั้งไว้ ~08:45
//  ถ้าครบทุกห้องแล้ว (sendCompleteSummary ส่งไปแล้วตอนห้องสุดท้ายส่ง) จะไม่ส่งข้อความนี้ซ้ำ
// ------------------------------------------------------------
function sendMissingRoomsReminder() {
  const now = new Date();
  if (!isSchoolDay(now)) return;

  const data = getDailySummaryData(todayStr());
  const missing = data.rooms.filter(r => !r.submitted).map(r => r.room);
  if (!missing.length) return; // ครบแล้ว ไม่ต้องเตือน

  const lines = [];
  lines.push('⏰ แจ้งเตือนการรายงานสถิติการมาเรียน');
  lines.push('');
  lines.push('ขณะนี้ยังไม่ได้รับข้อมูลจาก');
  missing.forEach(room => lines.push(`- ${room}`));
  lines.push('');
  lines.push('กรุณาส่งข้อมูลภายในเวลา 08.45 น.');

  sendLineNotify(lines.join('\n'));
}

// ------------------------------------------------------------
//  ส่งรายงานของห้องเดียวเข้า LINE ทันทีที่ห้องนั้นส่งรายงาน (เรียกจาก handleSubmit)
// ------------------------------------------------------------
function sendRoomReportToLine(room, total, presentTotal, absentTotal, personalTotal, sickTotal, activityTotal, reporterName, nowDate) {
  const lines = [];
  lines.push('📊 รายงานสถิติการมาเรียนประจำวัน');
  lines.push(thaiFullDateStr(nowDate));
  lines.push('');
  lines.push(`ชั้น ${room}`);
  lines.push(`👥 นักเรียนทั้งหมด ${total} คน`);
  lines.push(`✅ มาเรียน ${presentTotal} คน`);
  lines.push(`❌ ขาดเรียน ${absentTotal} คน`);
  lines.push(`📝 ลากิจ ${personalTotal} คน`);
  lines.push(`🤒 ลาป่วย ${sickTotal} คน`);
  lines.push(`🏃 ไปกิจกรรม ${activityTotal} คน`);
  lines.push('');
  lines.push(`ผู้รายงาน: ${reporterName || '-'}`);
  lines.push(`เวลาแจ้งข้อมูล: ${thaiTimeStr(nowDate)}`);
  sendLineNotify(lines.join('\n'));
}

// ------------------------------------------------------------
//  สร้างข้อความสรุปทั้งโรงเรียน แบบรายห้อง — ใช้ตอนส่งครบ 7/7 ห้อง
// ------------------------------------------------------------
function buildCompleteSummaryMessage(data) {
  const lines = ['📋 รายละเอียดรายห้อง', thaiFullDateStr(new Date(data.date + 'T00:00:00')), ''];
  data.rooms.forEach(r => lines.push(buildRoomLine(r)));
  return lines.join('\n');
}

// สร้างบรรทัดสรุป 1 ห้อง — ใช้ร่วมกันทั้ง buildCompleteSummaryMessage และ sendDailySummary
// แสดงเฉพาะหมวดที่มีคนมากกว่า 0 เท่านั้น (ไม่โชว์หมวดที่เป็น 0 กันข้อความยาวเกินจำเป็น)
function buildRoomLine(r) {
  if (!r.submitted) return `${r.room}: ⏳ ยังไม่ส่งข้อมูล`;
  const parts = [`ทั้งหมด ${r.roster.total}`, `มา ${r.present.total}`];
  if (r.absent.total > 0)        parts.push(`ขาด ${r.absent.total}`);
  if (r.personalLeave.total > 0) parts.push(`ลากิจ ${r.personalLeave.total}`);
  if (r.sickLeave.total > 0)     parts.push(`ป่วย ${r.sickLeave.total}`);
  if (r.activity.total > 0)      parts.push(`กิจกรรม ${r.activity.total}`);
  return `${r.room}: ${parts.join(' | ')}`;
}

// ------------------------------------------------------------
//  ส่งข้อความเข้ากลุ่ม LINE ผ่าน LINE Notify — จุดเดียวที่เรียก UrlFetchApp จริง
//  (ทุกฟังก์ชันที่ต้องส่ง LINE เรียกผ่านตัวนี้ทั้งหมด กันโค้ดซ้ำ)
// ------------------------------------------------------------
function sendLineNotify(message) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify({ messages: [{ type: 'text', text: message }] }),
    });
  } catch (err) {
    Logger.log('ส่ง LINE ไม่สำเร็จ: ' + err.toString());
  }
}

// ------------------------------------------------------------
//  แปลงวันเป็นข้อความไทยเต็มรูปแบบ "วันจันทร์ที่ 3 สิงหาคม 2569" (ปี พ.ศ.)
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

// แปลงค่าเซลล์ (อาจเป็น Date object หรือ string) ให้เป็น string รูปแบบ yyyy-MM-dd เสมอ
// จำเป็นเพราะ Google Sheets มักแปลงข้อความที่หน้าตาเหมือนวันที่ให้เป็นเซลล์ชนิด Date เองอัตโนมัติ
function cellDateStr(cell) {
  return cell instanceof Date ? Utilities.formatDate(cell, TZ, 'yyyy-MM-dd') : String(cell).trim();
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
