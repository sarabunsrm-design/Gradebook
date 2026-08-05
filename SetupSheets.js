/**
 * รันฟังก์ชันนี้ "1 ครั้งเดียว" จาก Script Editor เพื่อสร้างโครงสร้างชีตทั้งหมด
 * วิธีรัน: เปิด Apps Script Editor > เลือกฟังก์ชัน setupSheets > กด Run
 * (ครั้งแรกระบบจะขอ authorize สิทธิ์ ให้กด Allow)
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheetIfNotExists_(ss, 'Teachers',
    ['TeacherID', 'Name', 'Email', 'Role']);

  createSheetIfNotExists_(ss, 'Subjects',
    ['SubjectID', 'SubjectName', 'TeacherEmail', 'ClassRoom', 'AcademicYear', 'Semester',
     'SubjectCode', 'LearningArea', 'SubjectType', 'HoursPerWeek', 'Credits']);

  createSheetIfNotExists_(ss, 'Students',
    ['StudentID', 'StudentCode', 'Name', 'ClassRoom', 'RollNumber', 'AccessCode', 'Status']);

  createSheetIfNotExists_(ss, 'Enrollment',
    ['EnrollID', 'StudentID', 'SubjectID', 'GradeOverride', 'ClassRoom', 'RollNumber',
     'DesiredCharacteristics', 'ReadThinkWrite',
     'Char1', 'Char2', 'Char3', 'Char4', 'Char5', 'Char6', 'Char7', 'Char8',
     'Read1', 'Read2', 'Read3', 'Read4', 'Read5']);

  createSheetIfNotExists_(ss, 'ClassAdvisors',
    ['ClassRoom', 'AcademicYear', 'Semester', 'AdvisorName']);

  createSheetIfNotExists_(ss, 'ScoreItems',
    ['ItemID', 'SubjectID', 'Period', 'ItemName', 'MaxScore', 'Weight']);

  createSheetIfNotExists_(ss, 'Scores',
    ['ScoreID', 'StudentID', 'ItemID', 'RawScore', 'Timestamp', 'UpdatedBy']);

  createSheetIfNotExists_(ss, 'Config',
    ['Key', 'Value']);

  // เพิ่มผู้รันฟังก์ชันนี้เป็น admin คนแรกอัตโนมัติ (ถ้ายังไม่มีครูในระบบ)
  const teacherSheet = ss.getSheetByName('Teachers');
  if (teacherSheet.getLastRow() === 1) {
    teacherSheet.appendRow(['T001', 'ผู้ดูแลระบบ', Session.getActiveUser().getEmail(), 'admin']);
  }

  // ลบชีตเริ่มต้น "Sheet1" ถ้ายังไม่ได้ใช้งานอะไร
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.flush();
  Logger.log('✅ สร้างโครงสร้างชีตทั้งหมดเรียบร้อยแล้ว');
}

function createSheetIfNotExists_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f0f0f0');
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * (สำหรับคนที่ตั้งระบบไปแล้วก่อนมีฟีเจอร์ "เลขที่")
 * รันฟังก์ชันนี้ "1 ครั้งเดียว" เพื่อเพิ่มคอลัมน์ RollNumber (เลขที่) เข้าไปในชีต Students ที่มีอยู่แล้ว
 * ถ้าเป็นการติดตั้งใหม่ทั้งหมด (รัน setupSheets ครั้งแรก) ไม่ต้องรันฟังก์ชันนี้ เพราะสร้างคอลัมน์ไว้ให้แล้ว
 */
function migrateAddRollNumberColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  if (!sheet) throw new Error('ไม่พบชีต Students - กรุณารัน setupSheets ก่อน');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('RollNumber') !== -1) {
    Logger.log('มีคอลัมน์ RollNumber อยู่แล้ว ไม่ต้องทำอะไรเพิ่ม');
    return;
  }

  const accessCodeCol = headers.indexOf('AccessCode') + 1; // 1-indexed
  if (accessCodeCol === 0) throw new Error('ไม่พบคอลัมน์ AccessCode ในชีต Students ผิดปกติ กรุณาตรวจสอบชีตด้วยตนเอง');

  sheet.insertColumnBefore(accessCodeCol);
  sheet.getRange(1, accessCodeCol).setValue('RollNumber').setFontWeight('bold').setBackground('#f0f0f0');

  Logger.log('✅ เพิ่มคอลัมน์ RollNumber (เลขที่) เรียบร้อยแล้ว - นักเรียนเดิมจะยังไม่มีเลขที่ กรุณาเข้าไปกรอกเพิ่มที่หน้า "จัดการนักเรียน" หรือแก้ในชีตโดยตรง');
}

/**
 * (สำหรับคนที่ตั้งระบบไปแล้วก่อนมีฟีเจอร์เกรดพิเศษ "ร"/"มส")
 * รันฟังก์ชันนี้ "1 ครั้งเดียว" เพื่อเพิ่มคอลัมน์ GradeOverride เข้าไปในชีต Enrollment ที่มีอยู่แล้ว
 * ถ้าเป็นการติดตั้งใหม่ทั้งหมด (รัน setupSheets ครั้งแรก) ไม่ต้องรันฟังก์ชันนี้ เพราะสร้างคอลัมน์ไว้ให้แล้ว
 */
function migrateAddGradeOverrideColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enrollment');
  if (!sheet) throw new Error('ไม่พบชีต Enrollment - กรุณารัน setupSheets ก่อน');

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('GradeOverride') !== -1) {
    Logger.log('มีคอลัมน์ GradeOverride อยู่แล้ว ไม่ต้องทำอะไรเพิ่ม');
    return;
  }

  sheet.getRange(1, lastCol + 1).setValue('GradeOverride').setFontWeight('bold').setBackground('#f0f0f0');
  Logger.log('✅ เพิ่มคอลัมน์ GradeOverride เรียบร้อยแล้ว - ใช้สำหรับให้เกรดพิเศษ "ร" หรือ "มส" แทนการคำนวณอัตโนมัติ');
}

/**
 * (สำหรับคนที่ตั้งระบบไปแล้วก่อนมีฟีเจอร์สถานะนักเรียน)
 * รันฟังก์ชันนี้ "1 ครั้งเดียว" เพื่อเพิ่มคอลัมน์ Status เข้าไปในชีต Students ที่มีอยู่แล้ว
 * นักเรียนเดิมทั้งหมดจะถูกตั้งเป็น "กำลังเรียน" ให้อัตโนมัติ
 */
function migrateAddStudentStatusColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  if (!sheet) throw new Error('ไม่พบชีต Students - กรุณารัน setupSheets ก่อน');

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('Status') !== -1) {
    Logger.log('มีคอลัมน์ Status อยู่แล้ว ไม่ต้องทำอะไรเพิ่ม');
    return;
  }

  const statusCol = lastCol + 1;
  sheet.getRange(1, statusCol).setValue('Status').setFontWeight('bold').setBackground('#f0f0f0');

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const defaults = [];
    for (let i = 0; i < lastRow - 1; i++) defaults.push(['กำลังเรียน']);
    sheet.getRange(2, statusCol, lastRow - 1, 1).setValues(defaults);
  }

  Logger.log('✅ เพิ่มคอลัมน์ Status เรียบร้อยแล้ว - ตั้งนักเรียนเดิมทั้งหมดเป็น "กำลังเรียน" ให้อัตโนมัติ');
}

/**
 * (สำหรับคนที่ตั้งระบบไปแล้วก่อนมีฟีเจอร์ "แช่แข็ง" ห้อง/เลขที่ ณ ตอนลงทะเบียน)
 * รันฟังก์ชันนี้ "1 ครั้งเดียว" เพื่อเพิ่มคอลัมน์ ClassRoom/RollNumber เข้าไปในชีต Enrollment ที่มีอยู่แล้ว
 * ทำให้ห้อง/เลขที่ที่แสดงในแต่ละวิชาไม่เปลี่ยนตามการเลื่อนชั้นของนักเรียนในภายหลัง (ดูรายละเอียดใน Students.gs)
 * ข้อมูลเก่าที่เคยลงทะเบียนไว้ก่อนหน้านี้ จะถูก backfill ด้วยห้อง/เลขที่ "ปัจจุบัน" ของนักเรียนแต่ละคน ณ ตอนรันฟังก์ชันนี้
 * (เพราะไม่มีข้อมูลย้อนหลังที่แท้จริงเก็บไว้ก่อนหน้า จึงทำได้ดีที่สุดเท่านี้สำหรับข้อมูลเก่า)
 */
function migrateAddEnrollmentSnapshotColumns() {
  const enrollSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enrollment');
  if (!enrollSheet) throw new Error('ไม่พบชีต Enrollment - กรุณารัน setupSheets ก่อน');

  let headers = enrollSheet.getRange(1, 1, 1, enrollSheet.getLastColumn()).getValues()[0];
  const hasClassRoom = headers.indexOf('ClassRoom') !== -1;
  const hasRollNumber = headers.indexOf('RollNumber') !== -1;

  if (hasClassRoom && hasRollNumber) {
    Logger.log('มีคอลัมน์ ClassRoom/RollNumber ในชีต Enrollment อยู่แล้ว ไม่ต้องทำอะไรเพิ่ม');
    return;
  }

  if (!hasClassRoom) {
    enrollSheet.getRange(1, enrollSheet.getLastColumn() + 1).setValue('ClassRoom').setFontWeight('bold').setBackground('#f0f0f0');
  }
  if (!hasRollNumber) {
    enrollSheet.getRange(1, enrollSheet.getLastColumn() + 1).setValue('RollNumber').setFontWeight('bold').setBackground('#f0f0f0');
  }

  // อ่าน header ใหม่หลังเพิ่มคอลัมน์แล้ว เพื่อหาตำแหน่งคอลัมน์ที่แน่นอน
  headers = enrollSheet.getRange(1, 1, 1, enrollSheet.getLastColumn()).getValues()[0];
  const classRoomCol = headers.indexOf('ClassRoom') + 1;
  const rollNumberCol = headers.indexOf('RollNumber') + 1;
  const studentIdColIdx = headers.indexOf('StudentID');

  const studentMap = {};
  getSheetData_('Students').forEach(s => studentMap[s.StudentID] = s);

  const lastRow = enrollSheet.getLastRow();
  if (lastRow > 1) {
    const allData = enrollSheet.getRange(2, 1, lastRow - 1, enrollSheet.getLastColumn()).getValues();
    const classRoomValues = [];
    const rollNumberValues = [];
    allData.forEach(row => {
      const studentId = row[studentIdColIdx];
      const student = studentMap[studentId];
      classRoomValues.push([student ? (student.ClassRoom || '') : '']);
      rollNumberValues.push([student ? (student.RollNumber || '') : '']);
    });
    enrollSheet.getRange(2, classRoomCol, lastRow - 1, 1).setValues(classRoomValues);
    enrollSheet.getRange(2, rollNumberCol, lastRow - 1, 1).setValues(rollNumberValues);
  }

  Logger.log('✅ เพิ่มคอลัมน์ ClassRoom/RollNumber ในชีต Enrollment เรียบร้อยแล้ว (backfill ด้วยค่าปัจจุบันของนักเรียน ณ ตอนนี้)');
}

/**
 * (สำหรับคนที่ตั้งระบบไปแล้วก่อนมีฟีเจอร์ ปพ.5)
 * รันฟังก์ชันนี้ "1 ครั้งเดียว" เพื่อเพิ่มคอลัมน์รายละเอียดวิชา (รหัสวิชา, กลุ่มสาระ, ประเภทวิชา, เวลาเรียน, หน่วยกิต)
 * เข้าไปในชีต Subjects ที่มีอยู่แล้ว - วิชาเดิมจะยังไม่มีข้อมูลนี้ กรุณาเข้าไปแก้ไขเพิ่มทีหลังที่หน้า Dashboard
 */
function migrateAddSubjectDetailsColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subjects');
  if (!sheet) throw new Error('ไม่พบชีต Subjects - กรุณารัน setupSheets ก่อน');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newColumns = ['SubjectCode', 'LearningArea', 'SubjectType', 'HoursPerWeek', 'Credits'];
  const missing = newColumns.filter(c => headers.indexOf(c) === -1);

  if (!missing.length) {
    Logger.log('มีคอลัมน์รายละเอียดวิชาครบอยู่แล้ว ไม่ต้องทำอะไรเพิ่ม');
    return;
  }

  let nextCol = sheet.getLastColumn() + 1;
  missing.forEach(col => {
    sheet.getRange(1, nextCol).setValue(col).setFontWeight('bold').setBackground('#f0f0f0');
    nextCol++;
  });

  Logger.log('✅ เพิ่มคอลัมน์รายละเอียดวิชาเรียบร้อยแล้ว (' + missing.join(', ') + ') - กรุณาเข้าไปกรอกเพิ่มที่หน้า Dashboard ทีละวิชา');
}

/**
 * (สำหรับคนที่ตั้งระบบไปแล้วก่อนมีฟีเจอร์ ปพ.5)
 * รันฟังก์ชันนี้ "1 ครั้งเดียว" เพื่อเพิ่มคอลัมน์ประเมินคุณลักษณะอันพึงประสงค์ + การอ่านคิดวิเคราะห์เขียน
 * เข้าไปในชีต Enrollment ที่มีอยู่แล้ว
 */
function migrateAddCharacterAssessmentColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enrollment');
  if (!sheet) throw new Error('ไม่พบชีต Enrollment - กรุณารัน setupSheets ก่อน');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newColumns = ['DesiredCharacteristics', 'ReadThinkWrite'];
  const missing = newColumns.filter(c => headers.indexOf(c) === -1);

  if (!missing.length) {
    Logger.log('มีคอลัมน์ประเมินคุณลักษณะครบอยู่แล้ว ไม่ต้องทำอะไรเพิ่ม');
    return;
  }

  let nextCol = sheet.getLastColumn() + 1;
  missing.forEach(col => {
    sheet.getRange(1, nextCol).setValue(col).setFontWeight('bold').setBackground('#f0f0f0');
    nextCol++;
  });

  Logger.log('✅ เพิ่มคอลัมน์ประเมินคุณลักษณะเรียบร้อยแล้ว (' + missing.join(', ') + ')');
}

/**
 * (สำหรับคนที่ตั้งระบบไปแล้วก่อนมีตัวชี้วัดย่อยคุณลักษณะฯ/อ่านคิดฯ)
 * รันฟังก์ชันนี้ "1 ครั้งเดียว" เพื่อเพิ่มคอลัมน์ตัวชี้วัดย่อย 8 ข้อ (คุณลักษณะฯ) + 5 ข้อ (อ่านคิดฯ)
 * เข้าไปในชีต Enrollment ที่มีอยู่แล้ว
 */
function migrateAddDetailedCharacterColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enrollment');
  if (!sheet) throw new Error('ไม่พบชีต Enrollment - กรุณารัน setupSheets ก่อน');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newColumns = ['Char1', 'Char2', 'Char3', 'Char4', 'Char5', 'Char6', 'Char7', 'Char8',
    'Read1', 'Read2', 'Read3', 'Read4', 'Read5'];
  const missing = newColumns.filter(c => headers.indexOf(c) === -1);

  if (!missing.length) {
    Logger.log('มีคอลัมน์ตัวชี้วัดย่อยครบอยู่แล้ว ไม่ต้องทำอะไรเพิ่ม');
    return;
  }

  let nextCol = sheet.getLastColumn() + 1;
  missing.forEach(col => {
    sheet.getRange(1, nextCol).setValue(col).setFontWeight('bold').setBackground('#f0f0f0');
    nextCol++;
  });

  Logger.log('✅ เพิ่มคอลัมน์ตัวชี้วัดย่อยคุณลักษณะฯ/อ่านคิดฯ เรียบร้อยแล้ว (' + missing.join(', ') + ')');
}

/**
 * (ทางเลือก) เพิ่มครูคนใหม่เข้าระบบด้วยตัวเอง โดยรันฟังก์ชันนี้แล้วแก้ค่าด้านล่าง
 * ปกติแนะนำให้ใช้หน้าเว็บ "จัดการครู" (?page=teachers) แทน เพราะสะดวกกว่า
 * ฟังก์ชันนี้มีไว้เผื่อกรณีต้องเพิ่มครูก่อนที่จะมี admin คนแรกเข้าระบบ
 */
function addTeacherManually() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  const id = 'T' + new Date().getTime();
  sheet.appendRow([id, 'ชื่อครู (แก้ตรงนี้)', 'teacher-email@school.ac.th (แก้ตรงนี้)', 'teacher']);
}