// สถานะนักเรียนที่รองรับ - "กำลังเรียน" คือค่าเริ่มต้นเสมอสำหรับนักเรียนใหม่
var STUDENT_STATUSES_ = ['กำลังเรียน', 'จบการศึกษา', 'ย้ายออก'];
var STUDENT_STATUS_ACTIVE_ = 'กำลังเรียน';

/**
 * เรียงนักเรียนตามเลขที่ (ตัวเลขน้อยไปมาก) ถ้าไม่มีเลขที่หรือเลขที่เท่ากัน ให้เรียงตามรหัสนักเรียนแทน
 * ใช้ร่วมกันทุกจุดที่ต้องแสดงรายชื่อนักเรียน เพราะปกติครูกรอกคะแนนเรียงตามเลขที่
 */
function sortByRollNumber_(a, b) {
  const rollA = a.rollNumber === '' || a.rollNumber === undefined || a.rollNumber === null ? null : Number(a.rollNumber);
  const rollB = b.rollNumber === '' || b.rollNumber === undefined || b.rollNumber === null ? null : Number(b.rollNumber);

  if (rollA !== null && rollB !== null && rollA !== rollB) return rollA - rollB;
  if (rollA !== null && rollB === null) return -1; // มีเลขที่มาก่อนคนที่ไม่มีเลขที่
  if (rollA === null && rollB !== null) return 1;

  return String(a.studentCode).localeCompare(String(b.studentCode));
}

/**
 * ดึงรายชื่อนักเรียนทั้งหมดในฐานข้อมูลกลาง (ครูทุกคนเห็นชุดเดียวกัน)
 * รวมนักเรียนทุกสถานะ (กำลังเรียน/จบการศึกษา/ย้ายออก) - ฝั่งหน้าเว็บจะกรองแสดงผลเอง
 */
function getAllStudentsCentral() {
  requireTeacher_();
  return getSheetData_('Students')
    .map(s => ({
      studentId: s.StudentID,
      studentCode: s.StudentCode,
      name: s.Name,
      classRoom: s.ClassRoom,
      rollNumber: s.RollNumber,
      accessCode: s.AccessCode,
      status: s.Status || STUDENT_STATUS_ACTIVE_
    }))
    .sort((a, b) => String(a.classRoom).localeCompare(String(b.classRoom)) || sortByRollNumber_(a, b));
}

/**
 * ดึงรายชื่อห้องเรียนทั้งหมดที่มีนักเรียน "กำลังเรียน" อยู่ (ไม่ซ้ำ เรียงตามตัวอักษร)
 * ไม่รวมห้องที่มีแต่นักเรียนจบการศึกษา/ย้ายออกแล้ว เพื่อไม่ให้ตัวเลือกรกเกินจำเป็น
 */
function getDistinctClassRooms() {
  requireTeacher_();
  const rooms = getSheetData_('Students')
    .filter(s => (s.Status || STUDENT_STATUS_ACTIVE_) === STUDENT_STATUS_ACTIVE_)
    .map(s => String(s.ClassRoom || '').trim())
    .filter(r => r !== '');
  return Array.from(new Set(rooms)).sort();
}

/**
 * ดึงรายชื่อนักเรียนในห้องที่ระบุ (เฉพาะที่ "กำลังเรียน" อยู่) พร้อมระบุว่าคนไหนลงทะเบียนในวิชานี้แล้วบ้าง เรียงตามเลขที่
 * ใช้ตอนครูจะ "เลือกนักเรียนจากฐานข้อมูล" เข้าวิชาที่ตัวเองสอน
 */
function getStudentsByClassRoom(classRoom, subjectId) {
  requireSubjectAccess_(subjectId);

  const enrolledIds = {};
  getSheetData_('Enrollment')
    .filter(e => e.SubjectID === subjectId)
    .forEach(e => enrolledIds[e.StudentID] = true);

  return getSheetData_('Students')
    .filter(s => String(s.ClassRoom || '').trim() === String(classRoom).trim())
    .filter(s => (s.Status || STUDENT_STATUS_ACTIVE_) === STUDENT_STATUS_ACTIVE_)
    .map(s => ({
      studentId: s.StudentID,
      studentCode: s.StudentCode,
      name: s.Name,
      classRoom: s.ClassRoom,
      rollNumber: s.RollNumber,
      alreadyEnrolled: !!enrolledIds[s.StudentID]
    }))
    .sort(sortByRollNumber_);
}

/**
 * ลงทะเบียนนักเรียนที่เลือกไว้ (หลายคน) เข้าวิชานี้ทีเดียว
 * studentIds = [studentId, ...]
 */
function enrollSelectedStudents(subjectId, studentIds) {
  requireSubjectAccess_(subjectId);
  if (!studentIds || !studentIds.length) throw new Error('กรุณาเลือกนักเรียนอย่างน้อย 1 คน');

  studentIds.forEach(id => enrollStudent_(id, subjectId));

  return getStudentsInSubject(subjectId);
}

/**
 * เพิ่มนักเรียนใหม่เข้าฐานข้อมูลกลาง (ยังไม่ลงทะเบียนวิชาใดๆ) - สถานะเริ่มต้นเป็น "กำลังเรียน" เสมอ
 * form = { studentCode, name, classRoom, rollNumber }
 */
function addStudentCentral(form) {
  requireTeacher_();
  if (!form || !form.studentCode || !form.name) throw new Error('กรุณาระบุรหัสนักเรียนและชื่อ');

  const code = String(form.studentCode).trim();
  const existing = getSheetData_('Students').find(s => String(s.StudentCode) === code);
  if (existing) throw new Error('มีรหัสนักเรียนนี้อยู่ในระบบแล้ว (' + existing.Name + ')');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const studentId = 'ST' + new Date().getTime() + Math.floor(Math.random() * 1000);
  const accessCode = generateAccessCode_();
  sheet.appendRow([
    studentId, code, form.name.trim(), (form.classRoom || '').trim(),
    (form.rollNumber || '').toString().trim(), accessCode, STUDENT_STATUS_ACTIVE_
  ]);

  return getAllStudentsCentral();
}

/**
 * แก้ไขข้อมูลนักเรียนในฐานข้อมูลกลาง (รวมถึงสถานะ)
 * form = { studentId, studentCode, name, classRoom, rollNumber, status }
 */
function updateStudentCentral(form) {
  requireTeacher_();
  if (!form || !form.studentId) throw new Error('ไม่พบรหัสอ้างอิงนักเรียน');
  if (!form.studentCode || !form.name) throw new Error('กรุณาระบุรหัสนักเรียนและชื่อ');

  const status = STUDENT_STATUSES_.indexOf(form.status) !== -1 ? form.status : STUDENT_STATUS_ACTIVE_;

  const students = getSheetDataWithRow_('Students');
  const target = students.find(s => s.StudentID === form.studentId);
  if (!target) throw new Error('ไม่พบนักเรียนนี้ในระบบ');

  const code = String(form.studentCode).trim();
  const codeTaken = students.some(s => s.StudentID !== form.studentId && String(s.StudentCode) === code);
  if (codeTaken) throw new Error('มีนักเรียนอื่นใช้รหัสนี้อยู่แล้ว');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  sheet.getRange(target._row, 2, 1, 4).setValues([[
    code, form.name.trim(), (form.classRoom || '').trim(), (form.rollNumber || '').toString().trim()
  ]]);
  sheet.getRange(target._row, 7).setValue(status); // คอลัมน์ Status

  return getAllStudentsCentral();
}

/**
 * ลบนักเรียนออกจากฐานข้อมูลกลางทั้งหมด (cascade ลบการลงทะเบียนและคะแนนของนักเรียนคนนี้ในทุกวิชาด้วย)
 * ปกติแนะนำให้เปลี่ยนสถานะเป็น "จบการศึกษา"/"ย้ายออก" แทนการลบ เพื่อรักษาประวัติคะแนนไว้
 */
function deleteStudentCentral(studentId) {
  requireTeacher_();

  const students = getSheetDataWithRow_('Students');
  const target = students.find(s => s.StudentID === studentId);
  if (!target) throw new Error('ไม่พบนักเรียนนี้ในระบบ');

  // ลบการลงทะเบียนวิชาทั้งหมดของนักเรียนคนนี้ (ไล่ลบจากแถวท้ายไปแถวต้น กันเลขแถวเลื่อน)
  const enrollSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enrollment');
  const enrollRows = getSheetDataWithRow_('Enrollment').filter(e => e.StudentID === studentId);
  enrollRows.sort((a, b) => b._row - a._row).forEach(e => enrollSheet.deleteRow(e._row));

  // ลบคะแนนทั้งหมดของนักเรียนคนนี้
  const scoreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Scores');
  const scoreRows = getSheetDataWithRow_('Scores').filter(s => s.StudentID === studentId);
  scoreRows.sort((a, b) => b._row - a._row).forEach(s => scoreSheet.deleteRow(s._row));

  // ล้างแคชคะแนนของทุกวิชาที่นักเรียนคนนี้เคยลงทะเบียนไว้ กันแคชค้างข้อมูลที่ถูกลบไปแล้ว
  const cache = CacheService.getScriptCache();
  enrollRows.forEach(e => cache.remove('scoremap_' + e.SubjectID));

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  sheet.deleteRow(target._row);

  return getAllStudentsCentral();
}

/**
 * เลื่อนชั้นยกห้อง - เปลี่ยน "ห้องเรียนปัจจุบัน" ของนักเรียนทั้งห้องพร้อมกันทีเดียว (เฉพาะคนที่ยังกำลังเรียนอยู่)
 * หรือถ้าเป็นห้องที่จบการศึกษาแล้ว (เช่น ม.6) ให้ตั้ง graduate=true แทน จะเปลี่ยนสถานะเป็น "จบการศึกษา" ทั้งห้องโดยไม่ต้องมีห้องใหม่
 * ไม่กระทบห้อง/เลขที่ที่เคยแสดงในวิชาเก่าที่ลงทะเบียนไปแล้ว เพราะข้อมูลนั้น "แช่แข็ง" ไว้ในชีต Enrollment แยกต่างหาก
 * form = { oldRoom, newRoom, clearRollNumbers, graduate }
 */
function bulkPromoteClassRoom(form) {
  requireTeacher_();
  const oldRoom = (form && form.oldRoom || '').trim();
  const newRoom = (form && form.newRoom || '').trim();
  const graduate = !!(form && form.graduate);

  if (!oldRoom) throw new Error('กรุณาเลือกห้องเดิม');
  if (!graduate && !newRoom) throw new Error('กรุณาระบุห้องใหม่ หรือติ๊ก "จบการศึกษา" ถ้าเป็นห้องสุดท้าย (เช่น ม.6)');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const rows = getSheetDataWithRow_('Students').filter(s =>
    String(s.ClassRoom || '').trim() === oldRoom &&
    (s.Status || STUDENT_STATUS_ACTIVE_) === STUDENT_STATUS_ACTIVE_
  );

  if (!rows.length) throw new Error('ไม่พบนักเรียนที่ "กำลังเรียน" อยู่ในห้อง ' + oldRoom);

  rows.forEach(r => {
    if (graduate) {
      sheet.getRange(r._row, 7).setValue('จบการศึกษา'); // คอลัมน์ Status
    } else {
      sheet.getRange(r._row, 4).setValue(newRoom); // คอลัมน์ ClassRoom
      if (form.clearRollNumbers) sheet.getRange(r._row, 5).setValue(''); // คอลัมน์ RollNumber
    }
  });

  return { updated: rows.length, students: getAllStudentsCentral() };
}

/**
 * นำเข้านักเรียนหลายคนพร้อมกันจากไฟล์ Excel
 * rows = [{ rollNumber, studentCode, name, classRoom }, ...]
 * ถ้ารหัสนักเรียนซ้ำกับที่มีอยู่แล้ว จะอัปเดตชื่อ/ห้อง/เลขที่ให้ (ไม่แตะสถานะเดิม ไม่สร้างซ้ำ)
 * นักเรียนใหม่ที่นำเข้าจะได้สถานะ "กำลังเรียน" อัตโนมัติ
 */
function bulkImportStudentsCentral(rows) {
  requireTeacher_();
  if (!rows || !rows.length) throw new Error('ไม่มีข้อมูลให้นำเข้า');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const existing = getSheetDataWithRow_('Students');
  const existingMap = {};
  existing.forEach(s => existingMap[String(s.StudentCode)] = s);

  let imported = 0;
  let updated = 0;
  const errors = [];

  rows.forEach((row, idx) => {
    const rollNumber = (row.rollNumber || '').toString().trim();
    const code = (row.studentCode || '').toString().trim();
    const name = (row.name || '').toString().trim();
    const classRoom = (row.classRoom || '').toString().trim();

    if (!code || !name) {
      if (code || name || rollNumber) errors.push('แถวที่ ' + (idx + 2) + ': ข้อมูลไม่ครบ (ต้องมีรหัสนักเรียนและชื่อ)');
      return;
    }

    const found = existingMap[code];
    if (found) {
      sheet.getRange(found._row, 3, 1, 3).setValues([[name, classRoom, rollNumber]]);
      updated++;
    } else {
      const studentId = 'ST' + new Date().getTime() + Math.floor(Math.random() * 1000) + imported;
      const accessCode = generateAccessCode_();
      sheet.appendRow([studentId, code, name, classRoom, rollNumber, accessCode, STUDENT_STATUS_ACTIVE_]);
      existingMap[code] = { _row: sheet.getLastRow() };
      imported++;
    }
  });

  return { imported: imported, updated: updated, errors: errors, students: getAllStudentsCentral() };
}