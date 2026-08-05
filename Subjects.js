/**
 * ดึงรายวิชาที่ครูคนปัจจุบันดูแลอยู่ (admin จะเห็นวิชาของครูทุกคน)
 */
function getMySubjects() {
  const teacher = requireTeacher_();
  const email = teacher.Email;

  const subjects = getSheetDataWithRow_('Subjects');
  if (teacher.Role === 'admin') return subjects;
  return subjects.filter(s => String(s.TeacherEmail).toLowerCase() === String(email).toLowerCase());
}

/**
 * เพิ่มวิชาใหม่ - เจ้าของวิชาคืออีเมลของครูที่ล็อกอินอยู่เสมอ (กันการปลอมอีเมลจากฝั่ง client)
 * form = { subjectName, classRoom, academicYear, semester, subjectCode, learningArea, subjectType, hoursPerWeek, credits }
 */
function addSubject(form) {
  const teacher = requireTeacher_();
  if (!form || !form.subjectName) throw new Error('กรุณาระบุชื่อวิชา');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subjects');
  const id = 'S' + new Date().getTime();
  sheet.appendRow([
    id,
    form.subjectName,
    teacher.Email,
    form.classRoom || '',
    form.academicYear || '',
    form.semester || '',
    form.subjectCode || '',
    form.learningArea || '',
    form.subjectType || '',
    form.hoursPerWeek || '',
    form.credits || ''
  ]);

  return getMySubjects();
}

/**
 * แก้ไขวิชา - ต้องเป็นเจ้าของวิชาเดิม หรือเป็น admin เท่านั้น
 * form = { subjectId, subjectName, classRoom, academicYear, semester, subjectCode, learningArea, subjectType, hoursPerWeek, credits }
 */
function updateSubject(form) {
  const teacher = requireTeacher_();
  if (!form || !form.subjectId) throw new Error('ไม่พบรหัสวิชา');

  const subjects = getSheetDataWithRow_('Subjects');
  const target = subjects.find(s => s.SubjectID === form.subjectId);
  if (!target) throw new Error('ไม่พบวิชานี้ในระบบ');

  if (teacher.Role !== 'admin' &&
      String(target.TeacherEmail).toLowerCase() !== String(teacher.Email).toLowerCase()) {
    throw new Error('ไม่มีสิทธิ์แก้ไขวิชาของครูท่านอื่น');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subjects');
  // คอลัมน์: SubjectID(1) SubjectName(2) TeacherEmail(3) ClassRoom(4) AcademicYear(5) Semester(6)
  sheet.getRange(target._row, 2, 1, 4).setValues([[
    form.subjectName, target.TeacherEmail, form.classRoom || '', form.academicYear || ''
  ]]);
  sheet.getRange(target._row, 6).setValue(form.semester || '');
  // คอลัมน์: SubjectCode(7) LearningArea(8) SubjectType(9) HoursPerWeek(10) Credits(11)
  sheet.getRange(target._row, 7, 1, 5).setValues([[
    form.subjectCode || '', form.learningArea || '', form.subjectType || '',
    form.hoursPerWeek || '', form.credits || ''
  ]]);

  return getMySubjects();
}

/**
 * ลบวิชา - ต้องเป็นเจ้าของวิชาเดิม หรือเป็น admin เท่านั้น
 */
function deleteSubject(subjectId) {
  const teacher = requireTeacher_();

  const subjects = getSheetDataWithRow_('Subjects');
  const target = subjects.find(s => s.SubjectID === subjectId);
  if (!target) throw new Error('ไม่พบวิชานี้ในระบบ');

  if (teacher.Role !== 'admin' &&
      String(target.TeacherEmail).toLowerCase() !== String(teacher.Email).toLowerCase()) {
    throw new Error('ไม่มีสิทธิ์ลบวิชาของครูท่านอื่น');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subjects');
  sheet.deleteRow(target._row);

  return getMySubjects();
}

/**
 * ดึงข้อมูลวิชา 1 รายการจาก SubjectID (ไม่เช็คสิทธิ์ - ใช้ภายในเท่านั้น)
 */
function getSubjectById_(subjectId) {
  const subjects = getSheetData_('Subjects');
  return subjects.find(s => s.SubjectID === subjectId) || null;
}

/**
 * ตรวจสิทธิ์ว่าเป็นเจ้าของวิชานี้ หรือเป็น admin เท่านั้น ไม่เช่นนั้น throw error
 * คืนค่า { teacher, subject } ให้ฟังก์ชันที่เรียกใช้ต่อได้
 * ใช้เป็น guard ร่วมกันในทุกฟังก์ชันที่เกี่ยวกับนักเรียน/รายการคะแนนภายในวิชา
 */
function requireSubjectAccess_(subjectId) {
  const teacher = requireTeacher_();
  const subject = getSubjectById_(subjectId);
  if (!subject) throw new Error('ไม่พบวิชานี้ในระบบ');
  if (teacher.Role !== 'admin' &&
      String(subject.TeacherEmail).toLowerCase() !== String(teacher.Email).toLowerCase()) {
    throw new Error('ไม่มีสิทธิ์เข้าถึงวิชานี้');
  }
  return { teacher, subject };
}

/**
 * คัดลอกวิชาไปเป็นเทอม/ปีการศึกษาใหม่ - เอาเฉพาะ "โครงสร้างรายการคะแนน + น้ำหนัก" ไปด้วย
 * ไม่คัดลอกนักเรียนที่ลงทะเบียนหรือคะแนนเดิมมา (เริ่มต้นใหม่ทั้งหมดสำหรับเทอมใหม่)
 * form = { subjectName, classRoom, academicYear, semester }
 */
function duplicateSubject(originalSubjectId, form) {
  const access = requireSubjectAccess_(originalSubjectId); // ตรวจว่ามีสิทธิ์เข้าถึงวิชาต้นฉบับ
  if (!form || !form.subjectName) throw new Error('กรุณาระบุชื่อวิชา');

  const subjectSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subjects');
  const newId = 'S' + new Date().getTime() + Math.floor(Math.random() * 1000);

  // เจ้าของวิชาใหม่ = เจ้าของวิชาต้นฉบับเสมอ (แม้ admin จะเป็นคนกดคัดลอกให้ก็ตาม)
  // รายละเอียดวิชา (รหัส/กลุ่มสาระ/ประเภท/เวลาเรียน/หน่วยกิต) คัดลอกมาจากต้นฉบับด้วย เพราะมักไม่เปลี่ยนข้ามเทอม
  subjectSheet.appendRow([
    newId,
    form.subjectName,
    access.subject.TeacherEmail,
    form.classRoom || '',
    form.academicYear || '',
    form.semester || '',
    access.subject.SubjectCode || '',
    access.subject.LearningArea || '',
    access.subject.SubjectType || '',
    access.subject.HoursPerWeek || '',
    access.subject.Credits || ''
  ]);

  // คัดลอกรายการคะแนน + น้ำหนักทั้งหมดจากวิชาต้นฉบับ
  const itemSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ScoreItems');
  const originalItems = getSheetData_('ScoreItems').filter(i => i.SubjectID === originalSubjectId);
  originalItems.forEach((item, idx) => {
    const newItemId = 'I' + new Date().getTime() + Math.floor(Math.random() * 1000) + idx;
    itemSheet.appendRow([newItemId, newId, item.Period, item.ItemName, item.MaxScore, item.Weight]);
  });

  return getMySubjects();
}