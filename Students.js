/**
 * ดึงรายชื่อนักเรียนที่ลงทะเบียนในวิชานี้ (เรียงตามเลขที่)
 * ห้อง/เลขที่ที่แสดง ใช้ค่า "ณ ตอนลงทะเบียน" (เก็บไว้ในชีต Enrollment) เป็นหลักเสมอ
 * เพื่อไม่ให้เปลี่ยนไปตามการเลื่อนชั้น/ย้ายห้องของนักเรียนในภายหลัง (ข้อมูลวิชาเก่าจะคงที่ตลอดไป)
 * ถ้าเป็นข้อมูลเก่าก่อนมีฟีเจอร์นี้ (ยังไม่มี snapshot) จะ fallback ไปใช้ค่าปัจจุบันแทน
 */
function getStudentsInSubject(subjectId) {
  requireSubjectAccess_(subjectId);

  const enrollments = getSheetData_('Enrollment').filter(e => e.SubjectID === subjectId);
  const students = getSheetData_('Students');
  const studentMap = {};
  students.forEach(s => studentMap[s.StudentID] = s);

  return enrollments
    .map(e => {
      const s = studentMap[e.StudentID];
      if (!s) return null;
      return {
        enrollId: e.EnrollID,
        studentId: s.StudentID,
        studentCode: s.StudentCode,
        name: s.Name,
        classRoom: e.ClassRoom || s.ClassRoom,
        rollNumber: e.RollNumber || s.RollNumber,
        accessCode: s.AccessCode
      };
    })
    .filter(Boolean)
    .sort(sortByRollNumber_);
}

/**
 * นำนักเรียนออกจากวิชา (ยกเลิกการลงทะเบียนเท่านั้น ไม่ลบข้อมูลนักเรียนออกจากระบบกลาง)
 */
function removeStudentFromSubject(enrollId, subjectId) {
  requireSubjectAccess_(subjectId);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enrollment');
  const rows = getSheetDataWithRow_('Enrollment');
  const target = rows.find(r => r.EnrollID === enrollId);
  if (target) sheet.deleteRow(target._row);

  return getStudentsInSubject(subjectId);
}

// ---------- Helper ฟังก์ชันภายใน (ใช้ร่วมกับ StudentsCentral.gs) ----------

/**
 * ลงทะเบียนนักเรียนเข้าวิชา (ข้ามถ้าลงทะเบียนซ้ำอยู่แล้ว)
 * บันทึกห้อง/เลขที่ "ปัจจุบัน" ของนักเรียน ณ ตอนลงทะเบียนนี้ ติดไว้กับการลงทะเบียนแบบถาวร
 * (คอลัมน์: EnrollID, StudentID, SubjectID, GradeOverride, ClassRoom, RollNumber)
 */
function enrollStudent_(studentId, subjectId) {
  const already = getSheetData_('Enrollment')
    .some(e => e.StudentID === studentId && e.SubjectID === subjectId);
  if (already) return;

  const student = getSheetData_('Students').find(s => s.StudentID === studentId);
  const classRoomSnapshot = student ? (student.ClassRoom || '') : '';
  const rollNumberSnapshot = student ? (student.RollNumber || '') : '';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enrollment');
  const enrollId = 'E' + new Date().getTime() + Math.floor(Math.random() * 1000);
  sheet.appendRow([enrollId, studentId, subjectId, '', classRoomSnapshot, rollNumberSnapshot]);
}

/**
 * สุ่มรหัส 6 หลักสำหรับให้นักเรียนใช้ดูคะแนน (ตัดตัวอักษรที่สับสนง่ายออก เช่น 0,O,1,I)
 */
function generateAccessCode_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 6; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}