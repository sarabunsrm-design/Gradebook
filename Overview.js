/**
 * ดึงข้อมูลภาพรวมแบบ "เบา" (ไม่คำนวณคะแนนเฉลี่ย) สำหรับ admin เท่านั้น
 * ใช้แสดงการ์ดสรุปจำนวน + รายชื่อวิชาทั้งหมด (สำหรับสร้างตัวกรองปี/เทอมและตารางเบื้องต้น)
 * ไม่เรียก getGradeSummary() เลย จึงเร็วแม้มีวิชาสะสมหลายร้อย/พันวิชาจากหลายปี
 * คะแนนเฉลี่ยของแต่ละวิชาจะดึงเพิ่มทีหลังเฉพาะวิชาที่อยู่ในตัวกรองที่เลือกอยู่ ผ่าน getSubjectAvgStats()
 */
function getOverviewMeta() {
  requireAdmin_();

  const subjects = getSheetData_('Subjects');
  const enrollments = getSheetData_('Enrollment');
  const teachers = getSheetData_('Teachers');
  const students = getSheetData_('Students');

  const enrollCountBySubject = {};
  enrollments.forEach(e => {
    enrollCountBySubject[e.SubjectID] = (enrollCountBySubject[e.SubjectID] || 0) + 1;
  });

  const teacherNameByEmail = {};
  teachers.forEach(t => teacherNameByEmail[String(t.Email).toLowerCase()] = t.Name);

  const subjectRows = subjects.map(s => ({
    subjectId: s.SubjectID,
    subjectName: s.SubjectName,
    classRoom: s.ClassRoom,
    academicYear: s.AcademicYear,
    semester: s.Semester,
    teacherEmail: s.TeacherEmail,
    teacherName: teacherNameByEmail[String(s.TeacherEmail).toLowerCase()] || s.TeacherEmail,
    studentCount: enrollCountBySubject[s.SubjectID] || 0
  }));

  subjectRows.sort((a, b) => String(a.subjectName).localeCompare(String(b.subjectName)));

  return {
    totalSubjects: subjects.length,
    totalTeachers: teachers.length,
    totalStudents: students.length,
    subjects: subjectRows
  };
}

/**
 * คำนวณคะแนนเฉลี่ยของ "เฉพาะวิชาที่ระบุ" เท่านั้น (เรียก getGradeSummary จริงๆ ซึ่งค่อนข้างหนัก)
 * เรียกใช้หลังจากกรองด้วยปี/เทอมแล้วที่ฝั่งหน้าเว็บ เพื่อไม่ต้องคำนวณทุกวิชาในระบบทุกครั้งที่เปิดหน้า
 * subjectIds = [subjectId, ...] คืนค่าเป็น object { subjectId: avgTotal หรือ null }
 */
function getSubjectAvgStats(subjectIds) {
  requireAdmin_();
  const result = {};
  if (!subjectIds || !subjectIds.length) return result;

  subjectIds.forEach(id => {
    if (!id) return;
    try {
      const summary = getGradeSummary(id); // admin เข้าถึงได้ทุกวิชาอยู่แล้ว (requireSubjectAccess_ อนุญาต)
      if (summary.rows.length) {
        const sum = summary.rows.reduce((acc, r) => acc + r.total, 0);
        result[id] = Math.round((sum / summary.rows.length) * 100) / 100;
      } else {
        result[id] = null;
      }
    } catch (err) {
      result[id] = null; // เช่น ยังไม่มีรายการคะแนนในวิชานี้เลย
    }
  });

  return result;
}