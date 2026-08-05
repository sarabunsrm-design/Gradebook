/**
 * ตรวจสอบการเข้าสู่ระบบของนักเรียน แล้วคืนคะแนน/เกรดของทุกวิชาที่ลงทะเบียนไว้
 * ระบบล็อกอินของนักเรียนใช้ "รหัสประจำตัวนักเรียน" เป็นทั้ง Username และ Password
 * (เพื่อให้จำง่าย ไม่ต้องมีรหัสแยกอีกชุด) — Username และ Password ที่กรอกมาจึงต้องตรงกัน
 * และต้องตรงกับรหัสประจำตัวของนักเรียนคนนั้นในระบบด้วย
 * หมายเหตุ: ฟังก์ชันนี้ไม่ผ่าน requireTeacher_/requireSubjectAccess_ เพราะนักเรียนไม่ได้ login ด้วย Google
 * ⚠️ ข้อควรระวังด้านความปลอดภัย: เพราะ username/password เป็นค่าเดียวกัน (รหัสประจำตัว) ซึ่งมักไม่ใช่ความลับ
 * (เช่น เพื่อนอาจรู้รหัสประจำตัวกันเอง) ใครก็ตามที่รู้รหัสประจำตัวของนักเรียนคนใด จะดูคะแนนคนนั้นได้
 * เหมาะกับกรณีต้องการความสะดวกมากกว่าความปลอดภัยสูงสุด หากต้องการความปลอดภัยกว่านี้ แนะนำกลับไปใช้
 * ระบบรหัสดูคะแนน (AccessCode) แยกต่างหากแบบเดิม
 */
function getStudentPortalData(username, password) {
  if (!username || !password) {
    throw new Error('กรุณากรอกรหัสประจำตัวนักเรียนให้ครบทั้งสองช่อง');
  }

  const students = getSheetData_('Students');
  const student = students.find(s => String(s.StudentCode).trim() === String(username).trim());

  if (!student || String(student.StudentCode).trim() !== String(password).trim()) {
    throw new Error('รหัสประจำตัวนักเรียนไม่ถูกต้อง กรุณาตรวจสอบกับครูผู้สอน');
  }

  const enrollments = getSheetData_('Enrollment').filter(e => e.StudentID === student.StudentID);
  const subjects = getSheetData_('Subjects');
  const subjectMap = {};
  subjects.forEach(s => subjectMap[s.SubjectID] = s);

  const allItems = getSheetData_('ScoreItems');
  const allScores = getSheetData_('Scores');
  const scale = getGradeScale_(); // มาจาก Scores.gs

  const subjectsResult = enrollments.map(e => {
    const subject = subjectMap[e.SubjectID];
    if (!subject) return null;

    const items = allItems.filter(i => i.SubjectID === e.SubjectID);
    const periodScores = {};
    PERIODS_.forEach(p => periodScores[p] = 0); // PERIODS_ มาจาก ScoreItems.gs

    const itemDetails = items.map(item => {
      const scoreRow = allScores.find(sc => sc.StudentID === student.StudentID && sc.ItemID === item.ItemID);
      const raw = scoreRow ? scoreRow.RawScore : '';
      const rawNum = (raw === undefined || raw === null || raw === '') ? 0 : Number(raw);
      const maxScore = Number(item.MaxScore) || 1;
      const weight = Number(item.Weight) || 0;
      periodScores[item.Period] = (periodScores[item.Period] || 0) + (rawNum / maxScore) * weight;

      return {
        itemId: item.ItemID,
        itemName: item.ItemName,
        period: item.Period,
        maxScore: item.MaxScore,
        rawScore: (raw === undefined || raw === null) ? '' : raw
      };
    });

    let total = 0;
    Object.keys(periodScores).forEach(p => {
      periodScores[p] = Math.round(periodScores[p]);
      total += periodScores[p];
    });

    return {
      subjectId: subject.SubjectID,
      subjectName: subject.SubjectName,
      classRoom: subject.ClassRoom,
      periods: PERIODS_,
      items: itemDetails,
      periodScores: periodScores,
      total: total,
      grade: (e.GradeOverride || '') || calcGrade_(total, scale),
      hasItems: items.length > 0
    };
  }).filter(Boolean);

  return {
    studentName: student.Name,
    studentCode: student.StudentCode,
    subjects: subjectsResult
  };
}