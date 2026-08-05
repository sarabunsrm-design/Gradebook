/**
 * ดึงชื่อครูที่ปรึกษาของห้อง+ปีการศึกษา+เทอมที่ระบุ (รองรับสูงสุด 3 คน)
 * คืนค่าเป็น array ของชื่อ (ตัดชื่อว่างออก) - คืน array ว่างถ้ายังไม่เคยตั้ง
 */
function getClassAdvisor(classRoom, academicYear, semester) {
  requireTeacher_();
  const rows = getSheetData_('ClassAdvisors');
  const found = rows.find(r =>
    String(r.ClassRoom || '').trim() === String(classRoom || '').trim() &&
    String(r.AcademicYear || '') === String(academicYear || '') &&
    String(r.Semester || '') === String(semester || '')
  );
  const raw = found ? (found.AdvisorName || '') : '';
  return raw ? String(raw).split('|').map(s => s.trim()).filter(Boolean) : [];
}

/**
 * ตั้ง/แก้ไขชื่อครูที่ปรึกษาของห้อง+ปีการศึกษา+เทอม (ใช้ร่วมกันได้ทุกวิชาในห้องเดียวกัน)
 * advisorNames = array ของชื่อ สูงสุด 3 คน (ชื่อว่างจะถูกตัดออก)
 * เก็บในชีตเป็นข้อความเดียวคั่นด้วย "|" เพื่อไม่ต้องแก้โครงสร้างคอลัมน์เดิม
 */
function setClassAdvisor(classRoom, academicYear, semester, advisorNames) {
  requireTeacher_();
  if (!classRoom) throw new Error('ไม่พบห้องเรียน');

  const names = (advisorNames || []).map(n => (n || '').trim()).filter(Boolean).slice(0, 3);
  const joined = names.join('|');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ClassAdvisors');
  const rows = getSheetDataWithRow_('ClassAdvisors');
  const found = rows.find(r =>
    String(r.ClassRoom || '').trim() === String(classRoom || '').trim() &&
    String(r.AcademicYear || '') === String(academicYear || '') &&
    String(r.Semester || '') === String(semester || '')
  );

  if (found) {
    sheet.getRange(found._row, 4).setValue(joined);
  } else {
    sheet.appendRow([classRoom, academicYear || '', semester || '', joined]);
  }

  return names;
}