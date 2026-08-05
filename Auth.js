/**
 * ค้นหาครูจากอีเมล คืนค่า null ถ้าไม่พบ (แปลว่าไม่มีสิทธิ์เข้าระบบ)
 */
function getTeacherByEmail(email) {
  if (!email) return null;
  const data = getSheetData_('Teachers');
  const found = data.find(r => String(r.Email).toLowerCase() === String(email).toLowerCase());
  return found || null;
}

function isAdminEmail_(email) {
  const t = getTeacherByEmail(email);
  return !!t && t.Role === 'admin';
}

/**
 * ตรวจสอบว่าอีเมลปัจจุบันเป็น admin เท่านั้น ถ้าไม่ใช่จะ throw error
 * ใช้เรียกในทุกฟังก์ชันที่จัดการรายชื่อครู (เพิ่ม/แก้ไข/ลบ)
 */
function requireAdmin_() {
  const teacher = requireTeacher_();
  if (teacher.Role !== 'admin') {
    throw new Error('เฉพาะผู้ดูแลระบบ (admin) เท่านั้นที่จัดการรายชื่อครูได้');
  }
  return teacher;
}

/**
 * ตรวจสอบว่าอีเมลปัจจุบันเป็นครูที่ได้รับอนุญาต ถ้าไม่ใช่จะ throw error
 * ใช้เรียกซ้ำในทุกฟังก์ชันที่แก้ไขข้อมูล เพื่อกันการเรียก API ตรงๆ โดยไม่ผ่านหน้าเว็บ
 */
function requireTeacher_() {
  const email = Session.getActiveUser().getEmail();
  const teacher = getTeacherByEmail(email);
  if (!teacher) throw new Error('ไม่มีสิทธิ์เข้าถึง: อีเมลนี้ไม่ได้ถูกเพิ่มเป็นครูในระบบ');
  return teacher;
}

/**
 * อ่านข้อมูลทั้งชีต แปลงเป็น array ของ object โดยใช้แถวแรกเป็นชื่อคอลัมน์
 */
function getSheetData_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

/**
 * เหมือน getSheetData_ แต่แถมเลขแถวจริงในชีต (obj._row) ไว้ใช้ตอนแก้ไข/ลบแถวนั้น
 */
function getSheetDataWithRow_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map((row, i) => {
    const obj = {};
    headers.forEach((h, j) => obj[h] = row[j]);
    obj._row = i + 2; // แถวที่ 1 คือ header จึงเริ่มข้อมูลจริงที่แถว 2
    return obj;
  });
}