/**
 * ดึงรายชื่อครูทั้งหมดในระบบ (เฉพาะ admin เท่านั้น)
 */
function getAllTeachers() {
  requireAdmin_();
  return getSheetData_('Teachers').map(t => ({
    teacherId: t.TeacherID,
    name: t.Name,
    email: t.Email,
    role: t.Role
  }));
}

/**
 * เพิ่มครูใหม่เข้าระบบ (เฉพาะ admin เท่านั้น)
 * form = { name, email, role }  (role: 'teacher' หรือ 'admin')
 */
function addTeacher(form) {
  requireAdmin_();
  if (!form || !form.name || !form.email) throw new Error('กรุณาระบุชื่อและอีเมลให้ครบ');

  const email = String(form.email).trim().toLowerCase();
  const existing = getSheetData_('Teachers').find(t => String(t.Email).toLowerCase() === email);
  if (existing) throw new Error('มีอีเมลนี้อยู่ในระบบแล้ว');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  const id = 'T' + new Date().getTime();
  sheet.appendRow([id, form.name.trim(), form.email.trim(), form.role === 'admin' ? 'admin' : 'teacher']);

  return getAllTeachers();
}

/**
 * แก้ไขข้อมูลครู (เฉพาะ admin เท่านั้น)
 * form = { teacherId, name, email, role }
 */
function updateTeacher(form) {
  requireAdmin_();
  if (!form || !form.teacherId) throw new Error('ไม่พบรหัสครู');
  if (!form.name || !form.email) throw new Error('กรุณาระบุชื่อและอีเมลให้ครบ');

  const teachers = getSheetDataWithRow_('Teachers');
  const target = teachers.find(t => t.TeacherID === form.teacherId);
  if (!target) throw new Error('ไม่พบครูนี้ในระบบ');

  const newEmail = String(form.email).trim().toLowerCase();
  const emailTaken = teachers.some(t => t.TeacherID !== form.teacherId && String(t.Email).toLowerCase() === newEmail);
  if (emailTaken) throw new Error('มีครูอื่นใช้อีเมลนี้อยู่แล้ว');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  sheet.getRange(target._row, 2, 1, 3).setValues([[
    form.name.trim(), form.email.trim(), form.role === 'admin' ? 'admin' : 'teacher'
  ]]);

  return getAllTeachers();
}

/**
 * ลบครูออกจากระบบ (เฉพาะ admin เท่านั้น) - ห้ามลบบัญชีตัวเอง กันการล็อกตัวเองออกโดยไม่ตั้งใจ
 */
function deleteTeacher(teacherId) {
  const admin = requireAdmin_();

  const teachers = getSheetDataWithRow_('Teachers');
  const target = teachers.find(t => t.TeacherID === teacherId);
  if (!target) throw new Error('ไม่พบครูนี้ในระบบ');

  if (String(target.Email).toLowerCase() === String(admin.Email).toLowerCase()) {
    throw new Error('ไม่สามารถลบบัญชีของตัวเองได้');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  sheet.deleteRow(target._row);

  return getAllTeachers();
}

/**
 * นำเข้าครูหลายคนพร้อมกันจากไฟล์ Excel (แปลงเป็น array of object แล้วจากฝั่ง client ด้วย SheetJS)
 * rows = [{ name, email, role }, ...]
 * ถ้าอีเมลซ้ำกับที่มีอยู่แล้ว จะข้าม (นับเป็น skipped) ไม่เขียนทับข้อมูลเดิม
 */
function bulkImportTeachers(rows) {
  requireAdmin_();
  if (!rows || !rows.length) throw new Error('ไม่มีข้อมูลให้นำเข้า');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  const existing = getSheetData_('Teachers');
  const existingEmails = {};
  existing.forEach(t => existingEmails[String(t.Email).toLowerCase()] = true);

  let imported = 0;
  let skipped = 0;
  const errors = [];

  rows.forEach((row, idx) => {
    const name = (row.name || '').toString().trim();
    const email = (row.email || '').toString().trim();
    const role = (row.role || '').toString().trim().toLowerCase() === 'admin' ? 'admin' : 'teacher';

    if (!name || !email) {
      if (name || email) errors.push('แถวที่ ' + (idx + 2) + ': ข้อมูลไม่ครบ (ต้องมีชื่อและอีเมล)');
      return;
    }

    const emailKey = email.toLowerCase();
    if (existingEmails[emailKey]) {
      skipped++;
      return;
    }

    const id = 'T' + new Date().getTime() + Math.floor(Math.random() * 1000);
    sheet.appendRow([id, name, email, role]);
    existingEmails[emailKey] = true;
    imported++;
  });

  return { imported: imported, skipped: skipped, errors: errors, teachers: getAllTeachers() };
}

/**
 * ตรวจสอบว่าครูทุกคนที่ลงทะเบียนในระบบ มีสิทธิ์ "แก้ไข" สเปรดชีตต้นทางจริงหรือไม่ (เฉพาะ admin เท่านั้น)
 * สำคัญมาก: เพราะระบบ deploy แบบ "Execute as: User accessing the web app" ทุกคนต้องมีสิทธิ์เข้าถึงสเปรดชีตนี้เอง
 * ไม่งั้นจะเจอ error ตอนใช้งานจริง แม้จะมีชื่ออยู่ในชีต Teachers แล้วก็ตาม
 *
 * หมายเหตุ: ถ้าโรงเรียนแชร์สเปรดชีตแบบ "ทั้งโดเมนเข้าถึงได้" (ไม่ใช่แชร์รายบุคคล) getEditors() อาจไม่เห็นรายชื่อ
 * ครบทุกคน ทั้งที่จริงๆ เข้าถึงได้แล้ว กรณีนี้ควรตรวจสอบการตั้งค่าแชร์ด้วยตาเองเพิ่มเติม
 */
function checkTeacherAccess() {
  requireAdmin_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const editorEmails = ss.getEditors().map(u => u.getEmail().toLowerCase());
  const ownerEmail = ss.getOwner() ? ss.getOwner().getEmail().toLowerCase() : '';

  const teachers = getAllTeachers();
  const missing = teachers.filter(t => {
    const email = String(t.email).toLowerCase();
    return email !== ownerEmail && editorEmails.indexOf(email) === -1;
  });

  return {
    totalTeachers: teachers.length,
    totalEditors: editorEmails.length,
    missing: missing.map(t => ({ name: t.name, email: t.email }))
  };
}

/**
 * เพิ่มสิทธิ์ "แก้ไข" สเปรดชีตให้ครูที่ระบุ (เฉพาะ admin เท่านั้น)
 * emails = [email1, email2, ...] โดยปกติคือรายชื่อที่ checkTeacherAccess() รายงานว่าขาดสิทธิ์
 */
function grantAccessToTeachers(emails) {
  requireAdmin_();
  if (!emails || !emails.length) throw new Error('ไม่มีรายชื่อให้เพิ่มสิทธิ์');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = { granted: [], failed: [] };

  emails.forEach(email => {
    try {
      ss.addEditor(email);
      results.granted.push(email);
    } catch (err) {
      results.failed.push({ email: email, reason: err.message });
    }
  });

  return results;
}