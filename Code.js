/**
 * จุดเริ่มต้นของ Web App
 * URL รูปแบบ:
 *   .../exec                       -> หน้า Login กลาง (เลือกโหมด ครู/ผู้ดูแลระบบ หรือ นักเรียน)
 *   .../exec?page=login            -> หน้า Login กลาง (เหมือนกัน)
 *   .../exec?page=teacher          -> หน้าครู (ต้อง login ด้วย Google และมีชื่ออยู่ในชีต Teachers)
 *   .../exec?page=teachers         -> หน้าจัดการรายชื่อครู (เฉพาะ admin)
 *   .../exec?page=students         -> หน้าจัดการฐานข้อมูลนักเรียนกลาง (ครูทุกคนใช้ร่วมกัน)
 *   .../exec?page=settings         -> หน้าตั้งค่าระบบ (ชื่อ/โลโก้โรงเรียน) (เฉพาะ admin)
 *   .../exec?page=overview         -> หน้าภาพรวมทุกวิชา (เฉพาะ admin)
 *   .../exec?page=subject&id=xxx   -> หน้าจัดการวิชา (นักเรียน + รายการคะแนน)
 *   .../exec?page=student          -> หน้านักเรียน (กรอกรหัสดูคะแนน)
 */
function doGet(e) {
  try {
    return routeRequest_(e);
  } catch (err) {
    const tmpl = HtmlService.createTemplateFromFile('ErrorPage');
    tmpl.message = err && err.message ? err.message : String(err);
    tmpl.stack = err && err.stack ? err.stack : '';
    try {
      tmpl.webAppUrl = ScriptApp.getService().getUrl();
    } catch (_) {
      tmpl.webAppUrl = '';
    }
    return tmpl.evaluate().setTitle('เกิดข้อผิดพลาด');
  }
}

function routeRequest_(e) {
  const page = (e && e.parameter && e.parameter.page) || 'login';
  const webAppUrl = ScriptApp.getService().getUrl();

  if (page === 'login') {
    const email = Session.getActiveUser().getEmail();
    const teacher = getTeacherByEmail(email);
    const tmpl = HtmlService.createTemplateFromFile('Login');
    tmpl.webAppUrl = webAppUrl;
    tmpl.teacherName = teacher ? teacher.Name : '';
    return tmpl.evaluate()
      .setTitle('เข้าสู่ระบบ - ระบบบันทึกคะแนน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'student') {
    const tmpl = HtmlService.createTemplateFromFile('StudentLogin');
    tmpl.webAppUrl = webAppUrl;
    return tmpl.evaluate()
      .setTitle('ระบบดูคะแนนนักเรียน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // ----- โหมดครู (ทุกหน้าจากนี้ต้อง login ด้วย Google และมีชื่ออยู่ในชีต Teachers) -----
  const email = Session.getActiveUser().getEmail();
  const teacher = getTeacherByEmail(email);

  if (!teacher) {
    const tmpl = HtmlService.createTemplateFromFile('AccessDenied');
    tmpl.email = email || '(ไม่พบอีเมล - อาจยังไม่ได้อนุญาตสิทธิ์การเข้าถึง)';
    tmpl.webAppUrl = webAppUrl;
    return tmpl.evaluate().setTitle('ไม่มีสิทธิ์เข้าถึง');
  }

  if (page === 'teachers') {
    if (teacher.Role !== 'admin') {
      const tmpl = HtmlService.createTemplateFromFile('AccessDenied');
      tmpl.email = teacher.Email + ' (เฉพาะผู้ดูแลระบบเท่านั้นที่เข้าหน้านี้ได้)';
      tmpl.webAppUrl = webAppUrl;
      return tmpl.evaluate().setTitle('ไม่มีสิทธิ์เข้าถึง');
    }

    const tmpl = HtmlService.createTemplateFromFile('TeachersAdmin');
    tmpl.webAppUrl = webAppUrl;
    return tmpl.evaluate()
      .setTitle('จัดการรายชื่อครู')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'students') {
    const tmpl = HtmlService.createTemplateFromFile('StudentsAdmin');
    tmpl.webAppUrl = webAppUrl;
    return tmpl.evaluate()
      .setTitle('จัดการฐานข้อมูลนักเรียน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'settings') {
    if (teacher.Role !== 'admin') {
      const tmpl = HtmlService.createTemplateFromFile('AccessDenied');
      tmpl.email = teacher.Email + ' (เฉพาะผู้ดูแลระบบเท่านั้นที่เข้าหน้านี้ได้)';
      tmpl.webAppUrl = webAppUrl;
      return tmpl.evaluate().setTitle('ไม่มีสิทธิ์เข้าถึง');
    }

    const tmpl = HtmlService.createTemplateFromFile('SettingsAdmin');
    tmpl.webAppUrl = webAppUrl;
    return tmpl.evaluate()
      .setTitle('ตั้งค่าระบบ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'overview') {
    if (teacher.Role !== 'admin') {
      const tmpl = HtmlService.createTemplateFromFile('AccessDenied');
      tmpl.email = teacher.Email + ' (เฉพาะผู้ดูแลระบบเท่านั้นที่เข้าหน้านี้ได้)';
      tmpl.webAppUrl = webAppUrl;
      return tmpl.evaluate().setTitle('ไม่มีสิทธิ์เข้าถึง');
    }

    const tmpl = HtmlService.createTemplateFromFile('OverviewAdmin');
    tmpl.webAppUrl = webAppUrl;
    return tmpl.evaluate()
      .setTitle('ภาพรวมระบบ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'subject') {
    const subjectId = e.parameter.id;
    const subject = getSubjectById_(subjectId);
    const hasAccess = subject && (teacher.Role === 'admin' ||
      String(subject.TeacherEmail).toLowerCase() === String(teacher.Email).toLowerCase());

    if (!hasAccess) {
      const tmpl = HtmlService.createTemplateFromFile('AccessDenied');
      tmpl.email = teacher.Email + ' (ไม่พบวิชานี้ หรือไม่มีสิทธิ์เข้าถึง)';
      tmpl.webAppUrl = webAppUrl;
      return tmpl.evaluate().setTitle('ไม่มีสิทธิ์เข้าถึง');
    }

    const tmpl = HtmlService.createTemplateFromFile('SubjectDetail');
    tmpl.subjectId = subject.SubjectID;
    tmpl.subjectName = subject.SubjectName;
    tmpl.webAppUrl = webAppUrl;
    return tmpl.evaluate()
      .setTitle('จัดการวิชา: ' + subject.SubjectName)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const tmpl = HtmlService.createTemplateFromFile('Dashboard');
  tmpl.teacherName = teacher.Name;
  tmpl.teacherEmail = teacher.Email;
  tmpl.isAdmin = teacher.Role === 'admin';
  tmpl.webAppUrl = webAppUrl;

  return tmpl.evaluate()
    .setTitle('ระบบบันทึกคะแนน - หน้าครู')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ใช้ include ไฟล์ HTML/CSS/JS ย่อยเข้าไปในเทมเพลตหลัก (ไฟล์แบบ static ไม่มี scriptlet)
 * วิธีใช้ใน HTML: <?!= include('Stylesheet'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * เหมือน include() แต่ประมวลผล scriptlet (<?= ?> / <? ?>) ภายในไฟล์ที่ include ด้วย
 * ใช้กับไฟล์ที่ต้องเรียกฟังก์ชัน server ของตัวเอง เช่น BrandHeader ที่ต้องดึงชื่อ/โลโก้โรงเรียน
 * วิธีใช้ใน HTML: <?!= includeTemplate('BrandHeader'); ?>
 */
function includeTemplate(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}