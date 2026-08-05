/**
 * ดึงค่าตั้งค่าของโรงเรียน (ชื่อ + โลโก้ + ที่อยู่ + เขตพื้นที่การศึกษา)
 * ฟังก์ชันนี้ "ไม่ต้องผ่านการตรวจสิทธิ์" เพราะต้องใช้แสดงผลในหน้า Login/StudentLogin
 * ซึ่งเป็นหน้าสาธารณะที่ยังไม่ทราบว่าใครเข้ามา (ข้อมูลชื่อ/โลโก้โรงเรียนไม่ใช่ข้อมูลอ่อนไหว)
 */
function getSettings() {
  const config = getSheetData_('Config');
  const nameRow = config.find(c => c.Key === 'SchoolName');
  const logoRow = config.find(c => c.Key === 'SchoolLogo');
  const addressRow = config.find(c => c.Key === 'SchoolAddress');
  const areaRow = config.find(c => c.Key === 'EducationArea');

  return {
    schoolName: (nameRow && nameRow.Value) ? nameRow.Value : '',
    schoolLogo: (logoRow && logoRow.Value) ? logoRow.Value : '',
    schoolAddress: (addressRow && addressRow.Value) ? addressRow.Value : '',
    educationArea: (areaRow && areaRow.Value) ? areaRow.Value : ''
  };
}

/**
 * บันทึกชื่อโรงเรียน + โลโก้ + ที่อยู่ + เขตพื้นที่การศึกษา (เฉพาะ admin เท่านั้น)
 * form = { schoolName, schoolLogo, schoolAddress, educationArea }
 * schoolLogo เป็นได้ทั้ง URL ของรูป (https://...) หรือ Data URI แบบ base64 (data:image/png;base64,...)
 */
function saveSettings(form) {
  requireAdmin_();

  setConfigValue_('SchoolName', (form.schoolName || '').trim());
  setConfigValue_('SchoolLogo', (form.schoolLogo || '').trim());
  setConfigValue_('SchoolAddress', (form.schoolAddress || '').trim());
  setConfigValue_('EducationArea', (form.educationArea || '').trim());

  return getSettings();
}

/**
 * ตั้งค่า/อัปเดตค่าใน Config sheet (สร้างแถวใหม่ถ้ายังไม่มี key นี้)
 */
function setConfigValue_(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = getSheetDataWithRow_('Config');
  const row = data.find(r => r.Key === key);

  if (row) {
    sheet.getRange(row._row, 2).setValue(value);
  } else {
    sheet.appendRow([key, value]);
  }
}