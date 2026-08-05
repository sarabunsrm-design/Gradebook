/**
 * ตัวชี้วัดคุณลักษณะอันพึงประสงค์ 8 ข้อ ตามหลักสูตรแกนกลางฯ
 */
var CHARACTER_INDICATORS_ = [
  { key: 'Char1', label: '1.รักชาติ ศาสน์ กษัตริย์' },
  { key: 'Char2', label: '2.ซื่อสัตย์สุจริต' },
  { key: 'Char3', label: '3.มีวินัย' },
  { key: 'Char4', label: '4.ใฝ่เรียนรู้' },
  { key: 'Char5', label: '5.อยู่อย่างพอเพียง' },
  { key: 'Char6', label: '6.มุ่งมั่นในการทำงาน' },
  { key: 'Char7', label: '7.รักความเป็นไทย' },
  { key: 'Char8', label: '8.มีจิตสาธารณะ' }
];

/**
 * ตัวชี้วัดการอ่าน คิด วิเคราะห์และเขียน 5 ข้อ
 */
var READTHINK_INDICATORS_ = [
  { key: 'Read1', label: '1.การอ่าน' },
  { key: 'Read2', label: '2.การจับประเด็น' },
  { key: 'Read3', label: '3.การวิเคราะห์' },
  { key: 'Read4', label: '4.การประเมินค่า' },
  { key: 'Read5', label: '5.การเขียน' }
];

/**
 * ระดับผลการประเมิน: 3=ดีเยี่ยม, 2=ดี, 1=ผ่าน, 0=ไม่ผ่าน
 */
var RATING_LABEL_BY_NUMBER_ = { 3: 'ดีเยี่ยม', 2: 'ดี', 1: 'ผ่าน', 0: 'ไม่ผ่าน' };

function numOrBlank_(v) {
  return (v === undefined || v === null || v === '') ? '' : Number(v);
}

/**
 * ดึงข้อมูลตัวชี้วัดคุณลักษณะฯ/อ่านคิดฯ ของนักเรียนทุกคนในวิชานี้ สำหรับหน้ากรอกคะแนน
 */
function getCharacterAssessmentData(subjectId) {
  requireSubjectAccess_(subjectId);

  const students = getStudentsInSubject(subjectId);
  const enrollMap = {};
  getSheetData_('Enrollment').filter(e => e.SubjectID === subjectId).forEach(e => enrollMap[e.StudentID] = e);

  const rows = students.map(stu => {
    const e = enrollMap[stu.studentId] || {};
    const charValues = {};
    CHARACTER_INDICATORS_.forEach(ind => charValues[ind.key] = numOrBlank_(e[ind.key]));
    const readValues = {};
    READTHINK_INDICATORS_.forEach(ind => readValues[ind.key] = numOrBlank_(e[ind.key]));

    return {
      studentId: stu.studentId,
      studentCode: stu.studentCode,
      name: stu.name,
      rollNumber: stu.rollNumber,
      charValues: charValues,
      readValues: readValues,
      charOverall: e.DesiredCharacteristics || '',
      readOverall: e.ReadThinkWrite || ''
    };
  });

  return {
    students: rows,
    characterIndicators: CHARACTER_INDICATORS_,
    readThinkIndicators: READTHINK_INDICATORS_
  };
}

/**
 * คำนวณผลสรุปรวมด้วย "ฐานนิยม" (Mode) ตามหลักสูตรแกนกลางการศึกษาขั้นพื้นฐาน
 * ถ้ามีฐานนิยมมากกว่า 1 ค่า (จำนวนซ้ำเท่ากัน) ให้ใช้ค่าที่ต่ำกว่าเป็นเกณฑ์ตัดสิน
 */
function computeOverallRating_(values) {
  if (!values.length) return '';

  const freq = {};
  values.forEach(v => freq[v] = (freq[v] || 0) + 1);

  let maxFreq = 0;
  Object.keys(freq).forEach(k => { if (freq[k] > maxFreq) maxFreq = freq[k]; });

  const modes = Object.keys(freq).filter(k => freq[k] === maxFreq).map(Number);
  const chosen = Math.min.apply(null, modes);

  return RATING_LABEL_BY_NUMBER_[chosen];
}

/**
 * บันทึกตัวชี้วัดคุณลักษณะฯ/อ่านคิดฯ ของนักเรียนทุกคนทีเดียว
 * "สรุปรวม" ของแต่ละหมวดจะคำนวณอัตโนมัติด้วย "ฐานนิยม" (Mode) ของตัวชี้วัดที่กรอกไว้ ตามหลักสูตรแกนกลางฯ
 * ค่าที่คำนวณได้นี้จะถูกใช้แสดงในแท็บ "สรุปผล/เกรด" และรายงาน ปพ.5 ต่อไปด้วย
 * rows = [{ studentId, charValues:{Char1:3,...}, readValues:{Read1:3,...} }, ...]
 */
function saveCharacterAssessments(subjectId, rows) {
  requireSubjectAccess_(subjectId);
  if (!rows || !rows.length) throw new Error('ไม่มีข้อมูลให้บันทึก');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enrollment');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const enrollRows = getSheetDataWithRow_('Enrollment').filter(e => e.SubjectID === subjectId);
  const enrollMap = {};
  enrollRows.forEach(e => enrollMap[e.StudentID] = e);

  const charCol = headers.indexOf('DesiredCharacteristics') + 1;
  const readCol = headers.indexOf('ReadThinkWrite') + 1;

  rows.forEach(r => {
    const target = enrollMap[r.studentId];
    if (!target) return;

    CHARACTER_INDICATORS_.forEach(ind => {
      const col = headers.indexOf(ind.key) + 1;
      if (col > 0) sheet.getRange(target._row, col).setValue(r.charValues[ind.key] === '' ? '' : Number(r.charValues[ind.key]));
    });
    READTHINK_INDICATORS_.forEach(ind => {
      const col = headers.indexOf(ind.key) + 1;
      if (col > 0) sheet.getRange(target._row, col).setValue(r.readValues[ind.key] === '' ? '' : Number(r.readValues[ind.key]));
    });

    const charNums = CHARACTER_INDICATORS_
      .map(ind => r.charValues[ind.key])
      .filter(v => v !== '' && v !== undefined && v !== null)
      .map(Number);
    const readNums = READTHINK_INDICATORS_
      .map(ind => r.readValues[ind.key])
      .filter(v => v !== '' && v !== undefined && v !== null)
      .map(Number);

    const charOverall = computeOverallRating_(charNums);
    const readOverall = computeOverallRating_(readNums);

    if (charCol > 0) sheet.getRange(target._row, charCol).setValue(charOverall);
    if (readCol > 0) sheet.getRange(target._row, readCol).setValue(readOverall);
  });

  return getCharacterAssessmentData(subjectId);
}