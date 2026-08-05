// เกณฑ์เกรดเริ่มต้น (ใช้ถ้ายังไม่ได้ตั้งค่าเองในชีต Config คีย์ "GradeScale")
// เรียงจากคะแนนขั้นต่ำมากไปน้อย ระบบจะไล่เช็คจากบนลงล่าง เจอตัวแรกที่คะแนนถึงเกณฑ์จะใช้เกรดนั้น
var DEFAULT_GRADE_SCALE_ = [
  { min: 80, grade: '4' },
  { min: 75, grade: '3.5' },
  { min: 70, grade: '3' },
  { min: 65, grade: '2.5' },
  { min: 60, grade: '2' },
  { min: 55, grade: '1.5' },
  { min: 50, grade: '1' },
  { min: 0,  grade: '0' }
];

/**
 * ดึง "แผนที่คะแนนดิบ" เฉพาะของวิชานี้ (studentId -> itemId -> rawScore)
 * แคชไว้สั้นๆ (30 วินาที) แยกตามวิชา เพื่อไม่ต้องอ่านคะแนนทั้งโรงเรียนซ้ำทุกครั้งที่เปิด/สลับแท็บ
 * (เดิมฟังก์ชันนี้อ่านทั้งชีต Scores ทุกครั้ง ซึ่งรวมคะแนนของทุกวิชา/ทุกครู/ทุกปีที่เคยกรอกมา
 * ยิ่งใช้งานสะสมนาน ชีตยิ่งใหญ่ ยิ่งช้า - แคชนี้ช่วยได้เฉพาะการเปิดซ้ำในช่วงสั้นๆ เท่านั้น
 * ถ้าจะแก้ที่ต้นตอจริงๆ ต้องพิจารณาเก็บถาวรข้อมูลปีเก่าแยกออกไป)
 */
function getScoreMapForSubject_(subjectId, itemIds) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'scoremap_' + subjectId;

  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* แคชเสีย ก็อ่านใหม่ */ }
  }

  const scoreMap = {};
  getSheetData_('Scores').forEach(s => {
    if (!itemIds[s.ItemID]) return; // คะแนนของวิชาอื่น ไม่เกี่ยว
    if (!scoreMap[s.StudentID]) scoreMap[s.StudentID] = {};
    scoreMap[s.StudentID][s.ItemID] = s.RawScore;
  });

  try {
    cache.put(cacheKey, JSON.stringify(scoreMap), 30);
  } catch (e) {
    // ข้อมูลของวิชานี้ใหญ่เกิน 100KB (นักเรียน/รายการคะแนนเยอะมาก) ก็แค่ข้ามการแคช ไม่กระทบการทำงาน
  }

  return scoreMap;
}

/**
 * ดึงข้อมูลทั้งหมดที่ต้องใช้ในหน้ากรอกคะแนน: รายชื่อนักเรียน, รายการคะแนนทุกช่วง, คะแนนดิบที่เคยกรอกไว้
 */
function getScoreEntryData(subjectId) {
  requireSubjectAccess_(subjectId);

  const students = getStudentsInSubject(subjectId);
  const itemsData = getScoreItems(subjectId);

  const itemIds = {};
  itemsData.items.forEach(i => itemIds[i.ItemID] = true);

  const scoreMap = getScoreMapForSubject_(subjectId, itemIds);

  return {
    students: students,
    items: itemsData.items,
    periods: itemsData.periods,
    scores: scoreMap
  };
}

/**
 * บันทึกคะแนนดิบทีเดียวทั้งตาราง
 * scores = [{ studentId, itemId, rawScore }, ...]  (rawScore เป็น '' ได้ถ้าต้องการเว้นว่าง)
 * ตรวจสอบก่อนบันทึกทุกครั้งว่าคะแนนดิบต้องไม่ติดลบและไม่เกินคะแนนเต็มของรายการนั้น
 * ถ้าพบค่าไม่ถูกต้องแม้แต่รายการเดียว จะไม่บันทึกอะไรเลย (กันข้อมูลเสียครึ่งๆ กลางๆ)
 */
function saveScores(subjectId, scores) {
  requireSubjectAccess_(subjectId);
  if (!scores || !scores.length) throw new Error('ไม่มีข้อมูลคะแนนให้บันทึก');

  const items = getSheetData_('ScoreItems').filter(i => i.SubjectID === subjectId);
  const itemMap = {};
  items.forEach(i => itemMap[i.ItemID] = i);

  const errors = [];
  scores.forEach(sc => {
    const hasValue = sc.rawScore !== '' && sc.rawScore !== null && sc.rawScore !== undefined;
    if (!hasValue) return;

    const item = itemMap[sc.itemId];
    if (!item) return; // รายการนี้ไม่ได้อยู่ในวิชานี้ (ไม่ควรเกิดขึ้น แต่กันไว้)

    const maxScore = Number(item.MaxScore);
    const rawNum = Number(sc.rawScore);
    const EPSILON = 0.001; // กันปัญหา floating-point เช่น 20.00000000000002 ไม่ให้ถูกตัดสินว่าเกินคะแนนเต็ม

    if (isNaN(rawNum)) {
      errors.push(item.ItemName + ': ค่าที่กรอกไม่ใช่ตัวเลข');
    } else if (rawNum < -EPSILON) {
      errors.push(item.ItemName + ': คะแนนติดลบไม่ได้ (กรอก ' + rawNum + ')');
    } else if (rawNum > maxScore + EPSILON) {
      errors.push(item.ItemName + ': กรอก ' + rawNum + ' แต่คะแนนเต็มมีแค่ ' + maxScore);
    }
  });

  if (errors.length) {
    throw new Error('พบคะแนนที่ไม่ถูกต้อง กรุณาแก้ไขก่อนบันทึก:\n' + errors.join('\n'));
  }

  const email = Session.getActiveUser().getEmail();
  const now = new Date();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Scores');

  // อ่านข้อมูลทั้งบล็อกมาเป็น array ในหน่วยความจำครั้งเดียว (ถ้ามีข้อมูลอยู่แล้ว)
  const lastRow = sheet.getLastRow();
  const hasExistingData = lastRow >= 2;
  const fullRange = hasExistingData ? sheet.getRange(2, 1, lastRow - 1, 6) : null;
  const values = hasExistingData ? fullRange.getValues() : [];

  // ทำ index: "studentId_itemId" -> ตำแหน่งแถวใน array (เร็วกว่าไล่หาทีละรายการ)
  const indexMap = {};
  values.forEach((row, idx) => { indexMap[row[1] + '_' + row[2]] = idx; });

  const newRows = [];
  scores.forEach(sc => {
    const key = sc.studentId + '_' + sc.itemId;
    const hasValue = sc.rawScore !== '' && sc.rawScore !== null && sc.rawScore !== undefined;
    const rawValue = hasValue ? Number(sc.rawScore) : '';

    if (indexMap[key] !== undefined) {
      const idx = indexMap[key];
      values[idx][3] = rawValue; // คอลัมน์ RawScore
      values[idx][4] = now;      // Timestamp
      values[idx][5] = email;    // UpdatedBy
    } else if (hasValue) {
      const id = 'SC' + new Date().getTime() + Math.floor(Math.random() * 1000) + newRows.length;
      newRows.push([id, sc.studentId, sc.itemId, rawValue, now, email]);
    }
  });

  // เขียนข้อมูลเดิมที่แก้ไขแล้วกลับไปทีเดียวทั้งบล็อก (1 ครั้ง ไม่ว่าจะแก้กี่สิบกี่ร้อยรายการ แทนที่จะเขียนทีละแถว)
  if (hasExistingData) {
    fullRange.setValues(values);
  }

  // เพิ่มแถวใหม่ทั้งหมดทีเดียวเป็น batch เดียว แทนการ appendRow() วนลูป
  if (newRows.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
  }

  // ล้างแคชคะแนนของวิชานี้ทันที เพราะข้อมูลเปลี่ยนไปแล้ว (กันโหลดครั้งถัดไปเห็นข้อมูลเก่าค้างอยู่)
  CacheService.getScriptCache().remove('scoremap_' + subjectId);

  return getScoreEntryData(subjectId);
}

/**
 * เกรดพิเศษที่ครูตั้งเองได้ นอกเหนือจากเกรดที่คำนวณอัตโนมัติ
 * "" = ใช้ค่าที่คำนวณอัตโนมัติ, "ร" = ผลการเรียนไม่สมบูรณ์, "มส" = ไม่มีสิทธิ์สอบ
 */
var ALLOWED_GRADE_OVERRIDES_ = ['', 'ร', 'มส'];

/**
 * คำนวณคะแนนจริงตามน้ำหนัก + รวมคะแนนแต่ละช่วง + เกรด ของนักเรียนทุกคนในวิชา
 * ถ้าครูตั้งเกรดพิเศษ ("ร"/"มส") ไว้ให้คนไหน จะใช้เกรดพิเศษนั้นแสดงแทนเกรดที่คำนวณได้
 */
function getGradeSummary(subjectId) {
  requireSubjectAccess_(subjectId);

  const data = getScoreEntryData(subjectId);
  const scale = getGradeScale_();

  const overrideMap = {};
  const characterMap = {};
  const readThinkMap = {};
  getSheetData_('Enrollment')
    .filter(e => e.SubjectID === subjectId)
    .forEach(e => {
      overrideMap[e.StudentID] = e.GradeOverride || '';
      characterMap[e.StudentID] = e.DesiredCharacteristics || '';
      readThinkMap[e.StudentID] = e.ReadThinkWrite || '';
    });

  const rows = data.students.map(stu => {
    const periodScores = {};
    data.periods.forEach(p => periodScores[p] = 0);

    data.items.forEach(item => {
      const studentScores = data.scores[stu.studentId];
      const raw = studentScores ? studentScores[item.ItemID] : undefined;
      const rawNum = (raw === undefined || raw === null || raw === '') ? 0 : Number(raw);
      const maxScore = Number(item.MaxScore) || 1;
      const weight = Number(item.Weight) || 0;
      const earned = (rawNum / maxScore) * weight;
      periodScores[item.Period] = (periodScores[item.Period] || 0) + earned;
    });

    let total = 0;
    Object.keys(periodScores).forEach(p => {
      periodScores[p] = Math.round(periodScores[p]);
      total += periodScores[p];
    });

    const computedGrade = calcGrade_(total, scale);
    const gradeOverride = overrideMap[stu.studentId] || '';

    return {
      studentId: stu.studentId,
      studentCode: stu.studentCode,
      name: stu.name,
      rollNumber: stu.rollNumber,
      periodScores: periodScores,
      total: total,
      computedGrade: computedGrade,
      gradeOverride: gradeOverride,
      grade: gradeOverride || computedGrade,
      desiredCharacteristics: characterMap[stu.studentId] || '',
      readThinkWrite: readThinkMap[stu.studentId] || ''
    };
  });

  return { periods: data.periods, rows: rows, gradeScale: scale };
}

/**
 * ตั้ง/ยกเลิกเกรดพิเศษ ("ร" หรือ "มส") ให้นักเรียนคนใดคนหนึ่งในวิชานี้
 * ส่ง overrideValue เป็น '' เพื่อยกเลิกและกลับไปใช้เกรดที่คำนวณอัตโนมัติ
 */
function setGradeOverride(subjectId, studentId, overrideValue) {
  requireSubjectAccess_(subjectId);

  const value = (overrideValue || '').trim();
  if (ALLOWED_GRADE_OVERRIDES_.indexOf(value) === -1) {
    throw new Error('ค่าเกรดพิเศษไม่ถูกต้อง (ต้องเป็น "ร" หรือ "มส" หรือเว้นว่างเพื่อคำนวณอัตโนมัติ)');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enrollment');
  const rows = getSheetDataWithRow_('Enrollment');
  const target = rows.find(r => r.StudentID === studentId && r.SubjectID === subjectId);
  if (!target) throw new Error('ไม่พบการลงทะเบียนของนักเรียนคนนี้ในวิชานี้');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = headers.indexOf('GradeOverride') + 1;
  if (col === 0) {
    throw new Error('ไม่พบคอลัมน์ GradeOverride ในชีต Enrollment กรุณารันฟังก์ชัน migrateAddGradeOverrideColumn ใน SetupSheets.gs ก่อน (รันครั้งเดียว)');
  }

  sheet.getRange(target._row, col).setValue(value);

  return getGradeSummary(subjectId);
}

/**
 * ดึงข้อมูลทั้งหมดสำหรับออกรายงาน ปพ.5 ของวิชานี้ - รวมข้อมูลวิชา, ครูผู้สอน, ครูที่ปรึกษา,
 * การแจกแจงเกรด (พร้อมคะแนนเฉลี่ย + เกรดเฉลี่ย), การแจกแจงคุณลักษณะฯ และการอ่านคิดวิเคราะห์เขียน
 * โครงสร้างและสูตรคำนวณอ้างอิงจากไฟล์ต้นฉบับ ปพ.5 ของโรงเรียนโดยตรง
 */
function getPor5Report(subjectId) {
  const access = requireSubjectAccess_(subjectId);
  const subject = access.subject;
  const summary = getGradeSummary(subjectId);

  // ครูผู้สอน
  const teacherRow = getTeacherByEmail(subject.TeacherEmail);
  const teacherName = teacherRow ? teacherRow.Name : subject.TeacherEmail;

  // ครูที่ปรึกษา (ผูกกับห้อง+ปี+เทอม ใช้ร่วมกันได้ทุกวิชาของห้องนั้น)
  const advisorNames = getClassAdvisor(subject.ClassRoom, subject.AcademicYear, subject.Semester);

  // ระดับชั้น: ม.1-3 = มัธยมศึกษาตอนต้น (ปพ.5-ต), ม.4-6 = มัธยมศึกษาตอนปลาย (ปพ.5-ป)
  const gradeMatch = String(subject.ClassRoom || '').match(/\d+/);
  const gradeNum = gradeMatch ? parseInt(gradeMatch[0], 10) : null;
  const isUpperSecondary = gradeNum !== null && gradeNum >= 4;
  const formType = isUpperSecondary ? 'ปพ.5-ป (มัธยมศึกษาตอนปลาย)' : 'ปพ.5-ต (มัธยมศึกษาตอนต้น)';
  const shortFormLabel = isUpperSecondary ? 'ปพ.5 ป' : 'ปพ.5 ต';

  // ระดับผลการเรียนที่เป็นตัวเลข 8 ระดับ (ไม่รวม ร/มส ซึ่งแยกเป็นคนละกลุ่มในฟอร์มจริง)
  const numericGrades = ['4', '3.5', '3', '2.5', '2', '1.5', '1', '0'];
  const gradeCounts = {};
  numericGrades.forEach(g => gradeCounts[g] = 0);
  let countR = 0;   // "ร" ผลการเรียนไม่สมบูรณ์
  let countMS = 0;  // "มส" ไม่มีสิทธิ์สอบ

  const numericTotals = []; // คะแนนรวมของคนที่ได้เกรดตัวเลข (ไม่รวม ร/มส) ใช้คำนวณคะแนนเฉลี่ย
  summary.rows.forEach(r => {
    const g = String(r.grade);
    if (g === 'ร') { countR++; return; }
    if (g === 'มส') { countMS++; return; }
    if (gradeCounts[g] !== undefined) gradeCounts[g]++;
    numericTotals.push(r.total);
  });

  const totalStudents = summary.rows.length;
  const numericGradedCount = numericGrades.reduce((sum, g) => sum + gradeCounts[g], 0); // "รวม" ในฟอร์ม

  // คะแนน = ค่าเฉลี่ยคะแนนรวม (0-100) ของคนที่ได้เกรดตัวเลข
  const avgScore = numericTotals.length ? numericTotals.reduce((a, b) => a + b, 0) / numericTotals.length : 0;

  // ผลการเรียน (ค่าสถิติ) = ค่าเฉลี่ยเกรด (GPA แบบถ่วงน้ำหนักตามจำนวนคนในแต่ละระดับ) หารด้วยจำนวนคนที่ได้เกรดตัวเลขเท่านั้น
  const gradePointSum = numericGrades.reduce((sum, g) => sum + gradeCounts[g] * Number(g), 0);
  const avgGradePoint = numericGradedCount ? gradePointSum / numericGradedCount : 0;

  // ช่วยแปลง 0 ให้เป็น '-' ตามฟอร์มต้นฉบับ (นับ 0 คน จะโชว์ขีดแทนเลข 0)
  function countOrDash_(n) { return n > 0 ? n : '-'; }
  function percentOrDash_(count, denom) {
    return count > 0 && denom > 0 ? (count * 100 / denom).toFixed(2) : '-';
  }

  const gradeDistribution = numericGrades.map(g => ({
    grade: g,
    count: countOrDash_(gradeCounts[g]),
    percent: percentOrDash_(gradeCounts[g], numericGradedCount)
  }));

  const rCell = { count: countOrDash_(countR), percent: percentOrDash_(countR, numericGradedCount) };
  const msCell = { count: countOrDash_(countMS), percent: percentOrDash_(countMS, numericGradedCount) };

  // การแจกแจงคุณลักษณะฯ และ การอ่านคิดวิเคราะห์เขียน (ฐานร้อยละ = จำนวนนักเรียนทั้งหมด ไม่ใช่เฉพาะคนได้เกรดตัวเลข)
  const ratingOrder = ['ดีเยี่ยม', 'ดี', 'ผ่าน', 'ไม่ผ่าน'];
  function distributionOf(field) {
    const counts = {};
    ratingOrder.forEach(r => counts[r] = 0);
    summary.rows.forEach(r => {
      const val = r[field];
      if (val && counts[val] !== undefined) counts[val]++;
    });
    return ratingOrder.map(r => ({
      rating: r,
      count: countOrDash_(counts[r]),
      percent: percentOrDash_(counts[r], totalStudents)
    }));
  }

  return {
    subject: {
      subjectName: subject.SubjectName,
      subjectCode: subject.SubjectCode || '',
      learningArea: subject.LearningArea || '',
      subjectType: subject.SubjectType || '',
      hoursPerWeek: subject.HoursPerWeek || '',
      credits: subject.Credits || '',
      classRoom: subject.ClassRoom || '',
      academicYear: subject.AcademicYear || '',
      semester: subject.Semester || '',
      formType: formType,
      shortFormLabel: shortFormLabel
    },
    teacherName: teacherName,
    advisorName: advisorNames.join(', '),
    advisorNames: advisorNames,
    totalStudents: totalStudents,
    numericGradedCount: numericGradedCount,
    rCell: rCell,
    msCell: msCell,
    avgScore: Math.round(avgScore * 100) / 100,
    avgGradePoint: Math.round(avgGradePoint * 100) / 100,
    gradeDistribution: gradeDistribution,
    characterDistribution: distributionOf('desiredCharacteristics'),
    readThinkDistribution: distributionOf('readThinkWrite'),
    rows: summary.rows
  };
}
function getGradeScale_() {
  const config = getSheetData_('Config');
  const row = config.find(c => c.Key === 'GradeScale');
  if (row && row.Value) {
    try {
      const parsed = JSON.parse(row.Value);
      if (Array.isArray(parsed) && parsed.length) return parsed;
    } catch (e) {
      // ถ้า parse ไม่ได้ จะ fallback ไปใช้ค่าเริ่มต้นด้านล่าง
    }
  }
  return DEFAULT_GRADE_SCALE_;
}

function calcGrade_(totalScore, scale) {
  for (let i = 0; i < scale.length; i++) {
    if (totalScore >= scale[i].min) return scale[i].grade;
  }
  return scale[scale.length - 1].grade;
}