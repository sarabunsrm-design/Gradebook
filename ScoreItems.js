// ช่วงคะแนนทั้ง 4 ช่วงของระบบ (เรียงตามลำดับที่ต้องแสดงผลเสมอ)
var PERIODS_ = ['ก่อนกลางภาค', 'กลางภาค', 'หลังกลางภาค', 'ปลายภาค'];

/**
 * ดึงรายการคะแนนทั้งหมดของวิชา พร้อมสรุปน้ำหนักรวมแต่ละช่วง และรวมทั้งหมด
 */
function getScoreItems(subjectId) {
  requireSubjectAccess_(subjectId);

  const items = getSheetDataWithRow_('ScoreItems').filter(i => i.SubjectID === subjectId);

  const weightByPeriod = {};
  PERIODS_.forEach(p => weightByPeriod[p] = 0);
  items.forEach(i => {
    weightByPeriod[i.Period] = (weightByPeriod[i.Period] || 0) + Number(i.Weight || 0);
  });

  const totalWeight = Object.keys(weightByPeriod).reduce((sum, p) => sum + weightByPeriod[p], 0);

  return {
    items: items.map(i => ({
      ItemID: i.ItemID, SubjectID: i.SubjectID, Period: i.Period,
      ItemName: i.ItemName, MaxScore: i.MaxScore, Weight: i.Weight
    })),
    weightByPeriod: weightByPeriod,
    totalWeight: Math.round(totalWeight * 100) / 100,
    periods: PERIODS_
  };
}

/**
 * เพิ่มรายการคะแนนใหม่
 * form = { subjectId, period, itemName, maxScore, weight }
 */
function addScoreItem(form) {
  requireSubjectAccess_(form.subjectId);
  validateScoreItemForm_(form);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ScoreItems');
  const id = 'I' + new Date().getTime() + Math.floor(Math.random() * 1000);
  sheet.appendRow([id, form.subjectId, form.period, form.itemName, Number(form.maxScore), Number(form.weight)]);

  return getScoreItems(form.subjectId);
}

/**
 * แก้ไขรายการคะแนน
 * form = { itemId, subjectId, period, itemName, maxScore, weight }
 */
function updateScoreItem(form) {
  requireSubjectAccess_(form.subjectId);
  validateScoreItemForm_(form);

  const items = getSheetDataWithRow_('ScoreItems');
  const target = items.find(i => i.ItemID === form.itemId && i.SubjectID === form.subjectId);
  if (!target) throw new Error('ไม่พบรายการคะแนนนี้');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ScoreItems');
  // คอลัมน์: ItemID(1) SubjectID(2) Period(3) ItemName(4) MaxScore(5) Weight(6)
  sheet.getRange(target._row, 3, 1, 4).setValues([[
    form.period, form.itemName, Number(form.maxScore), Number(form.weight)
  ]]);

  return getScoreItems(form.subjectId);
}

/**
 * ลบรายการคะแนน
 */
function deleteScoreItem(itemId, subjectId) {
  requireSubjectAccess_(subjectId);

  const items = getSheetDataWithRow_('ScoreItems');
  const target = items.find(i => i.ItemID === itemId && i.SubjectID === subjectId);
  if (!target) throw new Error('ไม่พบรายการคะแนนนี้');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ScoreItems');
  sheet.deleteRow(target._row);

  return getScoreItems(subjectId);
}

function validateScoreItemForm_(form) {
  if (!form.itemName) throw new Error('กรุณาระบุชื่อรายการคะแนน');
  if (PERIODS_.indexOf(form.period) === -1) throw new Error('ช่วงคะแนนไม่ถูกต้อง');
  if (!form.maxScore || Number(form.maxScore) <= 0) throw new Error('คะแนนเต็มต้องมากกว่า 0');
  if (form.weight === undefined || form.weight === '' || Number(form.weight) < 0) {
    throw new Error('กรุณาระบุน้ำหนักคะแนน (%)');
  }
}