const SHEET_NAME = '시트2';
const MEMBERS = [
  '롱게임', '비비드드림', '꿈요셉', '사쁘니', '행복소금',
  '미쁨댁', '우부왕', '금만가', '플리퍼즈', '스테디 릴리',
  '올댓드림', '오늘날씨맑음', '꿀람쥐', '꿈을 현실로', '뚜연',
  '라난', '럭셔리 노블', '레리꼬', '부자될또정', '하예', '뿌래랠래', '자산일조', '산골지야'
];
const SPECIAL_ROWS = ['챌린지 성공', '챌린지 성공률'];

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'load';
  try {
    if (action === 'health') {
      return jsonOutput_({ ok: true, sheetName: SHEET_NAME, spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId() });
    }
    if (action === 'load') {
      return jsonOutput_({ ok: true, records: loadRecords_() });
    }
    return jsonOutput_({ ok: false, error: 'Unsupported action' });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const payload = parseBody_(e);
    if (!payload || !payload.date || !payload.record) {
      return jsonOutput_({ ok: false, error: 'Missing date or record' });
    }
    saveRecord_(payload.date, payload.record);
    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err.message || err) });
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureHeader_(sheet, requiredCols) {
  if (sheet.getMaxColumns() < requiredCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
  }
  if (sheet.getRange(1, 1).getValue() !== '닉네임') {
    sheet.getRange(1, 1).setValue('닉네임');
  }
}

function loadRecords_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const header = values[0] || [];
  const records = {};
  for (let c = 1; c < header.length; c++) {
    const label = header[c];
    if (!label || SPECIAL_ROWS.indexOf(label) !== -1) continue;
    const parts = String(label).split('/');
    if (parts.length !== 2) continue;
    const dateKey = `2026-${String(Number(parts[0])).padStart(2, '0')}-${String(Number(parts[1])).padStart(2, '0')}`;
    records[dateKey] = records[dateKey] || {};
    for (let r = 1; r < values.length; r++) {
      const name = values[r][0];
      if (!name || SPECIAL_ROWS.indexOf(name) !== -1) continue;
      if (MEMBERS.indexOf(name) === -1) continue;
      records[dateKey][name] = values[r][c] === 'O';
    }
  }
  return records;
}

function saveRecord_(date, record) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const header = values[0] || [];
  const parts = date.split('-');
  const dateLabel = `${Number(parts[1])}/${Number(parts[2])}`;
  let dateColIndex = header.indexOf(dateLabel);
  if (dateColIndex === -1) {
    dateColIndex = Math.max(header.length, 1);
  }
  ensureHeader_(sheet, dateColIndex + 1);
  sheet.getRange(1, dateColIndex + 1).setValue(dateLabel);

  MEMBERS.forEach(function(name, i) {
    const row = i + 2;
    sheet.getRange(row, 1).setValue(name);
    sheet.getRange(row, dateColIndex + 1).setValue(record[name] ? 'O' : 'X');
  });

  const checkedCount = MEMBERS.filter(function(name) { return !!record[name]; }).length;
  const rate = (checkedCount / MEMBERS.length * 100).toFixed(2) + '%';
  const successRow = MEMBERS.length + 2;
  const rateRow = MEMBERS.length + 3;
  sheet.getRange(successRow, 1).setValue('챌린지 성공');
  sheet.getRange(successRow, dateColIndex + 1).setValue(checkedCount);
  sheet.getRange(rateRow, 1).setValue('챌린지 성공률');
  sheet.getRange(rateRow, dateColIndex + 1).setValue(rate);
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return null;
  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    return null;
  }
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
