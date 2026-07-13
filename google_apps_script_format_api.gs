const SPREADSHEET_ID = '1dc-dIvVU8a4Kmozph0oULlcsc5iVZ5YPkJWTDWFUMok';
const DEFAULT_TEMPLATE_SHEET_NAME = '長府店_定型文';
const DEFAULT_STAFF_SHEET_NAME = '長府店_スタッフ';

function doGet(e) {
  const params = e.parameter || {};
  const callback = params.callback || 'callback';

  try {
    const action = params.action || 'list';
    let result;

    if (action === 'save') {
      const sheet = getSheet(params.sheet || DEFAULT_TEMPLATE_SHEET_NAME);
      result = saveTemplate(sheet, params.name || '', params.text || '');
    } else if (action === 'listStaff') {
      const sheet = getSheet(params.staffSheet || DEFAULT_STAFF_SHEET_NAME);
      result = listStaff(sheet);
    } else if (action === 'saveStaff') {
      const sheet = getSheet(params.staffSheet || DEFAULT_STAFF_SHEET_NAME);
      result = saveStaff(sheet, params.names || '');
    } else {
      const sheet = getSheet(params.sheet || DEFAULT_TEMPLATE_SHEET_NAME);
      result = listTemplates(sheet);
    }

    return jsonp(callback, Object.assign({ ok: true }, result));
  } catch (error) {
    return jsonp(callback, { ok: false, error: error.message });
  }
}

function getSheet(sheetName) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('シートが見つかりません: ' + sheetName);
  }

  return sheet;
}

function listTemplates(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 2).getValues()
    : [];

  return {
    templates: values
      .filter(row => row[0])
      .map(row => ({ name: String(row[0]), text: String(row[1] || '') }))
  };
}

function saveTemplate(sheet, name, text) {
  if (!name) {
    throw new Error('定型文名が空です');
  }

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const names = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat()
    : [];
  const index = names.findIndex(value => String(value) === name);
  const row = index >= 0 ? index + 2 : lastRow + 1;

  sheet.getRange(row, 1, 1, 3).setValues([[name, text, new Date()]]);

  return { row: row };
}

function listStaff(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    : [];

  return {
    staff: values
      .map(row => String(row[0] || '').trim())
      .filter(Boolean)
  };
}

function saveStaff(sheet, namesText) {
  const names = String(namesText || '')
    .split(String.fromCharCode(10))
    .map(name => name.trim())
    .filter(Boolean);

  if (names.length === 0) {
    throw new Error('スタッフ名が空です');
  }

  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 1, maxRows, 3).clearContent();

  const rows = names.map((name, index) => [name, index + 1, '']);
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  return { count: names.length };
}

function jsonp(callback, data) {
  const safeCallback = String(callback).replace(/[^a-zA-Z0-9_.$]/g, '');

  return ContentService
    .createTextOutput(safeCallback + '(' + JSON.stringify(data) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
