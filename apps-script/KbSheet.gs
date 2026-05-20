/**
 * Работа с листами КБ: колонки, вставка строк, подтягивание полей из БД.ОП.
 */

/**
 * @param {string} name
 * @returns {boolean}
 */
function isKbSheet_(name) {
  var n = String(name || '').trim();
  if (CONFIG.kbSheetPattern.test(n)) {
    return true;
  }
  if (/^КБ/i.test(n) && /\d/.test(n)) {
    return true;
  }
  if (CONFIG.kbSheetExtraNames && CONFIG.kbSheetExtraNames.indexOf(n) >= 0) {
    return true;
  }
  return false;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object}
 */
function resolveKbColumns_(sheet) {
  var map = getKbHeaderMap_(sheet);

  function col(title) {
    var c = map[title];
    if (!c) {
      throw new Error('На листе «' + sheet.getName() + '» (строка ' + CONFIG.kbHeaderRow +
        ') нет колонки «' + title + '».');
    }
    return c;
  }

  function colOptional(title) {
    return map[title] || 0;
  }

  var result = {
    sketch: col(CONFIG.kb.sketch),
    number: col(CONFIG.kb.number),
    n: col(CONFIG.kb.n),
    l: col(CONFIG.kb.l),
    op: colOptional(CONFIG.kb.op),
    timeTotal: colOptional(CONFIG.kb.timeTotal),
    timeMachineTotal: colOptional(CONFIG.kb.timeMachineTotal),
    price: colOptional(CONFIG.kb.price),
    opType: colOptional(CONFIG.kb.opType)
  };

  result.pullCols = {};
  Object.keys(CONFIG.kbDbPull).forEach(function (kbHeader) {
    if (map[kbHeader]) {
      result.pullCols[kbHeader] = map[kbHeader];
    }
  });

  return result;
}

/**
 * @returns {string[]}
 */
function listKbSheetNames_() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(function (s) {
      return s.getName();
    })
    .filter(isKbSheet_);
}

/**
 * @param {string} sheetName
 * @returns {number}
 */
function getNextKbRow_(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return Math.max(sheet.getLastRow() + 1, CONFIG.kbDataStartRow);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {Object} op
 */
function fillRowFromDb_(sheet, row, op) {
  if (!op || !op.rowIndex) {
    return;
  }

  var dbHeaderMap = getDbHeaderMap_();
  var kbHeaderMap = getKbHeaderMap_(sheet);
  var dbSheet = getDbSheet_();
  var dbRow = op.rowIndex;

  Object.keys(CONFIG.kbDbPull).forEach(function (kbHeader) {
    var dbTitle = CONFIG.kbDbPull[kbHeader];
    var kbCol = kbHeaderMap[kbHeader];
    var dbCol = dbHeaderMap[dbTitle];
    if (kbCol && dbCol) {
      sheet.getRange(row, kbCol).setValue(dbSheet.getRange(dbRow, dbCol).getValue());
    }
  });

  var skip = {};
  skip[CONFIG.kb.sketch] = true;
  skip[CONFIG.kb.number] = true;
  skip[CONFIG.kb.n] = true;
  skip[CONFIG.kb.l] = true;
  skip[CONFIG.kb.op] = true;
  skip[CONFIG.kb.timeTotal] = true;
  skip[CONFIG.kb.price] = true;
  Object.keys(CONFIG.kbDbPull).forEach(function (h) {
    skip[h] = true;
  });

  Object.keys(kbHeaderMap).forEach(function (title) {
    if (skip[title]) {
      return;
    }
    var dbCol = dbHeaderMap[title];
    if (dbCol) {
      sheet.getRange(row, kbHeaderMap[title])
        .setValue(dbSheet.getRange(dbRow, dbCol).getValue());
    }
  });
}

/**
 * @param {Object} payload
 * @returns {Object}
 */
function insertKbRows_(payload) {
  var sheetName = payload.sheetName;
  var sketch = payload.sketch;
  var number = payload.number;
  var n = payload.n;
  var l = payload.l;
  var opVal = payload.op;
  var count = Math.max(1, parseInt(payload.count, 10) || 1);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || !isKbSheet_(sheetName)) {
    throw new Error('Лист «' + sheetName + '» не найден или не является КБ.');
  }

  var dbOp = findOperation_(sketch, number);
  if (!dbOp) {
    throw new Error('Операция не найдена в БД.ОП: «' + sketch + '» / «' + number + '».');
  }

  var colMap = resolveKbColumns_(sheet);
  var startRow = getNextKbRow_(sheetName);
  var inserted = [];

  for (var i = 0; i < count; i++) {
    var row = startRow + i;
    sheet.getRange(row, colMap.sketch).setValue(sketch);
    sheet.getRange(row, colMap.number).setValue(number);
    sheet.getRange(row, colMap.n).setValue(n);
    sheet.getRange(row, colMap.l).setValue(l);
    if (colMap.op && opVal !== undefined && opVal !== '') {
      sheet.getRange(row, colMap.op).setValue(opVal);
    }
    fillRowFromDb_(sheet, row, dbOp);
    applyNumberValidationForRow_(sheet, row, colMap);
    recalcKbRow_(sheet, row, colMap);
    inserted.push(row);
  }

  return { rows: inserted, sheetName: sheetName };
}

/**
 * @param {string} sheetName
 * @param {number} count
 * @returns {Object}
 */
function duplicateLastKbRow_(sheetName, count) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var colMap = resolveKbColumns_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.kbDataStartRow) {
    throw new Error('Нет строк для копирования.');
  }

  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(lastRow, 1, lastRow, lastCol).getValues()[0];
  count = Math.max(1, count || 1);
  var start = getNextKbRow_(sheetName);
  var rows = [];

  for (var i = 0; i < count; i++) {
    var row = start + i;
    sheet.getRange(row, 1, row, lastCol).setValues([values]);
    applyNumberValidationForRow_(sheet, row, colMap);
    recalcKbRow_(sheet, row, colMap);
    rows.push(row);
  }

  return { rows: rows, sheetName: sheetName };
}
