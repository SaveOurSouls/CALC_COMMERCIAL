/**
 * Работа с листами КБ: колонки, вставка строк, подтягивание полей из БД.ОП.
 */

/**
 * @param {string} name
 * @returns {boolean}
 */
function isKbSheet_(name) {
  return CONFIG.kbSheetPattern.test(String(name || '').trim());
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object} ключ CONFIG.kb -> номер колонки
 */
function resolveKbColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    var key = String(h || '').trim();
    if (key) {
      map[key] = i + 1;
    }
  });

  function col(titleOrLetter) {
    if (/^[A-Z]+$/i.test(titleOrLetter)) {
      return columnLetterToIndex_(titleOrLetter);
    }
    var c = map[titleOrLetter];
    if (!c) {
      throw new Error('На листе «' + sheet.getName() + '» нет колонки «' + titleOrLetter + '».');
    }
    return c;
  }

  var result = {
    sketch: col(CONFIG.kb.sketch),
    number: col(CONFIG.kb.number),
    n: col(CONFIG.kb.n),
    l: col(CONFIG.kb.l),
    op: col(CONFIG.kb.op)
  };

  try {
    result.timeTotal = col(CONFIG.kb.timeTotal);
  } catch (e) {
    result.timeTotal = 0;
  }
  try {
    result.price = col(CONFIG.kb.price);
  } catch (e) {
    result.price = 0;
  }
  try {
    result.opType = col(CONFIG.kb.opType);
  } catch (e) {
    result.opType = 0;
  }

  result.pulled = {};
  CONFIG.kb.pulledFromDb.forEach(function (title) {
    try {
      result.pulled[title] = col(title);
    } catch (e2) {
      // опциональные колонки
    }
  });

  return result;
}

/**
 * @param {string} letters
 * @returns {number}
 */
function columnLetterToIndex_(letters) {
  var s = letters.toUpperCase();
  var n = 0;
  for (var i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
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
 * @returns {number} следующая свободная строка
 */
function getNextKbRow_(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return Math.max(sheet.getLastRow() + 1, CONFIG.kbDataStartRow);
}

/**
 * Подтянуть из БД.ОП все поля, которые есть и в КБ, и в справочнике.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {Object} op
 * @param {Object} colMap
 */
function fillRowFromDb_(sheet, row, op, colMap) {
  if (!op || !op.rowIndex) {
    return;
  }
  var headerMap = getDbHeaderMap_();
  var kbHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  kbHeaders.forEach(function (title, i) {
    var t = String(title || '').trim();
    if (!t || t === CONFIG.kb.sketch || t === CONFIG.kb.number ||
        t === CONFIG.kb.n || t === CONFIG.kb.l || t === CONFIG.kb.op ||
        t === CONFIG.kb.timeTotal || t === CONFIG.kb.price) {
      return;
    }
    if (headerMap[t]) {
      var dbCol = headerMap[t];
      var dbSheet = getDbSheet_();
      var dbRow = op.rowIndex;
      var value = dbSheet.getRange(dbRow, dbCol).getValue();
      sheet.getRange(row, i + 1).setValue(value);
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
  var op = payload.op;
  var count = Math.max(1, parseInt(payload.count, 10) || 1);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || !isKbSheet_(sheetName)) {
    throw new Error('Лист «' + sheetName + '» не найден или не является КБ.');
  }

  var dbOp = findOperation_(sketch, number);
  if (!dbOp) {
    throw new Error('Операция не найдена в БД.ОП: ' + sketch + ' / ' + number);
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
    if (colMap.op && op !== undefined && op !== '') {
      sheet.getRange(row, colMap.op).setValue(op);
    }
    fillRowFromDb_(sheet, row, dbOp, colMap);
    applyNumberValidationForRow_(sheet, row, colMap);
    recalcKbRow_(sheet, row, colMap);
    inserted.push(row);
  }

  return { rows: inserted, sheetName: sheetName };
}

/**
 * Дублировать последнюю заполненную строку (для «нарезки проводов»).
 *
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

  var sourceRange = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
  var values = sourceRange.getValues()[0];
  count = Math.max(1, count || 1);
  var start = getNextKbRow_(sheetName);
  var rows = [];

  for (var i = 0; i < count; i++) {
    var row = start + i;
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    applyNumberValidationForRow_(sheet, row, colMap);
    recalcKbRow_(sheet, row, colMap);
    rows.push(row);
  }

  return { rows: rows, sheetName: sheetName };
}
