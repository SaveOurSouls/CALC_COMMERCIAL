/**
 * Вставка строк в структурированную таблицу Google Sheets.
 * Колонки с типом «формула» / «выпадающий список таблицы» нельзя перезаписывать скриптом.
 */

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {GoogleAppsScript.Spreadsheet.Table|null}
 */
function getKbTable_(sheet) {
  if (typeof sheet.getTables !== 'function') {
    return null;
  }
  var tables = sheet.getTables();
  if (!tables || !tables.length) {
    return null;
  }
  for (var i = 0; i < tables.length; i++) {
    var r = tables[i].getRange();
    if (r.getRow() <= CONFIG.kbHeaderRow && r.getLastRow() >= CONFIG.kbDataStartRow) {
      return tables[i];
    }
  }
  return tables[0];
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {{firstDataRow: number, lastDataRow: number, table: Object|null, range: Object|null}}
 */
function getKbDataBounds_(sheet) {
  var table = getKbTable_(sheet);
  var firstDataRow = CONFIG.kbDataStartRow;
  var lastDataRow = CONFIG.kbDataStartRow - 1;
  var range = null;

  if (table) {
    range = table.getRange();
    firstDataRow = Math.max(range.getRow() + 1, CONFIG.kbDataStartRow);
    lastDataRow = findLastKbDataRow_(sheet, firstDataRow, range.getLastRow());
    if (lastDataRow < firstDataRow) {
      lastDataRow = firstDataRow - 1;
    }
    return { firstDataRow: firstDataRow, lastDataRow: lastDataRow, table: table, range: range };
  }

  lastDataRow = findLastKbDataRow_(sheet, firstDataRow, sheet.getLastRow());
  if (lastDataRow < firstDataRow) {
    lastDataRow = firstDataRow - 1;
  }
  return { firstDataRow: firstDataRow, lastDataRow: lastDataRow, table: null, range: null };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} fromRow
 * @param {number} toRow
 * @returns {number}
 */
function findLastKbDataRow_(sheet, fromRow, toRow) {
  var colMap;
  try {
    colMap = resolveKbColumns_(sheet);
  } catch (e) {
    return Math.max(toRow, fromRow - 1);
  }

  for (var r = toRow; r >= fromRow; r--) {
    var sketch = String(sheet.getRange(r, colMap.sketch).getValue() || '').trim();
    var number = String(sheet.getRange(r, colMap.number).getValue() || '').trim();
    var n = sheet.getRange(r, colMap.n).getValue();
    var l = sheet.getRange(r, colMap.l).getValue();
    if (sketch || number || n !== '' || l !== '') {
      return r;
    }
  }
  return fromRow - 1;
}

/**
 * Строка, после которой вставлять (учёт выделения в таблице).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number}
 */
function getKbInsertAfterRow_(sheet) {
  var bounds = getKbDataBounds_(sheet);
  var active = sheet.getActiveRange();
  if (active) {
    var ar = active.getRow();
    if (ar >= bounds.firstDataRow && ar <= Math.max(bounds.lastDataRow, bounds.firstDataRow)) {
      return ar;
    }
  }
  return Math.max(bounds.lastDataRow, bounds.firstDataRow - 1);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} count
 * @returns {number} первая вставленная строка
 */
function insertKbTableRows_(sheet, count) {
  var bounds = getKbDataBounds_(sheet);
  var afterRow = getKbInsertAfterRow_(sheet);
  if (afterRow < bounds.firstDataRow - 1) {
    afterRow = bounds.firstDataRow - 1;
  }

  sheet.insertRowsAfter(afterRow, count);
  return afterRow + 1;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {boolean}
 */
function shouldSkipCalculatedWrites_(sheet) {
  if (CONFIG.kbSkipWriteCalculatedInTable === false) {
    return false;
  }
  return !!getKbTable_(sheet);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {number} col
 * @param {*} value
 * @returns {boolean}
 */
function safeSetCellValue_(sheet, row, col, value) {
  if (!col) {
    return false;
  }
  try {
    sheet.getRange(row, col).setValue(value);
    return true;
  } catch (e) {
    console.warn('Пропуск записи R' + row + 'C' + col + ': ' + e.message);
    return false;
  }
}

/**
 * Колонки, в которые скрипт может писать (ввод оператором + подтягивание из БД.ОП).
 *
 * @param {Object} colMap
 * @returns {number[]}
 */
function getKbWritableColumnIndexes_(colMap) {
  var cols = [
    colMap.sketch, colMap.number, colMap.n, colMap.l, colMap.op
  ];
  Object.keys(colMap.pullCols || {}).forEach(function (k) {
    cols.push(colMap.pullCols[k]);
  });

  var skipCalc = [
    colMap.timeTotal, colMap.timeMachineTotal, colMap.price, colMap.opType
  ];
  var skip = {};
  skipCalc.forEach(function (c) {
    if (c) {
      skip[c] = true;
    }
  });

  var out = [];
  cols.forEach(function (c) {
    if (c && !skip[c] && out.indexOf(c) < 0) {
      out.push(c);
    }
  });
  return out;
}
