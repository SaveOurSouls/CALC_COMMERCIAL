/**
 * Вставка строк в структурированную таблицу Google Sheets.
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
    var sketch = String(sheet.getRange(r, colMap.sketch).getDisplayValue() || '').trim();
    var number = String(sheet.getRange(r, colMap.number).getDisplayValue() || '').trim();
    var n = sheet.getRange(r, colMap.n).getValue();
    var l = sheet.getRange(r, colMap.l).getValue();
    if (sketch || number || n !== '' || l !== '') {
      return r;
    }
  }
  return fromRow - 1;
}

/**
 * Строка-источник для «копировать» (выделенная или последняя заполненная).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number}
 */
function getKbSourceRowForCopy_(sheet) {
  var bounds = getKbDataBounds_(sheet);
  var active = sheet.getActiveRange();
  if (active) {
    var ar = active.getRow();
    if (ar >= bounds.firstDataRow && ar <= bounds.lastDataRow) {
      return ar;
    }
  }
  return bounds.lastDataRow;
}

/**
 * Строка, на которую вставлять (выделенная сдвигается вниз).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number}
 */
function getKbInsertBeforeRow_(sheet) {
  var bounds = getKbDataBounds_(sheet);
  var active = sheet.getActiveRange();
  if (active) {
    var ar = active.getRow();
    if (ar >= bounds.firstDataRow && ar <= Math.max(bounds.lastDataRow, bounds.firstDataRow)) {
      return ar;
    }
  }
  if (bounds.lastDataRow < bounds.firstDataRow) {
    return bounds.firstDataRow;
  }
  return bounds.lastDataRow + 1;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} count
 * @returns {number} первая вставленная строка
 */
function insertKbTableRows_(sheet, count) {
  var beforeRow = getKbInsertBeforeRow_(sheet);
  var bounds = getKbDataBounds_(sheet);
  if (beforeRow < bounds.firstDataRow) {
    beforeRow = bounds.firstDataRow;
  }
  sheet.insertRowsBefore(beforeRow, count);
  return beforeRow;
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
 * Строка вне таблицы для записи значений с последующим copyTo.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number}
 */
function getKbScratchRow_(sheet) {
  var table = getKbTable_(sheet);
  if (table) {
    return table.getRange().getLastRow() + 2;
  }
  return sheet.getLastRow() + 2;
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
    return false;
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} fromRow
 * @param {number} toRow
 * @param {number} col
 */
function copyCellValue_(sheet, fromRow, toRow, col) {
  sheet.getRange(fromRow, col).copyTo(
    sheet.getRange(toRow, col),
    SpreadsheetApp.CopyPasteType.PASTE_VALUES,
    false
  );
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} fromRow
 * @param {number} toRow
 * @param {Object} colMap
 */
function copyKbRowValues_(sheet, fromRow, toRow, colMap) {
  var cols = getKbWritableColumnIndexes_(colMap);
  cols.forEach(function (c) {
    try {
      copyCellValue_(sheet, fromRow, toRow, c);
    } catch (e) {
      var v = sheet.getRange(fromRow, c).getValue();
      safeSetCellValue_(sheet, toRow, c, v);
    }
  });
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} targetRow
 * @param {Object} colMap
 * @param {Object} data {sketch, number, n, l, op, dbOp}
 */
function writeKbRowData_(sheet, targetRow, colMap, data) {
  var scratch = getKbScratchRow_(sheet);
  var dbOp = data.dbOp;

  safeSetCellValue_(sheet, scratch, colMap.sketch, data.sketch);
  safeSetCellValue_(sheet, scratch, colMap.number, data.number);
  safeSetCellValue_(sheet, scratch, colMap.n, data.n);
  safeSetCellValue_(sheet, scratch, colMap.l, data.l);
  if (colMap.op && data.op !== undefined && data.op !== '') {
    safeSetCellValue_(sheet, scratch, colMap.op, data.op);
  }

  if (dbOp) {
    var dbHeaderMap = getDbHeaderMap_();
    var kbHeaderMap = getKbHeaderMap_(sheet);
    var dbSheet = getDbSheet_();
    var dbRow = dbOp.rowIndex;

    Object.keys(CONFIG.kbDbPull).forEach(function (kbHeader) {
      var dbTitle = CONFIG.kbDbPull[kbHeader];
      var kbCol = kbHeaderMap[kbHeader];
      var dbCol = dbHeaderMap[dbTitle];
      if (kbCol && dbCol) {
        safeSetCellValue_(sheet, scratch, kbCol, dbSheet.getRange(dbRow, dbCol).getValue());
      }
    });

    if (!shouldSkipCalculatedWrites_(sheet)) {
      Object.keys(kbHeaderMap).forEach(function (title) {
        var dbCol = dbHeaderMap[title];
        if (dbCol && kbHeaderMap[title]) {
          safeSetCellValue_(sheet, scratch, kbHeaderMap[title],
            dbSheet.getRange(dbRow, dbCol).getValue());
        }
      });
    }
  }

  copyKbRowValues_(sheet, scratch, targetRow, colMap);

  colsClearScratch_(sheet, scratch, colMap);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} scratch
 * @param {Object} colMap
 */
function colsClearScratch_(sheet, scratch, colMap) {
  getKbWritableColumnIndexes_(colMap).forEach(function (c) {
    try {
      sheet.getRange(scratch, c).clearContent();
    } catch (e) {
      // ignore
    }
  });
}

/**
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
