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
 * Подтянуть из БД.ОП поля 6, 7 и совпадающие заголовки (без эскиза/номера).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {Object} op
 */
function fillRowFromDb_(sheet, row, op) {
  if (!op || !op.rowIndex) {
    return;
  }

  var colMap = resolveKbColumns_(sheet);
  var dbHeaderMap = getDbHeaderMap_();
  var kbHeaderMap = getKbHeaderMap_(sheet);
  var dbSheet = getDbSheet_();
  var dbRow = op.rowIndex;
  var scratch = getKbScratchRow_(sheet);

  Object.keys(CONFIG.kbDbPull).forEach(function (kbHeader) {
    var dbTitle = CONFIG.kbDbPull[kbHeader];
    var kbCol = kbHeaderMap[kbHeader];
    var dbCol = dbHeaderMap[dbTitle];
    if (kbCol && dbCol) {
      safeSetCellValue_(sheet, scratch, kbCol, dbSheet.getRange(dbRow, dbCol).getValue());
      try {
        copyCellValue_(sheet, scratch, row, kbCol);
      } catch (e) {
        safeSetCellValue_(sheet, row, kbCol, sheet.getRange(scratch, kbCol).getValue());
      }
    }
  });

  if (!shouldSkipCalculatedWrites_(sheet)) {
    var skip = {};
    skip[CONFIG.kb.sketch] = true;
    skip[CONFIG.kb.number] = true;
    skip[CONFIG.kb.n] = true;
    skip[CONFIG.kb.l] = true;
    skip[CONFIG.kb.op] = true;
    Object.keys(CONFIG.kbDbPull).forEach(function (h) {
      skip[h] = true;
    });

    Object.keys(kbHeaderMap).forEach(function (title) {
      if (skip[title]) {
        return;
      }
      var dbCol = dbHeaderMap[title];
      if (dbCol) {
        safeSetCellValue_(sheet, scratch, kbHeaderMap[title],
          dbSheet.getRange(dbRow, dbCol).getValue());
        try {
          copyCellValue_(sheet, scratch, row, kbHeaderMap[title]);
        } catch (e2) {
          safeSetCellValue_(sheet, row, kbHeaderMap[title], sheet.getRange(scratch, kbCol).getValue());
        }
      }
    });
  }

  colsClearScratch_(sheet, scratch, colMap);
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
  var count = parseRowCount_(payload.count);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || !isKbSheet_(sheetName)) {
    throw new Error('Лист «' + sheetName + '» не найден или не является КБ.');
  }

  var dbOp = findOperation_(sketch, number);
  if (!CONFIG.kbInsertOnlyFiveColumns && !dbOp) {
    throw new Error('Операция не найдена в БД.ОП: «' + sketch + '» / «' + number + '».');
  }

  var colMap = resolveKbColumns_(sheet);
  var startRow = insertKbTableRows_(sheet, count);
  var inserted = [];
  var rowData = {
    sketch: sketch,
    number: number,
    n: n,
    l: l,
    op: opVal,
    dbOp: dbOp
  };

  writeKbRowData_(sheet, startRow, colMap, rowData);
  inserted.push(startRow);

  if (count > 1) {
    var destRows = [];
    for (var i = 1; i < count; i++) {
      destRows.push(startRow + i);
      inserted.push(startRow + i);
    }
    copyKbRowToMany_(sheet, startRow, destRows, colMap);
  }

  for (var r = 0; r < inserted.length; r++) {
    try {
      applyNumberValidationForRow_(sheet, inserted[r], colMap);
    } catch (eVal) {
      console.warn(eVal.message);
    }
    if (CONFIG.kbEnableScriptRecalc) {
      recalcKbRow_(sheet, inserted[r], colMap);
    }
  }

  return {
    rows: inserted,
    sheetName: sheetName,
    inTable: !!getKbTable_(sheet),
    hint: 'Вставлено ' + count + ' строк на место выделенной (строка ' + startRow + ').'
  };
}

/**
 * @param {string} sheetName
 * @param {number} count
 * @returns {Object}
 */
function duplicateLastKbRow_(sheetName, count) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var colMap = resolveKbColumns_(sheet);
  var bounds = getKbDataBounds_(sheet);
  var sourceRow = getKbSourceRowForCopy_(sheet);

  if (sourceRow < bounds.firstDataRow) {
    throw new Error('Нет строк для копирования. Выделите строку в таблице или заполните хотя бы одну.');
  }

  count = parseRowCount_(count);
  var startRow = insertKbTableRows_(sheet, count);
  var rows = [startRow];

  copyKbRowValues_(sheet, sourceRow, startRow, colMap);

  if (count > 1) {
    var destRows = [];
    for (var i = 1; i < count; i++) {
      destRows.push(startRow + i);
      rows.push(startRow + i);
    }
    copyKbRowToMany_(sheet, startRow, destRows, colMap);
  }

  for (var j = 0; j < rows.length; j++) {
    try {
      applyNumberValidationForRow_(sheet, rows[j], colMap);
    } catch (eVal) {
      console.warn(eVal.message);
    }
    if (CONFIG.kbEnableScriptRecalc) {
      recalcKbRow_(sheet, rows[j], colMap);
    }
  }

  return {
    rows: rows,
    sheetName: sheetName,
    sourceRow: sourceRow,
    hint: 'Скопировано ' + count + ' строк с строки ' + sourceRow + ' на позицию ' + startRow + '.'
  };
}

/**
 * Вставка очереди операций одним вызовом (один запрос к таблице).
 *
 * @param {Object} payload {sheetName, rows: Array}
 * @returns {Object}
 */
function insertKbQueue_(payload) {
  var sheetName = payload.sheetName;
  var rows = payload.rows || [];

  if (!rows.length) {
    throw new Error('Очередь пуста. Добавьте операции из справочника.');
  }
  if (rows.length > 50) {
    throw new Error('За один раз можно вставить не более 50 строк.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || !isKbSheet_(sheetName)) {
    throw new Error('Лист «' + sheetName + '» не найден или не является КБ.');
  }

  var colMap = resolveKbColumns_(sheet);
  var count = rows.length;
  var startRow = insertKbTableRows_(sheet, count);
  var inserted = [];
  for (var i = 0; i < count; i++) {
    var item = rows[i];
    var sketch = String(item.sketch || '').trim();
    var number = String(item.number || '').trim();
    var dbOp = findOperation_(sketch, number);
    if (!CONFIG.kbInsertOnlyFiveColumns && !dbOp) {
      throw new Error('Операция не найдена (позиция ' + (i + 1) + '): «' + sketch + '» / «' + number + '».');
    }

    var row = startRow + i;
    writeKbRowData_(sheet, row, colMap, {
      sketch: sketch,
      number: number,
      n: num_(item.n),
      l: num_(item.l),
      op: item.op || '',
      dbOp: dbOp
    });
    inserted.push(row);

    /* Пересчёт 8, 9, Цена — отключён (kbEnableScriptRecalc) */
    if (CONFIG.kbEnableScriptRecalc && !shouldSkipCalculatedWrites_(sheet)) {
      recalcKbRow_(sheet, row, colMap);
    }
  }

  SpreadsheetApp.flush();

  return {
    rows: inserted,
    sheetName: sheetName,
    count: count,
    startRow: startRow,
    hint: 'Вставлено ' + count + ' строк (Эскиз, Номер, N, L, OP) с ' + startRow + '.'
  };
}
