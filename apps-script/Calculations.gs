/**
 * Расчёт колонок «8» (суммарное время операции), «9» (суммарное время машины) и «Цена».
 */

/**
 * Суммарное время операции (кол. 8).
 *
 * @param {Object} op
 * @param {number} n
 * @param {number} l
 * @returns {number}
 */
function calculateOperationTime_(op, n, l) {
  var type = normalizeOpType_(op.opType);
  n = num_(n);
  l = num_(l);

  if (type === CONFIG.opTypes.linear) {
    var roll = op.rollSpeed;
    var toolSpeed = op.toolWorkSpeed;
    var toolOps = op.toolOpCount;
    if (!roll || roll <= 0) {
      return 0;
    }
    // (1/Скорость проката + Скорость работы инструмента * Кол-во операций инструмента) * L * N
    return (1 / roll + toolSpeed * toolOps) * l * n;
  }

  if (type === CONFIG.opTypes.variable || type === CONFIG.opTypes.static) {
    return op.timeHuman * n;
  }

  return 0;
}

/**
 * Суммарное время машины, сек (кол. 9).
 *
 * @param {Object} op
 * @param {number} n
 * @param {number} l
 * @returns {number}
 */
function calculateMachineTime_(op, n, l) {
  var type = normalizeOpType_(op.opType);
  n = num_(n);
  l = num_(l);

  if (type === CONFIG.opTypes.linear) {
    var roll = op.rollSpeed;
    if (!roll || roll <= 0) {
      return 0;
    }
    return (1 / roll + op.toolWorkSpeed * op.toolOpCount) * l * n;
  }

  if (type === CONFIG.opTypes.variable || type === CONFIG.opTypes.static) {
    return op.timeMachine * n;
  }

  return 0;
}

/**
 * @param {Object} op
 * @param {number} timeSec
 * @param {number} n
 * @returns {number}
 */
function calculateOperationPrice_(op, timeSec, n) {
  var type = normalizeOpType_(op.opType);
  timeSec = num_(timeSec);
  n = num_(n);

  if (type === CONFIG.opTypes.linear) {
    return timeSec * op.unitPriceMachine + timeSec * op.unitPriceHumanMag;
  }

  if (type === CONFIG.opTypes.variable || type === CONFIG.opTypes.static) {
    return timeSec * op.unitPriceHumanMag + op.timeMachine * n * op.unitPriceMachine;
  }

  return 0;
}

/**
 * @param {string} raw
 * @returns {string}
 */
function normalizeOpType_(raw) {
  var s = String(raw || '').trim();
  if (s === CONFIG.opTypes.linear ||
      s === CONFIG.opTypes.variable ||
      s === CONFIG.opTypes.static) {
    return s;
  }
  // допускаем суффиксы вроде «Погонный, сек»
  if (s.indexOf(CONFIG.opTypes.linear) === 0) {
    return CONFIG.opTypes.linear;
  }
  if (s.indexOf(CONFIG.opTypes.variable) === 0) {
    return CONFIG.opTypes.variable;
  }
  if (s.indexOf(CONFIG.opTypes.static) === 0) {
    return CONFIG.opTypes.static;
  }
  return s;
}

/**
 * Текст формулы для колонки времени (можно вставить в шапку и протянуть).
 * Подставьте буквы колонок под ваш лист КБ.
 *
 * @param {Object} cols { sketch, number, n, l, opType, rollSpeed, ... } буквы колонок
 * @returns {string}
 */
function buildTimeFormulaTemplate_(cols) {
  var linear = '(1/' + cols.rollSpeed + '+(' + cols.toolWorkSpeed + '*' + cols.toolOpCount + '))*' +
    cols.l + '*' + cols.n;
  var other = cols.timeHuman + '*' + cols.n;
  return '=IF(' + cols.opType + '="' + CONFIG.opTypes.linear + '",' + linear + ',' +
    'IF(OR(' + cols.opType + '="' + CONFIG.opTypes.variable + '",' +
    cols.opType + '="' + CONFIG.opTypes.static + '"),' + other + ',""))';
}

/**
 * @param {Object} cols
 * @returns {string}
 */
function buildPriceFormulaTemplate_(cols) {
  var timeRef = cols.timeTotal;
  var linear = timeRef + '*' + cols.unitPriceMachine + '+' + timeRef + '*' + cols.unitPriceHumanMag;
  var other = timeRef + '*' + cols.unitPriceHumanMag + '+' + cols.timeMachine + '*' + cols.n +
    '*' + cols.unitPriceMachine;
  return '=IF(' + cols.opType + '="' + CONFIG.opTypes.linear + '",' + linear + ',' +
    'IF(OR(' + cols.opType + '="' + CONFIG.opTypes.variable + '",' +
    cols.opType + '="' + CONFIG.opTypes.static + '"),' + other + ',""))';
}

/**
 * Пересчитать время и цену для одной строки листа КБ.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {Object} colMap
 */
function recalcKbRow_(sheet, row, colMap) {
  if (!CONFIG.kbEnableScriptRecalc) {
    return;
  }
  var sketch = String(sheet.getRange(row, colMap.sketch).getValue() || '').trim();
  var number = String(sheet.getRange(row, colMap.number).getValue() || '').trim();
  if (!sketch || !number) {
    return;
  }

  var op = findOperation_(sketch, number);
  if (!op) {
    return;
  }

  var n = sheet.getRange(row, colMap.n).getValue();
  var l = sheet.getRange(row, colMap.l).getValue();
  var time = calculateOperationTime_(op, n, l);
  var machineTime = calculateMachineTime_(op, n, l);
  var price = calculateOperationPrice_(op, time, n);

  if (shouldSkipCalculatedWrites_(sheet)) {
    return;
  }

  safeSetCellValue_(sheet, row, colMap.timeTotal, time);
  safeSetCellValue_(sheet, row, colMap.timeMachineTotal, machineTime);
  safeSetCellValue_(sheet, row, colMap.price, price);
}

/**
 * Пересчитать все заполненные строки на активном листе КБ.
 */
function recalcActiveKbSheet() {
  if (!CONFIG.kbEnableScriptRecalc) {
    SpreadsheetApp.getUi().alert(
      'Скриптовый пересчёт временно отключён (Config.kbEnableScriptRecalc = false).\n' +
        'Используйте формулы в таблице КБ.'
    );
    return;
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!isKbSheet_(sheet.getName())) {
    SpreadsheetApp.getUi().alert('Откройте лист кабельной книги (КБ1, КБ2, …).');
    return;
  }
  var colMap = resolveKbColumns_(sheet);
  var lastRow = sheet.getLastRow();
  for (var r = CONFIG.kbDataStartRow; r <= lastRow; r++) {
    recalcKbRow_(sheet, r, colMap);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Пересчёт кол. 8, 9 и Цена завершён.', 'КБ', 3);
}
