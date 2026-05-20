/**
 * Замена пар листов TL_КБ* — зависимый список «Номер» через FILTER по строке.
 */

/**
 * Имя скрытого листа-справочника для именованных диапазонов (опционально).
 */
var HELPER_SHEET_NAME = '_KB_HELPER';

/**
 * При изменении «Эскиз проекта» обновить проверку данных для «Номер» в той же строке.
 *
 * @param {Object} e
 */
function onEditKbHandler_(e) {
  if (!e || !e.range) {
    return;
  }
  var sheet = e.range.getSheet();
  if (!isKbSheet_(sheet.getName())) {
    return;
  }

  var colMap;
  try {
    colMap = resolveKbColumns_(sheet);
  } catch (err) {
    return;
  }

  var row = e.range.getRow();
  if (row < CONFIG.kbDataStartRow) {
    return;
  }

  var col = e.range.getColumn();
  if (col === colMap.sketch) {
    applyNumberValidationForRow_(sheet, row, colMap);
    var sketchVal = String(e.value || sheet.getRange(row, colMap.sketch).getValue() || '').trim();
    var number = String(sheet.getRange(row, colMap.number).getValue() || '').trim();
    if (sketchVal && number) {
      var op = findOperation_(sketchVal, number);
      if (op) {
        fillRowFromDb_(sheet, row, op, colMap);
      }
      recalcKbRow_(sheet, row, colMap);
    }
  }

  if (col === colMap.number || col === colMap.n || col === colMap.l) {
    recalcKbRow_(sheet, row, colMap);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {Object} colMap
 */
function applyNumberValidationForRow_(sheet, row, colMap) {
  var sketch = String(sheet.getRange(row, colMap.sketch).getValue() || '').trim();
  var numberCell = sheet.getRange(row, colMap.number);
  if (!sketch) {
    numberCell.clearDataValidations();
    return;
  }

  var numbers = getNumbersForSketch_(sketch);
  if (!numbers.length) {
    numberCell.clearDataValidations();
    return;
  }

  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(numbers, true)
    .setAllowInvalid(false)
    .build();
  numberCell.setDataValidation(validation);
}

/**
 * @param {string} sketch
 * @returns {string[]}
 */
function getNumbersForSketch_(sketch) {
  return listOperationsBySketch_(sketch).map(function (r) {
    return r.number;
  });
}

/**
 * Формула для ручной вставки в проверку данных (через UI):
 * Показывает, как заменить TL_КБ без отдельного листа на каждый КБ.
 *
 * @param {string} sketchColLetter например B
 * @param {number} row
 * @returns {string}
 */
function buildNumberDropdownFormula_(sketch, sketchColIndex, row, sheetName) {
  var db = CONFIG.dbSheet;
  var dbSketchCol = 'A'; // уточните после привязки к БД.ОП
  var dbNumberCol = 'B';
  return '=SORT(UNIQUE(FILTER(' + db + '!' + dbNumberCol + ':' + dbNumberCol + ',' +
    db + '!' + dbSketchCol + ':' + dbSketchCol + '=$' + columnIndexToLetter_(sketchColIndex) + row + ')))';
}

/**
 * @param {number} index 1-based
 * @returns {string}
 */
function columnIndexToLetter_(index) {
  var temp;
  var letter = '';
  while (index > 0) {
    temp = (index - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    index = (index - temp - 1) / 26;
  }
  return letter;
}

/**
 * Установить для всего листа КБ формулы FILTER вместо ссылок на TL_*.
 * Запускается один раз с меню после настройки букв колонок БД.ОП.
 */
function migrateKbSheetValidations() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!isKbSheet_(sheet.getName())) {
    SpreadsheetApp.getUi().alert('Откройте лист КБ (КБ1, КБ2, …).');
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Миграция без TL_КБ',
    'Введите букву колонки «Название» на листе БД.ОП (например A):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  var sketchDbCol = resp.getResponseText().trim().toUpperCase();
  resp = ui.prompt('Буква колонки «Номер» на БД.ОП (например B):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  var numberDbCol = resp.getResponseText().trim().toUpperCase();

  var colMap = resolveKbColumns_(sheet);
  var lastRow = Math.max(sheet.getLastRow(), CONFIG.kbDataStartRow + 50);
  var sketchLetter = columnIndexToLetter_(colMap.sketch);

  for (var r = CONFIG.kbDataStartRow; r <= lastRow; r++) {
    applyNumberValidationForRow_(sheet, r, colMap);
  }

  ui.alert(
    'Списки «Номер» обновлены по текущему эскизу в каждой строке.\n\n' +
      'Для полностью динамического списка без скрипта задайте вручную проверку данных:\n' +
      '=SORT(UNIQUE(FILTER(' + CONFIG.dbSheet + '!' + numberDbCol + ':' + numberDbCol + ',' +
      CONFIG.dbSheet + '!' + sketchDbCol + ':' + sketchDbCol + '=$' + sketchLetter + '2)))\n' +
      '(относительная ссылка на строку). Листы TL_* можно скрыть или удалить.'
  );
}
