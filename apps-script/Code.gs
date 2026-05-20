/**
 * Точка входа: меню, триггер onEdit, открытие Sidebar.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Техкарта КБ')
    .addItem('Панель быстрого набора', 'showKbSidebar')
    .addItem('Пересчитать время и цену (активный лист)', 'recalcActiveKbSheet')
    .addItem('Миграция: убрать TL_КБ (валидация FILTER)', 'migrateKbSheetValidations')
    .addSeparator()
    .addItem('Справка по формулам', 'showFormulaHelp')
    .addToUi();
}

/**
 * @param {Object} e
 */
function onEdit(e) {
  try {
    onEditKbHandler_(e);
  } catch (err) {
    console.error(err);
  }
}

function showKbSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Быстрый набор операций')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showFormulaHelp() {
  var msg =
    'Колонка 8 (время):\n' +
    '• Погонный: (1/Скорость проката + Скорость работы инструмента × Кол-во операций инструмента) × L × N\n' +
    '• Переменный / Статичный: Время Чел, сек/оп; сек/м × N\n\n' +
    'Цена:\n' +
    '• Погонный: время × Уд.Цена МШ + время × Уд.Цена ЧЛ_МАГ\n' +
    '• Переменный / Статичный: время × ЧЛ_МАГ + Время машины × N × Уд.Цена МШ\n\n' +
    'Зависимый список «Номер» без TL_КБ (вставить в Проверка данных → диапазон):\n' +
    '=SORT(UNIQUE(FILTER(БД.ОП!$B:$B,БД.ОП!$A:$A=$B2)))\n' +
    '(замените A/B на буквы «Название» и «Номер» в вашей книге; $B2 — ячейка «Эскиз проекта» в этой строке).';
  SpreadsheetApp.getUi().alert('Формулы', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/** --- API для Sidebar (google.script.run) --- */

function apiGetInitialData() {
  return {
    sheets: listKbSheetNames_(),
    sketches: listSketchNames_(),
    activeSheet: SpreadsheetApp.getActiveSheet().getName()
  };
}

/**
 * @param {string} sketch
 * @returns {Array.<Object>}
 */
function apiGetOperations(sketch) {
  return listOperationsBySketch_(sketch).map(function (op) {
    return {
      number: op.number,
      opType: op.opType,
      label: op.number + ' — ' + op.opType
    };
  });
}

/**
 * @param {Object} payload
 */
function apiInsertRows(payload) {
  return insertKbRows_(payload);
}

/**
 * @param {string} sheetName
 * @param {number} count
 */
function apiDuplicateLast(sheetName, count) {
  return duplicateLastKbRow_(sheetName, count);
}

/**
 * @param {string} sheetName
 * @param {number} row
 * @param {number} n
 * @param {number} l
 */
function apiUpdateRowDimensions(sheetName, row, n, l) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var colMap = resolveKbColumns_(sheet);
  sheet.getRange(row, colMap.n).setValue(n);
  sheet.getRange(row, colMap.l).setValue(l);
  recalcKbRow_(sheet, row, colMap);
  return { ok: true };
}
