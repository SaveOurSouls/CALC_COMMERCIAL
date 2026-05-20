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
  var r = CONFIG.kbDataStartRow;
  var msg =
    'Структура КБ: заголовки в строке ' + CONFIG.kbHeaderRow + ', данные с ' + r + '.\n' +
    'Эскиз проекта (A) = БД.ОП «Название» (B); Номер (B) = БД.ОП «Номер» (A).\n\n' +
    'Колонка «8» (время), пример для строки ' + r + ' (K' + r + '):\n' +
    '• Погонный: (1/I' + r + '+J' + r + '*по_кол-ву)×E' + r + '*D' + r + '\n' +
    '• Переменный/Статичный: I' + r + '×D' + r + ' (6=Время Чел, 7=Время машины — см. kbDbPull)\n\n' +
    'Цена (N' + r + '):\n' +
    '• Погонный: K' + r + '*P' + r + '+K' + r + '*Уд.Цена ЧЛ_МАГ из БД\n' +
    '• Переменный/Статичный: K' + r + '*Уд.Цена ЧЛ_МАГ+7' + r + '*D' + r + '*P' + r + '\n\n' +
    'Список «Номер» без TL_КБ:\n' +
    '=SORT(UNIQUE(FILTER(БД.ОП!$A:$A,БД.ОП!$B:$B=$A' + r + ')))';
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
