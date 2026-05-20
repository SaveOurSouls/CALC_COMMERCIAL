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
    'Кол. 6,7 — из БД.ОП (подготовка, расход на настройку).\n' +
    'Кол. 8 — суммарное время операции (K' + r + '), кол. 9 — время машины (L' + r + ').\n' +
    '• Погонный: 8 и 9 = (1/скорость_проката+скорость_инстр×кол-во)×L×N\n' +
    '• Переменный/Статичный: 8 = Время Чел×N; 9 = Время машины×N\n\n' +
    'Цена (N' + r + '):\n' +
    '• Погонный: K' + r + '*P' + r + '+K' + r + '*Уд.Цена ЧЛ_МАГ\n' +
    '• Переменный/Статичный: K' + r + '*Уд.Цена ЧЛ_МАГ+L' + r + '*P' + r + '\n\n' +
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
