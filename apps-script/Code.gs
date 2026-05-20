/**
 * Точка входа: меню, триггер onEdit, открытие Sidebar.
 */

function onOpen() {
  clearQueueDialogOpenFlag_();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Техкарта КБ')
    .addItem('Справочник (боковая панель)', 'showKbSidebar')
    .addItem('Очередь операций (плавающее окно)', 'showKbQueueDialog')
    .addItem('Сохранить активный лист в шаблон КБ', 'addActiveSheetToKbTemplate')
    .addItem('Пересчитать время и цену (активный лист)', 'recalcActiveKbSheet')
    .addItem('Миграция: убрать TL_КБ (валидация FILTER)', 'migrateKbSheetValidations')
    .addItem('Диагностика заголовков БД.ОП', 'showDbHeaderDiagnostics')
    .addSeparator()
    .addItem('Справка по формулам', 'showFormulaHelp')
    .addToUi();
}

/**
 * @param {Object} e
 */
function onEdit(e) {
  try {
    if (e && e.range) {
      var sheetName = normalizeHeaderKey_(e.range.getSheet().getName());
      if (sheetName === normalizeHeaderKey_(CONFIG.dbSheet)) {
        invalidateDbOpsCache_();
        clearDbHeaderCache_();
      }
    }
    onEditKbHandler_(e);
  } catch (err) {
    console.error(err);
  }
}

function showKbSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Справочник операций')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Плавающее окно очереди (не блокирует работу с листом).
 */
function showKbQueueDialog() {
  if (isQueueDialogOpen_()) {
    return { alreadyOpen: true };
  }
  if (!isAppsScriptHtmlFileValid_('QueueDialog')) {
    SpreadsheetApp.getUi().alert(
      'Ошибка: файл QueueDialog должен быть типа HTML, не .gs.\n\n' +
        'В редакторе Apps Script: «+» → HTML → имя QueueDialog → вставьте QueueDialog.html из репозитория.\n' +
        'Также нужны: SemifinishedPresets.gs, Sidebar.html.'
    );
    return { error: 'invalid_queue_dialog_html' };
  }
  var html = HtmlService.createHtmlOutputFromFile('QueueDialog')
    .setWidth(920)
    .setHeight(1360);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Очередь операций');
  return { alreadyOpen: false };
}

/**
 * @param {string} name имя HTML-файла без расширения
 * @returns {boolean}
 */
function isAppsScriptHtmlFileValid_(name) {
  try {
    var raw = HtmlService.createTemplateFromFile(name).getRawContent();
    var s = String(raw || '').trim();
    if (s.indexOf('function loadQueue_') >= 0 || s.indexOf('QUEUE_PROP_KEY') >= 0) {
      return false;
    }
    return s.indexOf('<') === 0 || s.indexOf('<!') === 0;
  } catch (e) {
    return false;
  }
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
  var active = SpreadsheetApp.getActiveSheet().getName();
  var sheets = [];
  try {
    sheets = listKbSheetNames_();
  } catch (e1) {
    sheets = [];
  }

  try {
    var info = getDbHeaderInfo_();
    var sidebarIndex = getDbOpsSidebarIndex_();
    var queueCount = loadQueue_().queue.length;

    var presets = getSemifinishedPresetsIndex_();

    return {
      sheets: sheets,
      sketches: sidebarIndex.sketches,
      opsBySketch: sidebarIndex.opsBySketch,
      presets: presets,
      queueCount: queueCount,
      activeSheet: active,
      dbHeaderRow: info.headerRow,
      hint: sheets.length ? '' : 'Нет листов КБ1, КБ2… Проверьте имена листов.'
    };
  } catch (err) {
    return {
      sheets: sheets,
      sketches: [],
      presets: { groups: [], byGroup: {} },
      activeSheet: active,
      error: err.message,
      hint: 'Меню → Техкарта КБ → Диагностика заголовков БД.ОП'
    };
  }
}

/**
 * @param {string} sketch
 * @returns {Array.<Object>}
 */
function apiGetOperations(sketch) {
  var index = getDbOpsSidebarIndex_();
  return index.opsBySketch[sketch] || [];
}

/**
 * @param {Object} payload
 */
function apiInsertRows(payload) {
  return insertKbRows_(payload);
}

/**
 * @param {Object} payload {sheetName, rows: Array}
 */
function apiInsertQueue(payload) {
  return insertKbQueue_(payload);
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
  if (CONFIG.kbEnableScriptRecalc) {
    recalcKbRow_(sheet, row, colMap);
  }
  return { ok: true };
}

/**
 * Копирует текущий открытый лист в архив шаблонов (КБ.ШАБЛ.N), формулы сохраняются.
 */
function apiAddActiveSheetToKbTemplate() {
  return addActiveSheetToKbTemplate_();
}
