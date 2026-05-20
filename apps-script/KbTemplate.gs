/**
 * Шаблоны КБ: копии листов с сохранением формул (отдельные вкладки КБ.ШАБЛ.N).
 */

/**
 * @param {string} name
 * @returns {boolean}
 */
function isKbTemplateSheet_(name) {
  var prefix = CONFIG.kbTemplatePrefix || 'КБ.ШАБЛ.';
  return String(name || '').indexOf(prefix) === 0;
}

/**
 * @returns {number}
 */
function getNextKbTemplateNumber_() {
  var prefix = CONFIG.kbTemplatePrefix || 'КБ.ШАБЛ.';
  var re = new RegExp('^' + escapeRegex_(prefix) + '(\\d+)$', 'i');
  var max = 0;
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (s) {
    var m = String(s.getName()).match(re);
    if (m) {
      max = Math.max(max, parseInt(m[1], 10));
    }
  });
  return max + 1;
}

/**
 * @param {string} s
 * @returns {string}
 */
function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * @returns {string}
 */
function getNextKbTemplateSheetName_() {
  return (CONFIG.kbTemplatePrefix || 'КБ.ШАБЛ.') + getNextKbTemplateNumber_();
}

/**
 * Уникальное имя вкладки (до 100 символов).
 *
 * @param {string} base
 * @returns {string}
 */
function ensureUniqueSheetName_(base) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = String(base).substring(0, 100);
  if (!ss.getSheetByName(name)) {
    return name;
  }
  for (var i = 2; i < 1000; i++) {
    var candidate = (String(base).substring(0, 95) + '_' + i).substring(0, 100);
    if (!ss.getSheetByName(candidate)) {
      return candidate;
    }
  }
  throw new Error('Не удалось подобрать уникальное имя листа.');
}

/**
 * Скопировать активный лист в шаблон (copyTo сохраняет формулы и оформление).
 *
 * @returns {{sheetName: string, sourceName: string, hint: string}}
 */
function addActiveSheetToKbTemplate_() {
  var source = SpreadsheetApp.getActiveSheet();
  if (!source) {
    throw new Error('Нет активного листа.');
  }

  var ss = source.getParent();
  var sourceName = source.getName();
  var newName = ensureUniqueSheetName_(getNextKbTemplateSheetName_());

  var copied = source.copyTo(ss);
  copied.setName(newName);
  ss.setActiveSheet(copied);

  return {
    sheetName: newName,
    sourceName: sourceName,
    hint: 'Лист «' + sourceName + '» сохранён как шаблон «' + newName + '» (формулы сохранены).'
  };
}

/**
 * @returns {string[]}
 */
function listKbTemplateSheetNames_() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(function (s) {
      return s.getName();
    })
    .filter(isKbTemplateSheet_)
    .sort(function (a, b) {
      return a.localeCompare(b, 'ru', { numeric: true });
    });
}
