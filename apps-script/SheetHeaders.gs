/**
 * Чтение строк заголовков КБ и БД.ОП.
 */

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object.<string, number>}
 */
function getKbHeaderMap_(sheet) {
  var row = CONFIG.kbHeaderRow;
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    throw new Error('Лист «' + sheet.getName() + '» пуст.');
  }
  var headers = sheet.getRange(row, 1, row, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    var key = String(h || '').trim();
    if (key) {
      map[key] = i + 1;
    }
  });
  return map;
}

/**
 * @returns {Object.<string, number>}
 */
function getDbHeaderMap_() {
  var sheet = getDbSheet_();
  var row = CONFIG.dbHeaderRow;
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    throw new Error('Лист ' + CONFIG.dbSheet + ' пуст.');
  }
  var headers = sheet.getRange(row, 1, row, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    var key = String(h || '').trim();
    if (key) {
      map[key] = i + 1;
    }
  });
  return map;
}
