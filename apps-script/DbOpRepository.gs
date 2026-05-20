/**
 * Чтение и индекс справочника БД.ОП.
 */

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getDbSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.dbSheet);
  if (!sheet) {
    throw new Error('Лист "' + CONFIG.dbSheet + '" не найден.');
  }
  return sheet;
}

/**
 * @returns {Object.<string, number>} заголовок -> индекс колонки (1-based)
 */
function getDbHeaderMap_() {
  var sheet = getDbSheet_();
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    throw new Error('Лист ' + CONFIG.dbSheet + ' пуст.');
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
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
 * @param {Object.<string, number>} headerMap
 * @param {string} fieldKey ключ из CONFIG.db
 * @returns {number}
 */
function dbCol_(headerMap, fieldKey) {
  var title = CONFIG.db[fieldKey];
  var col = headerMap[title];
  if (!col) {
    throw new Error('В БД.ОП нет колонки «' + title + '». Проверьте Config.gs.');
  }
  return col;
}

/**
 * @returns {Array.<Object>}
 */
function loadDbOperations_() {
  var sheet = getDbSheet_();
  var headerMap = getDbHeaderMap_();
  var start = CONFIG.dbDataStartRow;
  var lastRow = sheet.getLastRow();
  if (lastRow < start) {
    return [];
  }

  var cols = [
    'sketch', 'number', 'opType', 'rollSpeed', 'toolWorkSpeed', 'toolOpCount',
    'timeHuman', 'timeMachine', 'unitPriceHuman', 'unitPriceMachine',
    'unitPriceHumanMag', 'humanMag'
  ];
  var colIndexes = cols.map(function (k) {
    return dbCol_(headerMap, k);
  });
  var maxCol = Math.max.apply(null, colIndexes);
  var width = maxCol;
  var values = sheet.getRange(start, 1, lastRow - start + 1, width).getValues();

  var out = [];
  values.forEach(function (row, idx) {
    var sketch = String(row[dbCol_(headerMap, 'sketch') - 1] || '').trim();
    var number = String(row[dbCol_(headerMap, 'number') - 1] || '').trim();
    if (!sketch && !number) {
      return;
    }
    var rec = {
      rowIndex: start + idx,
      sketch: sketch,
      number: number,
      opType: String(row[dbCol_(headerMap, 'opType') - 1] || '').trim(),
      rollSpeed: num_(row[dbCol_(headerMap, 'rollSpeed') - 1]),
      toolWorkSpeed: num_(row[dbCol_(headerMap, 'toolWorkSpeed') - 1]),
      toolOpCount: num_(row[dbCol_(headerMap, 'toolOpCount') - 1]),
      timeHuman: num_(row[dbCol_(headerMap, 'timeHuman') - 1]),
      timeMachine: num_(row[dbCol_(headerMap, 'timeMachine') - 1]),
      unitPriceHuman: num_(row[dbCol_(headerMap, 'unitPriceHuman') - 1]),
      unitPriceMachine: num_(row[dbCol_(headerMap, 'unitPriceMachine') - 1]),
      unitPriceHumanMag: num_(row[dbCol_(headerMap, 'unitPriceHumanMag') - 1]),
      humanMag: num_(row[dbCol_(headerMap, 'humanMag') - 1])
    };
    rec.key = rec.sketch + '\u0001' + rec.number;
    out.push(rec);
  });
  return out;
}

/**
 * @returns {Object.<string, Object>}
 */
function buildDbIndex_() {
  var list = loadDbOperations_();
  var index = {};
  list.forEach(function (rec) {
    index[rec.key] = rec;
  });
  return index;
}

/**
 * @returns {string[]}
 */
function listSketchNames_() {
  var list = loadDbOperations_();
  var set = {};
  list.forEach(function (r) {
    if (r.sketch) {
      set[r.sketch] = true;
    }
  });
  return Object.keys(set).sort();
}

/**
 * @param {string} sketch
 * @returns {Array.<Object>}
 */
function listOperationsBySketch_(sketch) {
  return loadDbOperations_().filter(function (r) {
    return r.sketch === sketch;
  });
}

/**
 * @param {string} sketch
 * @param {string} number
 * @returns {Object|null}
 */
function findOperation_(sketch, number) {
  var index = buildDbIndex_();
  return index[sketch + '\u0001' + number] || null;
}

/**
 * @param {*} v
 * @returns {number}
 */
function num_(v) {
  if (v === '' || v === null || v === undefined) {
    return 0;
  }
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}
