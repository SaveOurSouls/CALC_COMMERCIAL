/**
 * Чтение и индекс справочника БД.ОП.
 */

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getDbSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = [CONFIG.dbSheet];
  if (CONFIG.dbSheetAltNames) {
    names = names.concat(CONFIG.dbSheetAltNames);
  }

  for (var i = 0; i < names.length; i++) {
    var sheet = ss.getSheetByName(names[i]);
    if (sheet) {
      return sheet;
    }
  }

  var target = normalizeHeaderKey_(CONFIG.dbSheet);
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    if (normalizeHeaderKey_(sheets[s].getName()) === target) {
      return sheets[s];
    }
  }

  throw new Error('Лист «' + CONFIG.dbSheet + '» не найден. Откройте «Диагностика заголовков» в меню.');
}

/**
 * @param {Object.<string, number>} headerMap
 * @param {string} fieldKey
 * @returns {number}
 */
function dbCol_(headerMap, fieldKey) {
  return resolveDbColumn_(headerMap, fieldKey);
}

/**
 * @returns {Array.<Object>}
 */
function loadDbOperations_() {
  var sheet = getDbSheet_();
  var headerMap = getDbHeaderMap_();
  var start = getDbDataStartRow_();
  var lastRow = sheet.getLastRow();
  if (lastRow < start) {
    return [];
  }

  var fields = [
    'sketch', 'number', 'opType', 'rollSpeed', 'toolWorkSpeed', 'toolOpCount',
    'timeHuman', 'timeMachine', 'unitPriceHuman', 'unitPriceMachine', 'unitPriceHumanMag',
    'prepTime', 'setupConsumption'
  ];

  var required = ['sketch', 'number'];
  var optional = fields.filter(function (f) {
    return required.indexOf(f) < 0;
  });

  var maxCol = sheet.getLastColumn();
  var values = sheet.getRange(start, 1, lastRow - start + 1, maxCol).getValues();
  var colIndex = { sketch: dbCol_(headerMap, 'sketch'), number: dbCol_(headerMap, 'number') };

  optional.forEach(function (fieldKey) {
    try {
      colIndex[fieldKey] = dbCol_(headerMap, fieldKey);
    } catch (e) {
      colIndex[fieldKey] = 0;
    }
  });

  var out = [];
  values.forEach(function (row, idx) {
    var sketch = String(row[colIndex.sketch - 1] || '').trim();
    var number = String(row[colIndex.number - 1] || '').trim();
    if (!sketch && !number) {
      return;
    }

    function cell(field) {
      var c = colIndex[field];
      return c ? row[c - 1] : '';
    }

    var rec = {
      rowIndex: start + idx,
      sketch: sketch,
      number: number,
      opType: String(cell('opType') || '').trim(),
      rollSpeed: num_(cell('rollSpeed')),
      toolWorkSpeed: num_(cell('toolWorkSpeed')),
      toolOpCount: num_(cell('toolOpCount')),
      timeHuman: num_(cell('timeHuman')),
      timeMachine: num_(cell('timeMachine')),
      unitPriceHuman: num_(cell('unitPriceHuman')),
      unitPriceMachine: num_(cell('unitPriceMachine')),
      unitPriceHumanMag: num_(cell('unitPriceHumanMag')),
      prepTime: num_(cell('prepTime')),
      setupConsumption: num_(cell('setupConsumption'))
    };
    rec.key = rec.sketch + '\u0001' + rec.number;
    out.push(rec);
  });
  return out;
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
  var list = loadDbOperations_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].sketch === sketch && list[i].number === number) {
      return list[i];
    }
  }
  return null;
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
