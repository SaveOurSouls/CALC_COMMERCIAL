/**
 * Поиск заголовков КБ и БД.ОП (с нормализацией и автоопределением строки).
 */

var DB_HEADER_CACHE_KEY = 'kbDbHeaderInfoV1';

/**
 * @param {*} s
 * @returns {string}
 */
function normalizeHeaderKey_(s) {
  return String(s || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

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
  return buildHeaderMapFromRow_(sheet, row, lastCol);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {number} lastCol
 * @returns {Object.<string, number>}
 */
function buildHeaderMapFromRow_(sheet, row, lastCol) {
  var headers = sheet.getRange(row, 1, row, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    var exact = String(h || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    if (exact) {
      map[exact] = i + 1;
    }
  });
  return map;
}

/**
 * @returns {{headerRow: number, dataStartRow: number, map: Object.<string, number>}}
 */
function getDbHeaderInfo_() {
  var cache = CacheService.getDocumentCache();
  var cached = cache.get(DB_HEADER_CACHE_KEY);
  if (cached) {
    return JSON.parse(cached);
  }
  var info = detectDbHeaderInfo_();
  cache.put(DB_HEADER_CACHE_KEY, JSON.stringify(info), 21600);
  return info;
}

function clearDbHeaderCache_() {
  CacheService.getDocumentCache().remove(DB_HEADER_CACHE_KEY);
  invalidateDbOpsCache_();
}

/**
 * @returns {Object.<string, number>}
 */
function getDbHeaderMap_() {
  return getDbHeaderInfo_().map;
}

/**
 * @returns {number}
 */
function getDbDataStartRow_() {
  return getDbHeaderInfo_().dataStartRow;
}

/**
 * @param {string[]} headers
 * @param {string[]} needles нормализованные подстроки
 * @returns {number} 1-based col or 0
 */
function findHeaderColumn_(headers, needles) {
  for (var i = 0; i < headers.length; i++) {
    var h = normalizeHeaderKey_(headers[i]);
    if (!h) {
      continue;
    }
    for (var j = 0; j < needles.length; j++) {
      if (h === needles[j] || h.indexOf(needles[j]) >= 0) {
        return i + 1;
      }
    }
  }
  return 0;
}

/**
 * @returns {{headerRow: number, dataStartRow: number, map: Object.<string, number>}}
 */
function detectDbHeaderInfo_() {
  var sheet = getDbSheet_();
  var lastCol = sheet.getLastColumn();
  var maxScan = CONFIG.dbHeaderScanMaxRow || 15;
  var sketchNeedles = getDbAliasNeedles_('sketch');
  var numberNeedles = getDbAliasNeedles_('number');

  for (var row = 1; row <= maxScan; row++) {
    var headers = sheet.getRange(row, 1, row, lastCol).getValues()[0];
    var numberCol = findHeaderColumn_(headers, numberNeedles);
    var sketchCol = findHeaderColumn_(headers, sketchNeedles);

    if (numberCol && sketchCol) {
      return {
        headerRow: row,
        dataStartRow: row + 1,
        map: buildHeaderMapFromRow_(sheet, row, lastCol),
        numberCol: numberCol,
        sketchCol: sketchCol
      };
    }

    if (numberCol && !sketchCol) {
      sketchCol = numberCol === 1 ? 2 : 1;
      return {
        headerRow: row,
        dataStartRow: row + 1,
        map: buildHeaderMapFromRow_(sheet, row, lastCol),
        numberCol: numberCol,
        sketchCol: sketchCol
      };
    }
  }

  var diag = getDbHeaderDiagnostics_();
  throw new Error(
    'Не найдена строка заголовков на листе «' + CONFIG.dbSheet + '» с колонкой «Номер».\n' + diag
  );
}

/**
 * @param {string} fieldKey
 * @returns {string[]}
 */
function getDbAliasNeedles_(fieldKey) {
  var list = [];
  if (CONFIG.db[fieldKey]) {
    list.push(normalizeHeaderKey_(CONFIG.db[fieldKey]));
  }
  var aliases = CONFIG.dbAliases && CONFIG.dbAliases[fieldKey];
  if (aliases) {
    aliases.forEach(function (a) {
      list.push(normalizeHeaderKey_(a));
    });
  }
  return list;
}

/**
 * @param {Object.<string, number>} headerMap
 * @param {string} fieldKey
 * @returns {number}
 */
function resolveDbColumn_(headerMap, fieldKey) {
  var candidates = [];
  if (CONFIG.db[fieldKey]) {
    candidates.push(String(CONFIG.db[fieldKey]).trim());
  }
  var aliases = CONFIG.dbAliases && CONFIG.dbAliases[fieldKey];
  if (aliases) {
    aliases.forEach(function (a) {
      candidates.push(String(a).trim());
    });
  }

  var i;
  var exact;
  for (i = 0; i < candidates.length; i++) {
    exact = candidates[i];
    if (headerMap[exact]) {
      return headerMap[exact];
    }
  }

  var normMap = {};
  Object.keys(headerMap).forEach(function (k) {
    normMap[normalizeHeaderKey_(k)] = headerMap[k];
  });

  for (i = 0; i < candidates.length; i++) {
    var norm = normalizeHeaderKey_(candidates[i]);
    if (norm && normMap[norm]) {
      return normMap[norm];
    }
  }

  for (i = 0; i < candidates.length; i++) {
    var needle = normalizeHeaderKey_(candidates[i]);
    if (!needle) {
      continue;
    }
    var keys = Object.keys(normMap);
    for (var k = 0; k < keys.length; k++) {
      if (keys[k].indexOf(needle) >= 0 || needle.indexOf(keys[k]) >= 0) {
        return normMap[keys[k]];
      }
    }
  }

  if (fieldKey === 'sketch') {
    var info = getDbHeaderInfo_();
    if (info.sketchCol) {
      return info.sketchCol;
    }
  }
  if (fieldKey === 'number') {
    var infoNum = getDbHeaderInfo_();
    if (infoNum.numberCol) {
      return infoNum.numberCol;
    }
  }

  throw new Error(
    'В БД.ОП нет колонки «' + CONFIG.db[fieldKey] + '». ' +
    'Меню: Техкарта КБ → Диагностика заголовков. ' +
    'Проверьте Config.gs → db / dbAliases.'
  );
}

/**
 * @returns {string}
 */
function getDbHeaderDiagnostics_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = ['Листы в книге:'];
  ss.getSheets().forEach(function (s) {
    lines.push(' • ' + s.getName());
  });

  var sheet = null;
  try {
    sheet = getDbSheet_();
  } catch (e) {
    lines.push('\nЛист «' + CONFIG.dbSheet + '» не найден.');
    return lines.join('\n');
  }

  lines.push('\nПервые ' + (CONFIG.dbHeaderScanMaxRow || 15) + ' строк листа «' + sheet.getName() + '»:');
  var maxRow = CONFIG.dbHeaderScanMaxRow || 15;
  var lastCol = Math.min(sheet.getLastColumn(), 12);
  for (var r = 1; r <= maxRow; r++) {
    var row = sheet.getRange(r, 1, r, lastCol).getValues()[0];
    var parts = row.map(function (c, i) {
      var v = String(c || '').trim();
      return v ? columnIndexToLetter_(i + 1) + ':' + v : '';
    }).filter(function (p) {
      return p;
    });
    if (parts.length) {
      lines.push('Строка ' + r + ': ' + parts.join(' | '));
    }
  }

  try {
    var info = detectDbHeaderInfo_();
    lines.push('\nАвтоопределение: заголовки в строке ' + info.headerRow +
      ', Номер=' + columnIndexToLetter_(info.numberCol) +
      ', Эскиз/Название=' + columnIndexToLetter_(info.sketchCol));
  } catch (e2) {
    lines.push('\nАвтоопределение не удалось: ' + e2.message);
  }

  return lines.join('\n');
}

function showDbHeaderDiagnostics() {
  clearDbHeaderCache_();
  SpreadsheetApp.getUi().alert('Диагностика БД.ОП', getDbHeaderDiagnostics_(),
    SpreadsheetApp.getUi().ButtonSet.OK);
}
