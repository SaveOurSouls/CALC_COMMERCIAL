/**
 * Кэш справочника БД.ОП (ускорение Sidebar).
 */

var DB_OPS_CACHE_KEY = 'kbDbOpsCacheV2';
var DB_OPS_CACHE_SEC = 600;

/**
 * @returns {Array.<Object>}
 */
function getCachedDbOperations_() {
  var cache = CacheService.getDocumentCache();
  var raw = cache.get(DB_OPS_CACHE_KEY);
  if (raw) {
    try {
      return JSON.parse(raw);
    } catch (e) {
      cache.remove(DB_OPS_CACHE_KEY);
    }
  }

  var list = loadDbOperations_();
  try {
    cache.put(DB_OPS_CACHE_KEY, JSON.stringify(list), DB_OPS_CACHE_SEC);
  } catch (e2) {
    console.warn('Кэш БД.ОП не записан: ' + e2.message);
  }
  return list;
}

/**
 * @returns {{sketches: string[], opsBySketch: Object.<string, Array>}}
 */
function getDbOpsSidebarIndex_() {
  var list = getCachedDbOperations_();
  var opsBySketch = {};
  list.forEach(function (op) {
    if (!op.sketch) {
      return;
    }
    if (!opsBySketch[op.sketch]) {
      opsBySketch[op.sketch] = [];
    }
    opsBySketch[op.sketch].push({
      number: op.number,
      opType: op.opType,
      label: op.number + ' — ' + op.opType
    });
  });

  var sketches = Object.keys(opsBySketch).sort();
  sketches.forEach(function (s) {
    opsBySketch[s].sort(function (a, b) {
      return String(a.number).localeCompare(String(b.number), 'ru');
    });
  });

  return { sketches: sketches, opsBySketch: opsBySketch };
}

function invalidateDbOpsCache_() {
  CacheService.getDocumentCache().remove(DB_OPS_CACHE_KEY);
}

/**
 * @param {*} value
 * @returns {number}
 */
function parseRowCount_(value) {
  var n = parseInt(String(value), 10);
  if (isNaN(n) || n < 1) {
    return 1;
  }
  return Math.min(n, 50);
}
