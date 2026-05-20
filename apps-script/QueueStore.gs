/**
 * Очередь операций (кэш + UserProperties) и состояние плавающего окна.
 */

var QUEUE_PROP_KEY = 'kbOpQueueV1';
var QUEUE_DIALOG_PING_KEY = 'kbQueueDialogPing';
/** Окно считается открытым, если диалог прислал ping не старше N мс */
var QUEUE_DIALOG_OPEN_MS = 8000;

function clearQueueDialogOpenFlag_() {
  PropertiesService.getUserProperties().deleteProperty(QUEUE_DIALOG_PING_KEY);
}

/**
 * @returns {{queue: Array, nextId: number, sheetName: string, sheetsCache: string[]}}
 */
function loadQueue_() {
  var cache = CacheService.getUserCache();
  var raw = cache.get(QUEUE_PROP_KEY);
  if (!raw) {
    raw = PropertiesService.getUserProperties().getProperty(QUEUE_PROP_KEY);
    if (raw) {
      cache.put(QUEUE_PROP_KEY, raw, 21600);
    }
  }
  if (!raw) {
    return { queue: [], nextId: 1, sheetName: '', sheetsCache: [] };
  }
  try {
    var data = JSON.parse(raw);
    return {
      queue: data.queue || [],
      nextId: data.nextId || 1,
      sheetName: data.sheetName || '',
      sheetsCache: data.sheetsCache || []
    };
  } catch (e) {
    return { queue: [], nextId: 1, sheetName: '', sheetsCache: [] };
  }
}

/**
 * @param {{queue: Array, nextId: number, sheetName: string, sheetsCache: string[]}} data
 */
function saveQueue_(data) {
  var payload = {
    queue: data.queue || [],
    nextId: data.nextId || 1,
    sheetName: data.sheetName || '',
    sheetsCache: data.sheetsCache || []
  };
  var json = JSON.stringify(payload);
  CacheService.getUserCache().put(QUEUE_PROP_KEY, json, 21600);
  PropertiesService.getUserProperties().setProperty(QUEUE_PROP_KEY, json);
}

function markQueueDialogPing_() {
  PropertiesService.getUserProperties().setProperty(QUEUE_DIALOG_PING_KEY, String(Date.now()));
}

/**
 * @returns {boolean}
 */
function isQueueDialogOpen_() {
  var raw = PropertiesService.getUserProperties().getProperty(QUEUE_DIALOG_PING_KEY);
  if (!raw) {
    return false;
  }
  var ts = parseInt(raw, 10);
  if (isNaN(ts)) {
    clearQueueDialogOpenFlag_();
    return false;
  }
  return (Date.now() - ts) < QUEUE_DIALOG_OPEN_MS;
}

/**
 * @returns {Object}
 */
function apiIsQueueDialogOpen() {
  return { open: isQueueDialogOpen_() };
}

/**
 * @returns {Object}
 */
function apiQueueDialogPing() {
  markQueueDialogPing_();
  return { open: true };
}

/**
 * @returns {Object}
 */
function apiQueueDialogClose() {
  clearQueueDialogOpenFlag_();
  return { open: false };
}

/**
 * Сброс «окно открыто» при F5 / открытии sidebar (диалог при этом уже закрыт).
 *
 * @returns {Object}
 */
function apiSidebarReady() {
  clearQueueDialogOpenFlag_();
  var data = loadQueue_();
  return { count: data.queue.length, dialogOpen: false };
}

/**
 * Только чтение очереди — без ping (для счётчика в sidebar).
 *
 * @returns {Object}
 */
function apiGetQueueCount() {
  var data = loadQueue_();
  return { count: data.queue.length };
}

/**
 * @returns {Object}
 */
function buildQueueStateResponse_(data) {
  var sheets = data.sheetsCache;
  if (!sheets || !sheets.length) {
    try {
      sheets = listKbSheetNames_();
    } catch (e) {
      sheets = [];
    }
  }
  var active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var sheetName = data.sheetName;
  if (!sheetName || sheets.indexOf(sheetName) < 0) {
    sheetName = sheets.indexOf(active) >= 0 ? active : (sheets[0] || active);
  }
  return {
    queue: data.queue,
    nextId: data.nextId,
    count: data.queue.length,
    sheetName: sheetName,
    sheets: sheets,
    activeSheet: active
  };
}

/**
 * @returns {Object}
 */
function apiGetQueueState() {
  return buildQueueStateResponse_(loadQueue_());
}

/**
 * Только для плавающего окна (ставит ping — окно реально загрузилось).
 *
 * @returns {Object}
 */
function apiQueueDialogBootstrap() {
  markQueueDialogPing_();
  var data = loadQueue_();
  if (!data.sheetsCache || !data.sheetsCache.length) {
    try {
      data.sheetsCache = listKbSheetNames_();
      saveQueue_(data);
    } catch (e) {
      data.sheetsCache = [];
    }
  }
  var resp = buildQueueStateResponse_(data);
  var index = getDbOpsSidebarIndex_();
  resp.sketches = index.sketches;
  resp.opsBySketch = index.opsBySketch;
  return resp;
}

/**
 * @param {Object} data
 * @returns {Object}
 */
function apiSaveQueue(data) {
  var current = loadQueue_();
  saveQueue_({
    queue: data.queue || [],
    nextId: data.nextId || 1,
    sheetName: data.sheetName !== undefined ? data.sheetName : current.sheetName,
    sheetsCache: current.sheetsCache
  });
  return { count: (data.queue || []).length };
}

/**
 * @param {Object} payload {items: Array, sheetName: string}
 * @returns {Object}
 */
function apiAppendToQueue(payload) {
  var data = loadQueue_();
  var items = payload.items || [];
  if (payload.sheetName) {
    data.sheetName = payload.sheetName;
  }
  items.forEach(function (item) {
    data.queue.push({
      id: data.nextId++,
      sketch: item.sketch,
      number: item.number,
      label: item.label || item.number,
      opType: item.opType || '',
      n: item.n,
      l: item.l,
      op: item.op || ''
    });
  });
  saveQueue_(data);
  return { count: data.queue.length, nextId: data.nextId };
}

/**
 * @returns {Object}
 */
function apiClearQueue() {
  var data = loadQueue_();
  saveQueue_({
    queue: [],
    nextId: 1,
    sheetName: data.sheetName,
    sheetsCache: data.sheetsCache
  });
  return { count: 0 };
}
