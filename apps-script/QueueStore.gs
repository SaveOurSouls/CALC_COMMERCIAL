/**
 * Очередь операций на сервере (общая для Sidebar и плавающего окна).
 */

var QUEUE_PROP_KEY = 'kbOpQueueV1';

/**
 * @returns {{queue: Array, nextId: number, sheetName: string}}
 */
function loadQueue_() {
  var raw = PropertiesService.getUserProperties().getProperty(QUEUE_PROP_KEY);
  if (!raw) {
    return { queue: [], nextId: 1, sheetName: '' };
  }
  try {
    var data = JSON.parse(raw);
    return {
      queue: data.queue || [],
      nextId: data.nextId || 1,
      sheetName: data.sheetName || ''
    };
  } catch (e) {
    return { queue: [], nextId: 1, sheetName: '' };
  }
}

/**
 * @param {{queue: Array, nextId: number, sheetName: string}} data
 */
function saveQueue_(data) {
  var payload = {
    queue: data.queue || [],
    nextId: data.nextId || 1,
    sheetName: data.sheetName || ''
  };
  PropertiesService.getUserProperties().setProperty(QUEUE_PROP_KEY, JSON.stringify(payload));
}

/**
 * @returns {Object}
 */
function apiGetQueueState() {
  var data = loadQueue_();
  var sheets = [];
  try {
    sheets = listKbSheetNames_();
  } catch (e) {
    sheets = [];
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
 * @param {Object} data
 * @returns {Object}
 */
function apiSaveQueue(data) {
  saveQueue_({
    queue: data.queue || [],
    nextId: data.nextId || 1,
    sheetName: data.sheetName || ''
  });
  return { count: (data.queue || []).length };
}

/**
 * @returns {Object}
 */
function apiClearQueue() {
  PropertiesService.getUserProperties().deleteProperty(QUEUE_PROP_KEY);
  return { count: 0 };
}
