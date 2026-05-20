/**
 * Наборы операций по полуфабрикатам (группа + артикул) — лист КБ.ПРЕСЕТЫ в книге.
 */

var PRESET_COL = {
  group: 1,
  article: 2,
  order: 3,
  sketch: 4,
  number: 5,
  n: 6,
  l: 7,
  op: 8,
  updated: 9
};

/**
 * @param {string} s
 * @returns {string}
 */
function normalizePresetKey_(s) {
  return String(s || '').replace(/\s+/g, ' ').trim();
}

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreatePresetsSheet_() {
  var name = CONFIG.kbPresetsSheet || 'КБ.ПРЕСЕТЫ';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, PRESET_COL.updated).setValues([[
      'Группа', 'Артикул', 'Порядок', 'Эскиз', 'Номер', 'N', 'L', 'OP', 'Изменено'
    ]]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    try {
      sheet.hideSheet();
    } catch (eHide) {
      // лист можно оставить видимым
    }
  }
  return sheet;
}

/**
 * @returns {{groups: string[], byGroup: Object.<string, string[]>}}
 */
function getSemifinishedPresetsIndex_() {
  var sheet = getOrCreatePresetsSheet_();
  var lastRow = sheet.getLastRow();
  var byGroup = {};
  if (lastRow < 2) {
    return { groups: [], byGroup: byGroup };
  }

  var values = sheet.getRange(2, 1, lastRow, 2).getValues();
  values.forEach(function (row) {
    var group = normalizePresetKey_(row[0]);
    var article = normalizePresetKey_(row[1]);
    if (!group || !article) {
      return;
    }
    if (!byGroup[group]) {
      byGroup[group] = [];
    }
    if (byGroup[group].indexOf(article) < 0) {
      byGroup[group].push(article);
    }
  });

  var groups = Object.keys(byGroup).sort(function (a, b) {
    return a.localeCompare(b, 'ru');
  });
  groups.forEach(function (g) {
    byGroup[g].sort(function (a, b) {
      return a.localeCompare(b, 'ru');
    });
  });

  return { groups: groups, byGroup: byGroup };
}

/**
 * @param {string} group
 * @param {string} article
 */
function deleteSemifinishedPreset_(group, article) {
  group = normalizePresetKey_(group);
  article = normalizePresetKey_(article);
  var sheet = getOrCreatePresetsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  for (var r = lastRow; r >= 2; r--) {
    var g = normalizePresetKey_(sheet.getRange(r, PRESET_COL.group).getValue());
    var a = normalizePresetKey_(sheet.getRange(r, PRESET_COL.article).getValue());
    if (g === group && a === article) {
      sheet.deleteRow(r);
    }
  }
}

/**
 * @param {string} group
 * @param {string} article
 * @param {Array.<Object>} items
 */
function saveSemifinishedPreset_(group, article, items) {
  group = normalizePresetKey_(group);
  article = normalizePresetKey_(article);
  if (!group) {
    throw new Error('Укажите тип / группу полуфабриката.');
  }
  if (!article) {
    throw new Error('Укажите название артикула полуфабриката.');
  }
  if (!items || !items.length) {
    throw new Error('Очередь пуста — добавьте операции перед записью.');
  }

  deleteSemifinishedPreset_(group, article);

  var sheet = getOrCreatePresetsSheet_();
  var now = new Date();
  var rows = items.map(function (item, i) {
    return [
      group,
      article,
      i + 1,
      String(item.sketch || '').trim(),
      String(item.number || '').trim(),
      num_(item.n),
      num_(item.l),
      String(item.op || ''),
      now
    ];
  });

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, PRESET_COL.updated).setValues(rows);
  SpreadsheetApp.flush();
}

/**
 * @param {string} group
 * @param {string} article
 * @returns {Array.<Object>}
 */
function loadSemifinishedPreset_(group, article) {
  group = normalizePresetKey_(group);
  article = normalizePresetKey_(article);
  var sheet = getOrCreatePresetsSheet_();
  var lastRow = sheet.getLastRow();
  var items = [];
  if (lastRow < 2) {
    return items;
  }

  for (var r = 2; r <= lastRow; r++) {
    var g = normalizePresetKey_(sheet.getRange(r, PRESET_COL.group).getValue());
    var a = normalizePresetKey_(sheet.getRange(r, PRESET_COL.article).getValue());
    if (g !== group || a !== article) {
      continue;
    }
    items.push({
      order: sheet.getRange(r, PRESET_COL.order).getValue(),
      sketch: String(sheet.getRange(r, PRESET_COL.sketch).getValue() || '').trim(),
      number: String(sheet.getRange(r, PRESET_COL.number).getValue() || '').trim(),
      n: sheet.getRange(r, PRESET_COL.n).getValue(),
      l: sheet.getRange(r, PRESET_COL.l).getValue(),
      op: String(sheet.getRange(r, PRESET_COL.op).getValue() || '')
    });
  }

  items.sort(function (x, y) {
    return (Number(x.order) || 0) - (Number(y.order) || 0);
  });

  return items.map(function (item) {
    return {
      sketch: item.sketch,
      number: item.number,
      n: item.n,
      l: item.l,
      op: item.op
    };
  });
}

/**
 * @returns {Object}
 */
function apiGetSemifinishedPresetsIndex() {
  return getSemifinishedPresetsIndex_();
}

/**
 * @param {Object} payload
 * @returns {Object}
 */
function apiSaveSemifinishedPreset(payload) {
  try {
    var group = payload.group;
    var article = payload.article;
    var items = payload.items || payload.queue || [];
    saveSemifinishedPreset_(group, article, items);
    var index = getSemifinishedPresetsIndex_();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Сохранено ' + items.length + ' операций',
      'КБ.ПРЕСЕТЫ',
      4
    );
    return {
      ok: true,
      hint: 'Набор «' + normalizePresetKey_(group) + '» / «' + normalizePresetKey_(article) +
        '» сохранён (' + items.length + ' оп.) на лист «' + (CONFIG.kbPresetsSheet || 'КБ.ПРЕСЕТЫ') + '».',
      presets: index
    };
  } catch (err) {
    throw new Error('Не удалось сохранить набор: ' + (err.message || err));
  }
}

/**
 * @param {Object} payload
 * @returns {Object}
 */
function apiDeleteSemifinishedPreset(payload) {
  var group = payload.group;
  var article = payload.article;
  deleteSemifinishedPreset_(group, article);
  return {
    ok: true,
    hint: 'Набор удалён.',
    presets: getSemifinishedPresetsIndex_()
  };
}

/**
 * @param {Object} payload
 * @returns {Object}
 */
function apiLoadSemifinishedPreset(payload) {
  var group = payload.group;
  var article = payload.article;
  var items = loadSemifinishedPreset_(group, article);
  if (!items.length) {
    throw new Error('Набор не найден: «' + normalizePresetKey_(group) + '» / «' +
      normalizePresetKey_(article) + '».');
  }
  return { items: items, count: items.length };
}

/**
 * @param {Object} payload
 * @returns {Object}
 */
function apiAppendSemifinishedToQueue(payload) {
  var loaded = apiLoadSemifinishedPreset(payload);
  var appendRes = apiAppendToQueue({
    sheetName: payload.sheetName,
    items: loaded.items
  });
  return {
    count: appendRes.count,
    added: loaded.count,
    hint: 'В очередь добавлено ' + loaded.count + ' операций («' +
      normalizePresetKey_(payload.group) + '» / «' + normalizePresetKey_(payload.article) + '»).'
  };
}

/**
 * Заменить текущую очередь сохранённым набором (для редактирования в окне очереди).
 *
 * @param {Object} payload
 * @returns {Object}
 */
function apiReplaceQueueWithSemifinished(payload) {
  var loaded = apiLoadSemifinishedPreset(payload);
  var data = loadQueue_();
  var queue = [];
  var nextId = 1;
  loaded.items.forEach(function (item) {
    queue.push({
      id: nextId++,
      sketch: item.sketch,
      number: item.number,
      label: item.number,
      opType: '',
      n: num_(item.n),
      l: num_(item.l),
      op: item.op || ''
    });
  });
  saveQueue_({
    queue: queue,
    nextId: nextId,
    sheetName: payload.sheetName !== undefined ? payload.sheetName : data.sheetName,
    sheetsCache: data.sheetsCache
  });
  return buildQueueStateResponse_(loadQueue_());
}
