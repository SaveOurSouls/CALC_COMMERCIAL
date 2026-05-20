/**
 * Подтягивание полей из БД.ОП формулами (вместо VLOOKUP по TL_КБ).
 * Пример для строки 2 листа КБ; буквы колонок замените на свои.
 */

/**
 * QUERY: все поля операции по паре (Название, Номер).
 * Вставьте в одну ячейку или используйте как основу для MAP.
 *
 * @returns {string}
 */
function exampleQueryPullFormula() {
  return '=QUERY(БД.ОП!A:Z,"select * where Col1 = \'"&$A2&"\' and Col2 = \'"&$B2&"\' limit 1",0)';
}

/**
 * MAP + FILTER: подтянуть одну колонку из БД.ОП по эскизу и номеру.
 *
 * @param {string} dbValueColumn буква колонки значения в БД.ОП, напр. "M"
 * @returns {string}
 */
function exampleMapPullFormula(dbValueColumn) {
  return '=MAP(A2:B2,LAMBDA(sk,nu,INDEX(FILTER(БД.ОП!' + dbValueColumn + ':' + dbValueColumn +
    ',(БД.ОП!$A:$A=sk)*(БД.ОП!$B:$B=nu)),1)))';
}
