/**
 * Структура книги CALC_COMMERCIAL (заголовки от оператора).
 *
 * КБ: строка заголовков = 4, данные с 5.
 * БД.ОП: строка заголовков = 1, данные с 2.
 * На КБ «Эскиз проекта» = в БД.ОП «Название»; «Номер» = «Номер».
 */
var CONFIG = {
  dbSheet: 'БД.ОП',
  kbSheetPattern: /^КБ\d+$/i,
  legacyTlPrefix: 'TL_',

  kbHeaderRow: 4,
  kbDataStartRow: 5,

  dbHeaderRow: 1,
  dbDataStartRow: 2,

  /** Буквы колонок БД.ОП для формул FILTER (Номер=A, Название=B) */
  dbColLetters: {
    number: 'A',
    sketch: 'B'
  },

  kb: {
    sketch: 'Эскиз проекта',
    number: 'Номер',
    component: 'Комплектующая',
    n: 'N',
    l: 'L',
    op: 'OP',
    summ: 'SUMM',
    requirements: 'Требования к задаче',
    col6: '6',
    col7: '7',
    timeTotal: '8',
    col9: '9',
    col10: '10',
    price: 'Цена',
    unitPriceHuman: 'Уд.Цена ЧЛ, сек',
    unitPriceMachine: 'Уд.Цена МШ, сек',
    opType: 'Тип операции'
  },

  /**
   * Колонки КБ с цифровыми заголовками → поле БД.ОП.
   * При необходимости поправьте соответствие под вашу книгу.
   */
  kbDbPull: {
    '6': 'Время Чел, сек/оп; сек/м',
    '7': 'Время машины, сек/оп; сек/м',
    '9': 'Скорость проката',
    '10': 'Скорость работы инструмента'
  },

  db: {
    number: 'Номер',
    sketch: 'Название',
    semifinished: 'Полуфабрикат',
    wire: 'Провод',
    connector: 'Разъем',
    tool: 'Инструмент',
    applicators: 'Апликаторы / модули',
    program: 'Программа',
    machine: 'Машина',
    opTime: 'Время Операции',
    prepTime: 'Время подготовки, сек',
    setupConsumption: 'Расход на настройку м; шт;',
    timeHuman: 'Время Чел, сек/оп; сек/м',
    timeMachine: 'Время машины, сек/оп; сек/м',
    unitPriceHuman: 'Уд.Цена ЧЛ, сек',
    unitPriceHumanMag: 'Уд.Цена ЧЛ_МАГ, сек',
    unitPriceMachine: 'Уд.Цена МШ, сек',
    opType: 'Тип операции',
    manualTakeTime: 'Время ручных работ для взятия полуфабриката',
    manualWorkTime: 'Время ручных работ',
    manualRemoveTime: 'Время ручных работ для снятия полуфабриката',
    rollSpeed: 'Скорость проката',
    toolWorkSpeed: 'Скорость работы инструмента',
    toolOpCount: 'Кол-во операций инструмента',
    extraOpTime: 'Время доп.операции'
  },

  opTypes: {
    linear: 'Погонный',
    variable: 'Переменный',
    static: 'Статичный'
  }
};
