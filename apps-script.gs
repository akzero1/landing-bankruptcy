/**
 * ══════════════════════════════════════════════════════════════════
 *  LEAD INTAKE — Юридическая компания «Императрица»
 *  Принимает данные с лендинга и складывает в Google Sheets
 *  с красивым оформлением и цветовым приоритетом лидов.
 * ══════════════════════════════════════════════════════════════════
 *
 * КАК ПОДКЛЮЧИТЬ (один раз, 5 минут):
 *
 * 1. Открой https://sheets.new — создаст чистую Google-таблицу
 * 2. Переименуй её, например: «Императрица — заявки»
 * 3. Меню: Расширения → Apps Script
 * 4. Удали шаблонный код, вставь ЭТОТ файл целиком
 * 5. Нажми «Сохранить» (дискета) → «Выполнить» рядом с функцией setup()
 *    — Google запросит разрешения → одобри (это твой же аккаунт)
 *    — функция создаст лист «Заявки» с готовыми заголовками
 * 6. Меню: Развернуть → Новое развертывание → Тип: Веб-приложение
 *    — Выполнять от: «Я» (от моего имени)
 *    — Кто имеет доступ: «Все» (нужно, чтобы лендинг мог слать)
 *    — Нажми «Развернуть»
 * 7. СКОПИРУЙ полученный URL (выглядит как
 *    https://script.google.com/macros/s/AKfycb.../exec)
 * 8. Вставь этот URL в index.html вместо `__PASTE_APPS_SCRIPT_URL__`
 *    (две строки в конце <script> — ищи SCRIPT_URL)
 * 9. Готово — заявки с квиза и формы падают в таблицу
 */

const SHEET_NAME = 'Заявки';
const COLOR_NAVY = '#0a0b14';
const COLOR_GOLD = '#c9a961';
const COLOR_GOLD_SOFT = '#f5ecd4';
const COLOR_ROW_ALT = '#fbfaf6';
const COLOR_HOT = '#ffe4e0';
const COLOR_WARM = '#fff4d1';
const COLOR_COLD = '#eaf0f7';
const COLOR_TEXT_COLD = '#6b7a8c';

const HEADERS = [
  'Дата / время',
  'Приоритет',
  'Источник',
  'Имя',
  'Телефон',
  'Город',
  'Сумма долга',
  'Тип долга',
  'Имущество',
  'Официальный доход',
  'Рекомендация',
  'Стоимость',
  'Срок',
  'Комментарий клиента',
  'Статус',
  'Заметки юриста'
];

const COLUMN_WIDTHS = [150, 115, 170, 130, 150, 115, 175, 170, 175, 145, 230, 105, 105, 230, 135, 260];

const STATUS_OPTIONS = [
  'Новая',
  'В работе',
  'На консультации',
  'Подписан договор',
  'Отказ клиента',
  'Не подходит',
  'Дубль'
];

/**
 * Запускается один раз вручную — готовит таблицу.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  applyFormatting_(sheet);
  SpreadsheetApp.getUi().alert(
    'Таблица готова ✓',
    'Лист «Заявки» настроен. Теперь разверни проект как веб-приложение (меню «Развернуть»).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Принимает данные с лендинга (GET-запрос через hidden iframe).
 */
function doGet(e) {
  try {
    const p = e.parameter || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      applyFormatting_(sheet);
    }
    if (sheet.getLastRow() === 0) applyFormatting_(sheet);

    const row = [
      new Date(),
      p.score || '',
      p.source || 'landing',
      p.name || '',
      p.phone || '',
      p.city || '',
      p.debt || '',
      p.type || '',
      p.property || '',
      p.income || '',
      p.verdict || '',
      p.cost || '',
      p.time_est || '',
      p.message || '',
      'Новая',
      ''
    ];
    sheet.appendRow(row);

    const lastRow = sheet.getLastRow();
    styleRow_(sheet, lastRow);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput('ERROR: ' + err.toString());
  }
}

/**
 * Форматирует лист: шапка, ширины, freeze, dropdown, условное форматирование.
 */
function applyFormatting_(sheet) {
  sheet.clear();
  sheet.clearConditionalFormatRules();

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange
    .setBackground(COLOR_NAVY)
    .setFontColor(COLOR_GOLD)
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontFamily('Inter, Arial')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 44);

  COLUMN_WIDTHS.forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  // Dropdown для колонки «Статус»
  const statusCol = HEADERS.indexOf('Статус') + 1;
  const statusRange = sheet.getRange(2, statusCol, 2000, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .setHelpText('Выберите статус заявки')
    .build();
  statusRange.setDataValidation(rule);

  // Условное форматирование для «Приоритет»
  const priCol = HEADERS.indexOf('Приоритет') + 1;
  const priRange = sheet.getRange(2, priCol, 2000, 1);

  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Горячий')
    .setBackground(COLOR_HOT)
    .setFontColor('#8b1f1f')
    .setBold(true)
    .setRanges([priRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Средний')
    .setBackground(COLOR_WARM)
    .setFontColor('#7a5a10')
    .setBold(true)
    .setRanges([priRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Холодный')
    .setBackground(COLOR_COLD)
    .setFontColor(COLOR_TEXT_COLD)
    .setRanges([priRange])
    .build());

  // Условное форматирование для «Статус»
  const statusColorMap = {
    'Новая': ['#dbeafe', '#1e4e9c'],
    'В работе': ['#fff4d1', '#7a5a10'],
    'На консультации': ['#e8e4ff', '#4b3a8e'],
    'Подписан договор': ['#d8f5d8', '#1e6b1e'],
    'Отказ клиента': ['#fce4e4', '#8b1f1f'],
    'Не подходит': ['#f0f0f0', '#555'],
    'Дубль': ['#f0f0f0', '#888']
  };
  Object.keys(statusColorMap).forEach(function(label) {
    const c = statusColorMap[label];
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(label)
      .setBackground(c[0])
      .setFontColor(c[1])
      .setBold(true)
      .setRanges([statusRange])
      .build());
  });

  sheet.setConditionalFormatRules(rules);

  // Формат даты в колонке А
  sheet.getRange(2, 1, 2000, 1).setNumberFormat('dd.MM.yyyy  HH:mm');

  // Вертикальное выравнивание всей области
  sheet.getRange(2, 1, 2000, HEADERS.length)
    .setVerticalAlignment('middle')
    .setFontFamily('Inter, Arial')
    .setFontSize(10)
    .setWrap(true);
}

/**
 * Стилизует новую строку: чередующийся фон, перенос текста.
 */
function styleRow_(sheet, row) {
  const rng = sheet.getRange(row, 1, 1, HEADERS.length);
  rng.setVerticalAlignment('middle')
    .setFontFamily('Inter, Arial')
    .setFontSize(10)
    .setWrap(true);
  if (row % 2 === 0) {
    rng.setBackground(COLOR_ROW_ALT);
  } else {
    rng.setBackground('#ffffff');
  }
  sheet.getRange(row, 1).setNumberFormat('dd.MM.yyyy  HH:mm');
  sheet.setRowHeight(row, 36);
}

/**
 * Утилита: отправить тестовый лид (чтобы проверить без запросов).
 * Запусти из редактора вручную, одну строку положит в таблицу.
 */
function testInsertFakeLead() {
  doGet({
    parameter: {
      score: '🔥 Горячий',
      source: 'test',
      name: 'Иван Тестов',
      phone: '+7 900 000 00 00',
      city: 'Пермь',
      debt: '800 000 – 3 000 000 ₽',
      type: 'Банковские кредиты',
      property: 'Автомобиль',
      income: '30 000 – 80 000 ₽',
      verdict: 'Судебное банкротство с защитой активов',
      cost: 'от 110 тыс.',
      time_est: '7–10 мес.',
      message: 'Это тестовая заявка'
    }
  });
}
