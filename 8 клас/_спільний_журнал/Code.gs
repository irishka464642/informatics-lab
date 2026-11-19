/**
 * ЄДИНИЙ ЖУРНАЛ ІНФОРМАТИКИ ДЛЯ 5–9 КЛАСІВ
 *
 * Одне розгортання Web App для всіх ігор і всіх уроків.
 * Кожен із 5 журналів — окрема приватна Google Таблиця.
 * На кожну дату уроку автоматично створюється вкладка DD.MM.YYYY.
 *
 * Налаштування ОДИН РАЗ у робочому журналі 8 класу:
 * 1. Розширення → Apps Script → вставити цей Code.gs → Зберегти.
 * 2. Запустити setupJournalRouter() вручну з редактора.
 *    Цей запуск збереже ID усіх 5 журналів у Script Properties.
 * 3. Розгорнути → Нове розгортання → Вебзастосунок.
 *    Виконувати як: власник; доступ відповідно до політики школи.
 * 4. URL /exec використовувати в КОЖНІЙ наступній грі.
 *
 * ID таблиць і дані учнів НЕ зберігаються у публічному GitHub.
 * «Anyone» не автентифікує учня. Для оцінювання потрібна перевірка.
 */
const ROUTER = Object.freeze({
  version: 'informatics-results-router-v2',
  registrySheet: 'Реєстр журналів',
  classGrades: [5, 6, 7, 8, 9],
  maxDateTabs: 300,
  propertyPrefix: 'INFORMATICS_JOURNAL_',
  headings: [
    'Час запису', 'Клас', 'Група', 'Прізвище', 'Ім’я',
    '№ уроку', 'Назва гри', 'Завдання', 'Прогрес',
    'Усього завдань', 'Бали', 'Макс. бал',
    'З першої спроби', 'Помилки в завданні', 'Помилки загалом',
    'ID події', 'ID сесії', 'Дата уроку'
  ]
});

/**
 * Натиснути Run у редакторі Google Apps Script ОДИН РАЗ.
 * Функція читає приховані від публічного GitHub адреси журналів
 * із вкладки «Реєстр журналів» в робочій таблиці 8 класу.
 *
 * Якщо колись потрібно поміняти журнал: відредагувати вкладку
 * «Реєстр журналів» і повторно запустити setupJournalRouter().
 * Вебзастосунок заново розгортати не потрібно.
 */
function setupJournalRouter() {
  const bound = SpreadsheetApp.getActiveSpreadsheet();
  if (!bound) {
    throw new Error('Запустіть setupJournalRouter з редактора Apps Script, прив’язаного до журналу 8 класу.');
  }

  const reg = bound.getSheetByName(ROUTER.registrySheet);
  if (!reg) {
    throw new Error('Відсутня вкладка «Реєстр журналів» у журналі 8 класу.');
  }
  const values = reg.getDataRange().getValues();
  const ids = {};
  values.slice(1).forEach(row => {
    const grade = Number(row[0]);
    const id = String(row[1] || '').trim();
    if (ROUTER.classGrades.includes(grade) && id) {
      if (!/^[a-zA-Z0-9_-]{20,150}$/.test(id)) {
        throw new Error('Некоректний ID журналу для ' + grade + ' класу.');
      }
      if (ids[grade] && ids[grade] !== id) {
        throw new Error('Клас ' + grade + ' дублюється у реєстрі.');
      }
      ids[grade] = id;
    }
  });

  const properties = {};
  ROUTER.classGrades.forEach(grade => {
    const id = ids[grade];
    if (!id) throw new Error('Немає журналу для ' + grade + ' класу.');
    const target = SpreadsheetApp.openById(id);
    if (!target.getName().includes('Інформатика ' + grade + ' клас')) {
      throw new Error('Перевірте відповідність журналу для ' + grade + ' класу.');
    }
    properties[ROUTER.propertyPrefix + grade] = id;
  });

  PropertiesService.getScriptProperties().setProperties(properties, false);
  Logger.log('Готово: підключено 5 журналів для 5, 6, 7, 8, 9 класів.');
  return 'Підключено 5 журналів. Тепер можна розгорнути вебзастосунок один раз.';
}

/**
 * Перевірка URL /exec (без показу ID або ПІБ).
 */
function doGet() {
  const props = PropertiesService.getScriptProperties();
  const ready = ROUTER.classGrades.every(g =>
    Boolean(props.getProperty(ROUTER.propertyPrefix + g))
  );
  return json_({
    ok: true,
    version: ROUTER.version,
    configured: ready,
    message: ready ? '5–9 класи підключені' : 'Виконайте setupJournalRouter()'
  });
}

/**
 * Запис одного завершеного завдання / контрольної точки.
 */
function doPost(e) {
  let lock = null;
  try {
    if (!e || !e.postData || typeof e.postData.contents !== 'string') {
      throw new Error('empty_payload');
    }
    const body = e.postData.contents;
    if (!body || body.length > 9000) throw new Error('bad_payload_size');
    const input = JSON.parse(body);
    const v = validate_(input);

    const props = PropertiesService.getScriptProperties();
    const targetId = props.getProperty(ROUTER.propertyPrefix + v.grade);
    if (!targetId) throw new Error('router_not_configured_for_grade');

    // Загальний lock захищає створення нових вкладок і запобігає дублям.
    lock = LockService.getScriptLock();
    lock.waitLock(25000);
    const spreadsheet = SpreadsheetApp.openById(targetId);

    const dateTab =
      v.lessonDate.slice(8, 10) + '.' +
      v.lessonDate.slice(5, 7) + '.' +
      v.lessonDate.slice(0, 4);

    let sheet = spreadsheet.getSheetByName(dateTab);
    if (!sheet) sheet = makeDateSheet_(spreadsheet, dateTab);

    // Колонка P (16) — стабільний ID події. Повторний запит не додає рядка.
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const duplicate = sheet.getRange(2, 16, lastRow - 1, 1)
        .createTextFinder(v.eventId)
        .matchEntireCell(true)
        .findNext();
      if (duplicate) {
        return json_({ok: true, duplicate: true, dateTab: dateTab});
      }
    }

    sheet.appendRow([
      new Date(), v.klass, v.group,
      safeCell_(v.surname), safeCell_(v.firstname),
      safeCell_(v.lessonId), safeCell_(v.gameTitle),
      safeCell_(v.taskNo), v.progress, v.maxProgress,
      v.score, v.maxScore, v.firstTry,
      v.questionMisses, v.totalMisses, v.eventId,
      v.sessionId, v.lessonDate
    ]);
    SpreadsheetApp.flush();
    return json_({
      ok: true, recorded: true,
      grade: v.grade, dateTab: dateTab,
      progress: v.progress
    });
  } catch (err) {
    console.error('Journal error', String(err));
    return json_({ok: false, error: 'request_rejected'});
  } finally {
    if (lock) lock.releaseLock();
  }
}

/**
 * Створення вкладки дати та її оформлення.
 */
function makeDateSheet_(spreadsheet, title) {
  if (spreadsheet.getSheets().length >= ROUTER.maxDateTabs + 2) {
    throw new Error('date_tabs_limit');
  }
  const sheet = spreadsheet.insertSheet(title);
  const colCount = ROUTER.headings.length;
  sheet.getRange(1, 1, 1, colCount).setValues([ROUTER.headings]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, colCount)
    .setBackground('#E9EEF5')
    .setFontColor('#15263E')
    .setFontWeight('bold')
    .setWrap(true)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 48);
  for (let i = 1; i <= colCount; i++) sheet.setColumnWidth(i, 140);
  sheet.setColumnWidth(4, 172);
  sheet.setColumnWidth(5, 155);
  sheet.setColumnWidth(7, 220);
  sheet.setColumnWidth(16, 225);
  sheet.setColumnWidth(17, 195);
  sheet.getRange(1, 1, sheet.getMaxRows(), colCount).createFilter();
  return sheet;
}

/**
 * Перевірка і нормалізація JSON усіх майбутніх ігор.
 */
function validate_(input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) {
    throw new Error('invalid_json_object');
  }
  const clean = (v, size) => String(v === null || v === undefined ? '' : v)
    .normalize('NFC').trim().replace(/\s+/g, ' ').slice(0, size);
  const validName = s => /^[\p{L}][\p{L}’' -]{1,59}$/u.test(s);
  const validId = s => /^[A-Za-z0-9._-]{8,120}$/.test(s);
  const getInteger = (v, min, max, label) => {
    if (v === undefined || v === null || v === '') {
      throw new Error('missing_' + label);
    }
    const n = Number(v);
    if (!Number.isInteger(n) || n < min || n > max) {
      throw new Error('invalid_' + label);
    }
    return n;
  };

  const surname = clean(input.surname, 60);
  const firstname = clean(input.firstname, 60);
  if (!validName(surname) || !validName(firstname)) {
    throw new Error('invalid_name');
  }

  // Українська абетка: 5-А, 5-Б, 5-В ... 9-А тощо.
  const klass = clean(input.klass || input.className, 12)
    .replace(/\s+/g, '').replace(/[–—]/g, '-');
  const classMatch = /^([5-9])-([А-ЯІЇЄҐ])$/u.exec(klass);
  if (!classMatch) throw new Error('invalid_class');
  const grade = Number(classMatch[1]);

  const group = clean(input.group, 10) || '—';
  if (!['1', '2', '3', '4', '—'].includes(group)) {
    throw new Error('invalid_group');
  }

  const lessonDate = clean(input.lessonDate, 10);
  const dateMatch = /^(\d{4})-(\d{2})-(\d{2})$/.exec(lessonDate);
  if (!dateMatch) throw new Error('invalid_lesson_date');
  const year = Number(dateMatch[1]);
  const month = Number(dateMatch[2]);
  const day = Number(dateMatch[3]);
  const dt = new Date(Date.UTC(year, month - 1, day));
  if (year < 2026 || year > 2035 ||
      dt.getUTCFullYear() !== year ||
      dt.getUTCMonth() !== month - 1 ||
      dt.getUTCDate() !== day) {
    throw new Error('impossible_date');
  }

  const lessonId = clean(input.lessonId, 45);
  const gameTitle = clean(input.gameTitle, 120);
  const taskNo = clean(
    input.taskNo !== undefined ? input.taskNo :
      (input.question !== undefined ? input.question : input.progress),
    80);
  if (!lessonId || !gameTitle || !taskNo) {
    throw new Error('missing_task_information');
  }

  const progress = getInteger(input.progress, 1, 500, 'progress');
  const maxProgress = getInteger(input.maxProgress, 1, 500, 'maxProgress');
  if (progress > maxProgress) throw new Error('progress_over_max');

  const score = Number(input.score === undefined ? 0 : input.score);
  const maxScore = Number(input.maxScore === undefined ? 1 : input.maxScore);
  if (!Number.isFinite(score) || !Number.isFinite(maxScore) ||
      score < 0 || maxScore <= 0 || maxScore > 1000 || score > maxScore) {
    throw new Error('invalid_score');
  }

  const firstTry = getInteger(input.firstTry === undefined ? 0 : input.firstTry, 0, progress, 'firstTry');
  const questionMisses = getInteger(
    input.questionMisses === undefined ? 0 : input.questionMisses,
    0, 500, 'questionMisses');
  const totalMisses = getInteger(
    input.totalMisses === undefined ? 0 : input.totalMisses,
    questionMisses, 10000, 'totalMisses');

  const eventId = clean(input.eventId, 120);
  const sessionId = clean(input.sessionId, 120);
  if (!validId(eventId) || !validId(sessionId)) {
    throw new Error('invalid_event_id');
  }

  return {
    grade, klass, group, surname, firstname, lessonDate,
    lessonId, gameTitle, taskNo, progress, maxProgress,
    score, maxScore, firstTry, questionMisses, totalMisses,
    eventId, sessionId
  };
}

function safeCell_(text) {
  const value = String(text);
  return /^[=+\-@\t\r]/.test(value) ? "'" + value : value;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
