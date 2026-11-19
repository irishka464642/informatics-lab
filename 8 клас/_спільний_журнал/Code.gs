/**
 * Інформатика 8 клас — єдиний приймач результатів ігор.
 * Працює для ВСІХ наступних уроків без повторного розгортання.
 *
 * ОДИН РАЗ:
 * 1. У робочій Google Таблиці «Загальний журнал ...» відкрийте
 *    Розширення → Apps Script.
 * 2. Вставте цей файл у Code.gs і збережіть.
 * 3. Розгорнути → Нове розгортання → Вебзастосунок.
 *    Виконувати від імені: Мене.
 *    Доступ: відповідно до шкільних правил (для анонімних
 *    телефонів зазвичай потрібен "Anyone", якщо це дозволено).
 * 4. Збережіть URL /exec як RESULTS_ENDPOINT для всіх ігор.
 *
 * Для НОВОЇ гри змінюється тільки JSON, що надсилає сайт:
 * lessonDate (YYYY-MM-DD), lessonId, gameTitle, taskNo,
 * progress/maxProgress, score/maxScore і поля учня.
 *
 * ВАЖЛИВО: анонімний endpoint НЕ автентифікує учня.
 * Дані не можна використовувати як єдиний доказ оцінки.
 */
const JOURNAL = Object.freeze({
  version: 'informatics-journal-v1',
  instructionsSheet: 'Інструкція',
  classes: ['8-А', '8-Б'],
  groups: ['1', '2'],
  maxDateTabs: 180,
  headings: [
    'Час запису', 'Клас', 'Група', 'Прізвище', 'Ім’я',
    '№ уроку', 'Назва гри', 'Завдання', 'Прогрес',
    'Усього завдань', 'Бали', 'Макс. бал', 'З першої спроби',
    'Помилки в завданні', 'Помилки загалом', 'ID події',
    'ID сесії', 'Дата уроку'
  ]
});

function doGet() {
  return asJson_({
    ok: true, version: JOURNAL.version,
    status: 'Приймач результатів працює'
  });
}

function doPost(e) {
  let lock = null;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('empty_payload');
    }
    const raw = e.postData.contents;
    if (raw.length > 9000) throw new Error('payload_too_large');

    const x = JSON.parse(raw);
    const v = validateEvent_(x);

    lock = LockService.getScriptLock();
    lock.waitLock(25000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('script_must_be_bound_to_journal');

    // Один аркуш на календарну дату уроку.
    // Наприклад: "20.11.2026", "23.11.2026".
    const tabName = v.lessonDate.slice(8, 10) + '.'
                  + v.lessonDate.slice(5, 7) + '.'
                  + v.lessonDate.slice(0, 4);
    let tab = ss.getSheetByName(tabName);
    if (!tab) tab = createDateSheet_(ss, tabName);

    const last = tab.getLastRow();
    if (last >= 2) {
      const duplicate = tab.getRange(2, 16, last - 1, 1)
        .createTextFinder(v.eventId)
        .matchEntireCell(true)
        .findNext();
      if (duplicate) {
        return asJson_({ok: true, duplicate: true, dateTab: tabName});
      }
    }

    tab.appendRow([
      new Date(), v.klass, v.group, safeCell_(v.surname),
      safeCell_(v.firstname), safeCell_(v.lessonId),
      safeCell_(v.gameTitle), safeCell_(v.taskNo),
      v.progress, v.maxProgress, v.score, v.maxScore,
      v.firstTry, v.questionMisses, v.totalMisses,
      v.eventId, v.sessionId, v.lessonDate
    ]);
    SpreadsheetApp.flush();

    return asJson_({
      ok: true, recorded: true,
      dateTab: tabName, progress: v.progress
    });
  } catch (err) {
    console.error('Journal error:', String(err));
    // Не віддаємо технічні деталі й дані учня у публічній відповіді.
    return asJson_({ok: false, error: 'request_rejected'});
  } finally {
    if (lock) lock.releaseLock();
  }
}

function createDateSheet_(ss, tabName) {
  if (ss.getSheets().length >= JOURNAL.maxDateTabs + 1) {
    throw new Error('too_many_date_tabs');
  }
  const tab = ss.insertSheet(tabName);
  const n = JOURNAL.headings.length;
  tab.getRange(1, 1, 1, n).setValues([JOURNAL.headings]);
  tab.setFrozenRows(1);
  tab.getRange(1, 1, 1, n)
    .setBackground('#e9eff6').setFontWeight('bold')
    .setWrap(true).setVerticalAlignment('middle');
  tab.setRowHeight(1, 48);
  for (let col = 1; col <= n; col++) tab.setColumnWidth(col, 138);
  tab.setColumnWidth(4, 170);
  tab.setColumnWidth(5, 160);
  tab.setColumnWidth(7, 225);
  tab.setColumnWidth(16, 220);
  tab.setColumnWidth(17, 200);
  tab.getRange(1, 1, tab.getMaxRows(), n).createFilter();
  return tab;
}

function validateEvent_(x) {
  if (!x || typeof x !== 'object' || Array.isArray(x)) {
    throw new Error('invalid_object');
  }
  const clean = (input, limit) => String(input ?? '')
    .normalize('NFC').trim().replace(/\s+/g, ' ').slice(0, limit);
  const isName = s => /^[\p{L}][\p{L}’' -]{1,59}$/u.test(s);
  const validId = s => /^[A-Za-z0-9._-]{8,120}$/.test(s);
  const intRange = (raw, lo, hi, label) => {
    const n = Number(raw);
    if (!Number.isInteger(n) || n < lo || n > hi) {
      throw new Error('bad_' + label);
    }
    return n;
  };
  const surname = clean(x.surname, 60);
  const firstname = clean(x.firstname, 60);
  const klass = clean(x.klass, 10);
  const group = clean(x.group, 10);
  if (!isName(surname) || !isName(firstname)) throw new Error('bad_name');
  if (!JOURNAL.classes.includes(klass)) throw new Error('bad_class');
  if (!JOURNAL.groups.includes(group)) throw new Error('bad_group');

  const lessonDate = clean(x.lessonDate, 10);
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(lessonDate);
  if (!match) throw new Error('bad_lesson_date');
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const dt = new Date(Date.UTC(year, month - 1, day));
  if (year < 2026 || year > 2035 ||
      dt.getUTCFullYear() !== year ||
      dt.getUTCMonth() !== month - 1 ||
      dt.getUTCDate() !== day) throw new Error('bad_date');

  const lessonId = clean(x.lessonId, 45);
  const gameTitle = clean(x.gameTitle, 120);
  const taskNo = clean(x.taskNo ?? x.question ?? x.progress, 80);
  if (!lessonId || !gameTitle || !taskNo) throw new Error('missing_lesson');

  const progress = intRange(x.progress, 1, 500, 'progress');
  const maxProgress = intRange(x.maxProgress, 1, 500, 'max_progress');
  if (progress > maxProgress) throw new Error('invalid_progress');
  const score = Number(x.score ?? 0);
  const maxScore = Number(x.maxScore ?? 1);
  if (!Number.isFinite(score) || !Number.isFinite(maxScore) ||
      score < 0 || maxScore <= 0 || maxScore > 1000 ||
      score > maxScore) throw new Error('bad_score');
  const firstTry = intRange(x.firstTry ?? 0, 0, progress, 'first_try');
  const questionMisses = intRange(x.questionMisses ?? 0, 0, 500, 'question_misses');
  const totalMisses = intRange(x.totalMisses ?? 0, questionMisses, 10000, 'total_misses');
  const eventId = clean(x.eventId, 120);
  const sessionId = clean(x.sessionId, 120);
  if (!validId(eventId) || !validId(sessionId)) throw new Error('bad_event_id');
  return {surname, firstname, klass, group, lessonDate,
    lessonId, gameTitle, taskNo, progress, maxProgress,
    score, maxScore, firstTry, questionMisses, totalMisses,
    eventId, sessionId};
}

function safeCell_(v) {
  const text = String(v);
  return /^[=+\-@\t\r]/.test(text) ? "'" + text : text;
}

function asJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
