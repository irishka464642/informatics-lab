/**
 * «Мій Word — 6 клас». Серверна частина для ПР № 2.
 * Розміщується у власному проєкті Google Apps Script вчителя, НЕ на GitHub Pages.
 *
 * Властивості скрипту (Project Settings -> Script properties):
 *   WORD6_TEACHER_EMAIL = адреса пошти вчителя
 *   WORD6_CLASS_PIN     = пароль/код доступу до надсилання (8+ символів)
 *
 * Deploy -> New deployment -> Web app
 * Execute as: Me
 * Who has access: Anyone (якщо це дозволяє політика шкільного акаунта)
 *
 * Призначення: один фіксований отримувач, тільки вкладення .docx,
 * без довільних адрес одержувачів та без зберігання робіт на Drive.
 */
const WORD6_LIMITS = Object.freeze({
  MAX_FILE_BYTES: 2 * 1024 * 1024,
  MAX_SUBMISSIONS_PER_CODE_PER_DAY: 3,
  MAX_SUBMISSIONS_PER_DAY: 80,
});

function doGet() {
  return renderResult_(
    'Приймач Word-документів',
    'Це службова сторінка для надсилання практичних робіт. Скористайтеся кнопкою «Надіслати Word на пошту» в редакторі.',
    false
  );
}

function doPost(e) {
  let lock = null;
  try {
    if (!e || !e.parameter) throw new Error('NO_REQUEST');
    const p = e.parameter;

    // Проти автоматизованих форм. Не є єдиним засобом захисту.
    if (String(p.website || '').trim()) throw new Error('INVALID_REQUEST');

    const props = PropertiesService.getScriptProperties();
    const teacherEmail = String(props.getProperty('WORD6_TEACHER_EMAIL') || '').trim();
    const classPin = String(props.getProperty('WORD6_CLASS_PIN') || '');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(teacherEmail) || classPin.length < 8) {
      throw new Error('NOT_CONFIGURED');
    }
    if (String(p.access_key || '') !== classPin) throw new Error('WRONG_CODE');

    const className = String(p.class_name || '').trim();
    const studentCode = String(p.student_code || '').trim();
    const documentName = String(p.document_name || '').trim();
    const payload = String(p.payload || '').trim();

    if (!/^6-[АБВ]$/.test(className)) throw new Error('BAD_CLASS');
    if (!/^[a-zA-Z0-9А-Яа-яІЇЄҐіїєґ_-]{2,24}$/.test(studentCode)) {
      throw new Error('BAD_STUDENT_CODE');
    }
    if (!/^[a-zA-Z0-9А-Яа-яІЇЄҐіїєґ_-]{2,70}\.docx$/i.test(documentName)) {
      throw new Error('BAD_FILENAME');
    }
    if (!payload || payload.length > Math.ceil(WORD6_LIMITS.MAX_FILE_BYTES * 4 / 3) + 8) {
      throw new Error('FILE_TOO_LARGE');
    }
    if (!/^[A-Za-z0-9+/]+={0,2}$/.test(payload)) throw new Error('BAD_DOCUMENT');

    const bytes = Utilities.base64Decode(payload);
    if (bytes.length < 200 || bytes.length > WORD6_LIMITS.MAX_FILE_BYTES) {
      throw new Error('BAD_DOCUMENT_SIZE');
    }
    // Документ Office Open XML є ZIP-контейнером.
    if ((bytes[0] & 255) !== 0x50 || (bytes[1] & 255) !== 0x4b) {
      throw new Error('BAD_DOCUMENT_FORMAT');
    }

    const day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    const counterProperty = 'WORD6_QUOTA';
    lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) throw new Error('BUSY');

    let state;
    try {
      state = JSON.parse(props.getProperty(counterProperty) || '{}');
    } catch (_) {
      state = {};
    }
    if (state.day !== day) state = { day: day, total: 0, byCode: {} };
    if (!state.byCode || typeof state.byCode !== 'object') state.byCode = {};

    const key = className + '-' + studentCode.toLowerCase();
    const count = Number(state.byCode[key] || 0);
    if (state.total >= WORD6_LIMITS.MAX_SUBMISSIONS_PER_DAY) throw new Error('DAILY_LIMIT');
    if (count >= WORD6_LIMITS.MAX_SUBMISSIONS_PER_CODE_PER_DAY) throw new Error('CODE_LIMIT');
    if (MailApp.getRemainingDailyQuota() < 1) throw new Error('MAIL_QUOTA');

    const attachment = Utilities.newBlob(
      bytes,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      documentName
    );
    MailApp.sendEmail({
      to: teacherEmail,
      subject: 'Інформатика 6 клас — ПР № 2 — ' + className + ' — ' + studentCode,
      body: [
        'Практична робота № 2: «Створення текстового документа з графічним зображенням і таблицею».',
        'Клас: ' + className,
        'Навчальний код: ' + studentCode,
        'Файл Word прикріплено до листа.',
        '',
        'Повідомлення створене застосунком «Мій Word — 6 клас».'
      ].join('\n'),
      attachments: [attachment],
      name: 'Мій Word — 6 клас'
    });

    state.total = Number(state.total || 0) + 1;
    state.byCode[key] = count + 1;
    props.setProperty(counterProperty, JSON.stringify(state));
    return renderResult_(
      'Роботу надіслано!',
      'Лист із документом Word (.docx) передано поштовому сервісу. Можна закрити цю вкладку й повернутися до редактора.',
      true
    );
  } catch (err) {
    console.error('WORD6 send failed: ' + String(err));
    const code = String(err && err.message || '');
    const messages = {
      NO_REQUEST: 'Дані запиту відсутні.',
      NOT_CONFIGURED: 'Пошта вчителя ще не налаштована.',
      WRONG_CODE: 'Неправильний код доступу. Уточніть код у вчителя.',
      BAD_CLASS: 'Перевірте обраний клас.',
      BAD_STUDENT_CODE: 'Уведіть навчальний код учня (2–24 літери чи цифри).',
      BAD_FILENAME: 'Назва файлу не підтримується.',
      FILE_TOO_LARGE: 'Документ завеликий. Спробуйте зменшити зображення.',
      BAD_DOCUMENT: 'Не вдалося прочитати файл Word.',
      BAD_DOCUMENT_SIZE: 'Помилка розміру документа.',
      BAD_DOCUMENT_FORMAT: 'Файл не схожий на документ Word.',
      DAILY_LIMIT: 'Досягнуто ліміту робіт за день. Зверніться до вчителя.',
      CODE_LIMIT: 'З цього навчального коду вже надіслано максимальну кількість робіт на сьогодні.',
      MAIL_QUOTA: 'Вичерпано добову квоту надсилання листів.',
      BUSY: 'Сервіс зайнятий. Спробуйте ще раз.'
    };
    return renderResult_(
      'Надсилання не вдалося',
      messages[code] || 'Не вдалося надіслати файл. Збережіть документ .docx і зверніться до вчителя.',
      false
    );
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (_) {}
    }
  }
}

function renderResult_(title, message, success) {
  const color = success ? '#146547' : '#9c2c35';
  const html = [
    '<!doctype html><html lang="uk"><head><meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width,initial-scale=1">',
    '<title>Мій Word — статус надсилання</title>',
    '<style>body{font:16px/1.55 system-ui,-apple-system,sans-serif;background:#eef3fb;color:#21334c;margin:0;padding:20px}',
    '.card{max-width:550px;margin:10vh auto;background:#fff;border-radius:16px;padding:30px;box-shadow:0 7px 24px #223c641f}',
    'h1{font-size:26px;color:', color, '}p{margin-bottom:20px}small{color:#6b778b}</style>',
    '</head><body><div class="card"><h1>', title, '</h1><p>', message, '</p>',
    '<small>Не надсилайте роботу повторно, якщо вже отримали підтвердження.</small>',
    '</div></body></html>'
  ].join('');
  return HtmlService.createHtmlOutput(html).setTitle('Мій Word — надсилання');
}
