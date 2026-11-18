/**
 * Урок №11 — результати для 5-А, 5-Б і 5-В класів.
 *
 * ВАЖЛИВО: цей файл призначений для ПРИВАТНОГО Google Apps Script,
 * створеного з Google Таблиці через Розширення → Apps Script.
 * Номер таблиці не зберігається в публічному GitHub.
 *
 * Публічний Web App без входу в обліковий запис НЕ перевіряє особу дитини.
 * Перед збиранням персональних даних дотримуйтеся шкільних правил приватності.
 */

const JOURNAL_PROP = "LESSON11_SHEET_ID";
const EXPECTED_TAB = "Результати";
const EXPECTED_CLASSES = ["5-А", "5-Б", "5-В"];
const EXPECTED_VARIANTS = ["А", "Б", "В"];

/**
 * Виконати ОДИН РАЗ у проєкті, створеному безпосередньо з журналу:
 * https://docs.google.com/spreadsheets/d/... -> Розширення -> Apps Script.
 */
function setupJournal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || !ss.getSheetByName(EXPECTED_TAB)) {
    throw new Error("Відкрийте Apps Script саме з журналу результатів уроку 11.");
  }
  PropertiesService.getScriptProperties().setProperty(JOURNAL_PROP, ss.getId());
  console.log("Журнал готовий: " + ss.getName());
}

function doGet() {
  return response_({ready:true,lesson:"11",note:"Receiver online. Send POST to record results."});
}

function doPost(e) {
  try {
    if (!e || !e.postData || typeof e.postData.contents !== "string" ||
        !e.postData.contents || e.postData.contents.length > 6000) throw new Error("empty/large");

    const x = JSON.parse(e.postData.contents);
    if (!x || x.lesson !== "11") throw new Error("lesson");
    const clean = (value, max) => String(value == null ? "" : value)
      .trim().replace(/\s+/g," ").slice(0,max);
    const klass=clean(x.klass,10), group=clean(x.group,2),variant=clean(x.variant,3);
    const surname=clean(x.surname,60), firstname=clean(x.firstname,60);
    const nameRe=/^[\p{L}][\p{L}’' -]{1,59}$/u;

    if (!EXPECTED_CLASSES.includes(klass) || !["1","2"].includes(group) ||
        !EXPECTED_VARIANTS.includes(variant) ||
        !nameRe.test(surname) || !nameRe.test(firstname)) throw new Error("identity");

    const num = (v,max) => {
      const n = Number(v);
      if (!Number.isInteger(n) || n < 0 || n > max) throw new Error("score");
      return n;
    };
    const score=num(x.score,12),gr1=num(x.gr1,4),gr3=num(x.gr3,4),gr4=num(x.gr4,4);
    const practice=num(x.practice,3),warm=num(x.warm,4),resources=num(x.resources,3);
    if (gr1+gr3+gr4!==score) throw new Error("score sum");
    const answers=clean(x.answers,25),mistakes=clean(x.mistakes,40);
    if (!/^[АБВ?]{12}$/.test(answers) || !/^\d{0,2}(,\d{1,2})*$/.test(mistakes)) throw new Error("answers");
    const attempt=num(x.attempt,20);
    if(attempt < 1)throw new Error("attempt");

    // Перетворюємо потенційні формули в безпечний звичайний текст.
    const safe = v => {
      const s=clean(v,160);
      return /^[=+\-@]/.test(s) ? "'"+s : s;
    };
    const row = [
      new Date(),klass,Number(group),safe(surname),safe(firstname),variant,
      score,gr1,gr3,gr4,practice,warm,resources,answers,mistakes,
      safe(x.ocean),safe(x.translation),safe(x.simulation),attempt,"Отримано через сайт"
    ];

    const sheetId=PropertiesService.getScriptProperties().getProperty(JOURNAL_PROP);
    if(!sheetId)throw new Error("Run setupJournal first");
    const sheet=SpreadsheetApp.openById(sheetId).getSheetByName(EXPECTED_TAB);
    if(!sheet)throw new Error("Sheet missing");

    const lock=LockService.getScriptLock();
    lock.waitLock(15000);
    try{sheet.appendRow(row)}
    finally{lock.releaseLock()}

    return response_({ok:true});
  } catch(error) {
    console.error(error);
    return response_({ok:false,error:"invalid_request"});
  }
}

function response_(object) {
  return ContentService.createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}
