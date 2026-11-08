/**
 * Informatics Lab — приймання результатів.
 * Створіть окремий проєкт у script.google.com і вставте цей файл.
 * НЕ додавайте ідентифікатори таблиць, токени чи дані учнів у публічний репозиторій.
 *
 * Увага: вебзастосунок з доступом "Anyone" не має достовірної автентифікації учнів.
 * Дані можуть бути помилковими або підробленими; не використовуйте їх як єдину підставу оцінювання.
 */
const GROUPS = ["6-А:1","6-А:2","6-Б:1","6-Б:2","6-В:1","6-В:2"];
const HEADERS = ["Дата отримання","Клас","Група","Прізвище","Ім’я","Вправа","Бали","Максимум","Спроба","Час (с)"];
const PROP_PREFIX = "RESULTS_SHEET_";

// Запустити один раз в Apps Script під обліковим записом учителя.
// Створює шість ОКРЕМИХ Google Таблиць, зберігає їх ID в Script Properties.
function setupJournals() {
 const props=PropertiesService.getScriptProperties();
 const links=[];
 GROUPS.forEach(key=>{
   const prop=PROP_PREFIX+key;
   let id=props.getProperty(prop);
   if(!id){
     const [klass,group]=key.split(":");
     const sheetFile=SpreadsheetApp.create("Informatics Lab — "+klass+" — "+group+" група");
     id=sheetFile.getId();
     const tab=sheetFile.getSheets()[0];
     tab.setName("Результати");
     tab.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
     tab.setFrozenRows(1);
     props.setProperty(prop,id);
   }
   links.push(key+": https://docs.google.com/spreadsheets/d/"+id+"/edit");
 });
 Logger.log(links.join("\n"));
}

function doPost(e){
 try{
  if(!e || !e.postData || !e.postData.contents)throw new Error("empty");
  if(e.postData.contents.length>3000)throw new Error("too large");
  const x=JSON.parse(e.postData.contents);
  const klass=String(x.klass||"");
  const group=String(x.group||"");
  const key=klass+":"+group;
  if(!GROUPS.includes(key))throw new Error("invalid group");
  const clean=(v,max)=>String(v||"").trim().replace(/\s+/g," ").slice(0,max);
  const surname=clean(x.surname,60);
  const firstname=clean(x.firstname,60);
  if(!/^[\p{L}’' -]{2,60}$/u.test(surname)||!/^[\p{L}’' -]{2,60}$/u.test(firstname))throw new Error("invalid name");
  if(x.exercise!=="table-ordering-v1")throw new Error("invalid exercise");
  const score=Number(x.score),total=Number(x.total),attempt=Number(x.attempt),duration=Number(x.durationSeconds);
  if(!Number.isInteger(score)||score<0||score>12||total!==12||!Number.isInteger(attempt)||attempt<1||attempt>100||!Number.isInteger(duration)||duration<0||duration>86400)throw new Error("invalid score");
  const id=PropertiesService.getScriptProperties().getProperty(PROP_PREFIX+key);
  if(!id)throw new Error("not configured");
  // Видаляємо початкові =, +, -, @, щоб ім'я не стало формулою в таблиці.
  const safe=s=>/^[=+\-@]/.test(s)?"'"+s:s;
  const row=[new Date(),klass,group,safe(surname),safe(firstname),x.exercise,score,total,attempt,duration];
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{SpreadsheetApp.openById(id).getSheetByName("Результати").appendRow(row)}
  finally{lock.releaseLock()}
  return output({ok:true});
 }catch(err){
  console.error(err);
  return output({ok:false,error:"invalid_request"});
 }
}
function doGet(){return output({status:"Informatics Lab receiver online"})}
function output(x){return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON)}
