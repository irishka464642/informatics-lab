/* Спільний профіль учня для всіх вправ Informatics Lab.
   Зберігається тільки в локальному браузері, не на сервері. */
(()=>{
"use strict";
const KEY="informatics_lab_student_profile_v1";
const ids=["klass","group","surname","firstname"];
const fields=ids.map(id=>document.getElementById(id));
if(fields.some(x=>!x))return;
const status=document.createElement("div");
status.setAttribute("role","status");
status.setAttribute("aria-live","polite");
status.style.cssText="margin-top:12px;padding:11px 12px;border-radius:10px;background:#edf3ff;color:#23416f;line-height:1.5";
const switchButton=document.createElement("button");
switchButton.type="button";
switchButton.textContent="Змінити учня";
switchButton.style.cssText="margin-top:9px;min-height:44px;padding:10px 13px;border:1px solid #a7b7d2;background:white;color:#27488a;border-radius:10px;font:inherit;font-weight:650;cursor:pointer";
const container=fields[0].closest("section")||fields[0].parentElement;
const note=document.createElement("p");
note.style.cssText="font-size:13px;color:#607089;margin:8px 0 0";
note.textContent="На спільному телефоні перед роботою перевір ім’я. Після передавання пристрою іншому учневі натисни «Змінити учня».";
container.append(status,switchButton,note);
let profile=null;
try{profile=JSON.parse(localStorage.getItem(KEY)||"null")}catch{}
if(profile && ids.every(id=>typeof profile[id]==="string")){
 ids.forEach((id,i)=>{fields[i].value=profile[id]});
}
function normalized(){
 const o={};
 ids.forEach((id,i)=>o[id]=fields[i].value.trim().replace(/\s+/g," "));
 return o;
}
function update(){
 const obj=normalized();
 const complete=ids.every(id=>obj[id]);
 if(complete){
  try{localStorage.setItem(KEY,JSON.stringify(obj))}catch{}
  status.textContent="👤 Учень: "+obj.surname+" "+obj.firstname+" · "+obj.klass+" · "+obj.group+" група. Дані підставляються в усі вправи на цьому пристрої.";
 }else{
  try{localStorage.removeItem(KEY)}catch{}
  status.textContent="Вкажи клас, групу, прізвище та ім’я один раз — у наступних вправах вони заповняться автоматично.";
 }
}
for(const f of fields){f.addEventListener("input",update);f.addEventListener("change",update)}
switchButton.addEventListener("click",()=>{
 if(!confirm("Змінити учня на цьому пристрої? Збережені результати попередніх вправ залишаться в історії браузера."))return;
 fields.forEach(f=>f.value="");
 try{localStorage.removeItem(KEY)}catch{}
 update();
 fields[0].focus();
});
update();
})();