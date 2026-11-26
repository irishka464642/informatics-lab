/* CODE QUEST 7 — автономний запуск після першого відвідання онлайн.
   Для роботи програми достатньо index.html: весь CSS і JavaScript уже всередині. */
const CACHE_NAME="codequest7-offline-v3-20261127";
const APP_URL=new URL("index.html",self.registration.scope).href;
const HOME_URL=new URL("./",self.registration.scope).href;

self.addEventListener("install",event=>{
 event.waitUntil(
  caches.open(CACHE_NAME).then(cache=>cache.addAll([APP_URL,HOME_URL]))
   .then(()=>self.skipWaiting())
 );
});
self.addEventListener("activate",event=>{
 event.waitUntil(
  caches.keys().then(keys=>Promise.all(keys
   .filter(name=>name.startsWith("codequest7-offline-")&&name!==CACHE_NAME)
   .map(name=>caches.delete(name))))
   .then(()=>self.clients.claim())
 );
});
self.addEventListener("fetch",event=>{
 const req=event.request;
 if(req.method!=="GET"||new URL(req.url).origin!==self.location.origin)return;
 if(!req.url.startsWith(self.registration.scope))return;
 if(req.mode!=="navigate"&&new URL(req.url).pathname!==new URL(APP_URL).pathname)return;
 event.respondWith(
  fetch(req).then(response=>{
   if(response.ok){
    const copy=response.clone();
    event.waitUntil(caches.open(CACHE_NAME).then(cache=>cache.put(APP_URL,copy)));
   }
   return response;
  }).catch(async()=>{
   const cache=await caches.open(CACHE_NAME);
   const fallback=await cache.match(req)||await cache.match(APP_URL)||await cache.match(HOME_URL);
   return fallback||Response.error();
  })
 );
});
