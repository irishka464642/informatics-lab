/* Академія рішень PRO — офлайн-копія гри (тільки власна сторінка). */
"use strict";
const CACHE_NAME="academy-branching-pro-10-offline-v1";
const SCOPE=new URL(self.registration.scope);
const OFFLINE_URL=new URL("index.html",SCOPE).toString();

self.addEventListener("install",event=>{
 event.waitUntil(
  caches.open(CACHE_NAME)
   .then(cache=>cache.add(OFFLINE_URL))
   .then(()=>self.skipWaiting())
 );
});

self.addEventListener("activate",event=>{
 event.waitUntil(
  caches.keys()
   .then(names=>Promise.all(names
    .filter(name=>name.startsWith("academy-branching-pro-10-offline-")&&name!==CACHE_NAME)
    .map(name=>caches.delete(name))))
   .then(()=>self.clients.claim())
 );
});

self.addEventListener("fetch",event=>{
 const request=event.request;
 if(request.method!=="GET"||request.mode!=="navigate")return;
 const url=new URL(request.url);
 if(url.origin!==SCOPE.origin||!url.pathname.startsWith(SCOPE.pathname))return;
 event.respondWith(
  fetch(request).then(response=>{
   if(response.ok){
    const copy=response.clone();
    event.waitUntil(caches.open(CACHE_NAME).then(cache=>cache.put(OFFLINE_URL,copy)));
   }
   return response;
  }).catch(async()=>{
   const cached=await caches.match(OFFLINE_URL);
   return cached||Response.error();
  })
 );
});
