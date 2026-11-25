const CACHE='notas-mobile-v6';
const FILES=['./index.html','./style.css?v=15','./app.js?v=15','./format.js?v=14','./manifest.webmanifest','./icon-192.png','./icon-512.png'];
const assets=new Set(FILES.map(path=>new URL(path,self.registration.scope).href));

self.addEventListener('install',event=>{
 event.waitUntil((async()=>{
  const cache=await caches.open(CACHE);
  await cache.addAll(FILES.map(path=>new Request(path,{cache:'reload'})));
  await self.skipWaiting();
 })());
});

self.addEventListener('activate',event=>{
 event.waitUntil((async()=>{
  const names=await caches.keys();
  await Promise.all(names.filter(name=>name.startsWith('notas-mobile-')&&name!==CACHE).map(name=>caches.delete(name)));
  await self.clients.claim();
 })());
});

self.addEventListener('fetch',event=>{
 const request=event.request;
 if(request.method!=='GET'||!request.url.startsWith(self.registration.scope))return;
 const navigation=request.mode==='navigate';
 if(!navigation&&!assets.has(request.url))return;
 event.respondWith((async()=>{
  let response;
  try{
   response=await fetch(request,{cache:'no-cache'});
   if(response.ok)return response;
  }catch{}
  const cache=await caches.open(CACHE);
  const fallback=await cache.match(navigation?'./index.html':request);
  return fallback||response||Response.error();
 })());
});
