// Retire the legacy app shell. Never cache or run the old broad-permission app.
self.addEventListener('install',event=>event.waitUntil(self.skipWaiting()));
self.addEventListener('activate',event=>event.waitUntil((async()=>{
 const names=await caches.keys();
 await Promise.all(names.filter(name=>name.startsWith('notas-mobile-')).map(name=>caches.delete(name)));
 await self.clients.claim();
})()));
self.addEventListener('fetch',event=>{
 if(event.request.mode!=='navigate')return;
 const url=new URL(event.request.url), root=new URL(self.registration.scope);
 if(url.origin===root.origin&&(url.pathname===root.pathname||url.pathname===root.pathname+'index.html')){
  event.respondWith(Response.redirect(new URL('restrito/',root).href,302));
 }
});
