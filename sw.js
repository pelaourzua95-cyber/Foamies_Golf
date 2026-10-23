// Página siempre desde la red (versión más nueva); si no hay señal, la última copia guardada.
const CACHE="tarjeta-golf";
const FILES=["./","./index.html","./manifest.webmanifest","./icons/icon-192.png","./icons/icon-512.png","./icons/apple-touch-icon.png"];
self.addEventListener("install",e=>{self.skipWaiting();e.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES)));});
self.addEventListener("activate",e=>{e.waitUntil(self.clients.claim());});
self.addEventListener("fetch",e=>{
  const r=e.request;if(r.method!=="GET")return;
  const save=res=>{if(res.ok||res.type==="opaque"){const copy=res.clone();caches.open(CACHE).then(c=>c.put(r,copy));}return res;};
  if(r.mode==="navigate"){
    e.respondWith(fetch(r,{cache:"no-store"}).then(save).catch(()=>caches.match(r,{ignoreSearch:true}).then(h=>h||caches.match("./index.html"))));
    return;}
  // resto (íconos, fuentes): rápido desde la copia y se refresca por detrás
  e.respondWith(caches.match(r).then(hit=>{const net=fetch(r).then(save).catch(()=>hit);return hit||net;}));
});
