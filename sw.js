const CACHE='john-quiz-v4';
const CORE=['./','index.html','manifest.webmanifest','icon-192.png','icon-512.png','apple-touch-icon.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(CORE)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const r=e.request; if(r.method!=='GET') return;
  const u=new URL(r.url);
  if(u.pathname.endsWith('.mp3')) return; // songs stream straight from the network
  if(u.origin!==location.origin){ if(u.hostname.endsWith('fonts.googleapis.com')||u.hostname.endsWith('fonts.gstatic.com')) e.respondWith(caches.open(CACHE).then(c=>c.match(r).then(m=>m||fetch(r).then(res=>{c.put(r,res.clone());return res})))); return; }
  // pages always ask the server for the newest version (no stale copies after an update)
  const fresh=(r.mode==='navigate'||u.pathname.endsWith('/')||u.pathname.endsWith('.html'))?fetch(r,{cache:'no-cache'}):fetch(r);
  e.respondWith(fresh.then(res=>{ if(res.ok){ const copy=res.clone(); caches.open(CACHE).then(c=>c.put(r,copy)); } return res; }).catch(()=>caches.match(r).then(m=>m||caches.match('index.html'))));
});
