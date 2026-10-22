const SHELL='personnel-shell-v4.2.2';
const FILES=['./','./index.html','./styles.css?v=4.2.2','./script.js?v=4.2.2','./report-ui.js?v=4.2.2','./saby-login.js?v=4.2.2','./workspace.js?v=4.2.2','./manifest.webmanifest'];
self.addEventListener('install',event=>{event.waitUntil(caches.open(SHELL).then(cache=>cache.addAll(FILES)));self.skipWaiting();});
self.addEventListener('activate',event=>{event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k.startsWith('personnel-shell-')&&k!==SHELL).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',event=>{
  const url=new URL(event.request.url);
  // Cache public UI only. Never persist tokens, API responses or reports.
  if(event.request.method!=='GET'||url.origin!==self.location.origin||!FILES.some(path=>new URL(path,self.registration.scope).href===url.href))return;
  event.respondWith(fetch(event.request).catch(()=>caches.match(event.request)));
});
