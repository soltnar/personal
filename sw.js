const SHELL='personnel-shell-v5.0.17';
const FILES=['./','./index.html','./styles.css?v=5.0.17','./script.js?v=5.0.17','./report-ui.js?v=5.0.17','./saby-login.js?v=5.0.17','./workspace.js?v=5.0.17','./access-ui.js?v=5.0.17','./role-workspace.css?v=5.0.17','./role-workspace.js?v=5.0.17','./vendor/feather.min.js','./manifest.webmanifest','./icon.svg?v=5.0.17','./personnel-icon-32.png','./personnel-icon-180.png','./personnel-icon-192.png','./personnel-icon-512.png','./personnel-icon-maskable-512.png'];
self.addEventListener('install',event=>{event.waitUntil(caches.open(SHELL).then(cache=>cache.addAll(FILES)));self.skipWaiting();});
self.addEventListener('activate',event=>{event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k.startsWith('personnel-shell-')&&k!==SHELL).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',event=>{
  const url=new URL(event.request.url);
  // Cache public UI only. Never persist tokens, API responses or reports.
  if(event.request.method!=='GET'||url.origin!==self.location.origin||!FILES.some(path=>new URL(path,self.registration.scope).href===url.href))return;
  event.respondWith(fetch(event.request).catch(()=>caches.match(event.request)));
});
