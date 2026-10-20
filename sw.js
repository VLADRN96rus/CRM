// Сервис-воркер: приложение открывается без интернета. Файлы приложения берутся из сети в первую очередь
// (новый релиз виден сразу, без чистки кэша), а без интернета — из кэша. Номер версии V меняйте при релизе.
const V = 'crm-v5';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url), same = u.origin === location.origin;
  const fonts = u.hostname === 'fonts.googleapis.com' || u.hostname === 'fonts.gstatic.com';
  if (!same && !fonts) return;
  e.respondWith((async () => {
    const c = await caches.open(V), nav = r.mode === 'navigate';
    const hit = await c.match(r, { ignoreSearch: nav });
    if (same) {   // сеть первой (в обход HTTP-кэша GitHub Pages), через 5 с — кэш
      try {
        const res = await Promise.race([fetch(r.url, { cache: 'no-cache' }), new Promise((_, no) => setTimeout(no, 5000))]);
        if (res.ok) { c.put(r, res.clone()); return res; }
        return hit || res;
      } catch (err) {
        return hit || (nav ? await c.match('index.html') : Response.error());
      }
    }
    const net = fetch(r).then(res => { if (res && (res.ok || res.type === 'opaque')) c.put(r, res.clone()); return res; }).catch(() => null);
    if (hit) { e.waitUntil(net); return hit; }
    return (await net) || Response.error();
  })());
});
