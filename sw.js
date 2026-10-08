// Service worker da app Finanças da Casa.
// Estratégia: rede primeiro, cache como reserva (para abrir sem internet).
// Os pedidos ao Firebase/Google nunca são intercetados.
const CACHE = 'financas-v2';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => {}));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (/(firestore|firebase|identitytoolkit|securetoken|googleapis\.com\/(identity|v1)|gstatic\.com\/firebasejs|accounts\.google)/.test(url.href)) return;
  if (url.pathname.includes('/__/auth/')) return;

  e.respondWith(
    fetch(req)
      .then(res => {
        if (res.ok && (url.origin === location.origin || /fonts\.(googleapis|gstatic)|cdn\.jsdelivr/.test(url.host))) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy));
        }
        return res;
      })
      .catch(() => caches.match(req).then(r => r || (req.mode === 'navigate' ? caches.match('./index.html') : undefined)))
  );
});
