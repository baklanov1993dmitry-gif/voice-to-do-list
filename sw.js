// Кэширует только оболочку приложения (HTML/иконки/манифест), чтобы форма
// открывалась мгновенно. Все обращения к API (сохранение, статистика) всегда
// идут в сеть напрямую — их кэшировать нельзя, данные должны быть свежими.
const CACHE = 'dnevnik-shell-v1';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  const isShellAsset = SHELL.some((path) => url.pathname.endsWith(path.replace('./', '')));
  if (isShellAsset) {
    e.respondWith(caches.match(e.request).then((cached) => cached || fetch(e.request)));
  }
  // Всё остальное (запросы к Apps Script) не перехватываем — идёт обычным путём в сеть.
});
