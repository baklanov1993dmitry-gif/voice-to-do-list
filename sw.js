// Стратегия: сама страница (index.html) всегда берётся из сети, пока есть
// интернет — так любое обновление кода сразу видно, без плясок с версиями
// кэша. Кэш для неё — только запасной вариант на случай отсутствия сети.
// Иконки и манифест меняются редко, их кэшируем сразу и отдаём мгновенно.
const CACHE = 'dnevnik-shell-v1';
const SHELL = ['./manifest.json', './icon-192.png', './icon-512.png'];

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
  const isHtml = url.pathname.endsWith('index.html') || url.pathname.endsWith('/');

  if (isHtml) {
    // Сеть в приоритете: всегда пытаемся получить свежую версию страницы.
    // Если сети нет — отдаём то, что успели закэшировать в прошлый раз.
    e.respondWith(
      fetch(e.request)
        .then((resp) => {
          caches.open(CACHE).then((c) => c.put(e.request, resp.clone()));
          return resp;
        })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  const isShellAsset = SHELL.some((path) => url.pathname.endsWith(path.replace('./', '')));
  if (isShellAsset) {
    e.respondWith(caches.match(e.request).then((cached) => cached || fetch(e.request)));
  }
  // Всё остальное (запросы к Apps Script) не перехватываем — идёт обычным путём в сеть.
});
