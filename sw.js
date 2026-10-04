// Версию поднимай при каждом обновлении оболочки (иконки/манифест). index.html и так грузится «сначала из сети».
const CACHE = "moidela-shell-v5";
const SHELL = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png", "./icon-512-maskable.png"];

self.addEventListener("install", e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => {})))));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;          // Supabase и всё чужое — мимо кэша, всегда напрямую
  if (url.pathname.endsWith("build.txt")) return;      // метка версии — всегда с сервера
  const isShell = req.mode === "navigate" || url.pathname.endsWith("/") || url.pathname.endsWith("index.html") || url.pathname.endsWith("manifest.json");
  if (isShell) {                                        // страница и манифест: сначала сеть, кэш — только если сети нет
    e.respondWith(fetch(req).then(res => {
      if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    }).catch(() => caches.match(req).then(r => r || caches.match("./index.html"))));
    return;
  }
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {   // иконки: кэш первым
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return res;
  })));
});

// ---- Web Push: приходит с сервера (Supabase Edge Function), работает при закрытом приложении ----
self.addEventListener("push", e => {
  let d = {}; try { d = e.data.json(); } catch (_) { d = { title: "Напоминание", body: e.data ? e.data.text() : "" }; }
  e.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(cs => {
    if (cs.some(c => c.visibilityState === "visible")) return;       // приложение открыто — оно само покажет всплывашку
    return self.registration.showNotification(d.title || "Напоминание", {
      body: d.body || "", tag: d.tag || "moidela", renotify: true, requireInteraction: true,
      icon: "icon-192.png", badge: "icon-192.png", vibrate: [250, 100, 250, 100, 250], data: { id: d.id || null }
    });
  }));
});
self.addEventListener("notificationclick", e => {
  e.notification.close();
  const id = e.notification.data && e.notification.data.id;
  e.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(cs => {
    if (cs.length) { cs[0].postMessage({ type: "open", id }); return cs[0].focus(); }
    return self.clients.openWindow("./" + (id ? "?t=" + encodeURIComponent(id) : ""));
  }));
});
