// Garde l'app en cache pour qu'elle s'ouvre instantanément (même avec un mauvais réseau).
// Changer VERSION à chaque mise à jour pour que l'iPhone récupère les nouveaux fichiers.
const VERSION = "sirob-mobile-1.3";
const SHELL = ["./", "index.html", "style.css", "app.js", "config.js", "manifest.webmanifest",
               "icons/icon-180.png", "icons/icon-192.png", "icons/icon-512.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// Fichiers de l'app : réseau d'abord (mises à jour immédiates), cache si hors ligne.
// Microsoft (connexion, Graph, OneDrive) n'est jamais mis en cache.
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  e.respondWith(
    // « no-cache » : toujours vérifier auprès de GitHub, sinon iOS garde l'ancienne version 10 min.
    fetch(e.request, { cache: "no-cache" })
      .then((resp) => {
        if (resp.ok && !url.search) {
          const copy = resp.clone();
          caches.open(VERSION).then((c) => c.put(e.request, copy));
        }
        return resp;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
