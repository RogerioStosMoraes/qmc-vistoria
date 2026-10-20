// Service worker: guarda o app no aparelho para funcionar sem internet.
const VERSION = 'v1.2.0';
const CACHE = 'qmc-vistoria-' + VERSION;
const SHELL = [
  './', 'index.html', 'styles.css', 'config.js', 'manifest.webmanifest',
  'js/app.js', 'js/model.js', 'js/db.js', 'js/cloud.js', 'js/pptx-builder.js',
  'vendor/jszip.min.js', 'vendor/msal-browser.min.js',
  'data/checklist.json', 'data/acompanhamento.json', 'assets/template.pptx', 'assets/qmc-logo.png',
  'assets/logos/vivo.png', 'assets/logos/claro.png', 'assets/logos/tim.png',
  'icons/icon-192.png', 'icons/icon-512.png',
];
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return; // Microsoft/Graph: sempre rede
  e.respondWith(
    caches.match(req, { ignoreSearch: true }).then((hit) => {
      const net = fetch(req).then((res) => {
        if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); }
        return res;
      }).catch(() => hit);
      return hit || net; // cache primeiro; atualiza em segundo plano
    })
  );
});
