/* Service worker：預先快取整個 App，之後完全離線也能使用。
 *
 * 更新方式：只要修改任何檔案，就把下面的 VERSION 加一（同時更新 app.js 的 APP_VERSION）。
 * tests/sw.test.mjs 會檢查 ASSETS 是否涵蓋 docs/ 內的所有檔案，以及兩處版本是否一致。
 */
const VERSION = '1.0.0';
const CACHE = `order-tool-${VERSION}`;

const ASSETS = [
  './',
  'index.html',
  'styles.css',
  'app.js',
  'core.js',
  'xlsx-io.js',
  'manifest.webmanifest',
  'vendor/exceljs.bare.min.js',
  'icons/icon.svg',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-512.png',
  'icons/apple-touch-icon.png',
];

self.addEventListener('install', (event) => {
  // 不自動 skipWaiting：由使用者在畫面上按「立即更新」後才切換，避免打斷處理中的工作。
  event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(ASSETS)));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith('order-tool-') && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener('message', (event) => {
  if (event.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    // 頁面導覽：一律回應快取的 index.html（忽略 ?query）
    const cached = await cache.match(request.mode === 'navigate' ? 'index.html' : request, { ignoreSearch: true });
    if (cached) return cached;
    try {
      return await fetch(request);
    } catch (err) {
      if (request.mode === 'navigate') return (await cache.match('index.html')) || Response.error();
      throw err;
    }
  })());
});
