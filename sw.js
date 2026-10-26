// Sổ Thu Chi — chạy offline + tự cập nhật.
// Có mạng: luôn lấy trang mới nhất từ web. Mất mạng (hoặc mạng quá chậm): dùng bản đã lưu trên máy.
// Khi sửa file, đổi số phiên bản dưới đây.
const CACHE = 'so-thu-chi-v4';
const ASSETS = ['./', './index.html', './manifest.webmanifest', './xlsx.full.min.js',
  './jszip.min.js', './template.xlsx',
  './icon-192.png', './icon-512.png', './apple-touch-icon.png'];
const NET_FIRST = /(\/|\.html|\.xlsx|\.webmanifest)$/;

self.addEventListener('install', e => {
  // cache:'reload' = bỏ qua bộ nhớ đệm trình duyệt, lấy đúng bản mới trên web
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS.map(u => new Request(u, { cache: 'reload' })))));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

async function networkFirst(req, key) {
  const cache = await caches.open(CACHE);
  try {
    const res = await Promise.race([
      fetch(req.url, { cache: 'no-store', credentials: 'same-origin' }),
      new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 5000))
    ]);
    if (res.ok) cache.put(key, res.clone());
    return res;
  } catch (err) {
    return (await cache.match(key, { ignoreSearch: true }))
        || (await cache.match('./index.html'))
        || Response.error();
  }
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;
  if (req.mode === 'navigate') { e.respondWith(networkFirst(req, './index.html')); return; }
  if (NET_FIRST.test(url.pathname)) { e.respondWith(networkFirst(req, req)); return; }
  e.respondWith(caches.match(req, { ignoreSearch: true }).then(hit => hit || fetch(req).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return res;
  })));
});
