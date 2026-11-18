// Mój budżet – service worker
// Strona: najpierw sieć (żeby nowa wersja wchodziła od razu), zdjęcia: z pamięci podręcznej.
// Dane z arkusza zawsze idą prosto do sieci.
const CACHE = 'budzet-v6';
const SHELL = ['./', 'index.html', 'manifest.json',
  'img/m01.jpg','img/m02.jpg','img/m03.jpg','img/m04.jpg','img/m05.jpg','img/m06.jpg',
  'img/m07.jpg','img/m08.jpg','img/m09.jpg','img/m10.jpg','img/m11.jpg','img/m12.jpg',
  'img/sky.jpg','img/seoul.jpg','img/silk.jpg'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;           // arkusz, fonty – bez ingerencji
  if (url.pathname.includes('/img/')) {
    e.respondWith(caches.match(req).then((hit) => hit || fetch(req).then((res) => {
      const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); return res;
    })));
    return;
  }
  e.respondWith(fetch(req).then((res) => {
    const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); return res;
  }).catch(() => caches.match(req).then((hit) => hit || caches.match('index.html'))));
});
