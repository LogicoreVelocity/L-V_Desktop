self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open('logicore-excel-manager-v1').then((cache) =>
      cache.addAll(['/', '/manifest.webmanifest', '/favicon.svg', '/icons.svg']),
    ),
  )
})

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached
      return fetch(event.request)
    }),
  )
})
