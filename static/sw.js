const CACHE_NAME = "rendiciones-shell-v1";

// No se almacenan páginas privadas, fotografías, datos ni archivos PDF.
self.addEventListener("install", event => {
    self.skipWaiting();
});

self.addEventListener("activate", event => {
    event.waitUntil((async () => {
        const keys = await caches.keys();

        await Promise.all(
            keys
                .filter(key => key.startsWith("rendiciones-"))
                .filter(key => key !== CACHE_NAME)
                .map(key => caches.delete(key))
        );

        await self.clients.claim();
    })());
});

// Las solicitudes se realizan directamente al servidor.
// No se interceptan ni se guardan respuestas en caché.
self.addEventListener("fetch", event => {
    // Sin estrategia de caché: la aplicación requiere conexión.
});
