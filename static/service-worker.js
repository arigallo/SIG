const SIG_CACHE = "sig-pwa-v1";
const APP_SHELL = [
    "/",
    "/portal",
    "/static/styles.css",
    "/static/app.js",
    "/static/img/logo.png",
];

self.addEventListener("install", (event) => {
    event.waitUntil(
        caches.open(SIG_CACHE)
            .then((cache) => cache.addAll(APP_SHELL))
            .then(() => self.skipWaiting())
            .catch(() => self.skipWaiting())
    );
});

self.addEventListener("activate", (event) => {
    event.waitUntil(
        caches.keys()
            .then((keys) => Promise.all(keys.filter((key) => key !== SIG_CACHE).map((key) => caches.delete(key))))
            .then(() => self.clients.claim())
    );
});

self.addEventListener("fetch", (event) => {
    if (event.request.method !== "GET") {
        return;
    }

    const url = new URL(event.request.url);
    const cacheable = (
        url.origin === self.location.origin
        && (
            url.pathname.startsWith("/static/")
            || url.pathname === "/manifest.webmanifest"
            || url.pathname === "/portal"
        )
    );

    event.respondWith(
        fetch(event.request)
            .then((response) => {
                if (cacheable && response.ok) {
                    const copy = response.clone();
                    caches.open(SIG_CACHE).then((cache) => cache.put(event.request, copy)).catch(() => {});
                }
                return response;
            })
            .catch(() => caches.match(event.request).then((cached) => cached || caches.match("/portal")))
    );
});

self.addEventListener("push", (event) => {
    let data = {};
    try {
        data = event.data ? event.data.json() : {};
    } catch (error) {
        data = { body: event.data ? event.data.text() : "" };
    }

    const title = data.title || "SIG";
    const options = {
        body: data.body || "Tenes una nueva notificacion.",
        icon: data.icon || "/static/img/logo.png",
        badge: data.badge || "/static/img/logo.png",
        data: {
            url: data.url || "/",
        },
    };

    event.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener("notificationclick", (event) => {
    event.notification.close();
    const url = event.notification.data && event.notification.data.url ? event.notification.data.url : "/";
    event.waitUntil(
        self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((clients) => {
            const existing = clients.find((client) => client.url.includes(url));
            if (existing) {
                return existing.focus();
            }
            return self.clients.openWindow(url);
        })
    );
});
