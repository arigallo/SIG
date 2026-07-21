const SIG_CACHE = "sig-pwa-v3";
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
    if (url.origin !== self.location.origin) {
        return;
    }

    if (url.pathname === "/manifest.webmanifest") {
        event.respondWith(fetch(event.request));
        return;
    }

    const isStatic = url.pathname.startsWith("/static/");
    if (isStatic) {
        event.respondWith(
            caches.match(event.request).then((cached) => {
                const network = fetch(event.request).then((response) => {
                    if (response.ok) {
                        caches.open(SIG_CACHE).then((cache) => cache.put(event.request, response.clone())).catch(() => {});
                    }
                    return response;
                });
                return cached || network;
            })
        );
        return;
    }

    if (event.request.mode === "navigate") {
        event.respondWith(
            fetch(event.request).catch(() => caches.match("/portal"))
        );
    }
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
