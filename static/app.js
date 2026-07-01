(() => {
    const setupNavigationGroups = () => {
        const groups = Array.from(document.querySelectorAll(".main-nav .nav-group"));
        if (!groups.length) {
            return;
        }

        const closeOthers = (activeGroup) => {
            groups.forEach((group) => {
                if (group !== activeGroup) {
                    group.removeAttribute("open");
                }
            });
        };

        groups.forEach((group) => {
            if (!group.classList.contains("current")) {
                group.removeAttribute("open");
            }
        });

        groups.forEach((group) => {
            group.addEventListener("toggle", () => {
                if (group.open) {
                    closeOthers(group);
                }
            });
        });

        document.querySelectorAll(".main-nav .nav-menu a").forEach((link) => {
            link.addEventListener("click", () => {
                closeOthers(null);
            });
        });

        document.addEventListener("click", (event) => {
            if (!event.target.closest(".main-nav")) {
                closeOthers(null);
            }
        });

        document.addEventListener("keydown", (event) => {
            if (event.key === "Escape") {
                closeOthers(null);
            }
        });
    };

    setupNavigationGroups();

    const csrfToken = document.querySelector('meta[name="csrf-token"]')?.content || "";
    const username = document.body?.dataset.username || "";

    const postPresenceHeartbeat = () => {
        if (!username || !csrfToken) {
            return;
        }
        fetch("/presencia/heartbeat", {
            method: "POST",
            headers: {
                "X-CSRF-Token": csrfToken,
            },
            credentials: "same-origin",
            keepalive: true,
        }).catch(() => {});
    };

    if (username) {
        postPresenceHeartbeat();
        window.setInterval(postPresenceHeartbeat, 30000);
    }

    const whatsappBadge = document.querySelector("[data-whatsapp-inbox-badge]");
    const notificationsBadge = document.querySelector("[data-notifications-badge]");
    let lastWhatsappInboxId = null;

    const setWhatsappBadge = (count) => {
        if (!whatsappBadge) {
            return;
        }
        const normalized = Number(count) || 0;
        whatsappBadge.textContent = String(normalized);
        whatsappBadge.hidden = normalized <= 0;
    };

    const setNotificationsBadgeFromWhatsapp = (whatsappCount) => {
        if (!notificationsBadge) {
            return;
        }
        const baseCount = Number(notificationsBadge.dataset.baseCount) || 0;
        const normalized = baseCount + (Number(whatsappCount) || 0);
        notificationsBadge.textContent = String(normalized);
        notificationsBadge.hidden = normalized <= 0;
    };

    const pollWhatsappInbox = async () => {
        if (!whatsappBadge) {
            return;
        }
        try {
            const response = await fetch("/comunicacion/whatsapp/estado", {
                credentials: "same-origin",
                headers: {"Accept": "application/json"},
            });
            if (!response.ok) {
                return;
            }
            const data = await response.json();
            const latestId = Number(data.ultimo_id) || 0;
            const previousId = lastWhatsappInboxId;
            lastWhatsappInboxId = latestId;
            setWhatsappBadge(data.sin_leer);
            setNotificationsBadgeFromWhatsapp(data.sin_leer);

            if (
                previousId !== null
                && latestId > previousId
                && document.body?.dataset.endpoint === "ver_whatsapp_inbox"
                && !document.hidden
            ) {
                window.location.reload();
            }

            if (
                previousId !== null
                && latestId > previousId
                && document.body?.dataset.endpoint === "ver_notificaciones"
                && !document.hidden
            ) {
                window.location.reload();
            }
        } catch (error) {
            // Best effort: the next poll will try again.
        }
    };

    if (whatsappBadge) {
        pollWhatsappInbox();
        window.setInterval(pollWhatsappInbox, 10000);
    }

    const prioritySelectAll = document.querySelector("[data-priority-select-all]");
    const priorityCheckboxes = Array.from(document.querySelectorAll("[data-priority-notification]"));
    if (prioritySelectAll && priorityCheckboxes.length) {
        const syncPrioritySelectAll = () => {
            const selected = priorityCheckboxes.filter((checkbox) => checkbox.checked).length;
            prioritySelectAll.checked = selected === priorityCheckboxes.length;
            prioritySelectAll.indeterminate = selected > 0 && selected < priorityCheckboxes.length;
        };

        prioritySelectAll.addEventListener("change", () => {
            priorityCheckboxes.forEach((checkbox) => {
                checkbox.checked = prioritySelectAll.checked;
            });
            syncPrioritySelectAll();
        });

        priorityCheckboxes.forEach((checkbox) => {
            checkbox.addEventListener("change", syncPrioritySelectAll);
        });
    }

    document.querySelectorAll("[data-whatsapp-reply]").forEach((button) => {
        button.addEventListener("click", () => {
            const textarea = document.querySelector("#mensaje");
            if (!textarea) {
                return;
            }
            textarea.value = button.dataset.whatsappReply || "";
            textarea.focus();
        });
    });

    const acceptedMimeTypes = new Set([
        "application/pdf",
        "image/jpeg",
        "image/png",
    ]);

    const acceptedExtensions = new Set(["pdf", "jpg", "jpeg", "png"]);

    const fileExtension = (name) => {
        const parts = (name || "").split(".");
        return parts.length > 1 ? parts.pop().toLowerCase() : "";
    };

    const parseAcceptedTypes = (input) => {
        const accept = (input?.getAttribute("accept") || "").trim();
        if (!accept) {
            return {
                mimeTypes: acceptedMimeTypes,
                extensions: acceptedExtensions,
            };
        }

        const mimeTypes = new Set();
        const extensions = new Set();
        accept.split(",").map((item) => item.trim().toLowerCase()).filter(Boolean).forEach((item) => {
            if (item.startsWith(".")) {
                extensions.add(item.slice(1));
            } else {
                mimeTypes.add(item);
            }
        });
        return { mimeTypes, extensions };
    };

    const isAcceptedFile = (file, input) => {
        if (!file) {
            return false;
        }

        const allowed = parseAcceptedTypes(input);
        return allowed.mimeTypes.has((file.type || "").toLowerCase()) || allowed.extensions.has(fileExtension(file.name));
    };

    const filesLabel = (files) => {
        const selected = Array.from(files || []);
        if (!selected.length) {
            return "Ningun archivo seleccionado";
        }

        if (selected.length > 1) {
            const totalMb = selected.reduce((total, file) => total + file.size, 0) / (1024 * 1024);
            return `${selected.length} archivos seleccionados (${totalMb.toFixed(totalMb >= 1 ? 1 : 2)} MB)`;
        }

        const file = selected[0];
        const sizeMb = file.size / (1024 * 1024);
        return `${file.name} (${sizeMb.toFixed(sizeMb >= 1 ? 1 : 2)} MB)`;
    };

    const normalizeClipboardFile = (file) => {
        if (!file) {
            return null;
        }
        if (file.name) {
            return file;
        }

        const extension = file.type === "image/jpeg" ? "jpg" : "png";
        return new File([file], `comprobante-portapapeles.${extension}`, {
            type: file.type || "image/png",
        });
    };

    const setInputFiles = (input, files) => {
        const transfer = new DataTransfer();
        Array.from(files || []).forEach((file) => transfer.items.add(file));
        input.files = transfer.files;
        input.dispatchEvent(new Event("change", { bubbles: true }));
    };

    const ensureDropzoneStatus = (dropzone) => {
        let status = dropzone.querySelector("[data-upload-file-name]");
        if (!status) {
            status = document.createElement("span");
            status.className = "upload-file";
            status.dataset.uploadFileName = "";
            status.textContent = "Ningun archivo seleccionado";
            dropzone.appendChild(status);
        }
        return status;
    };

    const setDropzoneStatus = (dropzone, message, isError = false) => {
        const status = ensureDropzoneStatus(dropzone);
        status.textContent = message;
        status.classList.toggle("upload-file-error", isError);
    };

    const applyFilesToDropzone = (dropzone, input, files) => {
        const selected = Array.from(files || [])
            .map(normalizeClipboardFile)
            .filter(Boolean);
        const nextFiles = input.multiple ? selected : selected.slice(0, 1);

        if (!nextFiles.length || nextFiles.some((file) => !isAcceptedFile(file, input))) {
            const accept = input.getAttribute("accept");
            const humanAccept = accept
                ? accept.replaceAll(",", ", ").replaceAll(".", "").toUpperCase()
                : "PDF, JPG o PNG";
            setDropzoneStatus(dropzone, `Solo se aceptan ${humanAccept}.`, true);
            return false;
        }

        setInputFiles(input, nextFiles);
        setDropzoneStatus(dropzone, filesLabel(nextFiles));
        dropzone.classList.add("has-file");
        return true;
    };

    const setupDropzone = (dropzone) => {
        const input = dropzone.querySelector('input[type="file"]');
        if (!input) {
            return;
        }

        ensureDropzoneStatus(dropzone);
        dropzone.__setUploadFiles = (files) => applyFilesToDropzone(dropzone, input, files);
        dropzone.__setUploadFile = (file) => applyFilesToDropzone(dropzone, input, [file]);

        input.addEventListener("change", () => {
            const files = input.files || [];
            setDropzoneStatus(dropzone, filesLabel(files));
            dropzone.classList.toggle("has-file", Boolean(files.length));
        });

        dropzone.addEventListener("dragover", (event) => {
            event.preventDefault();
            dropzone.classList.add("is-dragging");
        });

        dropzone.addEventListener("dragenter", (event) => {
            event.preventDefault();
            dropzone.classList.add("is-dragging");
        });

        dropzone.addEventListener("dragleave", () => {
            dropzone.classList.remove("is-dragging");
        });

        dropzone.addEventListener("drop", (event) => {
            event.preventDefault();
            dropzone.classList.remove("is-dragging");

            const files = event.dataTransfer && event.dataTransfer.files;
            if (files && files.length) {
                applyFilesToDropzone(dropzone, input, files);
            }
        });

        dropzone.addEventListener("keydown", (event) => {
            if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                input.click();
            }
        });
    };

    const dropzones = Array.from(document.querySelectorAll("[data-upload-dropzone]"));
    dropzones.forEach(setupDropzone);

    const findTargetDropzone = (button = null) => {
        if (button && button.dataset.uploadPaste) {
            const target = document.querySelector(button.dataset.uploadPaste);
            if (target) {
                return target;
            }
        }
        const activeDropzone = document.activeElement && document.activeElement.closest("[data-upload-dropzone]");
        if (activeDropzone) {
            return activeDropzone;
        }
        if (button) {
            const formDropzone = button.closest("form")?.querySelector("[data-upload-dropzone]");
            if (formDropzone) {
                return formDropzone;
            }
        }
        return dropzones.length === 1 ? dropzones[0] : null;
    };

    const applyClipboardItems = (items, targetDropzone) => {
        const files = Array.from(items || [])
            .filter((item) => item.kind === "file")
            .map((item) => item.getAsFile())
            .filter(Boolean);

        if (!files.length || !targetDropzone || typeof targetDropzone.__setUploadFiles !== "function") {
            return false;
        }
        return targetDropzone.__setUploadFiles(files);
    };

    document.addEventListener("paste", (event) => {
        if (!dropzones.length || !event.clipboardData) {
            return;
        }

        const targetDropzone = findTargetDropzone();
        if (applyClipboardItems(event.clipboardData.items, targetDropzone)) {
            event.preventDefault();
        }
    });

    document.querySelectorAll("[data-upload-paste]").forEach((button) => {
        button.addEventListener("click", async () => {
            const targetDropzone = findTargetDropzone(button);
            if (!targetDropzone) {
                return;
            }

            if (!navigator.clipboard || typeof navigator.clipboard.read !== "function") {
                setDropzoneStatus(targetDropzone, "Usa Ctrl+V para pegar una imagen del portapapeles.", true);
                targetDropzone.focus();
                return;
            }

            try {
                const clipboardItems = await navigator.clipboard.read();
                const files = [];
                for (const item of clipboardItems) {
                    const imageType = item.types.find((type) => type.startsWith("image/"));
                    if (!imageType) {
                        continue;
                    }
                    const blob = await item.getType(imageType);
                    const extension = imageType === "image/jpeg" ? "jpg" : "png";
                    files.push(new File([blob], `comprobante-portapapeles.${extension}`, { type: imageType }));
                }

                if (!files.length || !targetDropzone.__setUploadFiles(files)) {
                    setDropzoneStatus(targetDropzone, "No se encontro una imagen valida en el portapapeles.", true);
                }
            } catch (error) {
                setDropzoneStatus(targetDropzone, "No se pudo leer el portapapeles. Usa Ctrl+V sobre esta zona.", true);
                targetDropzone.focus();
            }
        });
    });

    const pwaActions = document.querySelector("[data-pwa-actions]");
    const installButton = document.querySelector("[data-pwa-install]");
    const inlineInstallButton = document.querySelector("[data-pwa-install-inline]");
    const pushButton = document.querySelector("[data-pwa-enable-push]");
    const testPushButton = document.querySelector("[data-pwa-test-push]");
    const pwaStatus = document.querySelector("[data-pwa-status]");
    const inlinePwaStatus = document.querySelector("[data-pwa-inline-status]");
    const pwaPermissionStatus = document.querySelector("[data-pwa-permission-status]");
    const pwaSubscriptionStatus = document.querySelector("[data-pwa-subscription-status]");
    const pwaWorkerStatus = document.querySelector("[data-pwa-worker-status]");
    const portalToken = document.body?.dataset.portalToken || "";
    const pwaStorageKey = `sig:pwa:test-ok:${portalToken || username || "default"}`;
    const pwaSavedKey = `sig:pwa:saved:v2:${portalToken || username || "default"}`;
    let deferredInstallPrompt = null;
    let pushConfig = null;
    let pwaSyncInFlight = false;

    const setPwaStatus = (message, isError = false) => {
        if (pwaStatus) {
            pwaStatus.textContent = message || "";
            pwaStatus.classList.toggle("text-danger", Boolean(isError));
        }
        if (inlinePwaStatus) {
            inlinePwaStatus.textContent = message || inlinePwaStatus.textContent;
            inlinePwaStatus.classList.toggle("text-danger", Boolean(isError));
        }
    };

    const showPwaActions = () => {
        if (pwaActions) {
            pwaActions.hidden = false;
        }
    };

    const urlBase64ToUint8Array = (base64String) => {
        const padding = "=".repeat((4 - base64String.length % 4) % 4);
        const base64 = (base64String + padding).replace(/-/g, "+").replace(/_/g, "/");
        const rawData = window.atob(base64);
        return Uint8Array.from([...rawData].map((char) => char.charCodeAt(0)));
    };

    const postPwaJson = async (url, body) => {
        const response = await fetch(url, {
            method: "POST",
            credentials: "same-origin",
            headers: {
                "Content-Type": "application/json",
                "X-CSRF-Token": csrfToken,
            },
            body: JSON.stringify(body || {}),
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok || data.ok === false) {
            throw new Error(data.error || "No se pudo completar la accion.");
        }
        return data;
    };

    const currentPushSubscription = async () => {
        if (!("serviceWorker" in navigator)) {
            return null;
        }
        const registration = await navigator.serviceWorker.ready;
        return registration.pushManager.getSubscription();
    };

    const updatePwaDiagnostics = async () => {
        if (pwaPermissionStatus) {
            pwaPermissionStatus.textContent = ("Notification" in window) ? Notification.permission : "no soportado";
        }
        if (pwaWorkerStatus) {
            pwaWorkerStatus.textContent = ("serviceWorker" in navigator) ? "disponible" : "no soportado";
        }
        if (pwaSubscriptionStatus) {
            try {
                const subscription = await currentPushSubscription();
                pwaSubscriptionStatus.textContent = subscription ? "suscripto en este navegador" : "sin suscripcion local";
            } catch (error) {
                pwaSubscriptionStatus.textContent = "no disponible";
            }
        }
    };

    const savePwaSubscription = async (subscription) => {
        if (!subscription) {
            return;
        }
        await postPwaJson("/pwa/push/subscribe", {
            subscription,
            portal_token: portalToken,
        });
        window.localStorage?.setItem(pwaSavedKey, "1");
    };

    const ensurePushSubscriptionSaved = async ({ requestPermission = false } = {}) => {
        if (pwaSyncInFlight || !pushConfig?.pushEnabled || !pushConfig?.vapidPublicKey) {
            return null;
        }
        if (!("serviceWorker" in navigator) || !("PushManager" in window) || !("Notification" in window)) {
            return null;
        }
        if (!portalToken && !username) {
            return null;
        }

        pwaSyncInFlight = true;
        try {
            let permission = Notification.permission;
            if (permission === "default" && requestPermission) {
                permission = await Notification.requestPermission();
            }
            if (permission !== "granted") {
                return null;
            }

            const registration = await navigator.serviceWorker.ready;
            let subscription = await registration.pushManager.getSubscription();
            if (!subscription && requestPermission) {
                subscription = await registration.pushManager.subscribe({
                    userVisibleOnly: true,
                    applicationServerKey: urlBase64ToUint8Array(pushConfig.vapidPublicKey),
                });
            }
            if (subscription) {
                await savePwaSubscription(subscription);
            }
            return subscription;
        } finally {
            pwaSyncInFlight = false;
        }
    };

    const refreshPwaButtons = async () => {
        if (!("serviceWorker" in navigator)) {
            return;
        }
        if (pushButton || testPushButton) {
            showPwaActions();
        }
        if (pushButton && "PushManager" in window && "Notification" in window && Notification.permission !== "denied") {
            const subscription = await currentPushSubscription();
            const testDone = window.localStorage?.getItem(pwaStorageKey) === "1";
            if (subscription) {
                await savePwaSubscription(subscription).catch(() => {});
            }
            pushButton.hidden = Boolean(subscription);
            testPushButton.hidden = !subscription || testDone;
            if (subscription && testDone && pwaActions) {
                pwaActions.hidden = true;
            }
        }
    };

    if ("serviceWorker" in navigator) {
        navigator.serviceWorker.register("/service-worker.js").then(async () => {
            showPwaActions();
            pushConfig = await fetch("/pwa/config", {
                credentials: "same-origin",
                headers: { "Accept": "application/json" },
            }).then((response) => response.json()).catch(() => null);
            await ensurePushSubscriptionSaved().catch(() => {});
            await refreshPwaButtons();
            await updatePwaDiagnostics();
        }).catch(() => {
            setPwaStatus("No se pudo preparar la app instalable.", true);
            updatePwaDiagnostics();
        });
    } else {
        updatePwaDiagnostics();
    }

    window.addEventListener("beforeinstallprompt", (event) => {
        event.preventDefault();
        deferredInstallPrompt = event;
        if (inlineInstallButton) {
            inlineInstallButton.hidden = false;
        } else if (installButton) {
            showPwaActions();
            installButton.hidden = false;
        }
    });

    const handleInstallClick = async () => {
        if (!deferredInstallPrompt) {
            setPwaStatus("Usa el menu del navegador para instalar la app.");
            return;
        }
        deferredInstallPrompt.prompt();
        await deferredInstallPrompt.userChoice.catch(() => {});
        deferredInstallPrompt = null;
        if (installButton) {
            installButton.hidden = true;
        }
        if (inlineInstallButton) {
            inlineInstallButton.hidden = true;
        }
    };

    installButton?.addEventListener("click", handleInstallClick);
    inlineInstallButton?.addEventListener("click", handleInstallClick);

    pushButton?.addEventListener("click", async () => {
        try {
            if (!pushConfig?.pushEnabled || !pushConfig?.vapidPublicKey) {
                setPwaStatus("Falta configurar la clave publica de notificaciones.", true);
                return;
            }
            const subscription = await ensurePushSubscriptionSaved({ requestPermission: true });
            if (!subscription) {
                setPwaStatus("Permiso de notificaciones no concedido.", true);
                return;
            }
            setPwaStatus("Notificaciones activadas.");
            await refreshPwaButtons();
            await updatePwaDiagnostics();
        } catch (error) {
            setPwaStatus(error.message || "No se pudieron activar las notificaciones.", true);
            await updatePwaDiagnostics();
        }
    });

    testPushButton?.addEventListener("click", async () => {
        try {
            const data = await postPwaJson("/pwa/push/test", { portal_token: portalToken });
            setPwaStatus(data.enviados ? "Notificacion de prueba enviada." : "No hay dispositivos activos para probar.", !data.enviados);
            if (data.enviados) {
                window.localStorage?.setItem(pwaStorageKey, "1");
                if (testPushButton) {
                    testPushButton.hidden = true;
                }
                if (pushButton) {
                    pushButton.hidden = true;
                }
                window.setTimeout(() => {
                    if (pwaActions) {
                        pwaActions.hidden = true;
                    }
                    setPwaStatus("");
                }, 1600);
            }
        } catch (error) {
            setPwaStatus(error.message || "No se pudo enviar la prueba.", true);
        }
        await updatePwaDiagnostics();
    });
})();
