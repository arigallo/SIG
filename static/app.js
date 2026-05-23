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
})();
