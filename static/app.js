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
            group.addEventListener("toggle", () => {
                if (group.open) {
                    closeOthers(group);
                }
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

    const isAcceptedFile = (file) => {
        if (!file) {
            return false;
        }

        return acceptedMimeTypes.has(file.type) || acceptedExtensions.has(fileExtension(file.name));
    };

    const filesLabel = (files) => {
        const selected = Array.from(files || []);
        if (!selected.length) {
            return "Ningún archivo seleccionado";
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
        if (!file || file.name) {
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

    const setDropzoneStatus = (dropzone, message, isError = false) => {
        const status = dropzone.querySelector("[data-upload-file-name]");
        if (status) {
            status.textContent = message;
            status.classList.toggle("upload-file-error", isError);
        }
    };

    const applyFilesToDropzone = (dropzone, input, files) => {
        const selected = Array.from(files || [])
            .map(normalizeClipboardFile)
            .filter(Boolean);
        const nextFiles = input.multiple ? selected : selected.slice(0, 1);

        if (!nextFiles.length || nextFiles.some((file) => !isAcceptedFile(file))) {
            setDropzoneStatus(dropzone, "Solo se aceptan PDF, JPG o PNG.", true);
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

    document.addEventListener("paste", (event) => {
        if (!dropzones.length || !event.clipboardData) {
            return;
        }

        const fileItem = Array.from(event.clipboardData.items || []).find((item) => item.kind === "file");
        if (!fileItem) {
            return;
        }

        const activeDropzone = document.activeElement && document.activeElement.closest("[data-upload-dropzone]");
        const targetDropzone = activeDropzone || (dropzones.length === 1 ? dropzones[0] : null);
        if (!targetDropzone || typeof targetDropzone.__setUploadFile !== "function") {
            return;
        }

        const file = fileItem.getAsFile();
        if (file && targetDropzone.__setUploadFile(file)) {
            event.preventDefault();
        }
    });
})();
