(() => {
    const form = document.querySelector(".entry-login-form");
    if (!form) return;

    const button = form.querySelector('button[type="submit"]');
    const error = form.querySelector("[data-login-error]");
    let sending = false;

    // Password managers can submit without firing input/change events.
    // Refresh on submit so restored pages use the current session token.
    form.addEventListener("submit", async (event) => {
        event.preventDefault();
        if (sending) return;
        sending = true;
        button.disabled = true;
        error.hidden = true;
        try {
            const response = await fetch(form.dataset.csrfUrl, {
                credentials: "same-origin",
                cache: "no-store",
                signal: AbortSignal.timeout(10000),
            });
            if (!response.ok) throw new Error("No se pudo renovar la sesión");
            const data = await response.json();
            if (typeof data.token !== "string" || !data.token) {
                throw new Error("Token inválido");
            }
            form.elements.namedItem("_csrf_token").value = data.token;
            HTMLFormElement.prototype.submit.call(form);
        } catch (exception) {
            error.textContent = "No se pudo conectar para ingresar. Volvé a intentarlo.";
            error.hidden = false;
            sending = false;
            button.disabled = false;
        }
    });
    window.addEventListener("pageshow", () => {
        sending = false;
        button.disabled = false;
    });
})();
