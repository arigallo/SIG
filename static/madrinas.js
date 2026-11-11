(() => {
    const rows = [];
    document.querySelectorAll('.quick-row').forEach(row => {
        const button = row.querySelector('[data-save]');
        if (!button) return;
        const fields = [...row.querySelectorAll('[data-field]')];
        const note = row.querySelector('[data-note]');
        const status = row.querySelector('[data-save-status]');
        const value = field => field.dataset.field === 'madrina_jugador_id'
            ? (field.value ? Number(field.value) : null) : field.value;
        let original = Object.fromEntries(fields.map(field => [field.dataset.field, value(field)]));
        let saving = false;
        let conflict = false;
        const paint = () => fields.forEach(field => { field.dataset.current = field.value; });
        paint();
        const dirty = () => fields.some(field => original[field.dataset.field] !== value(field)) || !!note.value.trim();
        rows.push(() => saving || dirty());
        const refresh = () => {
            paint();
            button.disabled = saving || conflict || !dirty();
            if (!saving && !conflict) status.textContent = dirty() ? 'Cambios sin guardar' : '';
        };
        fields.forEach(field => field.addEventListener('change', refresh));
        note.addEventListener('input', refresh);
        button.addEventListener('click', async () => {
            const cambios = {};
            const anteriores = {};
            fields.forEach(field => {
                const key = field.dataset.field;
                if (value(field) !== original[key]) {
                    cambios[key] = value(field);
                    anteriores[key] = original[key];
                }
            });
            saving = true;
            button.disabled = true;
            fields.forEach(field => { field.disabled = true; });
            note.disabled = true;
            status.textContent = 'Guardando…';
            try {
                const response = await fetch(row.dataset.saveUrl, {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json', 'X-CSRF-Token': document.querySelector('meta[name="csrf-token"]').content},
                    body: JSON.stringify({cambios, anteriores, nota: note.value.trim()})
                });
                const data = response.headers.get('content-type')?.includes('application/json')
                    ? await response.json() : {};
                if (!response.ok || !data.ok) {
                    conflict = response.status === 409;
                    throw new Error(data.error || 'No se confirmó el guardado. Revisá tu sesión antes de reintentar.');
                }
                const updateCount = (name, delta) => {
                    const counter = document.querySelector(`[data-counter="${name}"]`);
                    if (counter) counter.textContent = Math.max(0, Number(counter.textContent) + delta);
                };
                updateCount('sin_asignar', Number(data.valores.madrina_jugador_id === null) - Number(original.madrina_jugador_id === null));
                updateCount('pendientes_contacto', Number(data.valores.contacto_estado === 'pendiente') - Number(original.contacto_estado === 'pendiente'));
                original = data.valores;
                fields.forEach(field => { field.value = original[field.dataset.field] ?? ''; });
                paint();
                note.value = '';
                row.querySelector('[data-last-note]').textContent = data.detalle;
                row.querySelector('[data-last-author]').textContent = `${data.fecha_texto} · ${data.autor}`;
                status.textContent = 'Guardado ✓';
            } catch (error) {
                status.textContent = error.message || 'No se pudo guardar. Tus cambios siguen en la fila.';
            } finally {
                saving = false;
                fields.forEach(field => { field.disabled = conflict; });
                note.disabled = conflict;
                button.disabled = conflict || !dirty();
            }
        });
    });
    window.addEventListener('beforeunload', event => {
        if (rows.some(isDirty => isDirty())) {
            event.preventDefault();
            event.returnValue = '';
        }
    });
})();
