/* Shared navigation and section discovery; every operation remains server-rendered. */
(() => {
    const playerSearch = document.querySelector('[data-player-search]');
    if (playerSearch) {
        const input = playerSearch.querySelector('input');
        const panel = playerSearch.querySelector('.global-player-results');
        const status = playerSearch.querySelector('[data-search-status]');
        const results = playerSearch.querySelector('[data-search-results]');
        let timer, controller;
        const toggle = playerSearch.querySelector('.global-search-toggle');
        const close = () => { panel.hidden = true; input.setAttribute('aria-expanded', 'false'); };
        const dismiss = () => { close(); playerSearch.classList.remove('is-open'); toggle.setAttribute('aria-expanded', 'false'); };
        const open = () => { panel.hidden = false; input.setAttribute('aria-expanded', 'true'); };
        toggle.addEventListener('click', () => {
            const expanded = playerSearch.classList.toggle('is-open');
            toggle.setAttribute('aria-expanded', String(expanded));
            if (expanded) input.focus(); else close();
        });
        input.addEventListener('input', () => {
            clearTimeout(timer);
            controller?.abort();
            results.replaceChildren();
            if (input.value.trim().length < 2) { close(); return; }
            open(); status.textContent = 'Buscando…';
            timer = setTimeout(async () => {
                const current = new AbortController();
                controller = current;
                try {
                    const url = new URL(playerSearch.dataset.searchUrl, location.origin);
                    url.searchParams.set('q', input.value.trim());
                    const response = await fetch(url, {signal: current.signal, headers: {'Accept': 'application/json'}});
                    if (!response.ok || response.redirected) throw new Error('search');
                    const data = await response.json();
                    if (controller !== current || current.signal.aborted) return;
                    results.replaceChildren();
                    data.resultados.forEach(item => {
                        const row = document.createElement('article');
                        const name = document.createElement('a');
                        name.href = item.perfil; name.textContent = `${item.apellido}, ${item.nombre}`;
                        name.className = 'global-player-name'; row.append(name);
                        const meta = document.createElement('span');
                        meta.className = 'muted-line';
                        meta.textContent = [item.numero_socio ? `Socio ${item.numero_socio}` : '', item.categoria].filter(Boolean).join(' · ');
                        row.append(meta);
                        const links = document.createElement('div'); links.className = 'actions';
                        [['Finanzas', item.finanzas], ['Notas internas', item.notas]].forEach(([label, href]) => {
                            if (!href) return;
                            const link = document.createElement('a'); link.href = href; link.textContent = label; links.append(link);
                        });
                        row.append(links); results.append(row);
                    });
                    status.textContent = data.resultados.length ? `${data.resultados.length} coincidencia(s). Elegí una ficha o presioná Enter para ver el listado.` : 'No se encontraron jugadores o socios.';
                } catch (error) {
                    if (error.name !== 'AbortError' && controller === current) status.textContent = 'No se pudo buscar. Presioná Enter para abrir el listado.';
                }
            }, 250);
        });
        playerSearch.addEventListener('keydown', event => {
            if (event.key === 'Escape') { clearTimeout(timer); controller?.abort(); dismiss(); (getComputedStyle(toggle).display === 'none' ? input : toggle).focus(); }
            if (event.key === 'ArrowDown' && document.activeElement === input && !panel.hidden) {
                const first = results.querySelector('a'); if (first) { event.preventDefault(); first.focus(); }
            }
        });
        document.addEventListener('click', event => { if (!playerSearch.contains(event.target)) dismiss(); });
        playerSearch.addEventListener('focusout', () => setTimeout(() => { if (!playerSearch.contains(document.activeElement)) dismiss(); }, 0));
    }
    // Let the browser serialize the original submitter (including its name/value)
    // before disabling buttons. A second submit is rejected immediately.
    document.addEventListener('submit', event => {
        const form = event.target;
        if (form.method.toLowerCase() !== 'post' || form.target === '_blank' || event.defaultPrevented) return;
        if (form.dataset.saving === 'true') { event.preventDefault(); return; }
        form.dataset.saving = 'true';
        setTimeout(() => {
            if (event.defaultPrevented) { delete form.dataset.saving; return; }
            form.setAttribute('aria-busy', 'true');
            form.querySelectorAll('button[type="submit"], input[type="submit"], button:not([type])').forEach(button => {
                if (button.disabled) return;
                button.dataset.savingDisabled = 'true'; button.disabled = true;
                if (button === event.submitter && button.tagName === 'BUTTON') {
                    button.dataset.savingLabel = button.innerHTML; button.textContent = 'Procesando…';
                }
            });
        }, 0);
    });
    window.addEventListener('pageshow', () => {
        document.querySelectorAll('form[data-saving]').forEach(form => { delete form.dataset.saving; form.removeAttribute('aria-busy'); });
        document.querySelectorAll('[data-saving-disabled]').forEach(button => {
            button.disabled = false; delete button.dataset.savingDisabled;
            if (button.dataset.savingLabel) { button.innerHTML = button.dataset.savingLabel; delete button.dataset.savingLabel; }
        });
    });
    const navigation = document.querySelector('#admin-main-nav');
    const search = document.querySelector('#nav-search');
    const normalize = value => value.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
    if (navigation && search) {
        const groups = [...navigation.querySelectorAll('.nav-group')];
        const links = [...navigation.querySelectorAll('a')];
        const initial = new Map(groups.map(group => [group, group.open]));
        links.forEach(link => {
            if (link.getAttribute('href') === location.pathname) {
                link.classList.add('active');
                link.setAttribute('aria-current', 'page');
                const group = link.closest('.nav-group');
                if (group) { group.classList.add('current'); group.open = true; }
            }
        });
        search.addEventListener('input', () => {
            const query = normalize(search.value);
            navigation.classList.toggle('nav-searching', !!query);
            let found = 0;
            links.forEach(link => {
                const group = link.closest('.nav-group');
                const label = group?.querySelector('summary')?.textContent || '';
                const matches = !query || normalize(link.textContent + ' ' + label).includes(query);
                link.hidden = !matches;
                if (matches) found++;
            });
            groups.forEach(group => {
                group.hidden = ![...group.querySelectorAll('a')].some(link => !link.hidden);
                group.open = query ? !group.hidden : (initial.get(group) || group.classList.contains('current'));
            });
            navigation.querySelector('.nav-search-empty').hidden = found > 0;
        });
    }
    // Editors are explicit actions, not permanent blocks above the work list.
    document.addEventListener('click', event => {
        const link = event.target.closest('a[href^="#"]');
        if (!link) return;
        const target = document.getElementById(link.hash.slice(1));
        if (target?.matches('details.module-editor')) {
            target.open = true;
            target.querySelector('summary')?.focus();
        }
    });

    // Keep one safe, frequent action visible; retain secondary forms unchanged.
    document.querySelectorAll('.admin-body td.actions, .section-header > .actions, .admin-hero-actions').forEach(cell => {
        const units = [...cell.children].filter(element => element.matches('a.button, form, details'));
        if (units.length < 3) return;
        const primary = units.find(element => element.matches('.button:not(.secondary):not(.danger)')) || units.find(element => !element.querySelector('.danger') && !element.matches('.danger')
            && /marcar pago|revisar|ver ficha|ver jugador|editar|recibo|^ver$/i.test(element.textContent.trim()));
        const secondary = units.filter(element => element !== primary);
        if (secondary.length < 2) return;
        const group = document.createElement('details');
        group.className = 'row-actions';
        const summary = document.createElement('summary');
        summary.textContent = cell.matches('td') ? 'Más acciones' : 'Más opciones';
        const body = document.createElement('div');
        body.className = 'row-actions-content';
        group.append(summary, body);
        if (primary) primary.after(group);
        else secondary[0].before(group);
        secondary.forEach(element => body.append(element));
    });

    const collectSections = (root, mode) => {
        const sections = [];
        const containers = [];
        const marked = [...root.querySelectorAll(':scope > [data-tab-section]')];
        if (marked.length) return {sections: marked, containers};
        [...root.children].forEach(child => {
            if (child.matches('.detail-grid') || (mode === 'system' && child.matches('.dashboard-cards'))) {
                const cards = [...child.children].filter(e => e.matches('.card') && e.querySelector('h3'));
                sections.push(...cards);
                containers.push(child);
            } else if (child.matches('.card, .page-section, .danger-zone') && !child.matches('.dashboard-cards, .admin-hero, .dashboard-role-shortcuts') && child.querySelector('h3')) {
                sections.push(child);
            }
        });
        return {sections, containers};
    };
    const classify = (section, mode) => {
        const label = section.querySelector('h3')?.textContent.trim() || 'Contenido';
        const text = normalize(label);
        if (mode === 'profile') {
            if (section.matches('.danger-zone')) return 'Gestión';
            if (text.startsWith('resumen')) return 'Resumen';
            if (/finanz|financ|cuenta corriente|cuotas|planes|gastos|beca/.test(text)) return 'Finanzas';
            if (/salud|medica|lesion|bienestar/.test(text)) return 'Salud';
            if (/asistencia|deportivo|rendimiento|tests/.test(text)) return 'Deporte';
            if (/bitacora|portal/.test(text)) return 'Seguimiento';
            return 'Datos';
        }
        if (mode === 'system') {
            if (/planteles|postulacion/.test(text)) return 'Configuración';
            if (/automatizacion|alertas error/.test(text)) return 'Automatizaciones';
            if (/backup/.test(text)) return 'Backups';
            if (/integracion|cloud run/.test(text)) return 'Integraciones';
            return 'Estado';
        }
        if (mode === 'dashboard') {
            if (/salud|fichas|lesion/.test(text)) return 'Salud';
            if (/finanzas|deuda|cuotas|comprobantes/.test(text)) return 'Finanzas';
            return 'Operación';
        }
        if (mode === 'reports') {
            if (/deuda|moroso/.test(text)) return 'Deudas';
            if (/beca/.test(text)) return 'Becas';
            if (/asistencia/.test(text)) return 'Asistencia';
            return 'Ingresos y gastos';
        }
        return label;
    };
    document.querySelectorAll('[data-profile-tabs], [data-module-tabs]').forEach((root, rootIndex) => {
        const mode = root.matches('[data-profile-tabs]') ? 'profile' : root.dataset.moduleTabs;
        const {sections, containers} = collectSections(root, mode);
        const groups = new Map();
        sections.forEach(section => {
            const label = classify(section, mode);
            if (!groups.has(label)) groups.set(label, []);
            groups.get(label).push(section);
        });
        if (groups.size < 2) return;
        const orders = mode === 'profile' ? ['Resumen', 'Datos', 'Finanzas', 'Salud', 'Deporte', 'Seguimiento', 'Gestión']
            : mode === 'system' ? ['Estado', 'Configuración', 'Integraciones', 'Automatizaciones', 'Backups']
            : mode === 'reports' ? ['Ingresos y gastos', 'Deudas', 'Becas', 'Asistencia']
            : mode === 'dashboard' ? ['Operación', 'Finanzas', 'Salud'] : [...groups.keys()];
        const labels = orders.filter(label => groups.has(label));
        const nav = document.createElement('div');
        nav.className = 'workspace-tabs';
        nav.setAttribute('role', 'tablist');
        nav.setAttribute('aria-label', mode === 'profile' ? 'Secciones del jugador' : 'Secciones del módulo');
        const content = document.createElement('div');
        content.className = 'workspace-tab-content';
        const tabs = [];
        const panels = [];
        const key = 'sig-view:' + location.pathname;
        labels.forEach((label, index) => {
            const slug = normalize(label).replace(/[^a-z0-9]+/g, '-');
            const id = `vista-${rootIndex}-${slug}`;
            const tab = document.createElement('button');
            tab.type = 'button';
            tab.textContent = label;
            tab.id = id + '-tab';
            tab.setAttribute('role', 'tab');
            tab.setAttribute('aria-controls', id);
            const panel = document.createElement('section');
            panel.id = id;
            panel.className = 'workspace-tab-panel';
            panel.tabIndex = 0;
            panel.setAttribute('role', 'tabpanel');
            panel.setAttribute('aria-labelledby', tab.id);
            groups.get(label).forEach(section => {
                section.classList.add('workspace-section');
                if (section.querySelector('table, .detail-grid') || section.id === 'notas-finanzas') section.classList.add('workspace-section-wide');
                panel.append(section);
            });
            nav.append(tab);
            content.append(panel);
            tabs.push(tab);
            panels.push(panel);
            tab.addEventListener('click', () => activate(index, true));
            tab.addEventListener('keydown', event => {
                let next;
                if (event.key === 'ArrowRight') next = (index + 1) % tabs.length;
                if (event.key === 'ArrowLeft') next = (index - 1 + tabs.length) % tabs.length;
                if (event.key === 'Home') next = 0;
                if (event.key === 'End') next = tabs.length - 1;
                if (next !== undefined) { event.preventDefault(); activate(next, true); tabs[next].focus(); }
            });
        });
        function activate(index, remember = false) {
            tabs.forEach((tab, current) => {
                tab.setAttribute('aria-selected', String(current === index));
                tab.tabIndex = current === index ? 0 : -1;
                panels[current].hidden = current !== index;
            });
            if (remember) {
                try { sessionStorage.setItem(key, labels[index]); } catch (_) { /* Storage can be unavailable. */ }
            }
        }
        // Keep heading, indicators and filters outside the tabbed content.
        const anchor = containers[0] || sections[0];
        root.insertBefore(nav, anchor?.parentElement === root ? anchor : null);
        nav.after(content);
        containers.forEach(container => { if (!container.children.length) container.remove(); });
        let selected = 0;
        try { selected = Math.max(0, labels.indexOf(sessionStorage.getItem(key))); } catch (_) { /* Optional persistence. */ }
        function selectHash() {
            const target = document.getElementById(decodeURIComponent(location.hash.slice(1)));
            const index = target ? panels.findIndex(panel => panel === target || panel.contains(target)) : -1;
            if (index >= 0) { activate(index); target.scrollIntoView({block: 'start'}); return true; }
            return false;
        }
        activate(selected);
        selectHash();
        window.addEventListener('hashchange', selectHash);
    });
})();
