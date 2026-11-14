(() => {
    const form = document.getElementById('datos-plan');
    if (!form) return;
    const source = form.querySelector('[name="fuente"]');
    const updateSource = () => {
        form.querySelectorAll('[data-budget-source]').forEach(section => {
            section.hidden = !section.dataset.budgetSource.split(' ').includes(source.value);
        });
    };
    source.addEventListener('change', updateSource);
    updateSource();
    const markChanged = () => {
        document.querySelectorAll('[data-budget-save] button[type="submit"]').forEach(button => {
            button.disabled = true;
            button.textContent = 'Volvé a calcular para guardar';
        });
        document.querySelectorAll('[data-budget-dirty]').forEach(message => { message.hidden = false; });
    };
    form.addEventListener('input', markChanged);
    form.addEventListener('change', markChanged);
    // Also catches values restored by the browser after back/forward navigation.
    window.addEventListener('pageshow', () => {
        updateSource();
        if (Array.from(form.elements).some(element => {
            if (element.tagName === 'SELECT') {
                const original = Array.from(element.options).find(option => option.defaultSelected)
                    || element.options[0];
                return original && element.value !== original.value;
            }
            return 'defaultValue' in element && element.value !== element.defaultValue;
        })) markChanged();
    });
})();
