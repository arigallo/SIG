# Revisión de permisos y visuales por rol

## Alcance y límites

Revisión local de autorización de Operación, centro de tareas, navegación compartida, dashboard, ficha de jugador, notificaciones y circuito de Madrinas. Se reprodujo el render de Operación con admin, tesorero, médico, entrenador, Secretaría, Madrinas mínimo, Madrinas con comunicaciones y solo lectura. No se consultaron ni modificaron cuentas de producción. No es una validación visual exhaustiva de las 91 plantillas ni una certificación de seguridad.

Los perfiles base del código son admin, tesorero, médico y entrenador. Los demás perfiles y las modificaciones de permisos pueden residir en la base de datos. La demo anterior de Madrinas usaba solo aspirantes_ver y aspirantes_gestionar; por ello no representaba una cuenta real que también tuviera comunicaciones_ver.

## Ajustes implementados

1. Se agregaron permisos explícitos `tareas_ver` y `tareas_gestionar`; comunicaciones ya no habilita asignación de tareas.
2. Las tareas se aíslan por área tanto al listar como al crear y actualizar. Un rol no puede operar por ID una tarea de otro módulo.
3. Las notificaciones financieras, médicas, deportivas, de Secretaría, portal y ahijadxs se filtran según permisos del usuario.
4. Menú, dashboard y Operación muestran únicamente accesos, indicadores y checklist relevantes para el rol.
5. Las acciones de escritura en la ficha de jugador exigen permisos de gestión; los perfiles de consulta ya no ven Nueva cuota, Registrar pago, Editar ficha o Nueva lesión.
6. Madrinas tiene acceso directo a Ahijadxs, formulario público y comunicaciones, pero no a Operación ni al centro de tareas por defecto.
7. Se corrigieron contraste de etiquetas administrativas, navegación activa y presentación móvil de la tabla de ahijadxs.
8. El formulario público usa “Vení a Jugar a Ruda Macho”, exige fecha de nacimiento y ya no incluye categoría de interés ni disponibilidad.

## Reproducción local de Operación

| Perfil probado | Accede a Operación | Gestiona tareas |
|---|---|---|
| Admin | Sí | Sí |
| Tesorero base | Sí | Sí |
| Médico base | Sí | Sí |
| Entrenador base | Sí | Sí |
| Secretaría con consulta y gestión | Sí | Sí |
| Madrinas con consulta y gestión de ahijadxs | No | No |
| Madrinas + comunicaciones_ver | No | No |
| Solo lectura financiera, médica o deportiva | Sí, en su área | No |

Cada variante ve sólo los módulos derivados de sus permisos. La prueba no accede a datos reales ni ejecuta altas o cambios.

## Verificación

- 82 pruebas automatizadas completas: OK.
- `git diff --check`: OK.
- Revisión local en navegador de Admin, Tesorería, Médico, Entrenador, Secretaría y Madrinas.
- Prueba de rechazo al actualizar por ID una tarea de otro módulo.
- Prueba de navegación y dashboard específico de Madrinas.

Queda pendiente contrastar los permisos personalizados almacenados en la base productiva antes de desplegar, ya que pueden diferir de los perfiles base del código. No se hizo commit ni push.
