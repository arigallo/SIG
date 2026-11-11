# Organización del sistema interno

Se mantienen las áreas y los permisos existentes. El portal del jugador y las
páginas públicas conservan su diseño independiente.

| Área | Contenido | Identidad visual |
| --- | --- | --- |
| Inicio | Resumen, avisos y revisión diaria | Verde institucional |
| Club | Jugadores y socios, asistencia, tests, calendario y circulares | Verde |
| Madrinas | Ahijadxs y formulario de ingreso | Rosa |
| Finanzas | Cobranzas, cuotas, comprobantes, planes, gastos, caja, transferencias, facturas, presupuesto y ventas | Ámbar |
| Salud | Fichas médicas, lesiones y vencimientos | Coral |
| Comunicación | Mensajes, notificaciones de la app y WhatsApp | Azul |
| Secretaría | Documentación administrativa | Violeta |
| Administración | Usuarios, roles, sistema, auditoría y configuración | Gris azulado |

## Criterios compartidos

- Navegación lateral en escritorio y menú desplegable en pantallas pequeñas.
  Se conserva abierta la sección actual. La búsqueda filtra los enlaces que cada
  usuario tiene disponibles; no modifica los permisos de las rutas.
- El encabezado indica el área y la pantalla. La página presenta título, contexto
  y acciones; las acciones específicas van junto a su sección.
- Indicadores en franjas compactas con separadores, accesos breves con el color
  del área, fondos claros, estados legibles y
  botones principales institucionales. Los estados de error y peligro conservan
  su significado, independientemente del color del área.
- Filtros en un bloque común, tablas con encabezados uniformes y desplazamiento
  horizontal dentro de la tabla cuando sea necesario.
- La ficha del jugador ofrece accesos a datos, finanzas, salud, deporte y
  seguimiento. Las herramientas de renumeración de socios se agrupan aparte.
- La ayuda de los roles está en el menú de cuenta, sin repetir un bloque de ayuda
  encima de cada pantalla.

## Implementación

`templates/base.html` define la estructura, el área activa y la navegación según
permisos. `static/workspace.css` concentra el diseño administrativo compartido;
`static/workspace.js` agrega búsqueda y accesos por sección. `static/app.js`
gestiona la apertura y cierre del menú. El índice del perfil es una mejora
progresiva: los formularios y secciones siguen disponibles sin JavaScript.

Antes de publicar, ejecutar las pruebas unitarias y `scripts/local.ps1 check`.
Revisar también la navegación con teclado y las pantallas de escritorio y celular.

En la ficha financiera, el resumen es una franja sin tarjetas individuales. Las
cuotas aparecen primero; gastos, planes, historial y herramientas se consultan
en secciones desplegables para no desplazar el trabajo principal.

## Distribución por tarea

- Fichas de jugadores: pestañas Resumen, Datos, Finanzas, Salud, Deporte,
  Seguimiento y Gestión. Solo se crean las que contienen información permitida
  para el usuario. Los enlaces directos a notas abren Finanzas.
- Inicio: prioridades, tareas propias y accesos frecuentes según los permisos.
- Sistema: Estado, Configuración, Integraciones, Automatizaciones y Backups.
- Reportes: Ingresos y gastos, Deudas, Becas y Asistencia.
- Salud y vencimientos: una tabla temática por pestaña.
- Secretaría: archivo y filtros primero; carga por acción explícita, con el editor
  abierto si la validación devuelve datos.
- Ventas: catálogo e historial separados; registro y alta de producto desplegables.
- Presupuesto y tareas: consultas primero y creación desplegable.
- Caja: balance y movimientos del mes; totales históricos separados.
- Comunicación: destinatarios y vista previa como contenido principal; plantilla
  y edición agrupadas.

Las pestañas soportan flechas, Home y End, recuerdan la selección durante la sesión
y conservan los formularios originales. Sin JavaScript, el contenido permanece
visible en su orden de servidor. El diseño público del portal no cambia.

## Paquete de revisión de los recorridos diarios

- Inicio: tres bloques (requiere atención, tareas propias y accesos frecuentes).
  Los pendientes financieros se agrupan por persona y respetan los acuerdos
  futuros. Las consultas se limitan a los módulos permitidos para cada rol.
- Finanzas: recorridos de revisión de comprobantes, seguimiento de cobros y caja.
  Cada fila muestra situación, última gestión y siguiente acción. Un acuerdo
  futuro indica esperar; un comprobante recibido siempre requiere revisión.
- Notas: `_financial_notes.html` presenta la misma bitácora persistente desde
  el perfil y la cuenta financiera. La bitácora general excluye estas notas para
  evitar mostrarlas dos veces. Resolver conserva el registro y vuelve al origen.
- Acciones: las filas y encabezados con numerosas opciones conservan una acción
  frecuente visible y agrupan las secundarias en un desplegable nativo. Se
  mantienen los formularios, tokens CSRF y confirmaciones de las rutas originales.
- Vacíos: las tablas sin registros se reducen a un mensaje. Fichas faltantes,
  documentación vencida y seguimientos pendientes conservan alertas explícitas.

Los datos ficticios se crean únicamente desde `scripts/local_scenarios.py`,
invocado por el iniciador local. No se cargan desde `app.py` ni en producción.

## Paquete de continuidad y guardado

- Las notas financieras abiertas tienen una cola propia independiente de los
  saldos. También alimentan Inicio cuando corresponde retomarlas. Registrar
  un pago no cierra acuerdos; la pantalla invita a revisarlos.
- El buscador del encabezado consulta jugadores por nombre, DNI o socio, con
  accesos filtrados por permisos. Los resultados se construyen como texto y
  enlaces; las respuestas no se cachean ni incluyen DNI.
- Jugadores, cobranzas, excepciones y comprobantes tienen páginas de 25 filas.
  Los filtros se recuerdan en la sesión y la ficha ofrece volver al listado de
  origen. Cobranzas y Jugadores aplican LIMIT/OFFSET en PostgreSQL; comprobantes
  pagina después de combinar cuotas y gastos, sin el antiguo corte de 150.
- `services/operations.py` reserva en PostgreSQL el token de cada formulario
  protegido antes de ejecutar la ruta. La clave única por usuario impide ejecutar
  dos veces una solicitud. Guarda una huella de campos/archivos y el destino de
  los redirects. Las validaciones liberan la reserva; errores ambiguos la
  conservan para impedir reintentos automáticos. No sustituye los bloqueos de
  fila existentes en los pagos. Se requiere inicializar la tabla mediante el
  `init_db` habitual antes de habilitar estas rutas en otro entorno.
