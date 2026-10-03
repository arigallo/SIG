# Revisión de seguridad de SIG

Fecha: 3 de octubre de 2026, America/Buenos_Aires. Base Git: `6d14491` (`main`), incluyendo la corrección local del login todavía sin publicar.

## Resultado

Se identificaron **5 hallazgos de prioridad alta**, **9 de prioridad media** y **3 de prioridad baja**. El sistema necesita correcciones de autorización y gestión de sesiones antes de poder considerar resueltos sus riesgos principales. Pasar las pruebas funcionales no significa ausencia de vulnerabilidades.

La revisión agrega este informe y evidencia local. No modifica la lógica del sistema, no hace commits, push ni despliegues. Las modificaciones de login anteriores se conservan.

## Alcance y evidencia

- Inventario de 239 funciones de ruta en `app.py`; revisión de autenticación central, controles de rol/permiso y endpoints públicos. El servidor registra adicionalmente su ruta estática.
- 236 comprobaciones de rutas restringidas con un usuario sin permisos: no se alcanzó la base de datos en ninguna. Las dos rutas legacy respondieron 301 hacia páginas con controles propios. Esto comprueba rechazo general, no todas las combinaciones de roles y datos.
- 15 pruebas locales específicas con base de datos, email y Drive simulados. Sus aserciones documentan comportamientos vulnerables y controles vigentes: que pasen significa que la evidencia fue reproducida, no que el sistema sea seguro.
- 115 pruebas funcionales del proyecto: todas pasaron.
- Bandit 1.9.4: 23.753 líneas Python, 41 alertas, revisadas para separar problemas reales de falsos positivos.
- pip-audit 2.10.1: resolución de `requirements.txt`, 46 paquetes directos y transitivos, sin paquetes omitidos. Reportó 28 entradas correspondientes a 15 identificadores distintos en 3 paquetes; la consulta alternativa a OSV reportó 29 entradas, con los mismos 15 identificadores distintos. Son registros duplicados, no 28 vulnerabilidades distintas. Se agregó un aviso de Werkzeug verificado en la fuente oficial.
- Búsqueda de patrones de secretos en los 206 archivos versionados y 222 objetos de texto del historial disponible, correspondientes a 141 commits alcanzables: sin coincidencias en los patrones examinados. No constituye un escaneo exhaustivo de secretos ni certifica que nunca se hayan publicado credenciales.
- Revisión manual de SQL dinámico, plantillas, JavaScript, cargas/descargas, portal, recuperación de contraseñas, webhooks, automatizaciones, notificaciones, exportaciones y empaquetado.
- Las herramientas se instalaron temporalmente en `tmp/security-tools`, se ejecutaron sin modificar las dependencias de la aplicación y se retiraron después del análisis.

**Límite de alcance:** no se consultaron configuración efectiva de Cloud Run, IAM, Secret Manager, Cloud SQL, permisos de carpetas de Drive, imágenes publicadas, DNS, HTTPS ni logs de producción. El CLI `gcloud` no estaba disponible en este entorno. No se ejecutaron ataques contra producción. Los casos que dependen del proxy, red o servicios externos se indican expresamente. La resolución de dependencias se hizo desde Python 3.12 en Windows; el Dockerfile declara Python 3.13 en Linux, por lo que no equivale a auditar la imagen desplegada ni sus paquetes del sistema operativo.

## Hallazgos priorizados

| ID | Prioridad | Problema | Evidencia |
|---|---|---|---|
| SEC-01 | Alta | DNI suficiente para obtener el acceso completo al portal | Prueba local y código |
| SEC-02 | Alta | Sesiones sin revocación ni actualización de permisos | Pruebas locales y código |
| SEC-03 | Alta | Host no confiable en emails de recuperación | Prueba local; explotación en producción depende del ingreso HTTP |
| SEC-04 | Alta | Exportación de información médica con permiso de reportes | Archivo XLSX local verificado |
| SEC-05 | Alta | Destinos push arbitrarios y llamadas sin timeout definido por SIG | Almacenamiento reproducido y código de la versión exacta; salida de red no ejecutada |
| SEC-06 | Media | Lista vacía de permisos reactiva permisos predeterminados | Prueba local |
| SEC-07 | Media | Recuperación de claves: tokens coexistentes, consumo no atómico y falta de límites | Código y prueba parcial |
| SEC-08 | Media | Limitación de login evadible y atribución de IP no confiable | Código y prueba de cabecera |
| SEC-09 | Media | Baja de suscripciones push sin comprobar propietario | Prueba HTTP local |
| SEC-10 | Media | Fórmulas de Excel/CSV provenientes de datos de usuarios | Celda de fórmula reproducida |
| SEC-11 | Media | Tokens del portal persistidos en auditoría y alertas | Código y prueba parcial |
| SEC-12 | Media | Dependencias con avisos de seguridad vigentes | Escáner y avisos oficiales |
| SEC-13 | Media | Respuestas sensibles sin no-store y sin defensas HTTP adicionales | Descarga HTTP local |
| SEC-14 | Baja | Validación de adjuntos basada en extensión y nombre | Prueba local |
| SEC-15 | Baja | URLs y destinos internos con validación incompleta | Código; ejecución en navegador no probada |
| SEC-16 | Media | Docker incluye el contexto local completo sin .dockerignore | Dockerfile y archivos de exclusión |
| SEC-17 | Baja | CSRF no ASCII produce excepción; logout modifica sesión mediante GET | Prueba local y código |

### SEC-01 — DNI como única credencial del portal

`app.py:18810` (`portal_buscar`) busca jugadores por DNI y devuelve una redirección que contiene su `portal_token`. No exige un secreto adicional y tampoco usa el limitador público. El token permite acceder a datos personales, cuotas, información de bienestar y modificar datos de contacto; los controles de pertenencia de cada recurso se basan en ese token. La alta entropía de los tokens no compensa que el propio sistema los revele a cualquiera que conozca el DNI.

Se reprodujo un POST anónimo con DNI y CSRF propios, obteniendo la URL de un jugador simulado. El consentimiento para notificaciones no verifica identidad y puede omitirse mediante la función pública existente.

**Corrección:** verificar identidad mediante enlace u OTP enviado a un canal registrado, o una credencial personal. Usar el DNI como identificador de búsqueda; limitar intentos y dar respuestas uniformes. Definir expiración, rotación y revocación del acceso al portal. Esta decisión cambia el flujo de acceso y requiere acordar el mecanismo funcional antes de implementarlo.

### SEC-02 — Revocación y vigencia de sesiones

`app.py:9917`, `app.py:10202`, `app.py:11332`, `app.py:11541`, `app.py:23746`, `app.py:23792`. El login guarda usuario, rol y permisos en una cookie firmada; las rutas confían en esos datos sin verificar de nuevo el usuario ni una versión de sesión en el servidor. Eliminar un usuario, reducir su rol o cambiar su contraseña no invalida cookies emitidas anteriormente. Cerrar sesión limpia la cookie del navegador, pero una copia anterior sigue siendo aceptada.

Se verificó que una cookie que identifica a un administrador inexistente en el escenario simulado accede a `/usuarios` sin una consulta de validación de identidad. También se guardó una cookie, se ejecutó logout, se repuso esa cookie y se accedió nuevamente. La firma no fue vulnerada: se reutiliza una sesión previamente emitida. El límite de 30 minutos no garantiza revocación inmediata y Flask refresca por defecto las sesiones permanentes durante su uso.

**Corrección:** sesiones revocables en el servidor o identificador y versión de sesión comprobados contra la base de datos. Revocar al cerrar sesión, cambiar/resetear contraseña y eliminar/deshabilitar usuarios. Consultar permisos vigentes o invalidar sesiones cuando se modifican roles. Mantener el comportamiento autorizado de simulación de roles.

### SEC-03 — Enlace de recuperación controlado por Host

`app.py:11453`: el enlace se genera con `url_for(..., _external=True)` a partir del host de la solicitud. No hay validación de hosts confiables ni URL pública canónica para este email.

Una solicitud local con `Host: attacker.example`, cookie y token CSRF emitidos para ese mismo host produjo un email simulado con el enlace de recuperación apuntando a ese dominio. Si el ingreso de producción permite este Host, alguien puede pedir recuperación de una cuenta ajena y recibir el token cuando el usuario siga el enlace enviado por SIG. No se comprobó esa condición del proxy real.

**Corrección:** URL pública canónica HTTPS para enlaces sensibles y lista explícita de hosts aceptados tanto en el ingreso como en Flask. No confiar indiscriminadamente en cabeceras forwarded.

### SEC-04 — Reportes permite exportar salud

`app.py:20825`, `app.py:20915`, `app.py:21061`: `/exportar/datos` exige únicamente `reportes_ver`, pero consulta fichas médicas y genera una hoja con aptitud, contacto de emergencia y observaciones. A diferencia de la auditoría, cuya inclusión sí es condicional, no exige `salud_ver` para esos datos.

Un usuario simulado con solo `reportes_ver` descargó un XLSX cuyo campo K2 de la hoja «Fichas médicas» contenía una observación médica privada simulada. El chequeo general de acceso a la ruta pasa, pero la autorización de los datos incluidos es insuficiente.

**Corrección:** incluir cada sección únicamente cuando el usuario tenga el permiso de su dominio, o establecer un permiso explícito de exportación integral reservado a los roles apropiados. Comprobar también otras exportaciones con permisos amplios.

### SEC-05 — Suscripciones push y salida de red

`repositories/notificaciones.py:10`, `app.py:1359`, `app.py:21787`, `app.py:21821`. Al guardar una suscripción solo se exige un endpoint no vacío: no se verifica HTTPS, proveedor, dirección IP, claves ni pertenencia del endpoint en una reasignación. La prueba local aceptó `http://127.0.0.1:8080/private`.

La revisión del wheel oficial de `pywebpush==2.0.3` confirma que ese endpoint se usa como destino del POST. SIG tampoco pasa un `timeout`; el argumento por defecto de `webpush` es `None`. La explotación depende de que push esté configurado, de aportar claves válidas, del comportamiento de VAPID y de la salida de red del despliegue. No se hicieron solicitudes a direcciones internas. Son riesgos de SSRF, ocupación de workers y reasignación de dispositivos ajenos, no una prueba de acceso al metadata server.

**Corrección:** validar estructura y claves de la suscripción; aceptar HTTPS y proveedores explícitos; impedir destinos locales/privados, redirecciones y cambios de propietario arbitrarios. Establecer límites y timeout en los envíos, manejar errores y restringir salida de red cuando sea posible. Al validar DNS, considerar cambios de resolución y redirecciones, no solo el hostname inicial.

### SEC-06 — Permisos vacíos se interpretan como permisos faltantes

`app.py:10205`: `session.get("permisos") or permisos_default_rol(...)` sustituye una lista vacía explícita por el preset del rol. Quitar todos los permisos a un rol conocido no los elimina efectivamente.

Con rol `entrenador` y `permisos=[]`, `tiene_permiso("jugadores_gestionar")` devuelve True.

**Corrección:** distinguir `None`/clave inexistente de una lista vacía válida; obtener la autorización vigente del servidor.

### SEC-07 — Ciclo de vida y abuso de recuperación de claves

`app.py:10842`, `app.py:10860`, `app.py:11433`, `app.py:11483`. La emisión no revoca enlaces previos; el cambio mediante un enlace marca únicamente ese registro como usado. Los enlaces hermanos vigentes pueden seguir sirviendo para recuperar la cuenta. La comprobación y la actualización del registro son operaciones separadas y la actualización no condiciona `usado = 0`, por lo que no garantiza consumo único ante concurrencia.

Además, solicitud y validación no consumen un límite público. Buscar un token prueba hasta 100 hashes de contraseña costosos; peticiones inválidas pueden multiplicar CPU y emails. Se verificó localmente que el endpoint de solicitud no llama al limitador; las carreras concurrentes y el consumo de CPU no se explotaron.

**Corrección:** emitir tokens indexables mediante hash criptográfico apropiado, con expiración y consumo atómico; revocar todos los tokens del usuario tras cambios de contraseña. Limitar solicitudes y validaciones por IP y cuenta y mantener mensajes uniformes. Coordinarlo con la revocación de sesiones de SEC-02.

### SEC-08 — Límites de login y confianza en la IP

`app.py:7094`, `app.py:10880`. Se toma el primer valor de `X-Forwarded-For` sin comprobar la cadena de proxies. El bloqueo cuenta la combinación exacta usuario/IP, permitiendo distribuir intentos entre direcciones. Si el ingreso conserva un prefijo aportado por el cliente, cambiar esa cabecera evita el límite y falsea la auditoría.

La prueba local confirma que la aplicación usa la IP elegida por el solicitante. El comportamiento del proxy en producción queda pendiente.

**Corrección:** obtener la IP desde proxies verificados con una configuración acorde al despliegue; combinar límites por IP, cuenta y globales y hacerlos robustos ante concurrencia. Considerar segundo factor para administradores. Evitar que el bloqueo por cuenta se convierta en una forma sencilla de impedir el acceso a usuarios legítimos.

### SEC-09 — Unsubscribe no comprueba actor ni propiedad

`app.py:21810`, `repositories/notificaciones.py:40`. La ruta es pública y llama a la baja por endpoint sin autenticar actor ni filtrar por propietario. Tener un CSRF propio no acredita pertenencia de la suscripción.

Se confirmó una respuesta 200 desde un cliente anónimo con un CSRF propio y una dirección de dispositivo ajeno simulada. Las direcciones reales son difíciles de adivinar, por lo que el impacto requiere conocerlas; esa dificultad no reemplaza la autorización.

**Corrección:** resolver el actor como en subscribe y modificar únicamente sus suscripciones. Evitar reasignaciones de propietario a partir del mismo endpoint sin un mecanismo adecuado de comprobación.

### SEC-10 — Inyección de fórmulas en exportaciones

`app.py:20593`, `app.py:20599`, `app.py:18118`, `app.py:22240`, `app.py:23287`. Los datos de usuarios se escriben directamente en hojas y CSV. Se reprodujo que una cadena que empieza con `=HYPERLINK(...)` queda en Excel como una fórmula (`data_type = "f"`). Un jugador o postulante que pueda editar texto puede hacer que una persona que abra el reporte ejecute fórmulas; el efecto concreto depende del programa y sus restricciones.

**Corrección:** escribir texto externo como texto en XLSX y neutralizar los prefijos de fórmula en CSV, incluidos espacios/caracteres de control iniciales. Mantener campos numéricos legítimos como números. Verificar que el archivo exportado conserve texto, no fórmulas.

### SEC-11 — Tokens persistidos en auditoría y alertas

`app.py:7110`, `app.py:10176`, `app.py:10362`, `app.py:10643`. Se redactan contraseñas, pero no tokens. La auditoría genérica guarda `request.path` y los argumentos de ruta; en el portal ambos contienen la credencial bearer. Las alertas 500 también incluyen ruta y query, pudiendo copiar enlaces de recuperación o portal a logs y emails.

La prueba local confirma que `portal_token` no se redacta. La persistencia en ruta/argumentos se verificó por lectura del código; no se consultaron registros reales.

**Corrección:** redactar tokens y otras credenciales en todos los caminos de logging, auditoría y alertas; registrar identificadores de entidad en vez del bearer. Revisar retención y permisos de los registros existentes antes de decidir rotación de tokens.

### SEC-12 — Dependencias

Se encontraron **15 avisos distintos mediante escáner** y **1 adicional mediante verificación oficial**. La aplicabilidad depende del uso y de la plataforma, y no se confirmó explotación remota de estas dependencias en SIG.

| Dependencia fijada | Avisos | Versión corregida de referencia | Aplicabilidad observada |
|---|---|---|---|
| Flask 3.0.3 | CVE-2026-27205 / GHSA-68rp-wp8r-4726 | 3.1.3 | Puede omitir Vary: Cookie en ciertos accesos; requiere condiciones concretas de caché/proxy. El login local tiene no-store, otras respuestas requieren revisión. |
| Click 8.3.2 | CVE-2026-7246 / GHSA-47fr-3ffg-hgmw | 8.3.3 | Problema en click.edit(); no se identificó ese uso en el código de SIG. |
| Pillow 12.2.0 | 13 avisos únicos de lectores, codecs y otras API | 12.3.0 | SIG usa ImageReader con logo/firma locales; no se encontró procesamiento directo de imágenes cargadas por usuarios. No todos los avisos son alcanzables. |
| Werkzeug 3.1.8 | CVE-2026-102598 / GHSA-g6x2-hccm-hh4m | 3.1.9 | Denegación de servicio con dispositivos Windows/NTFS; no aplica a esa ruta de ataque en el contenedor Linux declarado. El aviso reciente no apareció en pip-audit. |

Fuentes: [Flask](https://github.com/pallets/flask/security/advisories/GHSA-68rp-wp8r-4726), [Click](https://github.com/tsigouris007/security-advisories/security/advisories/GHSA-47fr-3ffg-hgmw), [Pillow](https://github.com/python-pillow/Pillow/security/advisories), [Pillow EPS](https://github.com/python-pillow/Pillow/security/advisories/GHSA-pg7v-jwj7-p798), [Werkzeug](https://github.com/pallets/werkzeug/security/advisories/GHSA-g6x2-hccm-hh4m).

**Corrección:** actualizar esas versiones y resolver/probar el conjunto completo, especialmente el cambio de Flask. Fijar también las dependencias transitivas y auditar la imagen Linux y su sistema operativo; repetir la consulta de avisos al publicar.

### SEC-13 — Caché y cabeceras de respuestas sensibles

`app.py:130`, `app.py:10149`, `app.py:19869`, `app.py:19903`, `app.py:19598`. El login y el manifest personalizado incluyen no-store, pero no hay una política equivalente general para portal, reportes y documentos privados. El calendario personalizado declara caché pública de 900 segundos; las descargas heredan una configuración de archivos de hasta 6 horas.

La descarga local simulada de un comprobante privado respondió sin no-store, X-Content-Type-Options, X-Frame-Options, Content-Security-Policy ni Referrer-Policy. La falta de cabeceras por sí sola no demuestra una explotación. El riesgo de filtración depende de los navegadores y del proxy/CDN; una política de caché sensible explícita elimina esa ambigüedad.

**Corrección:** no-store/private en datos personales, recuperación y descargas privadas; nosniff; política de framing; política de referrer que proteja enlaces con tokens. Definir y probar CSP compatible con los scripts actuales; no imponer una CSP que rompa la aplicación. Verificar HTTPS/HSTS y cabeceras en el ingreso real.

### SEC-14 — Archivos: contenido no validado

`app.py:1984`, `app.py:2022`, `app.py:2783`. Se comprueban nombre, extensión, presencia y tamaño, pero no que el contenido corresponda al formato declarado. Un archivo `fake.pdf` con texto arbitrario fue aceptado como `application/pdf`.

No se comprobó ejecución de archivos ni un XSS mediante un adjunto. Los nombres se normalizan con secure_filename y no se encontró una ruta de descarga arbitraria controlada directamente por el solicitante. El riesgo principal es admitir contenido no esperado o dañino y confiar luego en él durante visualización/procesamiento.

**Corrección:** verificar contenido con validadores apropiados, tratar formatos complejos como no confiables, considerar análisis antimalware y restringir visualización inline. Para importaciones ZIP/XLSX, limitar tamaño descomprimido, filas y trabajo de procesamiento, además del tamaño HTTP.

### SEC-15 — URLs y redirecciones

`app.py:292`, `app.py:12791`, `templates/jugador_detalle.html:285`, `templates/portal_jugador.html:941`. La URL de un documento se almacena sin validar esquema y se usa en enlaces. El escape HTML no verifica que sea una URL apropiada. El filtro de destinos internos rechaza `//`, pero no normaliza barras invertidas ni otras formas que los navegadores pueden interpretar de manera diferente.

No se ejecutaron scripts ni se validó una redirección externa en navegador, por lo que se registra como validación incompleta, no como XSS explotado.

**Corrección:** permitir esquemas explícitos para enlaces externos, validar URLs tras normalización y usar destinos de navegación definidos por el servidor para las redirecciones internas.

### SEC-16 — Empaquetado del despliegue

`Dockerfile:9` usa `COPY . .` y no existe `.dockerignore`. `.gitignore` no protege un build Docker directo; `.gcloudignore` protege solo los flujos que efectivamente lo usan y no excluye todos los directorios locales, por ejemplo `.codex-venv`, `.git`, copias legacy y herramientas bajo `tmp`.

Una construcción desde este checkout puede incluir archivos locales, entornos, respaldos o credenciales ignorados por Git, según su contexto. No se construyó ni inspeccionó una imagen real y no se afirma que una imagen publicada contenga secretos.

**Corrección:** .dockerignore explícito y contexto de build mínimo; alinear exclusiones de Cloud Build y excluir datos, secretos, Git, entornos y salidas de auditoría. Ejecutar con usuario sin privilegios y revisar la imagen publicada y el IAM efectivo. Las credenciales encontradas en imágenes antiguas, si las hubiera, requieren rotación.

### SEC-17 — Errores evitables en CSRF y logout

`app.py:282`: compare_digest de cadenas no ASCII puede lanzar TypeError en vez de rechazar la solicitud. Se reprodujo con un CSRF `ñ`. Esto puede generar 500 y, según la configuración, alertas por email.

`app.py:11540`: logout es GET, por lo que una navegación externa puede provocar cierre de sesión sin comprobación CSRF. El replay de una cookie anterior es SEC-02 y tiene prioridad mayor.

**Corrección:** validar tipo/formato del CSRF y rechazar de manera uniforme; hacer logout mediante POST protegido. Mantener una respuesta amigable en el login para tokens vencidos.

## Controles que se verificaron y alertas descartadas

- Contraseñas con generate_password_hash/check_password_hash; no hay contraseña inicial fija: se requiere ADMIN_PASSWORD cuando debe crearse el administrador.
- SECRET_KEY obligatoria; cookies HttpOnly y Secure por defecto, SameSite=Lax; renovación de sesión/token tras login. Su configuración real en producción no se leyó.
- Validación CSRF central para POST. Exenciones limitadas a integraciones: WhatsApp verifica HMAC, la automatización exige un secreto no vacío con compare_digest y la solicitud de Meta tiene validación de firma.
- El cambio local del login conserva CSRF, no agrega una exención de autenticación ni expone contraseñas. `/login/csrf` devuelve el token de la propia sesión y no rota innecesariamente uno vigente. No se identificó una omisión de CSRF introducida por este cambio. No se probó con un gestor real de contraseñas.
- Descargas de cuotas/gastos del portal enlazan recurso y jugador mediante token y portal activo en SQL. Se revisaron sus consultas; no se detectó que bastara cambiar el ID para acceder a un recurso de otro portal.
- Las alertas SQL de Bandit revisadas construyen filtros/placeholders a partir de constantes y datos parametrizados; los ordenamientos identificados usan listas permitidas. No se confirmó inyección SQL. Esto no es una demostración formal de ausencia de SQLi.
- SHA1 en `participante_gasto_key` identifica invitados y no autentica usuarios ni protege credenciales. La alerta HIGH de Bandit en ese lugar no se clasifica como una vulnerabilidad criptográfica de autenticación.
- Las alertas de contraseña fija de Bandit corresponden a comparaciones de configuración, definiciones de esquema y None; no se encontró una credencial fija de producción en esas líneas.
- urlopen examinado para metadata/Meta/URBA usa fuentes definidas por el código. Los endpoints push recibidos del cliente son el caso diferente señalado en SEC-05.
- Jinja conserva escape automático; el constructor de encuestas usa innerHTML con una estructura fija y asigna el texto variable mediante value. No se confirmó XSS por ese constructor.
- Límite global de cuerpo HTTP y límites propios de adjuntos; secure_filename para nombres. No bastan para limitar descompresión ni procesamiento de formatos complejos.
- Modo de simulación de roles impide escrituras salvo la salida de simulación. El debug local depende de configuración; Docker usa Gunicorn.

## Orden de corrección propuesto

1. Corregir acceso al portal, revocación de sesiones, enlaces de recuperación, exportación médica y validación/destino de push (SEC-01 a SEC-05).
2. Corregir permisos vacíos, ciclo de tokens de recuperación, límites de acceso y pertenencia de suscripciones (SEC-06 a SEC-09).
3. Neutralizar fórmulas y credenciales en logs, actualizar dependencias y reforzar caché/cabeceras y empaquetado (SEC-10 a SEC-13 y SEC-16).
4. Completar validación de archivos/URLs y tratamiento de CSRF/logout (SEC-14, SEC-15 y SEC-17).
5. Repetir pruebas de seguridad con los cambios, probar flujos reales en un entorno de prueba y revisar configuración/imagen efectivas de producción antes de cerrar la auditoría operativa.

## Archivos de evidencia

- `probes.py` y `probes-results.txt`: pruebas reproducibles que documentan los comportamientos analizados con datos simulados.
- `routes.json` y `permission-probes.json`: inventario y controles generales de autorización.
- `bandit.json`: alertas originales; no deben interpretarse todas como vulnerabilidades confirmadas.
- `dependencies-direct.json`, `dependencies-osv.json`, `dependencies-resolved.json`: consultas originales de avisos.
- `secret-scan.json`: resultados de patrones sin valores de credenciales.
- `pywebpush-source-review.txt`: referencias de la versión exacta revisada.
- `installed-packages.json`: inventario del entorno local de pruebas, separado del conjunto resuelto para producción.

Para reproducir las pruebas locales desde la raíz: `.\.codex-venv\Scripts\python.exe output/security/probes.py`.
