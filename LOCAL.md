# SIG local funcional

Este entorno ejecuta `app.py`, sus rutas, permisos y formularios reales con una
base PostgreSQL propia. Los cambios persisten después de cerrar el navegador o
reiniciar el entorno. Las modificaciones del código se prueban aquí antes de
publicarlas en producción.

## Iniciar y entrar

En Windows, abrir `iniciar-local.cmd`. También se puede ejecutar:

```powershell
./scripts/local.ps1 start
```

Abrir <http://127.0.0.1:5080>. El iniciador muestra la contraseña inicial de los
usuarios `admin`, `tesorero`, `medico`, `entrenador`, `madrinas` y `secretaria`. Se guarda exclusivamente en
`.local/config.json`; es independiente de las contraseñas de producción. Si se
cambia desde SIG, se usa la nueva contraseña: el iniciador no la restablece.

El primer inicio descarga los binarios portátiles de PostgreSQL desde
[EDB](https://productsdl.enterprisedb.com/download-postgresql-binaries), crea la base
y carga tres jugadores, cuotas pendientes, una nota financiera, un ahijadx y un
evento ficticios. No instala un servicio global de Windows. Si faltan dependencias
Python, crea un entorno virtual dentro de `.local/venv`.

El usuario administrador puede usar todas las secciones, crear más usuarios y
roles, y abrir el portal de cada jugador desde su ficha. No hay login automático
ni permisos salteados.

## Guardado, reinicio y detención

- Los datos de PostgreSQL están en `.local/postgres-data` (puerto `55432`, solo loopback).
- Los archivos subidos, recibos y exportaciones están en `.local/files`.
- Los correos completos y mensajes simulados están en `.local/outbox`.
- Los registros del servidor están en `.local/web.log` y `.local/postgres.log`.

Para detener la aplicación y PostgreSQL, abrir `detener-local.cmd` o ejecutar
`./scripts/local.ps1 stop`. Para recargar cambios en Python, detener e iniciar.
Las migraciones de `init_db()` se aplican en cada inicio, conservando los datos.
No se activa el depurador de Flask.

Los cambios en plantillas, CSS y JavaScript se sirven al recargar la página.
La franja de SIG LOCAL muestra una revisión de interfaz para comparar lo que
está abierto en distintos navegadores. Los recursos usan una versión basada en
su contenido y las respuestas locales llevan `Cache-Control: no-store`.
El service worker local usa la red y limpia la caché de recursos de SIG de ese
origen, para evitar que una versión previa de los estilos o scripts siga visible.
Esto no borra sesiones, datos de PostgreSQL ni archivos subidos, y no cambia
el comportamiento de caché en producción.

```powershell
./scripts/local.ps1 status
./scripts/local.ps1 backup
```

`backup` genera un archivo PostgreSQL `.dump` en `.local/backups`. Para conservar
también los adjuntos, copiar `.local/files` junto con ese respaldo. No borrar
`.local` para actualizar el código: contiene todos los datos locales.

## Integraciones en las pruebas

Las operaciones de SIG guardan datos reales en la base local. El banner **SIG
LOCAL** identifica este entorno en todas las pantallas.

- **Archivos:** los comprobantes y documentos se cargan, descargan y eliminan en
  disco mediante un adaptador local que reemplaza Google Drive.
- **Correo:** el flujo real prepara el mensaje y los adjuntos; un transporte local
  los guarda como `.eml`. Se pueden revisar en **Ver envíos de prueba**.
- **WhatsApp y push:** se simula el envío. WhatsApp conserva sus registros en la
  base local y la bandeja muestra los mensajes generados. El estado enviado en este
  entorno significa captura local, no entrega a una persona.
- **OCR de Google, lectura de casillas IMAP y administración de Google Cloud:**
  requieren sus servicios externos y no se ejecutan aquí. Las fichas se pueden
  completar manualmente. La recepción real de webhooks, la entrega real de mensajes
  y estas integraciones necesitan una prueba separada en un entorno de integración.

El iniciador descarta configuraciones de conexión heredadas, fija la base
`sig_local` y bloquea conexiones salientes de Python excepto al puerto de su
PostgreSQL local. No lee ni copia la base de producción. `.local` queda excluido
de Git, Cloud Build y del contexto Docker.

## Verificar antes de publicar

```powershell
./scripts/local.ps1 check
```

Ejecuta verificaciones contra PostgreSQL y las rutas reales: pantallas principales,
login por rol, CSRF, alta de jugador, notas y seguimiento, resolución, privacidad del
portal, carga/descarga de archivos y captura de mensajes. Crea un jugador temporal
y lo elimina al terminar. Los registros de auditoría y envíos de prueba se conservan.
La comprobación usa las contraseñas iniciales de los usuarios de prueba.

Las pruebas unitarias siguen disponibles con:

```powershell
$env:INIT_DB = 'false'
$env:SECRET_KEY = 'test-only-secret-key'
./.codex-venv/Scripts/python.exe -m unittest discover -s tests -q
```

Flujo habitual: modificar código, reiniciar SIG local, ejecutar las verificaciones,
probar el flujo en el navegador con los roles involucrados y luego publicar usando
el procedimiento habitual de producción. Estos comandos nunca despliegan cambios.
