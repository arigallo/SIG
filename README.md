# SIG - Sistema integral de gestion

Aplicacion Flask desplegable en Cloud Run, con PostgreSQL/Cloud SQL, Drive,
Secret Manager, email y WhatsApp.

## Incluye
- Alta de jugadores
- Listado de jugadores
- Busqueda
- Edicion
- Eliminacion

## Versión preparada: 2.1.0

- La ficha del jugador enlaza su etapa previa como ahijadx cuando existe un DNI coincidente; el alta directa conserva el vínculo y evita duplicar jugadores con el mismo DNI.
- Cada evento del calendario tiene una vista central con confirmaciones, asistencia real, gastos vinculados y acciones disponibles según el rol.
- Finanzas muestra una cola de excepciones para comprobantes pendientes o rechazados y saldos vencidos.
- Comunicación muestra la cantidad de destinatarios, una vista previa y el último estado de WhatsApp por jugador.
- El panel prioriza accesos según permisos. El portal explica cuándo un comprobante sigue en revisión.
- Desde Sistema, un administrador puede simular otro rol para recorrer sus pantallas y permisos en modo de solo lectura, y volver a admin sin alterar usuarios.
- Madrinas dispone del manual del equipo además del instructivo de SIG. En Sistema se puede designar un usuario con email para recibir avisos de nuevas postulaciones públicas por correo y dentro de SIG; si tiene push habilitado, también recibe la notificación del navegador.

### Versiones publicadas

`Admin > Versiones` empieza a contar con la publicación de SIG 2.1.0 como entrega n.º 1. Después registra cada nueva entrega con fecha, nombre y revisión desplegada. El contador muestra 0 hasta registrar esa primera publicación; los commits no incrementan el número.

### Verificación

Con las dependencias instaladas, ejecutar `python -m unittest discover -s tests -q`. El proyecto requiere una `SECRET_KEY` de prueba y `INIT_DB=false` para ejecutar pruebas sin iniciar la base de producción.

## Requisitos
- Python 3.10 o superior

## Configuracion obligatoria

SIG no inicia sin una clave de sesion. En produccion, carga estos valores desde
Secret Manager en lugar de escribirlos directamente en Cloud Run:

```text
SECRET_KEY=<valor aleatorio largo>
ADMIN_PASSWORD=<solo para crear o recuperar el administrador inicial>
```

Si WhatsApp esta habilitado, tambien debe configurarse el secreto de la app de
Meta. Sin este valor el webhook responde `503` y no procesa eventos:

```text
WHATSAPP_APP_SECRET=<secreto de la app de Meta>
```

## Instalacion
Abri una terminal dentro de esta carpeta y ejecuta:

```bash
pip install -r requirements.txt
python app.py
```

Despues abri en tu navegador:

```text
http://127.0.0.1:5000
```

## Página institucional

La portada pública puede revisarse localmente en `/club`. Cuando el servicio
recibe una solicitud a `/` desde `rudamachorugby.com` o
`www.rudamachorugby.com`, muestra esa misma página sin requerir sesión.
El inicio de SIG en `sig.rudamachorugby.com` conserva su comportamiento.

Para publicar, configurar ambos dominios institucionales con DNS y HTTPS
hacia el servicio que ejecute esta versión. Este cambio de código no modifica
DNS ni despliega el servicio. Los botones usan la postulación y el portal
existentes; el acceso administrativo apunta al subdominio SIG.

## Automatizaciones del SIG

Desde `Admin > Sistema` se pueden habilitar:

- recordatorios de cuotas por email, con anticipacion configurable;
- sincronizacion automatica de facturas recibidas.

El endpoint para Cloud Scheduler es:

```text
POST /tasks/automatizaciones
X-Automation-Token: <AUTOMATION_TOKEN>
```

Debe configurarse `AUTOMATION_TOKEN` en Cloud Run y enviar el mismo valor desde
Cloud Scheduler. Cada recordatorio se registra de forma idempotente para evitar
duplicados durante el mismo dia.

## Formularios

Desde `Admin > Formularios` se pueden crear formularios de hasta 30 preguntas,
ordenarlas y marcar cuáles son obligatorias. Hay texto corto y largo, email,
número, fecha, opción única, selección múltiple, desplegable, escala y NPS.
Se conservan el acceso y los datos de las encuestas existentes.

Los borradores sin respuestas permiten editar preguntas y configuración.
Al publicar, el enlace público recibe respuestas durante las fechas configuradas;
cerrar el formulario detiene nuevos envíos. Los resultados incluyen distribuciones
y respuestas de texto, con exportación CSV que agrupa cada envío en una fila.
Los permisos existentes `encuestas_ver` y `encuestas_gestionar` siguen vigentes.
`init_db` amplía los tipos permitidos y registra `2026-10-09-formularios-v1`.

## Cuenta corriente

El perfil administrativo y el portal muestran una cuenta corriente unificada
con cuotas y gastos compartidos. Las deudas pendientes sobreviven al cierre del
gasto y pueden cobrarse posteriormente, generando un ingreso individual en caja.

Las actualizaciones de esquema se registran en `schema_migrations`; `init_db`
continua aplicando cambios compatibles durante el arranque.

## App instalable y notificaciones

SIG funciona como PWA: publica `manifest.webmanifest`, `service-worker.js` y
botones para instalar la app y activar notificaciones desde celulares.

Para Web Push deben configurarse claves VAPID en Cloud Run:

```text
PWA_VAPID_PUBLIC_KEY=...
PWA_VAPID_PRIVATE_KEY=...
PWA_VAPID_CLAIMS_SUB=mailto:admin@tudominio.com
```

Las suscripciones se guardan por usuario administrativo o por portal de jugador.
Desde la app se puede usar "Probar aviso" para validar que el celular recibe la
notificacion.

## Facturas recibidas por email

El modulo `Finanzas > Facturas recibidas` puede leer una casilla IMAP, aplicar filtros por remitente/asunto y guardar adjuntos PDF/JPG/PNG en Drive para luego registrarlos como egresos de caja.

La configuracion puede cargarse desde `Admin > Email facturas`. La contrasena se guarda como secreto en Google Secret Manager y SIG lee ese secreto al sincronizar.

Variables de entorno:

```text
FACTURA_EMAIL_IMAP_HOST=imap.example.com
FACTURA_EMAIL_IMAP_PORT=993
FACTURA_EMAIL_IMAP_USER=facturas@example.com
FACTURA_EMAIL_IMAP_PASSWORD=...
FACTURA_EMAIL_IMAP_FOLDER=INBOX
FACTURA_EMAIL_IMAP_USE_SSL=true
FACTURA_EMAIL_SEARCH_DAYS=45
FACTURA_EMAIL_MAX_MESSAGES=80
FACTURA_EMAIL_SECRET_NAME=sig-factura-email-imap-password
```

Los filtros se administran desde SIG. Por defecto se crean filtros iniciales para Meta/Facebook/Instagram y Canva, y se pueden agregar proveedores nuevos sin tocar codigo.
