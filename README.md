# SIG - Sistema integral de gestion

Aplicacion Flask desplegable en Cloud Run, con PostgreSQL/Cloud SQL, Drive,
Secret Manager, email y WhatsApp.

## Incluye
- Alta de jugadores
- Listado de jugadores
- Busqueda
- Edicion
- Eliminacion

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

## Automatizaciones

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
