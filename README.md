# Sistema simple de gestion de jugadores

Base inicial hecha con **Flask + SQLite**.

## Incluye
- Alta de jugadores
- Listado de jugadores
- Busqueda
- Edicion
- Eliminacion

## Requisitos
- Python 3.10 o superior

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

## Proximos pasos recomendados
1. Agregar cuotas y deuda
2. Agregar ficha medica
3. Agregar lesiones
4. Agregar panel de alertas
5. Agregar backup/exportacion

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
