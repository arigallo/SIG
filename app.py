from flask import Flask, render_template, request, redirect, url_for, flash, Response, send_file, session, has_request_context, abort, g
import psycopg
from psycopg.rows import dict_row
from pathlib import Path
import csv
import io
import json
import os
import posixpath
import re
import mimetypes
import smtplib
import secrets
import unicodedata
from datetime import datetime, timedelta
from email.message import EmailMessage
from urllib.parse import quote

from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

try:
    from google.auth import default as google_auth_default
    from googleapiclient.discovery import build as google_build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
except ImportError:
    google_auth_default = None
    google_build = None
    HttpError = None
    MediaIoBaseDownload = None
    MediaIoBaseUpload = None

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "cambiar-esto-por-una-clave-segura")

app.config["SESSION_PERMANENT"] = False
app.config["SESSION_COOKIE_NAME"] = "sig_session"
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SECURE"] = os.environ.get(
    "SESSION_COOKIE_SECURE",
    "true",
).lower() in {"1", "true", "yes", "on"}
app.config["SESSION_COOKIE_SAMESITE"] = os.environ.get("SESSION_COOKIE_SAMESITE", "Lax")
MAX_REQUEST_BYTES = max(
    int(os.environ.get("MAX_UPLOAD_BYTES", str(16 * 1024 * 1024))),
    int(os.environ.get("MAX_BATCH_UPLOAD_BYTES", str(64 * 1024 * 1024))),
)
app.config["MAX_CONTENT_LENGTH"] = MAX_REQUEST_BYTES

app.permanent_session_lifetime = timedelta(minutes=30)

def formato_moneda(valor):
    try:
        return "${:,.0f}".format(float(valor)).replace(",", ".")
    except:
        return "$0"

app.jinja_env.filters["moneda"] = formato_moneda

@app.context_processor
def inject_now():
    return {
        "now": datetime.now,
        "csrf_token": csrf_token,
        "puede": tiene_permiso,
        "mantenimiento": getattr(g, "mantenimiento", None) if has_request_context() else None,
        "notificaciones_count": obtener_contador_notificaciones() if has_request_context() else 0,
    }

BASE_DIR = Path(__file__).resolve().parent


def csrf_token():
    token = session.get("_csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["_csrf_token"] = token
    return token


def csrf_valido():
    token_session = session.get("_csrf_token")
    token_enviado = request.form.get("_csrf_token") or request.headers.get("X-CSRF-Token")
    return bool(
        token_session
        and token_enviado
        and secrets.compare_digest(token_session, token_enviado)
    )


def destino_interno(destino, fallback="index"):
    destino = (destino or "").strip()
    if destino.startswith("/") and not destino.startswith("//"):
        return destino
    return url_for(fallback)


def permisos_default_rol(rol):
    return list(ROLE_PRESETS.get(rol, []))


def normalizar_permisos(permisos):
    return sorted({permiso for permiso in permisos if permiso in PERMISOS})


def serializar_permisos(permisos):
    return json.dumps(normalizar_permisos(permisos), ensure_ascii=False)


def deserializar_permisos(valor, rol=None):
    if not valor:
        return permisos_default_rol(rol)
    try:
        permisos = json.loads(valor)
    except (TypeError, ValueError):
        return permisos_default_rol(rol)
    if not isinstance(permisos, list):
        return permisos_default_rol(rol)
    return normalizar_permisos(permisos)


def grupos_permisos():
    grupos = {}
    for clave, permiso in PERMISOS.items():
        grupos.setdefault(permiso["grupo"], []).append({
            "clave": clave,
            **permiso,
        })
    return grupos


def cargar_permisos_rol(conn, rol):
    fila = conn.execute("""
        SELECT permisos
        FROM roles
        WHERE nombre = %s
    """, (rol,)).fetchone()
    if not fila:
        return permisos_default_rol(rol)
    return deserializar_permisos(fila["permisos"], rol)

DATABASE_URL = os.environ.get("DATABASE_URL")
DB_NAME = os.environ.get("DB_NAME") or os.environ.get("POSTGRES_DB", "sig")
DB_USER = os.environ.get("DB_USER") or os.environ.get("POSTGRES_USER", "sig_user")
DB_PASS = os.environ.get("DB_PASS") or os.environ.get("DB_PASSWORD", "")
DB_HOST = os.environ.get("DB_HOST")
DB_PORT = os.environ.get("DB_PORT", "5432")
DB_SSLMODE = os.environ.get("DB_SSLMODE")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD")
FORCE_ADMIN_PASSWORD_UPDATE = os.environ.get(
    "FORCE_ADMIN_PASSWORD_UPDATE", ""
).lower() in {"1", "true", "yes", "on"}
CLOUD_SQL_CONNECTION_NAME = (
    os.environ.get("INSTANCE_CONNECTION_NAME")
    or os.environ.get("CLOUD_SQL_CONNECTION_NAME")
)
DB_SOCKET_DIR = os.environ.get("DB_SOCKET_DIR", "/cloudsql")
DB_CONNECT_TIMEOUT = int(os.environ.get("DB_CONNECT_TIMEOUT", "10"))
DRIVE_COMPROBANTES_FOLDER_ID = (
    os.environ.get("GOOGLE_DRIVE_COMPROBANTES_FOLDER_ID")
    or os.environ.get("DRIVE_COMPROBANTES_FOLDER_ID")
)
DRIVE_COMPROBANTES_SUBFOLDER = (
    os.environ.get("GOOGLE_DRIVE_COMPROBANTES_SUBFOLDER")
    or os.environ.get("DRIVE_COMPROBANTES_SUBFOLDER")
)
DRIVE_FICHAS_MEDICAS_FOLDER_ID = (
    os.environ.get("GOOGLE_DRIVE_FICHAS_MEDICAS_FOLDER_ID")
    or os.environ.get("DRIVE_FICHAS_MEDICAS_FOLDER_ID")
    or DRIVE_COMPROBANTES_FOLDER_ID
)
DRIVE_FICHAS_MEDICAS_SUBFOLDER = (
    os.environ.get("GOOGLE_DRIVE_FICHAS_MEDICAS_SUBFOLDER")
    or os.environ.get("DRIVE_FICHAS_MEDICAS_SUBFOLDER")
    or "Fichas medicas"
)
COMPROBANTE_MAX_BYTES = int(os.environ.get("COMPROBANTE_MAX_BYTES", str(10 * 1024 * 1024)))
COMPROBANTE_EXTENSIONS = {
    ".pdf": "application/pdf",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
}
FICHA_MEDICA_MAX_BYTES = int(os.environ.get("FICHA_MEDICA_MAX_BYTES", str(16 * 1024 * 1024)))
FICHA_MEDICA_EXTENSIONS = {
    ".pdf": "application/pdf",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
}
FICHA_MEDICA_OCR_LANGUAGE = os.environ.get("FICHA_MEDICA_OCR_LANGUAGE", "es")
MAX_LOGIN_ATTEMPTS = int(os.environ.get("MAX_LOGIN_ATTEMPTS", "5"))
LOGIN_ATTEMPT_WINDOW_MINUTES = int(os.environ.get("LOGIN_ATTEMPT_WINDOW_MINUTES", "15"))
PASSWORD_RESET_TOKEN_MINUTES = int(os.environ.get("PASSWORD_RESET_TOKEN_MINUTES", "45"))
SMTP_HOST = os.environ.get("SMTP_HOST", "").strip()
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "").strip()
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
SMTP_FROM = os.environ.get("SMTP_FROM", SMTP_USER).strip()
SMTP_USE_TLS = os.environ.get("SMTP_USE_TLS", "true").lower() in {"1", "true", "yes", "on"}
ASPIRANTE_ENTRENAMIENTOS_OBJETIVO = int(os.environ.get("ASPIRANTE_ENTRENAMIENTOS_OBJETIVO", "8"))
ASPIRANTE_ESTADOS = {"Aspirante", "Ingresado", "Baja"}
APP_VERSION = os.environ.get("APP_VERSION", "local")
CLOUD_SQL_BACKUP_WINDOW = os.environ.get("CLOUD_SQL_BACKUP_WINDOW", "12:00 a.m. - 4:00 a.m.")
CLOUD_SQL_BACKUP_RETENTION_DAYS = os.environ.get("CLOUD_SQL_BACKUP_RETENTION_DAYS", "7")
CLOUD_SQL_PITR_DAYS = os.environ.get("CLOUD_SQL_PITR_DAYS", "7")
COMPROBANTE_ESTADOS = {"sin_comprobante", "pendiente", "aceptado", "rechazado"}
BITACORA_TIPOS = {
    "general": "General",
    "finanzas": "Finanzas",
    "salud": "Salud",
    "deportivo": "Deportivo",
}
MAINTENANCE_DEFAULT_MESSAGE = os.environ.get(
    "MAINTENANCE_DEFAULT_MESSAGE",
    "El sistema esta en mantenimiento. Volve a intentar en unos minutos.",
)
MESES_ES = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
]
PERMISOS = {
    "jugadores_ver": {
        "grupo": "Jugadores",
        "nombre": "Ver jugadores",
        "descripcion": "Listado y detalle general de jugadores.",
    },
    "jugadores_gestionar": {
        "grupo": "Jugadores",
        "nombre": "Crear y editar jugadores",
        "descripcion": "Alta, edición, importación y acciones masivas.",
    },
    "jugadores_eliminar": {
        "grupo": "Jugadores",
        "nombre": "Eliminar jugadores",
        "descripcion": "Baja definitiva de jugadores.",
    },
    "aspirantes_ver": {
        "grupo": "Jugadores",
        "nombre": "Ver ahijadxs",
        "descripcion": "Listado, detalle y seguimiento de ahijadxs.",
    },
    "aspirantes_gestionar": {
        "grupo": "Jugadores",
        "nombre": "Gestionar ahijadxs",
        "descripcion": "Alta, edicion, asistencia y conversion a jugador activo.",
    },
    "cuotas_ver": {
        "grupo": "Cuotas y caja",
        "nombre": "Ver cuotas",
        "descripcion": "Consulta de cuotas, recibos y comprobantes.",
    },
    "cuotas_gestionar": {
        "grupo": "Cuotas y caja",
        "nombre": "Gestionar cuotas",
        "descripcion": "Crear, generar, cobrar, eliminar y adjuntar comprobantes.",
    },
    "planes_pago_ver": {
        "grupo": "Cuotas y caja",
        "nombre": "Ver planes de pago",
        "descripcion": "Consulta de acuerdos y planes de regularizacion.",
    },
    "planes_pago_gestionar": {
        "grupo": "Cuotas y caja",
        "nombre": "Gestionar planes de pago",
        "descripcion": "Alta, seguimiento y cierre de planes de pago.",
    },
    "caja_ver": {
        "grupo": "Cuotas y caja",
        "nombre": "Ver caja",
        "descripcion": "Consulta y exportación de movimientos de caja.",
    },
    "caja_gestionar": {
        "grupo": "Cuotas y caja",
        "nombre": "Gestionar caja",
        "descripcion": "Crear, editar, anular movimientos y cerrar meses.",
    },
    "reportes_ver": {
        "grupo": "Reportes",
        "nombre": "Ver reportes",
        "descripcion": "Reportes y exportaciones generales.",
    },
    "comunicaciones_ver": {
        "grupo": "Reportes",
        "nombre": "Ver comunicaciones",
        "descripcion": "Morosos, plantillas y exportaciones de comunicación.",
    },
    "salud_ver": {
        "grupo": "Salud",
        "nombre": "Ver salud",
        "descripcion": "Fichas médicas, lesiones y alertas médicas.",
    },
    "salud_gestionar": {
        "grupo": "Salud",
        "nombre": "Gestionar salud",
        "descripcion": "Editar fichas médicas, crear, editar y eliminar lesiones.",
    },
    "documentos_ver": {
        "grupo": "Salud",
        "nombre": "Ver documentos",
        "descripcion": "Consulta de documentos, vencimientos y faltantes.",
    },
    "documentos_gestionar": {
        "grupo": "Salud",
        "nombre": "Gestionar documentos",
        "descripcion": "Carga, edicion y baja de documentos manuales.",
    },
    "calendario_ver": {
        "grupo": "Deportivo",
        "nombre": "Ver calendario",
        "descripcion": "Consulta del calendario.",
    },
    "calendario_gestionar": {
        "grupo": "Deportivo",
        "nombre": "Gestionar calendario",
        "descripcion": "Crear eventos de calendario.",
    },
    "asistencia_ver": {
        "grupo": "Deportivo",
        "nombre": "Ver asistencia",
        "descripcion": "Consulta de eventos e historial de asistencia.",
    },
    "asistencia_gestionar": {
        "grupo": "Deportivo",
        "nombre": "Tomar asistencia",
        "descripcion": "Crear eventos y registrar asistencia.",
    },
    "tests_ver": {
        "grupo": "Deportivo",
        "nombre": "Ver tests deportivos",
        "descripcion": "Consulta de mediciones, rankings y graficos deportivos.",
    },
    "tests_gestionar": {
        "grupo": "Deportivo",
        "nombre": "Gestionar tests deportivos",
        "descripcion": "Crear tests, cargar puntajes e importar mediciones.",
    },
    "alertas_finanzas": {
        "grupo": "Alertas",
        "nombre": "Alertas financieras",
        "descripcion": "Cuotas vencidas, caja y movimientos inusuales.",
    },
    "alertas_salud": {
        "grupo": "Alertas",
        "nombre": "Alertas médicas",
        "descripcion": "Fichas vencidas y lesiones activas.",
    },
    "auditoria_ver": {
        "grupo": "Administración",
        "nombre": "Ver auditoría",
        "descripcion": "Consulta de actividad del sistema.",
    },
    "backup_ver": {
        "grupo": "Administración",
        "nombre": "Ver backup",
        "descripcion": "Acceso a información de backup.",
    },
    "portal_jugador_gestionar": {
        "grupo": "Administracion",
        "nombre": "Gestionar portal externo",
        "descripcion": "Generar y desactivar enlaces externos para jugadores y familias.",
    },
    "seguridad_ver": {
        "grupo": "Administracion",
        "nombre": "Ver seguridad",
        "descripcion": "Consulta de logins, bloqueos y actividad sensible.",
    },
}
TODOS_LOS_PERMISOS = list(PERMISOS.keys())
ROLE_PRESETS = {
    "admin": TODOS_LOS_PERMISOS,
    "tesorero": [
        "jugadores_ver",
        "cuotas_ver",
        "cuotas_gestionar",
        "planes_pago_ver",
        "planes_pago_gestionar",
        "caja_ver",
        "caja_gestionar",
        "reportes_ver",
        "comunicaciones_ver",
        "calendario_ver",
        "calendario_gestionar",
        "asistencia_ver",
        "asistencia_gestionar",
        "alertas_finanzas",
    ],
    "medico": [
        "jugadores_ver",
        "salud_ver",
        "salud_gestionar",
        "documentos_ver",
        "documentos_gestionar",
        "alertas_salud",
    ],
    "entrenador": [
        "jugadores_ver",
        "jugadores_gestionar",
        "aspirantes_ver",
        "aspirantes_gestionar",
        "calendario_ver",
        "calendario_gestionar",
        "asistencia_ver",
        "asistencia_gestionar",
        "tests_ver",
        "tests_gestionar",
    ],
}


def get_db_connect_args():
    kwargs = {
        "row_factory": dict_row,
        "connect_timeout": DB_CONNECT_TIMEOUT,
    }

    if DB_SSLMODE:
        kwargs["sslmode"] = DB_SSLMODE

    if DATABASE_URL:
        return (DATABASE_URL,), kwargs

    host = DB_HOST
    if not host and CLOUD_SQL_CONNECTION_NAME:
        host = posixpath.join(DB_SOCKET_DIR, CLOUD_SQL_CONNECTION_NAME)

    kwargs.update({
        "dbname": DB_NAME,
        "user": DB_USER,
        "password": DB_PASS,
        "host": host or "localhost",
        "port": DB_PORT,
    })

    return (), kwargs


class DBConnection:
    def __init__(self):
        args, kwargs = get_db_connect_args()
        self.conn = psycopg.connect(*args, **kwargs)

    def execute(self, query, params=None):
        cur = self.conn.cursor()
        cur.execute(query, params or ())
        return cur

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()

    def close(self):
        if self.conn and not self.conn.closed:
            self.conn.close()


def get_connection():
    return DBConnection()


def get_columns(conn, table_name):
    columnas = conn.execute("""
        SELECT column_name
        FROM information_schema.columns
        WHERE table_schema = 'public'
          AND table_name = %s
    """, (table_name,)).fetchall()
    return [col["column_name"] for col in columnas]


def parse_bool_setting(value):
    return str(value or "").strip().lower() in {"1", "true", "yes", "on"}


def obtener_config_mantenimiento(conn=None):
    own_conn = conn is None
    conn = conn or get_connection()
    rows = conn.execute("""
        SELECT clave, valor, actualizado_en, actualizado_por
        FROM app_settings
        WHERE clave IN ('maintenance_mode', 'maintenance_message')
    """).fetchall()
    if own_conn:
        conn.close()

    settings = {row["clave"]: row for row in rows}
    mode_row = settings.get("maintenance_mode") or {}
    message_row = settings.get("maintenance_message") or {}
    return {
        "activo": parse_bool_setting(mode_row.get("valor")),
        "mensaje": message_row.get("valor") or MAINTENANCE_DEFAULT_MESSAGE,
        "actualizado_en": mode_row.get("actualizado_en") or message_row.get("actualizado_en"),
        "actualizado_por": mode_row.get("actualizado_por") or message_row.get("actualizado_por"),
    }


def guardar_app_setting(conn, clave, valor, usuario=None):
    conn.execute("""
        INSERT INTO app_settings (clave, valor, actualizado_en, actualizado_por)
        VALUES (%s, %s, CURRENT_TIMESTAMP, %s)
        ON CONFLICT(clave) DO UPDATE SET
            valor = EXCLUDED.valor,
            actualizado_en = EXCLUDED.actualizado_en,
            actualizado_por = EXCLUDED.actualizado_por
    """, (clave, valor, usuario))


def drive_service():
    if (
        google_auth_default is None
        or google_build is None
        or MediaIoBaseDownload is None
        or MediaIoBaseUpload is None
    ):
        raise RuntimeError(
            "Faltan dependencias de Google Drive. Instalá google-api-python-client y google-auth."
        )

    credentials, _ = google_auth_default(
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return google_build("drive", "v3", credentials=credentials, cache_discovery=False)


def require_drive_comprobantes_folder():
    if not DRIVE_COMPROBANTES_FOLDER_ID:
        raise RuntimeError("Falta configurar GOOGLE_DRIVE_COMPROBANTES_FOLDER_ID.")
    return DRIVE_COMPROBANTES_FOLDER_ID


def require_drive_fichas_medicas_folder():
    if not DRIVE_FICHAS_MEDICAS_FOLDER_ID:
        raise RuntimeError(
            "Falta configurar GOOGLE_DRIVE_FICHAS_MEDICAS_FOLDER_ID o GOOGLE_DRIVE_COMPROBANTES_FOLDER_ID."
        )
    return DRIVE_FICHAS_MEDICAS_FOLDER_ID


def drive_query_escape(value):
    return str(value or "").replace("\\", "\\\\").replace("'", "\\'")


def get_or_create_drive_subfolder(service, parent_id, name):
    safe_name = drive_query_escape(name)
    safe_parent = drive_query_escape(parent_id)
    query = (
        "mimeType = 'application/vnd.google-apps.folder' "
        f"and name = '{safe_name}' "
        f"and '{safe_parent}' in parents "
        "and trashed = false"
    )
    result = service.files().list(
        q=query,
        fields="files(id, name)",
        pageSize=1,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = result.get("files", [])
    if files:
        return files[0]["id"]

    folder = service.files().create(
        body={
            "name": name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        },
        fields="id",
        supportsAllDrives=True,
    ).execute()
    return folder["id"]


def get_drive_periodo_folder(service, root_folder, periodo):
    try:
        fecha_periodo = datetime.strptime(periodo, "%Y-%m")
    except (TypeError, ValueError):
        fecha_periodo = datetime.now()

    year_folder = get_or_create_drive_subfolder(
        service,
        root_folder,
        str(fecha_periodo.year),
    )
    return get_or_create_drive_subfolder(
        service,
        year_folder,
        MESES_ES[fecha_periodo.month - 1],
    )


def get_drive_comprobantes_base_folder(service):
    root_folder = require_drive_comprobantes_folder()
    subfolder = (DRIVE_COMPROBANTES_SUBFOLDER or "").strip()
    if not subfolder:
        return root_folder
    return get_or_create_drive_subfolder(service, root_folder, subfolder)


def get_drive_fichas_medicas_base_folder(service):
    root_folder = require_drive_fichas_medicas_folder()
    subfolder = (DRIVE_FICHAS_MEDICAS_SUBFOLDER or "").strip()
    if not subfolder:
        return root_folder
    return get_or_create_drive_subfolder(service, root_folder, subfolder)


def get_drive_jugador_folder(service, root_folder, jugador):
    jugador_slug = secure_filename(
        f"{jugador['apellido']}_{jugador['nombre']}_{jugador['id']}"
    ) or f"jugador_{jugador['id']}"
    return get_or_create_drive_subfolder(service, root_folder, jugador_slug)


def subir_ficha_medica_batch_pendiente(validado, batch_id):
    if not validado:
        return None

    service = drive_service()
    root_folder = get_drive_fichas_medicas_base_folder(service)
    pendientes_folder = get_or_create_drive_subfolder(service, root_folder, "_batch_pendientes")
    batch_folder = get_or_create_drive_subfolder(service, pendientes_folder, batch_id)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = validado["filename"] or f"ficha_medica{validado['ext']}"
    drive_name = f"pendiente_{timestamp}_{filename}"
    media = MediaIoBaseUpload(
        io.BytesIO(validado["content"]),
        mimetype=validado["mime_type"],
        resumable=False,
    )

    uploaded = service.files().create(
        body={"name": drive_name, "parents": [batch_folder]},
        media_body=media,
        fields="id, name, mimeType, size, webViewLink",
        supportsAllDrives=True,
    ).execute()

    return {
        "file_id": uploaded["id"],
        "nombre": uploaded.get("name") or filename,
        "mime_type": uploaded.get("mimeType") or validado["mime_type"],
        "tamano": int(uploaded.get("size") or len(validado["content"])),
        "web_url": uploaded.get("webViewLink"),
        "folder_id": batch_folder,
    }


def mover_ficha_medica_batch_a_jugador(file_id, jugador, ext, source_folder_id=None):
    service = drive_service()
    root_folder = get_drive_fichas_medicas_base_folder(service)
    target_folder = get_drive_jugador_folder(service, root_folder, jugador)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    jugador_slug = secure_filename(f"{jugador['apellido']}_{jugador['nombre']}") or f"jugador_{jugador['id']}"
    drive_name = f"ficha_medica_{jugador['id']}_{jugador_slug}_{timestamp}{ext or ''}"

    update_kwargs = {
        "fileId": file_id,
        "body": {"name": drive_name},
        "addParents": target_folder,
        "fields": "id, name, mimeType, size, webViewLink",
        "supportsAllDrives": True,
    }
    if source_folder_id:
        update_kwargs["removeParents"] = source_folder_id

    updated = service.files().update(**update_kwargs).execute()
    return {
        "file_id": updated["id"],
        "nombre": updated.get("name") or drive_name,
        "mime_type": updated.get("mimeType"),
        "tamano": int(updated.get("size") or 0),
        "web_url": updated.get("webViewLink"),
        "folder_id": target_folder,
    }


def validar_comprobante_upload(file_storage):
    if not file_storage or not file_storage.filename:
        return None

    filename = secure_filename(file_storage.filename)
    ext = Path(filename).suffix.lower()
    if ext not in COMPROBANTE_EXTENSIONS:
        raise ValueError("El comprobante debe ser PDF, JPG o PNG.")

    content = file_storage.read()
    if not content:
        raise ValueError("El comprobante está vacío.")
    if len(content) > COMPROBANTE_MAX_BYTES:
        max_mb = max(1, COMPROBANTE_MAX_BYTES // (1024 * 1024))
        raise ValueError(f"El comprobante supera el tamaño máximo permitido ({max_mb} MB).")

    mime_type = COMPROBANTE_EXTENSIONS[ext]
    guessed_mime = mimetypes.guess_type(filename)[0]
    if guessed_mime in COMPROBANTE_EXTENSIONS.values():
        mime_type = guessed_mime

    return filename, ext, content, mime_type


def subir_comprobante_a_drive(file_storage, cuota):
    validado = validar_comprobante_upload(file_storage)
    if not validado:
        return None

    filename, ext, content, mime_type = validado
    service = drive_service()
    root_folder = get_drive_comprobantes_base_folder(service)
    periodo = cuota["periodo"] or datetime.now().strftime("%Y-%m")
    folder_id = get_drive_periodo_folder(service, root_folder, periodo)
    jugador_slug = secure_filename(f"{cuota['apellido']}_{cuota['nombre']}") or f"jugador_{cuota['jugador_id']}"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    drive_name = f"cuota_{cuota['id']}_{periodo}_{jugador_slug}_{timestamp}{ext}"
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype=mime_type, resumable=False)
    body = {"name": drive_name, "parents": [folder_id]}

    existing_file_id = cuota.get("comprobante_drive_file_id")
    if existing_file_id:
        uploaded = service.files().update(
            fileId=existing_file_id,
            body={"name": drive_name},
            media_body=media,
            fields="id, name, mimeType, size, webViewLink",
            supportsAllDrives=True,
        ).execute()
    else:
        uploaded = service.files().create(
            body=body,
            media_body=media,
            fields="id, name, mimeType, size, webViewLink",
            supportsAllDrives=True,
        ).execute()

    return {
        "file_id": uploaded["id"],
        "nombre": uploaded.get("name") or filename,
        "mime_type": uploaded.get("mimeType") or mime_type,
        "tamano": int(uploaded.get("size") or len(content)),
        "web_url": uploaded.get("webViewLink"),
    }


def validar_ficha_medica_upload(file_storage):
    if not file_storage or not file_storage.filename:
        return None

    filename = secure_filename(file_storage.filename)
    ext = Path(filename).suffix.lower()
    if ext not in FICHA_MEDICA_EXTENSIONS:
        raise ValueError("La ficha medica debe ser PDF, JPG o PNG.")

    content = file_storage.read()
    if not content:
        raise ValueError("La ficha medica esta vacia.")
    if len(content) > FICHA_MEDICA_MAX_BYTES:
        max_mb = max(1, FICHA_MEDICA_MAX_BYTES // (1024 * 1024))
        raise ValueError(f"La ficha medica supera el tamano maximo permitido ({max_mb} MB).")

    mime_type = FICHA_MEDICA_EXTENSIONS[ext]
    guessed_mime = mimetypes.guess_type(filename)[0]
    if guessed_mime in FICHA_MEDICA_EXTENSIONS.values():
        mime_type = guessed_mime

    return {
        "filename": filename,
        "ext": ext,
        "content": content,
        "mime_type": mime_type,
    }


def subir_ficha_medica_a_drive(validado, jugador, ficha=None):
    if not validado:
        return None

    service = drive_service()
    root_folder = get_drive_fichas_medicas_base_folder(service)
    folder_id = get_drive_jugador_folder(service, root_folder, jugador)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    jugador_slug = secure_filename(f"{jugador['apellido']}_{jugador['nombre']}") or f"jugador_{jugador['id']}"
    drive_name = f"ficha_medica_{jugador['id']}_{jugador_slug}_{timestamp}{validado['ext']}"
    media = MediaIoBaseUpload(
        io.BytesIO(validado["content"]),
        mimetype=validado["mime_type"],
        resumable=False,
    )

    existing_file_id = ficha.get("documento_drive_file_id") if ficha else None
    if existing_file_id:
        uploaded = service.files().update(
            fileId=existing_file_id,
            body={"name": drive_name},
            media_body=media,
            fields="id, name, mimeType, size, webViewLink",
            supportsAllDrives=True,
        ).execute()
    else:
        uploaded = service.files().create(
            body={"name": drive_name, "parents": [folder_id]},
            media_body=media,
            fields="id, name, mimeType, size, webViewLink",
            supportsAllDrives=True,
        ).execute()

    return {
        "file_id": uploaded["id"],
        "nombre": uploaded.get("name") or validado["filename"],
        "mime_type": uploaded.get("mimeType") or validado["mime_type"],
        "tamano": int(uploaded.get("size") or len(validado["content"])),
        "web_url": uploaded.get("webViewLink"),
        "folder_id": folder_id,
    }


def extraer_texto_ocr_drive(validado, jugador, folder_id=None):
    if not validado:
        return ""

    service = drive_service()
    if folder_id is None:
        root_folder = get_drive_fichas_medicas_base_folder(service)
        folder_id = get_drive_jugador_folder(service, root_folder, jugador)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    doc_name = f"ocr_tmp_ficha_medica_{jugador['id']}_{timestamp}"
    media = MediaIoBaseUpload(
        io.BytesIO(validado["content"]),
        mimetype=validado["mime_type"],
        resumable=False,
    )

    doc_id = None
    try:
        doc = service.files().create(
            body={
                "name": doc_name,
                "mimeType": "application/vnd.google-apps.document",
                "parents": [folder_id],
            },
            media_body=media,
            fields="id",
            supportsAllDrives=True,
            ocrLanguage=FICHA_MEDICA_OCR_LANGUAGE,
        ).execute()
        doc_id = doc["id"]
        exported = service.files().export(
            fileId=doc_id,
            mimeType="text/plain",
        ).execute()
        if isinstance(exported, bytes):
            return exported.decode("utf-8", errors="replace").strip()
        return str(exported or "").strip()
    finally:
        if doc_id:
            try:
                service.files().delete(
                    fileId=doc_id,
                    supportsAllDrives=True,
                ).execute()
            except Exception as cleanup_error:
                status = getattr(getattr(cleanup_error, "resp", None), "status", None)
                if status != 404:
                    app.logger.warning("No se pudo eliminar el documento temporal OCR %s.", doc_id)


def normalizar_ocr_texto(texto):
    texto = (texto or "").replace("\r\n", "\n").replace("\r", "\n")
    texto = re.sub(r"[ \t]+", " ", texto)
    return texto.strip()


def normalizar_fecha_ocr(valor):
    valor = (valor or "").strip()
    for formato in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(valor, formato).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def buscar_fecha_cercana(texto, palabras_clave):
    patron_fecha = r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2})"
    lineas = [linea.strip() for linea in texto.splitlines() if linea.strip()]
    for linea in lineas:
        linea_min = linea.lower()
        if any(palabra in linea_min for palabra in palabras_clave):
            match = re.search(patron_fecha, linea)
            if match:
                fecha = normalizar_fecha_ocr(match.group(1))
                if fecha:
                    return fecha
    return ""


def extraer_anio_bduar(texto):
    texto_min = texto.lower()
    patrones = (
        r"evaluaci[oó]n\s+pre\s*competitiva\s+(20\d{2})",
        r"evaluaci[oó]n\s+precompetitiva\s+(20\d{2})",
        r"pre\s*competitiva\s+(20\d{2})",
        r"\bbduar\b.*?(20\d{2})",
    )
    for patron in patrones:
        match = re.search(patron, texto_min, re.DOTALL)
        if match:
            return match.group(1)
    return ""


def extraer_telefono_ocr(texto):
    lineas = [linea.strip() for linea in texto.splitlines() if linea.strip()]
    for linea in lineas:
        linea_min = linea.lower()
        if any(palabra in linea_min for palabra in ("telefono", "tel", "celular", "emergencia")):
            match = re.search(r"(\+?\d[\d\s().-]{6,}\d)", linea)
            if match:
                return re.sub(r"\s+", " ", match.group(1)).strip()
    return ""


def extraer_contacto_emergencia_ocr(texto):
    lineas = [linea.strip() for linea in texto.splitlines() if linea.strip()]
    for index, linea in enumerate(lineas):
        linea_min = linea.lower()
        if "contacto" in linea_min and "emergencia" in linea_min:
            partes = re.split(r":|-", linea, maxsplit=1)
            if len(partes) > 1:
                contacto = re.sub(r"\+?\d[\d\s().-]{6,}\d", "", partes[1]).strip(" -:")
                if contacto:
                    return contacto[:120]
            if index + 1 < len(lineas):
                return lineas[index + 1][:120]
    return ""


def datos_ficha_desde_ocr(texto):
    texto = normalizar_ocr_texto(texto)
    texto_min = texto.lower()
    apto_detectado = None
    if re.search(r"\bno\s+apto\b", texto_min):
        apto_detectado = 0
    elif re.search(r"\bapto\b", texto_min):
        apto_detectado = 1

    fecha_vencimiento = buscar_fecha_cercana(
        texto,
        ("vencimiento", "vence", "vigencia", "validez", "valido hasta", "valida hasta"),
    )
    if not fecha_vencimiento:
        anio_bduar = extraer_anio_bduar(texto)
        if anio_bduar:
            fecha_vencimiento = f"{anio_bduar}-12-31"

    return {
        "fecha_vencimiento": fecha_vencimiento,
        "apto_fisico": apto_detectado,
        "contacto_emergencia": extraer_contacto_emergencia_ocr(texto),
        "telefono_emergencia": extraer_telefono_ocr(texto),
    }


def normalizar_texto_match(valor):
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(char for char in texto if not unicodedata.combining(char))
    texto = texto.casefold()
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def normalizar_documento(valor):
    digitos = re.sub(r"\D+", "", str(valor or ""))
    return digitos.lstrip("0")


def extraer_dni_ocr(texto):
    texto = normalizar_ocr_texto(texto)
    lineas = [linea.strip() for linea in texto.splitlines() if linea.strip()]
    patron = r"(\d{1,2}\.?\d{3}\.?\d{3}|\d{7,9})"
    for linea in lineas:
        linea_min = linea.lower()
        if any(palabra in linea_min for palabra in ("dni", "documento", "d.n.i", "du ")):
            match = re.search(patron, linea)
            if match:
                return normalizar_documento(match.group(1))

    for match in re.finditer(patron, texto):
        documento = normalizar_documento(match.group(1))
        if 7 <= len(documento) <= 9:
            return documento
    return ""


def sugerir_jugador_ficha_ocr(texto, jugadores):
    texto_norm = normalizar_texto_match(texto)
    dni_detectado = extraer_dni_ocr(texto)
    if dni_detectado:
        for jugador in jugadores:
            if normalizar_documento(jugador.get("dni")) == dni_detectado:
                return jugador, "alta", f"DNI {dni_detectado}"

    candidatos = []
    for jugador in jugadores:
        nombre_tokens = normalizar_texto_match(jugador.get("nombre")).split()
        apellido_tokens = normalizar_texto_match(jugador.get("apellido")).split()
        tokens = [token for token in apellido_tokens + nombre_tokens if len(token) >= 3]
        if not tokens:
            continue

        coincidencias = sum(1 for token in tokens if token in texto_norm)
        if coincidencias == len(tokens):
            candidatos.append((90, jugador, "media", "Nombre y apellido detectados"))
        elif apellido_tokens and all(token in texto_norm for token in apellido_tokens):
            candidatos.append((70, jugador, "baja", "Apellido detectado"))

    if not candidatos:
        return None, "sin_coincidencia", "Sin coincidencia"

    candidatos.sort(key=lambda item: item[0], reverse=True)
    mejor = candidatos[0]
    if len(candidatos) > 1 and candidatos[1][0] == mejor[0]:
        return None, "baja", "Coincidencia ambigua"
    return mejor[1], mejor[2], mejor[3]


def obtener_jugadores_selector(conn):
    return conn.execute("""
        SELECT id, apellido, nombre, dni, categoria, estado
        FROM jugadores
        WHERE COALESCE(estado, 'Activo') <> 'Baja'
        ORDER BY apellido, nombre
    """).fetchall()


def obtener_fichas_medicas_batch_recientes(conn):
    return conn.execute("""
        SELECT
            batch_id,
            MIN(creado_en) AS creado_en,
            MAX(creado_por) AS creado_por,
            COUNT(*) AS total,
            SUM(CASE WHEN estado = 'pendiente' THEN 1 ELSE 0 END) AS pendientes,
            SUM(CASE WHEN estado = 'procesado' THEN 1 ELSE 0 END) AS procesadas,
            MAX(id) AS ultimo_id
        FROM fichas_medicas_batch
        GROUP BY batch_id
        ORDER BY ultimo_id DESC
        LIMIT 10
    """).fetchall()


def procesar_ocr_ficha_medica_batch_item(conn, item):
    archivo = descargar_drive_file(item["drive_file_id"])
    content = archivo.getvalue()
    ext = item["extension"] or Path(item["archivo_original"] or item["documento_nombre"] or "").suffix.lower()
    mime_type = item["documento_mime_type"] or FICHA_MEDICA_EXTENSIONS.get(ext) or mimetypes.guess_type(item["documento_nombre"] or "")[0] or "application/pdf"
    filename = item["archivo_original"] or item["documento_nombre"] or f"ficha_medica_batch_{item['id']}{ext}"
    validado = {
        "filename": filename,
        "ext": ext,
        "content": content,
        "mime_type": mime_type,
    }

    ocr_texto = normalizar_ocr_texto(
        extraer_texto_ocr_drive(
            validado,
            {"id": "batch", "apellido": "batch", "nombre": item["batch_id"]},
            item["drive_folder_id"],
        )
    )
    datos_ocr = datos_ficha_desde_ocr(ocr_texto)
    jugadores = obtener_jugadores_selector(conn)
    jugador_sugerido, confianza, motivo = sugerir_jugador_ficha_ocr(ocr_texto, jugadores)

    conn.execute("""
        UPDATE fichas_medicas_batch
        SET ocr_texto = %s,
            ocr_fecha = %s,
            ocr_usuario = %s,
            jugador_sugerido_id = %s,
            confianza = %s,
            motivo = %s,
            fecha_vencimiento_sugerida = %s,
            apto_sugerido = %s,
            contacto_emergencia_sugerido = %s,
            telefono_emergencia_sugerido = %s,
            error = NULL
        WHERE id = %s
    """, (
        ocr_texto or None,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S") if ocr_texto else None,
        session.get("username") if ocr_texto else None,
        jugador_sugerido["id"] if jugador_sugerido else None,
        confianza,
        motivo,
        datos_ocr.get("fecha_vencimiento") or None,
        datos_ocr.get("apto_fisico"),
        datos_ocr.get("contacto_emergencia") or None,
        datos_ocr.get("telefono_emergencia") or None,
        item["id"],
    ))
    return ocr_texto


def descargar_drive_file(file_id):
    service = drive_service()
    drive_request = service.files().get_media(
        fileId=file_id,
        supportsAllDrives=True,
    )
    file_buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(file_buffer, drive_request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    file_buffer.seek(0)
    return file_buffer


def mensaje_error_drive(error, carpeta="Cuotas", accion="subir el comprobante"):
    if HttpError is not None and isinstance(error, HttpError):
        status = getattr(getattr(error, "resp", None), "status", None)
        detalle = ""
        try:
            data = json.loads(error.content.decode("utf-8"))
            detalle = data.get("error", {}).get("message", "")
        except Exception:
            detalle = str(error)

        detalle_lower = detalle.lower()
        if "storage quota" in detalle_lower or "service accounts do not have storage quota" in detalle_lower:
            return (
                "Google Drive rechazó la subida porque la cuenta de servicio no puede "
                "guardar archivos en una carpeta de Mi unidad. Usá una unidad compartida "
                "de Google Drive o una integración OAuth con un usuario real."
            )
        if status == 403:
            return (
                f"Google Drive rechazó el acceso. Revisá que la carpeta {carpeta} esté compartida "
                "como editor con la service account de Cloud Run."
            )
        if status == 404:
            return (
                f"Google Drive no encontró la carpeta configurada. Revisá el ID de la carpeta {carpeta}."
            )
        if detalle:
            return f"Google Drive rechazó la operación: {truncate_audit_value(detalle, 180)}"

    return f"No se pudo {accion} en Google Drive."


def fecha_movimiento_default(mes=None):
    hoy = datetime.now().strftime("%Y-%m-%d")
    if not mes:
        return hoy

    try:
        datetime.strptime(mes, "%Y-%m")
    except ValueError:
        return hoy

    if mes == hoy[:7]:
        return hoy

    return f"{mes}-01"


def validar_fecha_movimiento(fecha):
    try:
        datetime.strptime(fecha, "%Y-%m-%d")
    except (TypeError, ValueError):
        return None
    return fecha


def validar_mes_beca(valor):
    valor = (valor or "").strip()
    if not valor:
        return ""
    try:
        datetime.strptime(valor, "%Y-%m")
    except ValueError:
        return None
    return valor


def porcentaje_beca(valor):
    try:
        porcentaje = float(valor or 0)
    except (TypeError, ValueError):
        return None
    if porcentaje < 0 or porcentaje > 100:
        return None
    return round(porcentaje, 2)


def datos_beca_form():
    activa = 1 if request.form.get("beca_activa") == "on" else 0

    if not activa:
        return {
            "beca_activa": 0,
            "beca_porcentaje": 0,
            "beca_desde": "",
            "beca_hasta": "",
            "beca_motivo": "",
        }, None

    porcentaje = porcentaje_beca(request.form.get("beca_porcentaje", "0"))
    desde = validar_mes_beca(request.form.get("beca_desde", ""))
    hasta = validar_mes_beca(request.form.get("beca_hasta", ""))
    motivo = request.form.get("beca_motivo", "").strip()
    if porcentaje is None:
        return None, "El porcentaje de beca debe estar entre 0 y 100."
    if desde is None or hasta is None:
        return None, "La vigencia de la beca debe tener formato YYYY-MM."
    if desde and hasta and hasta < desde:
        return None, "La fecha hasta de la beca no puede ser anterior a la fecha desde."
    if activa and porcentaje <= 0:
        return None, "Para activar una beca indicá un porcentaje mayor a 0."

    return {
        "beca_activa": activa,
        "beca_porcentaje": porcentaje,
        "beca_desde": desde,
        "beca_hasta": hasta,
        "beca_motivo": motivo,
    }, None


def beca_vigente(jugador, periodo):
    if not jugador or not jugador.get("beca_activa"):
        return False

    porcentaje = porcentaje_beca(jugador.get("beca_porcentaje"))
    if not porcentaje or porcentaje <= 0:
        return False

    periodo = (periodo or "").strip()
    try:
        datetime.strptime(periodo, "%Y-%m")
    except ValueError:
        return False

    desde = (jugador.get("beca_desde") or "").strip()
    hasta = (jugador.get("beca_hasta") or "").strip()
    if desde and periodo < desde:
        return False
    if hasta and periodo > hasta:
        return False
    return True


def calcular_importe_con_beca(jugador, periodo, importe_base):
    importe_original = round(float(importe_base or 0), 2)
    if not beca_vigente(jugador, periodo):
        return {
            "importe_original": importe_original,
            "importe": importe_original,
            "descuento_beca": 0,
            "beca_porcentaje": 0,
            "beca_motivo": "",
            "becada": 0,
            "beca_total": 0,
        }

    porcentaje = porcentaje_beca(jugador.get("beca_porcentaje")) or 0
    descuento = round(importe_original * porcentaje / 100, 2)
    importe = max(0, round(importe_original - descuento, 2))
    return {
        "importe_original": importe_original,
        "importe": importe,
        "descuento_beca": descuento,
        "beca_porcentaje": porcentaje,
        "beca_motivo": jugador.get("beca_motivo") or "",
        "becada": 1,
        "beca_total": 1 if importe <= 0 else 0,
    }


def snapshot_beca(jugador):
    if not jugador:
        return {
            "beca_activa": 0,
            "beca_porcentaje": 0,
            "beca_desde": "",
            "beca_hasta": "",
            "beca_motivo": "",
        }
    return {
        "beca_activa": 1 if jugador.get("beca_activa") else 0,
        "beca_porcentaje": porcentaje_beca(jugador.get("beca_porcentaje")) or 0,
        "beca_desde": jugador.get("beca_desde") or "",
        "beca_hasta": jugador.get("beca_hasta") or "",
        "beca_motivo": jugador.get("beca_motivo") or "",
    }


def beca_modificada(jugador, data):
    anterior = snapshot_beca(jugador)
    nuevo = snapshot_beca(data)
    return anterior != nuevo


def registrar_historial_beca(conn, jugador_id, data, accion, detalle=None):
    snapshot = snapshot_beca(data)
    conn.execute("""
        INSERT INTO becas_historial (
            jugador_id, accion, beca_activa, beca_porcentaje, beca_desde,
            beca_hasta, beca_motivo, detalle, creado_en, creado_por
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, CURRENT_TIMESTAMP, %s)
    """, (
        jugador_id,
        accion,
        snapshot["beca_activa"],
        snapshot["beca_porcentaje"],
        snapshot["beca_desde"] or None,
        snapshot["beca_hasta"] or None,
        snapshot["beca_motivo"] or None,
        json.dumps(detalle or {}, ensure_ascii=False, default=str),
        session.get("username") if has_request_context() else "sistema",
    ))


def recalcular_cuotas_becadas(conn, jugador, periodo_desde="", periodo_hasta=""):
    condiciones = ["jugador_id = %s", "pagado = 0"]
    parametros = [jugador["id"]]

    if periodo_desde:
        condiciones.append("periodo >= %s")
        parametros.append(periodo_desde)
    if periodo_hasta:
        condiciones.append("periodo <= %s")
        parametros.append(periodo_hasta)

    cuotas = conn.execute(f"""
        SELECT *
        FROM cuotas
        WHERE {" AND ".join(condiciones)}
        ORDER BY periodo ASC, id ASC
    """, parametros).fetchall()

    resultado = {
        "revisadas": len(cuotas),
        "actualizadas": 0,
        "becas_totales": 0,
        "becas_parciales": 0,
        "sin_beca": 0,
    }

    hoy = datetime.now().strftime("%Y-%m-%d")
    for cuota in cuotas:
        importe_base = cuota.get("importe_original")
        if importe_base is None:
            importe_base = cuota.get("importe") or 0

        cuota_calculada = calcular_importe_con_beca(jugador, cuota["periodo"], importe_base)
        pagado = 1 if cuota_calculada["beca_total"] else 0
        fecha_pago = hoy if pagado else None
        metodo_pago = "Beca" if pagado else None
        referencia_pago = (
            f"Beca total {cuota_calculada['beca_porcentaje']:g}%"
            if pagado else None
        )

        conn.execute("""
            UPDATE cuotas
            SET importe = %s,
                importe_original = %s,
                descuento_beca = %s,
                beca_porcentaje = %s,
                beca_motivo = %s,
                becada = %s,
                pagado = %s,
                fecha_pago = %s,
                metodo_pago = %s,
                referencia_pago = %s
            WHERE id = %s
        """, (
            cuota_calculada["importe"],
            cuota_calculada["importe_original"],
            cuota_calculada["descuento_beca"],
            cuota_calculada["beca_porcentaje"],
            cuota_calculada["beca_motivo"],
            cuota_calculada["becada"],
            pagado,
            fecha_pago,
            metodo_pago,
            referencia_pago,
            cuota["id"],
        ))

        resultado["actualizadas"] += 1
        if cuota_calculada["beca_total"]:
            resultado["becas_totales"] += 1
        elif cuota_calculada["becada"]:
            resultado["becas_parciales"] += 1
        else:
            resultado["sin_beca"] += 1

    return resultado


def generar_portal_token():
    return secrets.token_urlsafe(32)


def normalizar_mes(valor, default):
    valor = (valor or "").strip()
    try:
        datetime.strptime(valor, "%Y-%m")
        return valor
    except ValueError:
        return default


def meses_entre(desde, hasta):
    actual = datetime.strptime(desde, "%Y-%m")
    final = datetime.strptime(hasta, "%Y-%m")
    meses = []

    while actual <= final:
        meses.append(actual.strftime("%Y-%m"))
        if actual.month == 12:
            actual = actual.replace(year=actual.year + 1, month=1)
        else:
            actual = actual.replace(month=actual.month + 1)

    return meses


def filtros_reportes():
    hoy = datetime.now()
    default_desde = f"{hoy.year}-01"
    default_hasta = hoy.strftime("%Y-%m")

    desde = normalizar_mes(request.args.get("desde"), default_desde)
    hasta = normalizar_mes(request.args.get("hasta"), default_hasta)

    if desde > hasta:
        desde, hasta = hasta, desde

    return {
        "desde": desde,
        "hasta": hasta,
    }


def obtener_reportes(desde, hasta):
    conn = get_connection()

    resumen_caja = conn.execute("""
        SELECT
            COALESCE(SUM(CASE WHEN tipo = 'ingreso' THEN monto ELSE 0 END), 0) AS ingresos,
            COALESCE(SUM(CASE WHEN tipo = 'egreso' THEN monto ELSE 0 END), 0) AS egresos
        FROM movimientos
        WHERE COALESCE(anulado, 0) = 0
          AND substring(fecha from 1 for 7) BETWEEN %s AND %s
    """, (desde, hasta)).fetchone()

    resumen_cuotas = conn.execute("""
        SELECT
            COALESCE(SUM(CASE WHEN pagado = 1 THEN importe ELSE 0 END), 0) AS cuotas_cobradas,
            COUNT(CASE WHEN pagado = 1 THEN 1 END) AS cuotas_pagadas,
            COUNT(CASE WHEN pagado = 0 AND COALESCE(importe, 0) > 0 THEN 1 END) AS cuotas_pendientes_periodo
        FROM cuotas
        WHERE periodo BETWEEN %s AND %s
    """, (desde, hasta)).fetchone()

    deuda_total = conn.execute("""
        SELECT
            COALESCE(SUM(importe), 0) AS deuda,
            COALESCE(SUM(
                CASE
                    WHEN fecha_vencimiento IS NOT NULL
                     AND fecha_vencimiento <> ''
                     AND fecha_vencimiento::date < CURRENT_DATE
                    THEN importe
                    ELSE 0
                END
            ), 0) AS deuda_vencida,
            COUNT(*) AS cuotas_pendientes
        FROM cuotas
        WHERE pagado = 0
          AND COALESCE(importe, 0) > 0
    """).fetchone()

    jugadores = conn.execute("""
        SELECT
            COUNT(*) AS total,
            COUNT(CASE WHEN estado = 'Activo' THEN 1 END) AS activos
        FROM jugadores
    """).fetchone()

    asistencia_resumen = conn.execute("""
        SELECT
            COUNT(DISTINCT e.id) AS eventos,
            COUNT(a.id) AS registros,
            COALESCE(SUM(CASE WHEN a.presente = 1 THEN 1 ELSE 0 END), 0) AS presentes
        FROM eventos_asistencia e
        LEFT JOIN asistencias a ON a.evento_id = e.id
        WHERE substring(e.fecha from 1 for 7) BETWEEN %s AND %s
    """, (desde, hasta)).fetchone()

    movimientos_mensuales = conn.execute("""
        SELECT
            substring(fecha from 1 for 7) AS mes,
            COALESCE(SUM(CASE WHEN tipo = 'ingreso' THEN monto ELSE 0 END), 0) AS ingresos,
            COALESCE(SUM(CASE WHEN tipo = 'egreso' THEN monto ELSE 0 END), 0) AS egresos,
            COUNT(*) AS movimientos
        FROM movimientos
        WHERE COALESCE(anulado, 0) = 0
          AND substring(fecha from 1 for 7) BETWEEN %s AND %s
        GROUP BY substring(fecha from 1 for 7)
        ORDER BY mes ASC
    """, (desde, hasta)).fetchall()

    cuotas_mensuales = conn.execute("""
        SELECT
            periodo AS mes,
            COUNT(*) AS cuotas_emitidas,
            COUNT(CASE WHEN pagado = 1 THEN 1 END) AS cuotas_pagadas,
            COUNT(CASE WHEN pagado = 0 AND COALESCE(importe, 0) > 0 THEN 1 END) AS cuotas_pendientes,
            COALESCE(SUM(importe), 0) AS total_emitido,
            COALESCE(SUM(CASE WHEN pagado = 1 THEN importe ELSE 0 END), 0) AS total_cobrado
        FROM cuotas
        WHERE periodo BETWEEN %s AND %s
        GROUP BY periodo
        ORDER BY periodo ASC
    """, (desde, hasta)).fetchall()

    becas_resumen = conn.execute("""
        SELECT
            COUNT(CASE WHEN COALESCE(c.becada, 0) = 1 THEN 1 END) AS cuotas_becadas,
            COUNT(CASE WHEN COALESCE(c.becada, 0) = 1 AND COALESCE(c.importe, 0) = 0 THEN 1 END) AS becas_totales,
            COUNT(CASE WHEN COALESCE(c.becada, 0) = 1 AND COALESCE(c.importe, 0) > 0 THEN 1 END) AS becas_parciales,
            COALESCE(SUM(CASE WHEN COALESCE(c.becada, 0) = 1 THEN c.descuento_beca ELSE 0 END), 0) AS total_bonificado,
            COALESCE(SUM(CASE WHEN COALESCE(c.becada, 0) = 1 THEN COALESCE(c.importe_original, c.importe) ELSE 0 END), 0) AS total_original_becado,
            COALESCE(SUM(CASE WHEN COALESCE(c.becada, 0) = 1 THEN c.importe ELSE 0 END), 0) AS total_neto_becado
        FROM cuotas c
        WHERE c.periodo BETWEEN %s AND %s
    """, (desde, hasta)).fetchone()

    becas_mensuales = conn.execute("""
        SELECT
            periodo AS mes,
            COUNT(*) AS cuotas_becadas,
            COUNT(CASE WHEN COALESCE(importe, 0) = 0 THEN 1 END) AS becas_totales,
            COUNT(CASE WHEN COALESCE(importe, 0) > 0 THEN 1 END) AS becas_parciales,
            COALESCE(SUM(descuento_beca), 0) AS total_bonificado
        FROM cuotas
        WHERE periodo BETWEEN %s AND %s
          AND COALESCE(becada, 0) = 1
        GROUP BY periodo
        ORDER BY periodo ASC
    """, (desde, hasta)).fetchall()

    becas_jugadores = conn.execute("""
        SELECT
            j.id,
            j.apellido,
            j.nombre,
            j.categoria,
            j.beca_porcentaje,
            j.beca_desde,
            j.beca_hasta,
            j.beca_motivo,
            COUNT(c.id) AS cuotas_becadas,
            COALESCE(SUM(c.descuento_beca), 0) AS total_bonificado
        FROM jugadores j
        LEFT JOIN cuotas c
          ON c.jugador_id = j.id
         AND c.periodo BETWEEN %s AND %s
         AND COALESCE(c.becada, 0) = 1
        WHERE COALESCE(j.beca_activa, 0) = 1
        GROUP BY
            j.id, j.apellido, j.nombre, j.categoria, j.beca_porcentaje,
            j.beca_desde, j.beca_hasta, j.beca_motivo
        ORDER BY j.apellido, j.nombre
    """, (desde, hasta)).fetchall()

    deuda_por_categoria = conn.execute("""
        SELECT
            COALESCE(NULLIF(j.categoria, ''), 'Sin categoria') AS categoria,
            COUNT(DISTINCT j.id) AS jugadores,
            COUNT(c.id) AS cuotas_pendientes,
            COALESCE(SUM(c.importe), 0) AS deuda,
            COALESCE(SUM(
                CASE
                    WHEN c.fecha_vencimiento IS NOT NULL
                     AND c.fecha_vencimiento <> ''
                     AND c.fecha_vencimiento::date < CURRENT_DATE
                    THEN c.importe
                    ELSE 0
                END
            ), 0) AS deuda_vencida
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
        GROUP BY COALESCE(NULLIF(j.categoria, ''), 'Sin categoria')
        ORDER BY deuda DESC, categoria ASC
    """).fetchall()

    egresos_por_concepto = conn.execute("""
        SELECT
            COALESCE(NULLIF(concepto, ''), 'Sin concepto') AS concepto,
            COUNT(*) AS cantidad,
            COALESCE(SUM(monto), 0) AS total
        FROM movimientos
        WHERE tipo = 'egreso'
          AND COALESCE(anulado, 0) = 0
          AND substring(fecha from 1 for 7) BETWEEN %s AND %s
        GROUP BY COALESCE(NULLIF(concepto, ''), 'Sin concepto')
        ORDER BY total DESC, concepto ASC
        LIMIT 15
    """, (desde, hasta)).fetchall()

    morosos_recurrentes = conn.execute("""
        SELECT
            j.id,
            j.apellido,
            j.nombre,
            j.categoria,
            j.telefono,
            j.email,
            COUNT(c.id) AS cuotas_pendientes,
            SUM(
                CASE
                    WHEN c.fecha_vencimiento IS NOT NULL
                     AND c.fecha_vencimiento <> ''
                     AND c.fecha_vencimiento::date < CURRENT_DATE
                    THEN 1
                    ELSE 0
                END
            ) AS cuotas_vencidas,
            COALESCE(SUM(c.importe), 0) AS deuda,
            MIN(NULLIF(c.fecha_vencimiento, '')) AS primer_vencimiento
        FROM jugadores j
        JOIN cuotas c ON c.jugador_id = j.id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
        GROUP BY j.id, j.apellido, j.nombre, j.categoria, j.telefono, j.email
        HAVING COUNT(c.id) >= 2
        ORDER BY deuda DESC, cuotas_vencidas DESC, j.apellido, j.nombre
        LIMIT 20
    """).fetchall()

    asistencia_por_categoria = conn.execute("""
        SELECT
            COALESCE(NULLIF(j.categoria, ''), 'Sin categoria') AS categoria,
            COUNT(DISTINCT e.id) AS eventos,
            COUNT(a.id) AS registros,
            COALESCE(SUM(CASE WHEN a.presente = 1 THEN 1 ELSE 0 END), 0) AS presentes,
            COALESCE(SUM(CASE WHEN a.presente = 0 THEN 1 ELSE 0 END), 0) AS ausentes
        FROM eventos_asistencia e
        JOIN asistencias a ON a.evento_id = e.id
        JOIN jugadores j ON j.id = a.jugador_id
        WHERE substring(e.fecha from 1 for 7) BETWEEN %s AND %s
        GROUP BY COALESCE(NULLIF(j.categoria, ''), 'Sin categoria')
        ORDER BY categoria ASC
    """, (desde, hasta)).fetchall()

    conn.close()

    movimientos_por_mes = {fila["mes"]: fila for fila in movimientos_mensuales}
    cuotas_por_mes = {fila["mes"]: fila for fila in cuotas_mensuales}
    becas_por_mes = {fila["mes"]: fila for fila in becas_mensuales}
    mensual = []

    for mes in meses_entre(desde, hasta):
        caja = movimientos_por_mes.get(mes, {})
        cuotas = cuotas_por_mes.get(mes, {})
        becas = becas_por_mes.get(mes, {})
        ingresos = caja.get("ingresos", 0) or 0
        egresos = caja.get("egresos", 0) or 0

        mensual.append({
            "mes": mes,
            "ingresos": ingresos,
            "egresos": egresos,
            "resultado": ingresos - egresos,
            "movimientos": caja.get("movimientos", 0) or 0,
            "cuotas_emitidas": cuotas.get("cuotas_emitidas", 0) or 0,
            "cuotas_pagadas": cuotas.get("cuotas_pagadas", 0) or 0,
            "cuotas_pendientes": cuotas.get("cuotas_pendientes", 0) or 0,
            "total_emitido": cuotas.get("total_emitido", 0) or 0,
            "total_cobrado": cuotas.get("total_cobrado", 0) or 0,
            "cuotas_becadas": becas.get("cuotas_becadas", 0) or 0,
            "becas_totales": becas.get("becas_totales", 0) or 0,
            "becas_parciales": becas.get("becas_parciales", 0) or 0,
            "total_bonificado": becas.get("total_bonificado", 0) or 0,
        })

    asistencia_registros = asistencia_resumen["registros"] or 0
    asistencia_presentes = asistencia_resumen["presentes"] or 0
    asistencia_porcentaje = (
        round((asistencia_presentes / asistencia_registros) * 100, 1)
        if asistencia_registros else 0
    )

    for fila in asistencia_por_categoria:
        registros = fila["registros"] or 0
        presentes = fila["presentes"] or 0
        fila["porcentaje"] = round((presentes / registros) * 100, 1) if registros else 0

    ingresos = resumen_caja["ingresos"] or 0
    egresos = resumen_caja["egresos"] or 0

    return {
        "resumen": {
            "ingresos": ingresos,
            "egresos": egresos,
            "resultado": ingresos - egresos,
            "cuotas_cobradas": resumen_cuotas["cuotas_cobradas"] or 0,
            "cuotas_pagadas": resumen_cuotas["cuotas_pagadas"] or 0,
            "cuotas_pendientes_periodo": resumen_cuotas["cuotas_pendientes_periodo"] or 0,
            "deuda": deuda_total["deuda"] or 0,
            "deuda_vencida": deuda_total["deuda_vencida"] or 0,
            "cuotas_pendientes": deuda_total["cuotas_pendientes"] or 0,
            "cuotas_becadas": becas_resumen["cuotas_becadas"] or 0,
            "becas_totales": becas_resumen["becas_totales"] or 0,
            "becas_parciales": becas_resumen["becas_parciales"] or 0,
            "total_bonificado_becas": becas_resumen["total_bonificado"] or 0,
            "total_original_becado": becas_resumen["total_original_becado"] or 0,
            "total_neto_becado": becas_resumen["total_neto_becado"] or 0,
            "jugadores_total": jugadores["total"] or 0,
            "jugadores_activos": jugadores["activos"] or 0,
            "asistencia_eventos": asistencia_resumen["eventos"] or 0,
            "asistencia_presentes": asistencia_presentes,
            "asistencia_registros": asistencia_registros,
            "asistencia_porcentaje": asistencia_porcentaje,
        },
        "mensual": mensual,
        "deuda_por_categoria": deuda_por_categoria,
        "egresos_por_concepto": egresos_por_concepto,
        "morosos_recurrentes": morosos_recurrentes,
        "asistencia_por_categoria": asistencia_por_categoria,
        "becas_jugadores": becas_jugadores,
    }


def obtener_alertas():
    conn = get_connection()

    cuotas_vencidas = conn.execute("""
        SELECT
            c.id,
            c.jugador_id,
            c.periodo,
            c.importe,
            c.fecha_vencimiento,
            CURRENT_DATE - c.fecha_vencimiento::date AS dias_vencida,
            j.apellido,
            j.nombre,
            j.categoria
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
          AND c.fecha_vencimiento IS NOT NULL
          AND c.fecha_vencimiento <> ''
          AND c.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND c.fecha_vencimiento::date < CURRENT_DATE
        ORDER BY c.fecha_vencimiento ASC, j.apellido, j.nombre
        LIMIT 25
    """).fetchall()

    cuotas_por_vencer = conn.execute("""
        SELECT
            c.id,
            c.jugador_id,
            c.periodo,
            c.importe,
            c.fecha_vencimiento,
            c.fecha_vencimiento::date - CURRENT_DATE AS dias_restantes,
            j.apellido,
            j.nombre,
            j.categoria
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
          AND c.fecha_vencimiento IS NOT NULL
          AND c.fecha_vencimiento <> ''
          AND c.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND c.fecha_vencimiento::date BETWEEN CURRENT_DATE AND CURRENT_DATE + INTERVAL '7 days'
        ORDER BY c.fecha_vencimiento ASC, j.apellido, j.nombre
        LIMIT 25
    """).fetchall()

    fichas_vencidas = conn.execute("""
        SELECT
            j.id AS jugador_id,
            j.apellido,
            j.nombre,
            j.categoria,
            f.fecha_vencimiento,
            CURRENT_DATE - f.fecha_vencimiento::date AS dias_vencida
        FROM fichas_medicas f
        JOIN jugadores j ON j.id = f.jugador_id
        WHERE f.fecha_vencimiento IS NOT NULL
          AND f.fecha_vencimiento <> ''
          AND f.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND f.fecha_vencimiento::date < CURRENT_DATE
        ORDER BY f.fecha_vencimiento ASC, j.apellido, j.nombre
        LIMIT 25
    """).fetchall()

    fichas_por_vencer = conn.execute("""
        SELECT
            j.id AS jugador_id,
            j.apellido,
            j.nombre,
            j.categoria,
            f.fecha_vencimiento,
            f.fecha_vencimiento::date - CURRENT_DATE AS dias_restantes
        FROM fichas_medicas f
        JOIN jugadores j ON j.id = f.jugador_id
        WHERE f.fecha_vencimiento IS NOT NULL
          AND f.fecha_vencimiento <> ''
          AND f.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND f.fecha_vencimiento::date BETWEEN CURRENT_DATE AND CURRENT_DATE + INTERVAL '30 days'
        ORDER BY f.fecha_vencimiento ASC, j.apellido, j.nombre
        LIMIT 25
    """).fetchall()

    lesiones_activas = conn.execute("""
        SELECT
            l.id,
            l.jugador_id,
            l.fecha_lesion,
            l.tipo_lesion,
            l.zona_cuerpo,
            l.estado,
            j.apellido,
            j.nombre,
            j.categoria
        FROM lesiones l
        JOIN jugadores j ON j.id = l.jugador_id
        WHERE l.estado IN ('Activa', 'En recuperación')
        ORDER BY
            CASE
                WHEN l.estado = 'Activa' THEN 0
                ELSE 1
            END,
            l.fecha_lesion DESC,
            j.apellido,
            j.nombre
        LIMIT 25
    """).fetchall()

    meses_sin_cerrar = conn.execute("""
        WITH mensual AS (
            SELECT
                substring(m.fecha from 1 for 7) AS mes,
                COALESCE(SUM(CASE WHEN m.tipo = 'ingreso' THEN m.monto ELSE 0 END), 0) AS ingresos,
                COALESCE(SUM(CASE WHEN m.tipo = 'egreso' THEN m.monto ELSE 0 END), 0) AS egresos,
                COUNT(*) AS movimientos
            FROM movimientos m
            WHERE COALESCE(m.anulado, 0) = 0
              AND substring(m.fecha from 1 for 7) < to_char(CURRENT_DATE, 'YYYY-MM')
            GROUP BY substring(m.fecha from 1 for 7)
        )
        SELECT *
        FROM mensual
        WHERE NOT EXISTS (
            SELECT 1
            FROM cierres_mensuales c
            WHERE c.mes = mensual.mes
        )
        ORDER BY mes DESC
        LIMIT 12
    """).fetchall()

    movimientos_altos = conn.execute("""
        WITH stats AS (
            SELECT COALESCE(AVG(monto), 0) AS promedio
            FROM movimientos
            WHERE COALESCE(anulado, 0) = 0
              AND fecha ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
              AND fecha::date >= CURRENT_DATE - INTERVAL '90 days'
        )
        SELECT
            m.id,
            m.fecha,
            m.tipo,
            m.concepto,
            m.referencia,
            m.monto,
            stats.promedio
        FROM movimientos m
        CROSS JOIN stats
        WHERE COALESCE(m.anulado, 0) = 0
          AND m.fecha ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND m.fecha::date >= CURRENT_DATE - INTERVAL '90 days'
          AND stats.promedio > 0
          AND m.monto >= stats.promedio * 2
        ORDER BY m.monto DESC, m.fecha DESC
        LIMIT 10
    """).fetchall()

    conn.close()

    for fila in meses_sin_cerrar:
        fila["resultado"] = (fila["ingresos"] or 0) - (fila["egresos"] or 0)

    deuda_vencida = sum((fila["importe"] or 0) for fila in cuotas_vencidas)
    deuda_proxima = sum((fila["importe"] or 0) for fila in cuotas_por_vencer)

    criticas = (
        len(cuotas_vencidas)
        + len(fichas_vencidas)
        + len(meses_sin_cerrar)
        + len(movimientos_altos)
    )

    proximas = len(cuotas_por_vencer) + len(fichas_por_vencer)
    seguimiento = len(lesiones_activas)

    return {
        "resumen": {
            "criticas": criticas,
            "proximas": proximas,
            "seguimiento": seguimiento,
            "total": criticas + proximas + seguimiento,
            "deuda_vencida": deuda_vencida,
            "deuda_proxima": deuda_proxima,
        },
        "cuotas_vencidas": cuotas_vencidas,
        "cuotas_por_vencer": cuotas_por_vencer,
        "fichas_vencidas": fichas_vencidas,
        "fichas_por_vencer": fichas_por_vencer,
        "lesiones_activas": lesiones_activas,
        "meses_sin_cerrar": meses_sin_cerrar,
        "movimientos_altos": movimientos_altos,
    }


def filtrar_alertas_por_permisos(alertas, puede_ver_finanzas, puede_ver_salud):
    if not puede_ver_finanzas:
        alertas["cuotas_vencidas"] = []
        alertas["cuotas_por_vencer"] = []
        alertas["meses_sin_cerrar"] = []
        alertas["movimientos_altos"] = []

    if not puede_ver_salud:
        alertas["fichas_vencidas"] = []
        alertas["fichas_por_vencer"] = []
        alertas["lesiones_activas"] = []

    deuda_vencida = sum((fila["importe"] or 0) for fila in alertas["cuotas_vencidas"])
    deuda_proxima = sum((fila["importe"] or 0) for fila in alertas["cuotas_por_vencer"])
    criticas = (
        len(alertas["cuotas_vencidas"])
        + len(alertas["fichas_vencidas"])
        + len(alertas["meses_sin_cerrar"])
        + len(alertas["movimientos_altos"])
    )
    proximas = len(alertas["cuotas_por_vencer"]) + len(alertas["fichas_por_vencer"])
    seguimiento = len(alertas["lesiones_activas"])

    alertas["resumen"] = {
        "criticas": criticas,
        "proximas": proximas,
        "seguimiento": seguimiento,
        "total": criticas + proximas + seguimiento,
        "deuda_vencida": deuda_vencida,
        "deuda_proxima": deuda_proxima,
    }
    return alertas


def obtener_panel_salud():
    conn = get_connection()

    resumen = conn.execute("""
        SELECT
            COUNT(*) AS jugadores_activos,
            SUM(CASE WHEN f.id IS NULL OR COALESCE(f.presentada, 0) = 0 THEN 1 ELSE 0 END) AS fichas_faltantes,
            SUM(CASE WHEN COALESCE(f.apto_fisico, 0) = 1 THEN 1 ELSE 0 END) AS aptos,
            SUM(CASE WHEN f.id IS NOT NULL AND COALESCE(f.apto_fisico, 0) = 0 THEN 1 ELSE 0 END) AS no_aptos,
            SUM(
                CASE
                    WHEN f.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                     AND f.fecha_vencimiento::date < CURRENT_DATE
                    THEN 1 ELSE 0
                END
            ) AS vencidas,
            SUM(
                CASE
                    WHEN f.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                     AND f.fecha_vencimiento::date BETWEEN CURRENT_DATE AND CURRENT_DATE + INTERVAL '30 days'
                    THEN 1 ELSE 0
                END
            ) AS por_vencer
        FROM jugadores j
        LEFT JOIN fichas_medicas f ON f.jugador_id = j.id
        WHERE j.estado = 'Activo'
    """).fetchone()

    fichas_por_categoria = conn.execute("""
        SELECT
            COALESCE(NULLIF(j.categoria, ''), 'Sin categoria') AS categoria,
            COUNT(*) AS jugadores,
            SUM(CASE WHEN COALESCE(f.apto_fisico, 0) = 1 THEN 1 ELSE 0 END) AS aptos,
            SUM(CASE WHEN f.id IS NOT NULL AND COALESCE(f.apto_fisico, 0) = 0 THEN 1 ELSE 0 END) AS no_aptos,
            SUM(CASE WHEN f.id IS NULL OR COALESCE(f.presentada, 0) = 0 THEN 1 ELSE 0 END) AS faltantes,
            SUM(
                CASE
                    WHEN f.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                     AND f.fecha_vencimiento::date < CURRENT_DATE
                    THEN 1 ELSE 0
                END
            ) AS vencidas,
            SUM(
                CASE
                    WHEN f.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                     AND f.fecha_vencimiento::date BETWEEN CURRENT_DATE AND CURRENT_DATE + INTERVAL '30 days'
                    THEN 1 ELSE 0
                END
            ) AS por_vencer
        FROM jugadores j
        LEFT JOIN fichas_medicas f ON f.jugador_id = j.id
        WHERE j.estado = 'Activo'
        GROUP BY COALESCE(NULLIF(j.categoria, ''), 'Sin categoria')
        ORDER BY categoria
    """).fetchall()

    lesiones_por_estado = conn.execute("""
        SELECT estado, COUNT(*) AS cantidad
        FROM lesiones
        GROUP BY estado
        ORDER BY
            CASE estado
                WHEN 'Activa' THEN 0
                ELSE CASE WHEN estado ILIKE 'En recuperaci%' THEN 1 WHEN estado ILIKE 'Alta%' THEN 2 ELSE 3 END
            END,
            estado
    """).fetchall()

    seguimiento_retorno = conn.execute("""
        SELECT
            l.id,
            l.jugador_id,
            l.fecha_lesion,
            l.fecha_alta,
            l.tipo_lesion,
            l.zona_cuerpo,
            l.estado,
            j.apellido,
            j.nombre,
            j.categoria
        FROM lesiones l
        JOIN jugadores j ON j.id = l.jugador_id
        WHERE l.estado = 'Activa'
           OR l.estado ILIKE 'En recuperaci%'
        ORDER BY
            CASE WHEN l.fecha_alta IS NULL OR l.fecha_alta = '' THEN 1 ELSE 0 END,
            l.fecha_alta ASC NULLS LAST,
            l.fecha_lesion DESC
        LIMIT 40
    """).fetchall()

    documentos_por_vencer = conn.execute("""
        SELECT
            d.id,
            d.jugador_id,
            d.tipo,
            d.nombre,
            d.fecha_vencimiento,
            j.apellido,
            j.nombre,
            j.categoria
        FROM documentos_jugadores d
        JOIN jugadores j ON j.id = d.jugador_id
        WHERE d.fecha_vencimiento IS NOT NULL
          AND d.fecha_vencimiento <> ''
          AND d.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND d.fecha_vencimiento::date <= CURRENT_DATE + INTERVAL '30 days'
        ORDER BY d.fecha_vencimiento ASC, j.apellido, j.nombre
        LIMIT 40
    """).fetchall()

    conn.close()

    return {
        "resumen": resumen,
        "fichas_por_categoria": fichas_por_categoria,
        "lesiones_por_estado": lesiones_por_estado,
        "seguimiento_retorno": seguimiento_retorno,
        "documentos_por_vencer": documentos_por_vencer,
    }


def obtener_estado_sistema_admin():
    estado = {
        "db_ok": False,
        "db_error": None,
        "db_time": None,
        "conteos": {},
        "mantenimiento": {
            "activo": False,
            "mensaje": MAINTENANCE_DEFAULT_MESSAGE,
            "actualizado_en": None,
            "actualizado_por": None,
        },
        "cloud_run": {
            "service": os.environ.get("K_SERVICE", "local"),
            "revision": os.environ.get("K_REVISION", APP_VERSION),
            "configuration": os.environ.get("K_CONFIGURATION", "local"),
            "project": os.environ.get("GOOGLE_CLOUD_PROJECT", ""),
            "region": os.environ.get("CLOUD_RUN_REGION", "us-central1"),
            "version": APP_VERSION,
        },
        "backup": {
            "nivel": "Cloud SQL automatico",
            "retencion_dias": CLOUD_SQL_BACKUP_RETENTION_DAYS,
            "pitr_dias": CLOUD_SQL_PITR_DAYS,
            "ventana": CLOUD_SQL_BACKUP_WINDOW,
            "ultimo_backup": os.environ.get("CLOUD_SQL_LAST_BACKUP", "Visible desde Google Cloud SQL"),
        },
    }

    conn = None
    try:
        conn = get_connection()
        estado["db_ok"] = True
        estado["db_time"] = conn.execute("SELECT CURRENT_TIMESTAMP AS ahora").fetchone()["ahora"]
        estado["mantenimiento"] = obtener_config_mantenimiento(conn)
        estado["conteos"] = conn.execute("""
            SELECT
                (SELECT COUNT(*) FROM jugadores) AS jugadores,
                (SELECT COUNT(*) FROM cuotas WHERE pagado = 0 AND COALESCE(importe, 0) > 0) AS cuotas_pendientes,
                (
                    SELECT COUNT(*)
                    FROM cuotas
                    WHERE comprobante_drive_file_id IS NOT NULL
                      AND COALESCE(NULLIF(comprobante_estado, ''), 'sin_comprobante') IN ('pendiente', 'sin_comprobante')
                ) AS comprobantes_pendientes,
                (
                    SELECT COUNT(*)
                    FROM fichas_medicas
                    WHERE fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                      AND fecha_vencimiento::date < CURRENT_DATE
                ) AS fichas_vencidas,
                (SELECT COUNT(*) FROM auditoria) AS auditoria_registros
        """).fetchone()
    except Exception as error:
        app.logger.exception("No se pudo obtener el estado del sistema.")
        estado["db_error"] = str(error)
    finally:
        if conn is not None:
            conn.close()

    return estado


def rango_mes(mes):
    inicio = datetime.strptime(mes, "%Y-%m")
    if inicio.month == 12:
        siguiente = inicio.replace(year=inicio.year + 1, month=1)
    else:
        siguiente = inicio.replace(month=inicio.month + 1)

    return inicio.strftime("%Y-%m-%d"), siguiente.strftime("%Y-%m-%d")


def obtener_calendario(mes):
    desde, hasta = rango_mes(mes)
    conn = get_connection()

    eventos_manuales = conn.execute("""
        SELECT
            id,
            fecha,
            tipo,
            titulo,
            descripcion,
            ubicacion,
            categoria
        FROM calendario_eventos
        WHERE fecha >= %s AND fecha < %s
        ORDER BY fecha ASC, id ASC
    """, (desde, hasta)).fetchall()

    eventos_asistencia = conn.execute("""
        SELECT
            id,
            fecha,
            tipo,
            descripcion
        FROM eventos_asistencia
        WHERE fecha >= %s AND fecha < %s
        ORDER BY fecha ASC, id ASC
    """, (desde, hasta)).fetchall()

    cuotas = conn.execute("""
        SELECT
            c.id,
            c.jugador_id,
            c.periodo,
            c.importe,
            c.fecha_vencimiento,
            j.apellido,
            j.nombre,
            j.categoria
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
          AND c.fecha_vencimiento IS NOT NULL
          AND c.fecha_vencimiento <> ''
          AND c.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND c.fecha_vencimiento >= %s
          AND c.fecha_vencimiento < %s
        ORDER BY c.fecha_vencimiento ASC, j.apellido, j.nombre
        LIMIT 80
    """, (desde, hasta)).fetchall()

    fichas = conn.execute("""
        SELECT
            j.id AS jugador_id,
            j.apellido,
            j.nombre,
            j.categoria,
            f.fecha_vencimiento
        FROM fichas_medicas f
        JOIN jugadores j ON j.id = f.jugador_id
        WHERE f.fecha_vencimiento IS NOT NULL
          AND f.fecha_vencimiento <> ''
          AND f.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND f.fecha_vencimiento >= %s
          AND f.fecha_vencimiento < %s
        ORDER BY f.fecha_vencimiento ASC, j.apellido, j.nombre
        LIMIT 80
    """, (desde, hasta)).fetchall()

    conn.close()

    eventos = []

    for evento in eventos_manuales:
        eventos.append({
            "fecha": evento["fecha"],
            "tipo": evento["tipo"],
            "titulo": evento["titulo"],
            "detalle": evento["descripcion"] or "",
            "ubicacion": evento["ubicacion"] or "",
            "categoria": evento["categoria"] or "",
            "origen": "calendario",
            "url": None,
            "prioridad": "normal",
        })

    for evento in eventos_asistencia:
        eventos.append({
            "fecha": evento["fecha"],
            "tipo": evento["tipo"],
            "titulo": f"Asistencia: {evento['tipo']}",
            "detalle": evento["descripcion"] or "",
            "ubicacion": "",
            "categoria": "",
            "origen": "asistencia",
            "url": url_for("tomar_asistencia", evento_id=evento["id"]),
            "prioridad": "normal",
        })

    for cuota in cuotas:
        eventos.append({
            "fecha": cuota["fecha_vencimiento"],
            "tipo": "Cuota",
            "titulo": f"Vence cuota {cuota['periodo']}",
            "detalle": f"{cuota['apellido']}, {cuota['nombre']} - {formato_moneda(cuota['importe'])}",
            "ubicacion": "",
            "categoria": cuota["categoria"] or "",
            "origen": "cuota",
            "url": url_for("ver_cuotas", jugador_id=cuota["jugador_id"]),
            "prioridad": "warning",
        })

    for ficha in fichas:
        eventos.append({
            "fecha": ficha["fecha_vencimiento"],
            "tipo": "Ficha médica",
            "titulo": "Vence ficha médica",
            "detalle": f"{ficha['apellido']}, {ficha['nombre']}",
            "ubicacion": "",
            "categoria": ficha["categoria"] or "",
            "origen": "ficha",
            "url": url_for("ver_ficha_medica", jugador_id=ficha["jugador_id"]),
            "prioridad": "danger",
        })

    eventos.sort(key=lambda item: (item["fecha"], item["tipo"], item["titulo"]))

    eventos_por_dia = {}
    for evento in eventos:
        eventos_por_dia.setdefault(evento["fecha"], []).append(evento)

    return {
        "mes": mes,
        "desde": desde,
        "hasta": hasta,
        "eventos": eventos,
        "eventos_por_dia": eventos_por_dia,
        "resumen": {
            "total": len(eventos),
            "manuales": len(eventos_manuales),
            "asistencia": len(eventos_asistencia),
            "cuotas": len(cuotas),
            "fichas": len(fichas),
        },
    }


def normalizar_telefono_whatsapp(telefono):
    digitos = re.sub(r"\D+", "", telefono or "")
    return digitos


def mensaje_moroso(template, jugador):
    contexto = {
        "nombre": jugador["nombre"] or "",
        "apellido": jugador["apellido"] or "",
        "deuda": formato_moneda(jugador["deuda"] or 0),
        "cuotas_pendientes": jugador["cuotas_pendientes"] or 0,
        "cuotas_vencidas": jugador["cuotas_vencidas"] or 0,
        "primer_vencimiento": jugador["primer_vencimiento"] or "-",
    }

    mensaje = template
    for clave, valor in contexto.items():
        mensaje = mensaje.replace("{" + clave + "}", str(valor))
    return mensaje


def obtener_morosos_para_comunicacion():
    conn = get_connection()
    morosos = conn.execute("""
        SELECT
            j.id,
            j.apellido,
            j.nombre,
            j.categoria,
            j.telefono,
            j.email,
            j.telefono_tutor,
            j.email_tutor,
            j.contacto_tutor,
            COUNT(c.id) AS cuotas_pendientes,
            SUM(
                CASE
                    WHEN c.fecha_vencimiento IS NOT NULL
                     AND c.fecha_vencimiento <> ''
                     AND c.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                     AND c.fecha_vencimiento::date < CURRENT_DATE
                    THEN 1
                    ELSE 0
                END
            ) AS cuotas_vencidas,
            COALESCE(SUM(c.importe), 0) AS deuda,
            MIN(NULLIF(c.fecha_vencimiento, '')) AS primer_vencimiento
        FROM jugadores j
        JOIN cuotas c ON c.jugador_id = j.id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
        GROUP BY
            j.id, j.apellido, j.nombre, j.categoria, j.telefono, j.email,
            j.telefono_tutor, j.email_tutor, j.contacto_tutor
        HAVING COALESCE(SUM(c.importe), 0) > 0
        ORDER BY deuda DESC, cuotas_vencidas DESC, j.apellido, j.nombre
    """).fetchall()
    conn.close()
    return morosos


def obtener_notificaciones_operativas():
    conn = get_connection()
    cuotas_vencidas = conn.execute("""
        SELECT
            c.id AS cuota_id,
            c.jugador_id,
            c.periodo,
            c.importe,
            c.fecha_vencimiento,
            j.apellido,
            j.nombre,
            j.telefono,
            j.telefono_tutor
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
          AND c.fecha_vencimiento IS NOT NULL
          AND c.fecha_vencimiento <> ''
          AND c.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND c.fecha_vencimiento::date < CURRENT_DATE
        ORDER BY c.fecha_vencimiento ASC, j.apellido, j.nombre
        LIMIT 50
    """).fetchall()

    cuotas_por_vencer = conn.execute("""
        SELECT
            c.id AS cuota_id,
            c.jugador_id,
            c.periodo,
            c.importe,
            c.fecha_vencimiento,
            j.apellido,
            j.nombre,
            j.telefono,
            j.telefono_tutor
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
          AND c.fecha_vencimiento IS NOT NULL
          AND c.fecha_vencimiento <> ''
          AND c.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND c.fecha_vencimiento::date >= CURRENT_DATE
          AND c.fecha_vencimiento::date <= CURRENT_DATE + INTERVAL '7 days'
        ORDER BY c.fecha_vencimiento ASC, j.apellido, j.nombre
        LIMIT 50
    """).fetchall()

    fichas = conn.execute("""
        SELECT
            j.id AS jugador_id,
            j.apellido,
            j.nombre,
            j.telefono,
            j.telefono_tutor,
            f.fecha_vencimiento,
            CASE
                WHEN f.fecha_vencimiento IS NULL OR f.fecha_vencimiento = '' THEN 'faltante'
                WHEN f.fecha_vencimiento !~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' THEN 'faltante'
                WHEN f.fecha_vencimiento::date < CURRENT_DATE THEN 'vencida'
                WHEN f.fecha_vencimiento::date <= CURRENT_DATE + INTERVAL '30 days' THEN 'por_vencer'
                ELSE 'vigente'
            END AS estado_documento
        FROM jugadores j
        LEFT JOIN fichas_medicas f ON f.jugador_id = j.id
        WHERE j.estado = 'Activo'
          AND (
              f.id IS NULL
              OR f.fecha_vencimiento IS NULL
              OR f.fecha_vencimiento = ''
              OR (
                  f.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                  AND f.fecha_vencimiento::date <= CURRENT_DATE + INTERVAL '30 days'
              )
          )
        ORDER BY estado_documento, f.fecha_vencimiento ASC NULLS FIRST, j.apellido, j.nombre
        LIMIT 80
    """).fetchall()

    asistencia_baja = conn.execute("""
        SELECT
            j.id AS jugador_id,
            j.apellido,
            j.nombre,
            j.telefono,
            j.telefono_tutor,
            COUNT(a.id) AS registros,
            SUM(CASE WHEN a.presente = 1 THEN 1 ELSE 0 END) AS presentes
        FROM jugadores j
        JOIN asistencias a ON a.jugador_id = j.id
        JOIN eventos_asistencia e ON e.id = a.evento_id
        WHERE j.estado = 'Activo'
          AND e.fecha::date >= CURRENT_DATE - INTERVAL '60 days'
        GROUP BY j.id, j.apellido, j.nombre, j.telefono, j.telefono_tutor
        HAVING COUNT(a.id) >= 3
           AND (SUM(CASE WHEN a.presente = 1 THEN 1 ELSE 0 END)::numeric / COUNT(a.id)) < 0.6
        ORDER BY (SUM(CASE WHEN a.presente = 1 THEN 1 ELSE 0 END)::numeric / COUNT(a.id)) ASC,
                 j.apellido, j.nombre
        LIMIT 50
    """).fetchall()

    comprobantes_pendientes = conn.execute("""
        SELECT
            c.id AS cuota_id,
            c.jugador_id,
            c.periodo,
            c.importe,
            c.comprobante_fecha,
            c.comprobante_usuario,
            c.comprobante_nombre,
            c.comprobante_web_url,
            j.apellido,
            j.nombre
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.comprobante_drive_file_id IS NOT NULL
          AND COALESCE(NULLIF(c.comprobante_estado, ''), 'sin_comprobante') IN ('pendiente', 'sin_comprobante')
        ORDER BY c.comprobante_fecha DESC NULLS LAST, c.id DESC
        LIMIT 50
    """).fetchall()

    ahijadxs_objetivo = conn.execute("""
        SELECT
            a.id AS aspirante_id,
            a.apellido,
            a.nombre,
            a.categoria,
            COALESCE(a.entrenamientos_objetivo, %s) AS entrenamientos_objetivo,
            COUNT(aa.id) FILTER (WHERE COALESCE(aa.presente, 0) = 1) AS entrenamientos_presentes
        FROM aspirantes a
        LEFT JOIN aspirante_asistencias aa ON aa.aspirante_id = a.id
        WHERE a.estado = 'Aspirante'
        GROUP BY a.id, a.apellido, a.nombre, a.categoria, a.entrenamientos_objetivo
        HAVING COUNT(aa.id) FILTER (WHERE COALESCE(aa.presente, 0) = 1) >= COALESCE(a.entrenamientos_objetivo, %s)
        ORDER BY a.apellido, a.nombre
        LIMIT 50
    """, (ASPIRANTE_ENTRENAMIENTOS_OBJETIVO, ASPIRANTE_ENTRENAMIENTOS_OBJETIVO)).fetchall()
    conn.close()

    return {
        "cuotas_vencidas": cuotas_vencidas,
        "cuotas_por_vencer": cuotas_por_vencer,
        "fichas": fichas,
        "asistencia_baja": asistencia_baja,
        "comprobantes_pendientes": comprobantes_pendientes,
        "ahijadxs_objetivo": ahijadxs_objetivo,
    }


def obtener_contador_notificaciones():
    if "user_id" not in session or not tiene_permiso("comunicaciones_ver"):
        return 0

    try:
        conn = get_connection()
        resumen = conn.execute("""
            SELECT
                (
                    SELECT COUNT(*)
                    FROM cuotas c
                    WHERE c.pagado = 0
                      AND COALESCE(c.importe, 0) > 0
                      AND c.fecha_vencimiento IS NOT NULL
                      AND c.fecha_vencimiento <> ''
                      AND c.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                      AND c.fecha_vencimiento::date <= CURRENT_DATE + INTERVAL '7 days'
                ) AS cuotas,
                (
                    SELECT COUNT(*)
                    FROM fichas_medicas f
                    JOIN jugadores j ON j.id = f.jugador_id
                    WHERE j.estado = 'Activo'
                      AND f.fecha_vencimiento IS NOT NULL
                      AND f.fecha_vencimiento <> ''
                      AND f.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                      AND f.fecha_vencimiento::date <= CURRENT_DATE + INTERVAL '30 days'
                ) AS fichas,
                (
                    SELECT COUNT(*)
                    FROM cuotas c
                    WHERE c.comprobante_drive_file_id IS NOT NULL
                      AND COALESCE(NULLIF(c.comprobante_estado, ''), 'sin_comprobante') IN ('pendiente', 'sin_comprobante')
                ) AS comprobantes,
                (
                    SELECT COUNT(*)
                    FROM aspirantes a
                    WHERE a.estado = 'Aspirante'
                      AND (
                          SELECT COUNT(*)
                          FROM aspirante_asistencias aa
                          WHERE aa.aspirante_id = a.id
                            AND COALESCE(aa.presente, 0) = 1
                      ) >= COALESCE(a.entrenamientos_objetivo, %s)
                ) AS ahijadxs
        """, (ASPIRANTE_ENTRENAMIENTOS_OBJETIVO,)).fetchone()
        conn.close()
    except Exception:
        app.logger.exception("No se pudo calcular el contador de notificaciones.")
        return 0

    return sum(int(resumen[campo] or 0) for campo in ("cuotas", "fichas", "comprobantes", "ahijadxs"))


def whatsapp_mensaje(telefono, mensaje):
    telefono_whatsapp = normalizar_telefono_whatsapp(telefono)
    if telefono_whatsapp:
        return f"https://wa.me/{telefono_whatsapp}?text={quote(mensaje)}"
    return f"https://wa.me/?text={quote(mensaje)}"


def normalizar_header_simple(valor):
    texto = unicodedata.normalize("NFKD", str(valor or "").strip().lower())
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "_", texto).strip("_")


def extraer_valor_fila(fila, opciones):
    for opcion in opciones:
        if opcion in fila and str(fila[opcion] or "").strip():
            return str(fila[opcion] or "").strip()
    return ""


def parsear_importe(valor):
    texto = str(valor or "").strip()
    if not texto:
        return None
    texto = re.sub(r"[^0-9,.-]", "", texto)
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return round(float(texto), 2)
    except ValueError:
        return None


def leer_csv_conciliacion(archivo):
    nombre = (archivo.filename or "").lower()
    if nombre.endswith(".xlsx"):
        try:
            wb = load_workbook(archivo, read_only=True, data_only=True)
            ws = wb.active
        except Exception as error:
            raise ValueError("No se pudo leer el archivo Excel.") from error

        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []

        headers = [normalizar_header_simple(valor) for valor in rows[0]]
        filas = []
        for row in rows[1:]:
            normalizada = {}
            for index, header in enumerate(headers):
                if not header:
                    continue
                valor = row[index] if index < len(row) else None
                normalizada[header] = limpiar_valor_excel(valor)
            if any(str(v or "").strip() for v in normalizada.values()):
                filas.append(normalizada)
        return filas

    raw = archivo.read()
    texto = None
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            texto = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if texto is None:
        raise ValueError("No se pudo leer el archivo CSV.")

    sample = texto[:2048]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,")
    except csv.Error:
        dialect = csv.excel
        dialect.delimiter = ";"

    reader = csv.DictReader(io.StringIO(texto), dialect=dialect)
    filas = []
    for row in reader:
        normalizada = {normalizar_header_simple(k): v for k, v in row.items() if k is not None}
        if any(str(v or "").strip() for v in normalizada.values()):
            filas.append(normalizada)
    return filas


def buscar_match_conciliacion(conn, fila):
    dni = extraer_valor_fila(fila, ["dni", "documento", "nro_documento", "numero_documento"])
    periodo = extraer_valor_fila(fila, ["periodo", "mes", "cuota", "periodo_cuota"])
    referencia = extraer_valor_fila(fila, ["referencia", "comprobante", "id", "operacion", "transaccion"])
    fecha_pago = validar_fecha_movimiento(extraer_valor_fila(fila, ["fecha", "fecha_pago", "fecha_de_pago"]))
    importe = parsear_importe(extraer_valor_fila(fila, ["importe", "monto", "valor", "total"]))

    if periodo:
        periodo = periodo[:7]

    if not dni or importe is None:
        return {
            "estado": "error",
            "motivo": "Falta DNI o importe.",
            "dni": dni,
            "periodo": periodo,
            "importe": importe,
            "referencia": referencia,
            "fecha_pago": fecha_pago,
        }

    condiciones = [
        "j.dni = %s",
        "c.pagado = 0",
        "COALESCE(c.importe, 0) > 0",
        "ABS(COALESCE(c.importe, 0) - %s) < 1",
    ]
    parametros = [dni, importe]
    if periodo:
        condiciones.append("c.periodo = %s")
        parametros.append(periodo)

    matches = conn.execute(f"""
        SELECT
            c.id AS cuota_id,
            c.jugador_id,
            c.periodo,
            c.importe,
            j.apellido,
            j.nombre,
            j.dni
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE {" AND ".join(condiciones)}
        ORDER BY c.periodo ASC, c.id ASC
        LIMIT 3
    """, parametros).fetchall()

    base = {
        "dni": dni,
        "periodo": periodo,
        "importe": importe,
        "referencia": referencia,
        "fecha_pago": fecha_pago,
    }
    if len(matches) == 1:
        match = matches[0]
        return {
            **base,
            "estado": "match",
            "cuota_id": match["cuota_id"],
            "jugador_id": match["jugador_id"],
            "jugador": f"{match['apellido']}, {match['nombre']}",
            "cuota_periodo": match["periodo"],
            "cuota_importe": match["importe"],
        }
    if len(matches) > 1:
        return {**base, "estado": "multiple", "motivo": "Hay mas de una cuota posible."}
    return {**base, "estado": "sin_match", "motivo": "No se encontro cuota pendiente compatible."}


def siguiente_numero_recibo(conn):
    conn.execute("SELECT pg_advisory_xact_lock(hashtext('sig_numero_recibo'))")
    ultimo = conn.execute("""
        SELECT MAX(numero_recibo) AS maximo
        FROM cuotas
    """).fetchone()["maximo"] or 0
    return ultimo + 1


def aplicar_matches_conciliacion(conn, matches):
    aplicados = 0
    omitidos = 0

    for match in matches:
        cuota_id = match.get("cuota_id")
        if match.get("estado") != "match" or not cuota_id:
            omitidos += 1
            continue

        revalidado = buscar_match_conciliacion(conn, match)
        try:
            cuota_id = int(cuota_id)
        except (TypeError, ValueError):
            omitidos += 1
            continue

        if revalidado.get("estado") != "match" or revalidado.get("cuota_id") != cuota_id:
            omitidos += 1
            continue

        cuota = conn.execute("""
            SELECT
                c.*,
                j.apellido,
                j.nombre
            FROM cuotas c
            JOIN jugadores j ON j.id = c.jugador_id
            WHERE c.id = %s
            FOR UPDATE
        """, (cuota_id,)).fetchone()
        if not cuota or cuota["pagado"]:
            omitidos += 1
            continue

        fecha_pago = match.get("fecha_pago") or datetime.now().strftime("%Y-%m-%d")
        referencia = match.get("referencia") or "Conciliacion importada"
        numero_recibo = cuota["numero_recibo"] or siguiente_numero_recibo(conn)
        conn.execute("""
            UPDATE cuotas
            SET pagado = 1,
                fecha_pago = %s,
                numero_recibo = %s,
                metodo_pago = 'Conciliacion',
                referencia_pago = %s
            WHERE id = %s
        """, (fecha_pago, numero_recibo, referencia, cuota["id"]))

        conn.execute("""
            INSERT INTO movimientos (tipo, concepto, monto, fecha, referencia)
            VALUES (%s, %s, %s, %s, %s)
        """, (
            "ingreso",
            f"Cuota {cuota['periodo']} - {cuota['apellido']}, {cuota['nombre']}",
            cuota["importe"],
            fecha_pago,
            "Cuota Social (Conciliacion)",
        ))
        aplicados += 1

    return {"aplicados": aplicados, "omitidos": omitidos}


def normalizar_header_excel(valor):
    texto = str(valor or "").strip().lower()
    reemplazos = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ñ": "n",
    }
    for origen, destino in reemplazos.items():
        texto = texto.replace(origen, destino)
    texto = re.sub(r"[^a-z0-9]+", "_", texto).strip("_")
    return texto


def limpiar_valor_excel(valor):
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%Y-%m-%d")
    return str(valor).strip()


def mapear_fila_jugador(headers, row):
    aliases = {
        "nombre": "nombre",
        "apellido": "apellido",
        "dni": "dni",
        "documento": "dni",
        "fecha_nacimiento": "fecha_nacimiento",
        "nacimiento": "fecha_nacimiento",
        "telefono": "telefono",
        "telefono_jugador": "telefono",
        "tel_fono": "telefono",
        "celular": "telefono",
        "email": "email",
        "mail": "email",
        "categoria": "categoria",
        "fecha_ingreso": "fecha_ingreso",
        "ingreso": "fecha_ingreso",
        "estado": "estado",
        "contacto_tutor": "contacto_tutor",
        "tutor": "contacto_tutor",
        "responsable": "contacto_tutor",
        "parentesco_tutor": "parentesco_tutor",
        "parentesco": "parentesco_tutor",
        "telefono_tutor": "telefono_tutor",
        "telefono_del_tutor": "telefono_tutor",
        "tel_tutor": "telefono_tutor",
        "tel_fono_tutor": "telefono_tutor",
        "email_tutor": "email_tutor",
        "direccion": "direccion",
        "obra_social": "obra_social",
        "numero_socio": "numero_socio",
        "nro_socio": "numero_socio",
        "documentos": "documentos",
        "observaciones": "observaciones",
    }

    data = {
        "nombre": "",
        "apellido": "",
        "dni": "",
        "fecha_nacimiento": "",
        "telefono": "",
        "email": "",
        "categoria": "",
        "fecha_ingreso": "",
        "estado": "Activo",
        "contacto_tutor": "",
        "parentesco_tutor": "",
        "telefono_tutor": "",
        "email_tutor": "",
        "direccion": "",
        "obra_social": "",
        "numero_socio": "",
        "documentos": "",
        "observaciones": "",
    }

    for index, header in enumerate(headers):
        campo = aliases.get(header)
        if campo:
            data[campo] = limpiar_valor_excel(row[index] if index < len(row) else None)

    if not data["estado"]:
        data["estado"] = "Activo"

    return data


def parsear_puntaje_test(valor):
    texto = str(valor or "").strip()
    if not texto:
        return None
    texto = re.sub(r"[^0-9,.-]", "", texto)
    if not texto:
        return None
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def normalizar_fecha_test(valor):
    if isinstance(valor, datetime):
        return valor.strftime("%Y-%m-%d")
    texto = str(valor or "").strip()
    if not texto:
        return datetime.now().strftime("%Y-%m-%d")
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(texto[:10], formato).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return validar_fecha_movimiento(texto) or datetime.now().strftime("%Y-%m-%d")


def obtener_test_tipos(conn, solo_activos=True):
    where = "WHERE activo = 1" if solo_activos else ""
    return conn.execute(f"""
        SELECT *
        FROM test_tipos
        {where}
        ORDER BY nombre
    """).fetchall()


def sugerir_jugador_por_nombre_test(datos, jugadores):
    nombre_completo = str(
        datos.get("jugador") or
        datos.get("nombre_completo") or
        datos.get("jugador_nombre") or
        ""
    ).strip()
    nombre = str(datos.get("nombre") or "").strip()
    apellido = str(datos.get("apellido") or "").strip()
    texto = " ".join(parte for parte in (nombre_completo, apellido, nombre) if parte).strip()

    if not texto:
        return None, "sin_coincidencia", "Sin nombre"

    texto_norm = normalizar_texto_match(texto)
    candidatos = []
    for jugador in jugadores:
        nombre_tokens = normalizar_texto_match(jugador.get("nombre")).split()
        apellido_tokens = normalizar_texto_match(jugador.get("apellido")).split()
        tokens = [token for token in apellido_tokens + nombre_tokens if len(token) >= 3]
        if not tokens:
            continue

        coincidencias = sum(1 for token in tokens if token in texto_norm)
        jugador_nombre_completo = normalizar_texto_match(f"{jugador.get('apellido')} {jugador.get('nombre')}")
        jugador_nombre_inverso = normalizar_texto_match(f"{jugador.get('nombre')} {jugador.get('apellido')}")

        if texto_norm in {jugador_nombre_completo, jugador_nombre_inverso}:
            candidatos.append((100, jugador, "alta", "Nombre completo exacto"))
        elif coincidencias == len(tokens):
            candidatos.append((90, jugador, "media", "Nombre y apellido detectados"))
        elif apellido_tokens and all(token in texto_norm for token in apellido_tokens):
            candidatos.append((70, jugador, "baja", "Apellido detectado"))

    if not candidatos:
        return None, "sin_coincidencia", "Sin coincidencia"

    candidatos.sort(key=lambda item: item[0], reverse=True)
    mejor = candidatos[0]
    if len(candidatos) > 1 and candidatos[1][0] == mejor[0]:
        return None, "baja", "Coincidencia ambigua"
    return mejor[1], mejor[2], mejor[3]


def obtener_test_importaciones_batch_recientes(conn):
    return conn.execute("""
        SELECT
            batch_id,
            MIN(creado_en) AS creado_en,
            MAX(creado_por) AS creado_por,
            COUNT(*) AS total,
            SUM(CASE WHEN estado = 'pendiente' THEN 1 ELSE 0 END) AS pendientes,
            SUM(CASE WHEN estado = 'procesado' THEN 1 ELSE 0 END) AS procesadas,
            MAX(id) AS ultimo_id
        FROM test_importaciones_batch
        GROUP BY batch_id
        ORDER BY ultimo_id DESC
        LIMIT 10
    """).fetchall()


def construir_grafico_tests(resultados):
    if not resultados:
        return {
            "series": [],
            "fechas": [],
            "minimo": 0,
            "maximo": 0,
            "svg_width": 920,
            "svg_height": 320,
            "eje_y": [],
        }

    fechas = sorted({fila["fecha"] for fila in resultados if fila["fecha"]})
    valores = [float(fila["puntaje"] or 0) for fila in resultados]
    minimo = min(valores)
    maximo = max(valores)
    if minimo == maximo:
        minimo -= 1
        maximo += 1

    width = 920
    height = 320
    pad_left = 54
    pad_right = 24
    pad_top = 22
    pad_bottom = 42
    inner_w = width - pad_left - pad_right
    inner_h = height - pad_top - pad_bottom
    palette = ["#2563eb", "#16a34a", "#dc2626", "#9333ea", "#ea580c", "#0891b2", "#be123c", "#4f46e5"]

    fecha_x = {}
    total_fechas = max(1, len(fechas) - 1)
    for index, fecha in enumerate(fechas):
        fecha_x[fecha] = pad_left + (inner_w * index / total_fechas if len(fechas) > 1 else inner_w / 2)

    agrupado = {}
    for fila in resultados:
        clave = fila["jugador_id"]
        agrupado.setdefault(clave, {
            "jugador_id": clave,
            "nombre": f"{fila['apellido']}, {fila['nombre']}",
            "categoria": fila["categoria"] or "-",
            "puntos": [],
        })
        valor = float(fila["puntaje"] or 0)
        y = pad_top + inner_h - ((valor - minimo) / (maximo - minimo) * inner_h)
        agrupado[clave]["puntos"].append({
            "fecha": fila["fecha"],
            "valor": valor,
            "x": round(fecha_x[fila["fecha"]], 2),
            "y": round(y, 2),
        })

    series = []
    for index, serie in enumerate(agrupado.values()):
        serie["puntos"].sort(key=lambda item: item["fecha"])
        serie["polyline"] = " ".join(f"{p['x']},{p['y']}" for p in serie["puntos"])
        serie["color"] = palette[index % len(palette)]
        serie["ultimo"] = serie["puntos"][-1]["valor"] if serie["puntos"] else None
        series.append(serie)

    eje_y = []
    for index in range(5):
        valor = maximo - ((maximo - minimo) * index / 4)
        y = pad_top + inner_h * index / 4
        eje_y.append({"valor": round(valor, 2), "y": round(y, 2)})

    return {
        "series": series,
        "fechas": fechas,
        "minimo": round(minimo, 2),
        "maximo": round(maximo, 2),
        "svg_width": width,
        "svg_height": height,
        "pad_left": pad_left,
        "pad_right": pad_right,
        "pad_top": pad_top,
        "pad_bottom": pad_bottom,
        "inner_w": inner_w,
        "inner_h": inner_h,
        "eje_y": eje_y,
    }


AUDIT_ENDPOINTS = {
    "nuevo_movimiento": ("crear", "movimiento_caja"),
    "editar_movimiento": ("editar", "movimiento_caja"),
    "eliminar_movimiento": ("anular", "movimiento_caja"),
    "cerrar_mes": ("cerrar", "caja"),
    "nuevo_jugador": ("crear", "jugador"),
    "editar_jugador": ("editar", "jugador"),
    "eliminar_jugador": ("eliminar", "jugador"),
    "nueva_bitacora_jugador": ("crear", "bitacora_jugador"),
    "importar_jugadores": ("importar", "jugadores"),
    "acciones_masivas_jugadores": ("accion_masiva", "jugadores"),
    "nuevo_aspirante": ("crear", "ahijadx"),
    "editar_aspirante": ("editar", "ahijadx"),
    "convertir_aspirante": ("convertir", "ahijadx"),
    "eliminar_aspirante": ("eliminar", "ahijadx"),
    "nueva_cuota": ("crear", "cuota"),
    "pagar_cuota": ("pagar", "cuota"),
    "subir_comprobante_cuota": ("subir", "comprobante_cuota"),
    "revisar_comprobante_cuota": ("revisar", "comprobante_cuota"),
    "eliminar_cuota": ("eliminar", "cuota"),
    "generar_cuotas": ("generar", "cuotas"),
    "recalcular_becas_jugador": ("recalcular_beca", "cuotas"),
    "nuevo_plan_pago": ("crear", "plan_pago"),
    "actualizar_plan_pago": ("actualizar", "plan_pago"),
    "conciliar_pagos": ("conciliar", "cuotas"),
    "nuevo_test_tipo": ("crear", "test_deportivo"),
    "cargar_test_resultados": ("cargar", "test_deportivo"),
    "importar_test_resultados": ("importar", "test_deportivo"),
    "revisar_test_importacion": ("confirmar_importacion", "test_deportivo"),
    "editar_ficha_medica": ("editar", "ficha_medica"),
    "cargar_fichas_medicas_batch": ("cargar_batch", "ficha_medica"),
    "revisar_fichas_medicas_batch": ("confirmar_batch", "ficha_medica"),
    "nuevo_documento_jugador": ("crear", "documento"),
    "eliminar_documento_jugador": ("eliminar", "documento"),
    "nueva_lesion": ("crear", "lesion"),
    "editar_lesion": ("editar", "lesion"),
    "eliminar_lesion": ("eliminar", "lesion"),
    "nuevo_usuario": ("crear", "usuario"),
    "editar_usuario": ("editar", "usuario"),
    "eliminar_usuario": ("eliminar", "usuario"),
    "nuevo_rol": ("crear", "rol"),
    "editar_rol": ("editar", "rol"),
    "eliminar_rol": ("eliminar", "rol"),
    "configurar_mantenimiento": ("configurar", "mantenimiento"),
    "cambiar_mi_password": ("cambiar_password", "usuario"),
    "resetear_password_usuario": ("resetear_password", "usuario"),
    "solicitar_recuperacion_password": ("solicitar_reset", "usuario"),
    "restablecer_password": ("restablecer_password", "usuario"),
    "descartar_onboarding": ("descartar", "onboarding"),
    "generar_portal_jugador": ("generar", "portal_jugador"),
    "desactivar_portal_jugador": ("desactivar", "portal_jugador"),
    "nuevo_evento_asistencia": ("crear", "evento_asistencia"),
    "tomar_asistencia": ("registrar", "asistencia"),
    "nuevo_evento_calendario": ("crear", "evento_calendario"),
}

SENSITIVE_AUDIT_FIELDS = {
    "password",
    "pass",
    "clave",
    "contraseña",
    "contrasena",
    "admin_password",
    "db_pass",
    "secret_key",
}


def audit_request_ip():
    forwarded_for = request.headers.get("X-Forwarded-For", "")
    if forwarded_for:
        return forwarded_for.split(",")[0].strip()
    return request.remote_addr


def truncate_audit_value(value, max_length=300):
    if value is None:
        return None
    value = str(value)
    if len(value) <= max_length:
        return value
    return value[:max_length] + "...[truncado]"


def sanitized_audit_form():
    data = {}
    for key, values in request.form.lists():
        key_lower = key.lower()
        if any(sensitive in key_lower for sensitive in SENSITIVE_AUDIT_FIELDS):
            data[key] = "[redactado]"
            continue

        clean_values = [truncate_audit_value(value) for value in values]
        data[key] = clean_values[0] if len(clean_values) == 1 else clean_values
    return data


def audit_entity_id():
    if not request.view_args:
        return None
    for key in (
        "jugador_id", "cuota_id", "plan_id", "documento_id", "lesion_id",
        "usuario_id", "rol_id", "movimiento_id", "evento_id"
    ):
        if key in request.view_args:
            return str(request.view_args[key])
    return None


def registrar_auditoria(
    accion,
    entidad="sistema",
    entidad_id=None,
    detalle=None,
    usuario_id=None,
    username=None,
    rol=None,
):
    conn = None
    try:
        if has_request_context():
            usuario_id = usuario_id if usuario_id is not None else session.get("user_id")
            username = username if username is not None else session.get("username")
            rol = rol if rol is not None else session.get("rol")
            ip = audit_request_ip()
            user_agent = request.headers.get("User-Agent", "")
        else:
            ip = None
            user_agent = None

        if isinstance(detalle, str):
            detalle_texto = detalle
        else:
            detalle_texto = json.dumps(detalle or {}, ensure_ascii=False, default=str)

        conn = get_connection()
        conn.execute("""
            INSERT INTO auditoria (
                usuario_id, username, rol, accion, entidad, entidad_id,
                detalle, ip, user_agent
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            usuario_id,
            username,
            rol,
            accion,
            entidad,
            entidad_id,
            detalle_texto,
            ip,
            truncate_audit_value(user_agent, 500),
        ))
        conn.commit()
        conn.close()
    except Exception:
        try:
            if conn:
                conn.close()
        except Exception:
            pass


def init_db():
    conn = get_connection()
    conn.execute("SELECT pg_advisory_xact_lock(hashtext('sig_init_db'))")

    conn.execute("""
    CREATE TABLE IF NOT EXISTS jugadores (
        id SERIAL PRIMARY KEY,
        nombre TEXT NOT NULL,
        apellido TEXT NOT NULL,
        dni TEXT,
        fecha_nacimiento TEXT,
        telefono TEXT,
        email TEXT,
        categoria TEXT,
        fecha_ingreso TEXT,
        estado TEXT NOT NULL DEFAULT 'Activo',
        observaciones TEXT
    )
""")

    columnas_jugadores = get_columns(conn, "jugadores")
    columnas_extra_jugador = {
        "telefono_tutor": "TEXT",
        "email_tutor": "TEXT",
        "contacto_tutor": "TEXT",
        "parentesco_tutor": "TEXT",
        "direccion": "TEXT",
        "obra_social": "TEXT",
        "numero_socio": "TEXT",
        "documentos": "TEXT",
        "beca_activa": "INTEGER DEFAULT 0",
        "beca_porcentaje": "REAL DEFAULT 0",
        "beca_desde": "TEXT",
        "beca_hasta": "TEXT",
        "beca_motivo": "TEXT",
        "portal_token": "TEXT",
        "portal_activo": "INTEGER DEFAULT 0",
        "portal_actualizado_en": "TEXT",
    }

    for columna, tipo_columna in columnas_extra_jugador.items():
        if columna not in columnas_jugadores:
            conn.execute(f"ALTER TABLE jugadores ADD COLUMN {columna} {tipo_columna}")

    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_jugadores_portal_token
        ON jugadores (portal_token)
        WHERE portal_token IS NOT NULL
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS aspirantes (
            id SERIAL PRIMARY KEY,
            nombre TEXT NOT NULL,
            apellido TEXT NOT NULL,
            dni TEXT,
            fecha_nacimiento TEXT,
            telefono TEXT,
            email TEXT,
            categoria TEXT,
            fecha_postulacion TEXT,
            estado TEXT NOT NULL DEFAULT 'Aspirante',
            madrina_jugador_id INTEGER,
            entrenamientos_objetivo INTEGER DEFAULT 8,
            fecha_ingreso_club TEXT,
            jugador_id INTEGER,
            observaciones TEXT,
            FOREIGN KEY (madrina_jugador_id) REFERENCES jugadores(id),
            FOREIGN KEY (jugador_id) REFERENCES jugadores(id)
        )
    """)

    columnas_aspirantes = get_columns(conn, "aspirantes")
    columnas_extra_aspirante = {
        "dni": "TEXT",
        "fecha_nacimiento": "TEXT",
        "telefono": "TEXT",
        "email": "TEXT",
        "categoria": "TEXT",
        "fecha_postulacion": "TEXT",
        "estado": "TEXT NOT NULL DEFAULT 'Aspirante'",
        "madrina_jugador_id": "INTEGER",
        "entrenamientos_objetivo": "INTEGER DEFAULT 8",
        "fecha_ingreso_club": "TEXT",
        "jugador_id": "INTEGER",
        "observaciones": "TEXT",
    }

    for columna, tipo_columna in columnas_extra_aspirante.items():
        if columna not in columnas_aspirantes:
            conn.execute(f"ALTER TABLE aspirantes ADD COLUMN {columna} {tipo_columna}")

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_aspirantes_estado
        ON aspirantes (estado)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS app_settings (
            clave TEXT PRIMARY KEY,
            valor TEXT,
            actualizado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
            actualizado_por TEXT
        )
    """)

    columnas_app_settings = get_columns(conn, "app_settings")
    if "actualizado_en" not in columnas_app_settings:
        conn.execute("ALTER TABLE app_settings ADD COLUMN actualizado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP")
    if "actualizado_por" not in columnas_app_settings:
        conn.execute("ALTER TABLE app_settings ADD COLUMN actualizado_por TEXT")

    conn.execute("""
        INSERT INTO app_settings (clave, valor, actualizado_por)
        VALUES ('maintenance_mode', 'false', 'sistema')
        ON CONFLICT(clave) DO NOTHING
    """)

    conn.execute("""
        INSERT INTO app_settings (clave, valor, actualizado_por)
        VALUES ('maintenance_message', %s, 'sistema')
        ON CONFLICT(clave) DO NOTHING
    """, (MAINTENANCE_DEFAULT_MESSAGE,))

    conn.execute("""
        CREATE TABLE IF NOT EXISTS cuotas (
            id SERIAL PRIMARY KEY,
            jugador_id INTEGER,
            periodo TEXT,
            importe REAL,
            pagado INTEGER DEFAULT 0,
            fecha_vencimiento TEXT,
            fecha_pago TEXT
        )
    """)


    columnas_cuotas = get_columns(conn, "cuotas")
    if "numero_recibo" not in columnas_cuotas:
        conn.execute("ALTER TABLE cuotas ADD COLUMN numero_recibo INTEGER")



    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_cuotas_jugador_periodo
        ON cuotas (jugador_id, periodo)
    """)

    columnas_cuotas = get_columns(conn, "cuotas")

    if "fecha_vencimiento" not in columnas_cuotas:
        conn.execute("ALTER TABLE cuotas ADD COLUMN fecha_vencimiento TEXT")

    if "referencia_pago" not in columnas_cuotas:
        conn.execute("ALTER TABLE cuotas ADD COLUMN referencia_pago TEXT")

    if "metodo_pago" not in columnas_cuotas:
        conn.execute("ALTER TABLE cuotas ADD COLUMN metodo_pago TEXT")

    columnas_cuotas = get_columns(conn, "cuotas")
    columnas_beca_cuota = {
        "importe_original": "REAL",
        "descuento_beca": "REAL DEFAULT 0",
        "beca_porcentaje": "REAL DEFAULT 0",
        "beca_motivo": "TEXT",
        "becada": "INTEGER DEFAULT 0",
    }

    for columna, tipo_columna in columnas_beca_cuota.items():
        if columna not in columnas_cuotas:
            conn.execute(f"ALTER TABLE cuotas ADD COLUMN {columna} {tipo_columna}")

    conn.execute("""
        UPDATE cuotas
        SET importe_original = importe
        WHERE importe_original IS NULL
          AND importe IS NOT NULL
    """)

    columnas_cuotas = get_columns(conn, "cuotas")
    columnas_comprobante_cuota = {
        "comprobante_drive_file_id": "TEXT",
        "comprobante_nombre": "TEXT",
        "comprobante_mime_type": "TEXT",
        "comprobante_tamano": "INTEGER",
        "comprobante_fecha": "TEXT",
        "comprobante_usuario": "TEXT",
        "comprobante_web_url": "TEXT",
        "comprobante_estado": "TEXT DEFAULT 'sin_comprobante'",
        "comprobante_revisado_en": "TEXT",
        "comprobante_revisado_por": "TEXT",
        "comprobante_observaciones": "TEXT",
    }

    for columna, tipo_columna in columnas_comprobante_cuota.items():
        if columna not in columnas_cuotas:
            conn.execute(f"ALTER TABLE cuotas ADD COLUMN {columna} {tipo_columna}")

    conn.execute("""
        UPDATE cuotas
        SET comprobante_estado = CASE
            WHEN comprobante_drive_file_id IS NOT NULL AND pagado = 1 THEN 'aceptado'
            WHEN comprobante_drive_file_id IS NOT NULL THEN 'pendiente'
            ELSE 'sin_comprobante'
        END
        WHERE comprobante_estado IS NULL
           OR comprobante_estado = ''
           OR (
               comprobante_drive_file_id IS NOT NULL
               AND comprobante_estado = 'sin_comprobante'
           )
    """)

    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_cuotas_jugador_periodo
        ON cuotas (jugador_id, periodo)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS becas_historial (
            id SERIAL PRIMARY KEY,
            jugador_id INTEGER NOT NULL,
            accion TEXT NOT NULL,
            beca_activa INTEGER DEFAULT 0,
            beca_porcentaje REAL DEFAULT 0,
            beca_desde TEXT,
            beca_hasta TEXT,
            beca_motivo TEXT,
            detalle TEXT,
            creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
            creado_por TEXT,
            FOREIGN KEY (jugador_id) REFERENCES jugadores(id)
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_becas_historial_jugador
        ON becas_historial (jugador_id, creado_en DESC)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS jugador_bitacora (
            id SERIAL PRIMARY KEY,
            jugador_id INTEGER NOT NULL,
            tipo TEXT NOT NULL DEFAULT 'general',
            nota TEXT NOT NULL,
            creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
            creado_por TEXT,
            FOREIGN KEY (jugador_id) REFERENCES jugadores(id) ON DELETE CASCADE
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_jugador_bitacora_jugador
        ON jugador_bitacora (jugador_id, creado_en DESC)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS planes_pago (
            id SERIAL PRIMARY KEY,
            jugador_id INTEGER NOT NULL,
            fecha_inicio TEXT,
            monto_total REAL NOT NULL,
            cantidad_cuotas INTEGER DEFAULT 1,
            monto_cuota REAL,
            estado TEXT NOT NULL DEFAULT 'Activo',
            descripcion TEXT,
            observaciones TEXT,
            creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
            creado_por TEXT,
            cerrado_en TEXT,
            cerrado_por TEXT,
            FOREIGN KEY (jugador_id) REFERENCES jugadores(id)
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_planes_pago_jugador
        ON planes_pago (jugador_id, estado)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS documentos_jugadores (
            id SERIAL PRIMARY KEY,
            jugador_id INTEGER NOT NULL,
            tipo TEXT NOT NULL,
            nombre TEXT,
            fecha_presentacion TEXT,
            fecha_vencimiento TEXT,
            url TEXT,
            observaciones TEXT,
            creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
            creado_por TEXT,
            FOREIGN KEY (jugador_id) REFERENCES jugadores(id)
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_documentos_jugadores_vencimiento
        ON documentos_jugadores (fecha_vencimiento)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS fichas_medicas (
            id SERIAL PRIMARY KEY,
            jugador_id INTEGER NOT NULL UNIQUE,
            presentada INTEGER DEFAULT 0,
            fecha_vencimiento TEXT,
            apto_fisico INTEGER DEFAULT 0,
            contacto_emergencia TEXT,
            telefono_emergencia TEXT,
            observaciones TEXT,
            FOREIGN KEY (jugador_id) REFERENCES jugadores (id)
        )
    """)

    columnas_fichas_medicas = get_columns(conn, "fichas_medicas")
    columnas_extra_ficha_medica = {
        "documento_drive_file_id": "TEXT",
        "documento_nombre": "TEXT",
        "documento_mime_type": "TEXT",
        "documento_tamano": "INTEGER",
        "documento_fecha": "TEXT",
        "documento_usuario": "TEXT",
        "documento_web_url": "TEXT",
        "ocr_texto": "TEXT",
        "ocr_fecha": "TEXT",
        "ocr_usuario": "TEXT",
    }

    for columna, tipo_columna in columnas_extra_ficha_medica.items():
        if columna not in columnas_fichas_medicas:
            conn.execute(f"ALTER TABLE fichas_medicas ADD COLUMN {columna} {tipo_columna}")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS fichas_medicas_batch (
            id SERIAL PRIMARY KEY,
            batch_id TEXT NOT NULL,
            estado TEXT NOT NULL DEFAULT 'pendiente',
            archivo_original TEXT,
            drive_file_id TEXT,
            drive_folder_id TEXT,
            documento_nombre TEXT,
            documento_mime_type TEXT,
            documento_tamano INTEGER,
            documento_web_url TEXT,
            extension TEXT,
            ocr_texto TEXT,
            ocr_fecha TEXT,
            ocr_usuario TEXT,
            jugador_sugerido_id INTEGER,
            confianza TEXT,
            motivo TEXT,
            fecha_vencimiento_sugerida TEXT,
            apto_sugerido INTEGER,
            contacto_emergencia_sugerido TEXT,
            telefono_emergencia_sugerido TEXT,
            jugador_id INTEGER,
            procesado_en TEXT,
            procesado_por TEXT,
            error TEXT,
            creado_en TEXT,
            creado_por TEXT,
            FOREIGN KEY (jugador_sugerido_id) REFERENCES jugadores (id),
            FOREIGN KEY (jugador_id) REFERENCES jugadores (id)
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_fichas_medicas_batch_batch_id
        ON fichas_medicas_batch (batch_id)
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_fichas_medicas_batch_estado
        ON fichas_medicas_batch (estado)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS lesiones (
            id SERIAL PRIMARY KEY,
            jugador_id INTEGER NOT NULL,
            fecha_lesion TEXT,
            tipo_lesion TEXT,
            zona_cuerpo TEXT,
            diagnostico TEXT,
            tratamiento TEXT,
            estado TEXT NOT NULL DEFAULT 'Activa',
            fecha_alta TEXT,
            observaciones TEXT,
            FOREIGN KEY (jugador_id) REFERENCES jugadores (id)
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS roles (
            id SERIAL PRIMARY KEY,
            nombre TEXT UNIQUE NOT NULL,
            descripcion TEXT,
            permisos TEXT NOT NULL DEFAULT '[]',
            sistema INTEGER DEFAULT 0
        )
    """)

    columnas_roles = get_columns(conn, "roles")
    if "descripcion" not in columnas_roles:
        conn.execute("ALTER TABLE roles ADD COLUMN descripcion TEXT")
    if "permisos" not in columnas_roles:
        conn.execute("ALTER TABLE roles ADD COLUMN permisos TEXT NOT NULL DEFAULT '[]'")
    if "sistema" not in columnas_roles:
        conn.execute("ALTER TABLE roles ADD COLUMN sistema INTEGER DEFAULT 0")

    roles_base = {
        "admin": "Acceso completo al sistema.",
        "tesorero": "Gestión financiera, cuotas, caja y reportes.",
        "medico": "Gestión de fichas médicas y lesiones.",
        "entrenador": "Gestión deportiva, jugadores, calendario y asistencia.",
    }

    for nombre_rol, descripcion in roles_base.items():
        conn.execute("""
            INSERT INTO roles (nombre, descripcion, permisos, sistema)
            VALUES (%s, %s, %s, 1)
            ON CONFLICT(nombre) DO UPDATE SET
                descripcion = EXCLUDED.descripcion,
                permisos = CASE
                    WHEN roles.sistema = 1 THEN EXCLUDED.permisos
                    ELSE roles.permisos
                END,
                sistema = CASE
                    WHEN roles.sistema = 1 THEN 1
                    ELSE roles.sistema
                END
        """, (
            nombre_rol,
            descripcion,
            serializar_permisos(permisos_default_rol(nombre_rol)),
        ))

    columnas_usuarios = get_columns(conn, "usuarios")
    if "rol" not in columnas_usuarios:
        conn.execute("ALTER TABLE usuarios ADD COLUMN rol TEXT DEFAULT 'admin'")
    columnas_usuarios = get_columns(conn, "usuarios")
    onboarding_columna_creada = "onboarding_visto" not in columnas_usuarios
    columnas_extra_usuarios = {
        "email": "TEXT",
        "debe_cambiar_password": "INTEGER DEFAULT 0",
        "onboarding_visto": "INTEGER DEFAULT 0",
        "creado_en": "TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP",
        "ultimo_login": "TIMESTAMPTZ",
    }
    for columna, tipo_columna in columnas_extra_usuarios.items():
        if columna not in columnas_usuarios:
            conn.execute(f"ALTER TABLE usuarios ADD COLUMN {columna} {tipo_columna}")

    if onboarding_columna_creada:
        conn.execute("""
            UPDATE usuarios
            SET onboarding_visto = 1
            WHERE onboarding_visto IS NULL
               OR onboarding_visto = 0
        """)

    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_usuarios_email
        ON usuarios (email)
        WHERE email IS NOT NULL AND email <> ''
    """)

    usuario_admin = conn.execute("""
        SELECT * FROM usuarios WHERE username = 'admin'
    """).fetchone()

    if not usuario_admin:
        admin_password = ADMIN_PASSWORD or "admin123"
        conn.execute("""
    INSERT INTO usuarios (username, password, rol, debe_cambiar_password, onboarding_visto)
    VALUES (%s, %s, %s, %s, %s)
""", ("admin", generate_password_hash(admin_password), "admin", 1, 0))
    elif ADMIN_PASSWORD and (
        FORCE_ADMIN_PASSWORD_UPDATE
        or check_password_hash(usuario_admin["password"], "admin123")
    ):
        conn.execute("""
            UPDATE usuarios
            SET password = %s, rol = 'admin'
            WHERE username = 'admin'
        """, (generate_password_hash(ADMIN_PASSWORD),))

    conn.execute("""
        CREATE TABLE IF NOT EXISTS login_attempts (
            id SERIAL PRIMARY KEY,
            username TEXT,
            ip TEXT,
            success INTEGER DEFAULT 0,
            fecha TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_login_attempts_user_ip_fecha
        ON login_attempts (username, ip, fecha DESC)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS password_reset_tokens (
            id SERIAL PRIMARY KEY,
            usuario_id INTEGER NOT NULL,
            token_hash TEXT NOT NULL,
            creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
            expira_en TIMESTAMPTZ NOT NULL,
            usado INTEGER DEFAULT 0,
            usado_en TIMESTAMPTZ,
            FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_password_reset_tokens_usuario
        ON password_reset_tokens (usuario_id, usado, expira_en DESC)
    """)

    conn.execute("""
        DO $$
        BEGIN
            IF EXISTS (
                SELECT 1
                FROM information_schema.table_constraints
                WHERE constraint_schema = 'public'
                  AND table_name = 'password_reset_tokens'
                  AND constraint_name = 'password_reset_tokens_usuario_id_fkey'
                  AND constraint_type = 'FOREIGN KEY'
            ) THEN
                ALTER TABLE password_reset_tokens
                DROP CONSTRAINT password_reset_tokens_usuario_id_fkey;
            END IF;

            ALTER TABLE password_reset_tokens
            ADD CONSTRAINT password_reset_tokens_usuario_id_fkey
            FOREIGN KEY (usuario_id)
            REFERENCES usuarios(id)
            ON DELETE CASCADE;
        END $$;
    """)

    conn.execute("""
    CREATE TABLE IF NOT EXISTS movimientos (
        id SERIAL PRIMARY KEY,
        tipo TEXT NOT NULL,                -- ingreso / egreso
        concepto TEXT,
        monto REAL NOT NULL,
        fecha TEXT NOT NULL,
        referencia TEXT,
        anulado INTEGER DEFAULT 0,
        fecha_anulacion TEXT,
        usuario_anulacion TEXT,
        motivo_anulacion TEXT
    )
""")

    columnas_movimientos = get_columns(conn, "movimientos")
    if "anulado" not in columnas_movimientos:
        conn.execute("ALTER TABLE movimientos ADD COLUMN anulado INTEGER DEFAULT 0")
    if "fecha_anulacion" not in columnas_movimientos:
        conn.execute("ALTER TABLE movimientos ADD COLUMN fecha_anulacion TEXT")
    if "usuario_anulacion" not in columnas_movimientos:
        conn.execute("ALTER TABLE movimientos ADD COLUMN usuario_anulacion TEXT")
    if "motivo_anulacion" not in columnas_movimientos:
        conn.execute("ALTER TABLE movimientos ADD COLUMN motivo_anulacion TEXT")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS cierres_mensuales (
            id SERIAL PRIMARY KEY,
            mes TEXT UNIQUE NOT NULL,
            ingresos REAL NOT NULL,
            egresos REAL NOT NULL,
            resultado REAL NOT NULL,
            fecha_cierre TEXT NOT NULL,
            usuario TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS eventos_asistencia (
            id SERIAL PRIMARY KEY,
            fecha TEXT NOT NULL,
            tipo TEXT NOT NULL,
            descripcion TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS calendario_eventos (
            id SERIAL PRIMARY KEY,
            fecha TEXT NOT NULL,
            tipo TEXT NOT NULL,
            titulo TEXT NOT NULL,
            descripcion TEXT,
            ubicacion TEXT,
            categoria TEXT
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_calendario_eventos_fecha
        ON calendario_eventos (fecha)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS asistencias (
            id SERIAL PRIMARY KEY,
            evento_id INTEGER NOT NULL,
            jugador_id INTEGER NOT NULL,
            presente INTEGER DEFAULT 0,
            observaciones TEXT,
            FOREIGN KEY (evento_id) REFERENCES eventos_asistencia(id),
            FOREIGN KEY (jugador_id) REFERENCES jugadores(id),
            UNIQUE(evento_id, jugador_id)
        )
    """)

    columnas_asistencias = get_columns(conn, "asistencias")
    if "estado_asistencia" not in columnas_asistencias:
        conn.execute("ALTER TABLE asistencias ADD COLUMN estado_asistencia TEXT")

    conn.execute("""
        UPDATE asistencias
        SET estado_asistencia = CASE
            WHEN presente = 1 THEN 'a_tiempo'
            ELSE 'ausente'
        END
        WHERE estado_asistencia IS NULL
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS aspirante_asistencias (
            id SERIAL PRIMARY KEY,
            evento_id INTEGER NOT NULL,
            aspirante_id INTEGER NOT NULL,
            presente INTEGER DEFAULT 0,
            estado_asistencia TEXT,
            observaciones TEXT,
            FOREIGN KEY (evento_id) REFERENCES eventos_asistencia(id),
            FOREIGN KEY (aspirante_id) REFERENCES aspirantes(id),
            UNIQUE(evento_id, aspirante_id)
        )
    """)

    columnas_aspirante_asistencias = get_columns(conn, "aspirante_asistencias")
    if "estado_asistencia" not in columnas_aspirante_asistencias:
        conn.execute("ALTER TABLE aspirante_asistencias ADD COLUMN estado_asistencia TEXT")

    conn.execute("""
        UPDATE aspirante_asistencias
        SET estado_asistencia = CASE
            WHEN presente = 1 THEN 'a_tiempo'
            ELSE 'ausente'
        END
        WHERE estado_asistencia IS NULL
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS test_tipos (
            id SERIAL PRIMARY KEY,
            nombre TEXT UNIQUE NOT NULL,
            descripcion TEXT,
            unidad TEXT,
            puntaje_min REAL,
            puntaje_max REAL,
            mayor_es_mejor INTEGER DEFAULT 1,
            activo INTEGER DEFAULT 1,
            creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
            creado_por TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS test_resultados (
            id SERIAL PRIMARY KEY,
            test_id INTEGER NOT NULL,
            jugador_id INTEGER NOT NULL,
            fecha TEXT NOT NULL,
            puntaje REAL NOT NULL,
            observaciones TEXT,
            creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
            creado_por TEXT,
            FOREIGN KEY (test_id) REFERENCES test_tipos(id),
            FOREIGN KEY (jugador_id) REFERENCES jugadores(id)
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_test_resultados_test_fecha
        ON test_resultados (test_id, fecha DESC)
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_test_resultados_jugador
        ON test_resultados (jugador_id, fecha DESC)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS test_importaciones_batch (
            id SERIAL PRIMARY KEY,
            batch_id TEXT NOT NULL,
            estado TEXT DEFAULT 'pendiente',
            fila INTEGER,
            test_id INTEGER,
            test_nombre TEXT,
            jugador_sugerido_id INTEGER,
            jugador_id INTEGER,
            confianza TEXT,
            motivo TEXT,
            nombre_excel TEXT,
            apellido_excel TEXT,
            nombre_completo_excel TEXT,
            fecha TEXT,
            puntaje REAL,
            observaciones TEXT,
            error TEXT,
            creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
            creado_por TEXT,
            procesado_en TIMESTAMPTZ,
            procesado_por TEXT,
            FOREIGN KEY (test_id) REFERENCES test_tipos(id),
            FOREIGN KEY (jugador_sugerido_id) REFERENCES jugadores(id),
            FOREIGN KEY (jugador_id) REFERENCES jugadores(id)
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_test_importaciones_batch_id
        ON test_importaciones_batch (batch_id)
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_test_importaciones_estado
        ON test_importaciones_batch (estado)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS auditoria (
            id SERIAL PRIMARY KEY,
            fecha TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
            usuario_id INTEGER,
            username TEXT,
            rol TEXT,
            accion TEXT NOT NULL,
            entidad TEXT,
            entidad_id TEXT,
            detalle TEXT,
            ip TEXT,
            user_agent TEXT
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_auditoria_fecha
        ON auditoria (fecha DESC)
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_auditoria_usuario
        ON auditoria (username)
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_auditoria_entidad
        ON auditoria (entidad)
    """)

    conn.commit()
    conn.close()

@app.route("/caja")
def ver_caja():
    check = permiso_requerido("caja_ver")
    if check:
        return check

    mes = request.args.get("mes")
    mes_actual = mes or datetime.now().strftime("%Y-%m")

    conn = get_connection()

    movimientos = conn.execute("""
        SELECT *
        FROM movimientos
        ORDER BY fecha DESC, id DESC
        LIMIT 50
    """).fetchall()

    total_ingresos = conn.execute("""
        SELECT COALESCE(SUM(monto), 0) AS total
        FROM movimientos
        WHERE tipo = 'ingreso'
          AND COALESCE(anulado, 0) = 0
    """).fetchone()["total"]

    total_egresos = conn.execute("""
        SELECT COALESCE(SUM(monto), 0) AS total
        FROM movimientos
        WHERE tipo = 'egreso'
          AND COALESCE(anulado, 0) = 0
    """).fetchone()["total"]

    ingresos_mes = conn.execute("""
        SELECT COALESCE(SUM(monto), 0) AS total
        FROM movimientos
        WHERE tipo = 'ingreso'
          AND COALESCE(anulado, 0) = 0
          AND substring(fecha from 1 for 7) = %s
    """, (mes_actual,)).fetchone()["total"]

    egresos_mes = conn.execute("""
        SELECT COALESCE(SUM(monto), 0) AS total
        FROM movimientos
        WHERE tipo = 'egreso'
          AND COALESCE(anulado, 0) = 0
          AND substring(fecha from 1 for 7) = %s
    """, (mes_actual,)).fetchone()["total"]

    movimientos_mes = conn.execute("""
        SELECT *
        FROM movimientos
        WHERE substring(fecha from 1 for 7) = %s
        ORDER BY fecha DESC, id DESC
    """, (mes_actual,)).fetchall()

    cierre_mes = conn.execute("""
        SELECT *
        FROM cierres_mensuales
        WHERE mes = %s
    """, (mes_actual,)).fetchone()

    conn.close()

    saldo = total_ingresos - total_egresos
    resultado_mes = ingresos_mes - egresos_mes

    return render_template(
        "caja.html",
        movimientos=movimientos,
        total_ingresos=total_ingresos,
        total_egresos=total_egresos,
        saldo=saldo,
        mes_actual=mes_actual,
        ingresos_mes=ingresos_mes,
        egresos_mes=egresos_mes,
        resultado_mes=resultado_mes,
        movimientos_mes=movimientos_mes,
        cierre_mes=cierre_mes
    )

@app.route("/movimientos/nuevo", methods=["GET", "POST"])
def nuevo_movimiento():
    check = permiso_requerido("caja_gestionar")
    if check:
        return check

    if request.method == "POST":
        tipo = request.form.get("tipo")
        concepto = request.form.get("concepto")
        monto = request.form.get("monto")
        fecha = validar_fecha_movimiento(request.form.get("fecha", "").strip())
        referencia = request.form.get("referencia")

        if not fecha:
            flash("La fecha del movimiento no es válida.", "error")
            return render_template(
                "movimiento_form.html",
                movimiento={
                    "tipo": tipo,
                    "concepto": concepto,
                    "monto": monto,
                    "fecha": request.form.get("fecha", "").strip(),
                    "referencia": referencia,
                },
            )

        mes_movimiento = fecha[:7]
        if mes_esta_cerrado(mes_movimiento):
            flash("No se puede agregar un movimiento en un mes cerrado.", "error")
            return render_template(
                "movimiento_form.html",
                movimiento={
                    "tipo": tipo,
                    "concepto": concepto,
                    "monto": monto,
                    "fecha": fecha,
                    "referencia": referencia,
                },
            )

        conn = get_connection()

        conn.execute("""
            INSERT INTO movimientos (tipo, concepto, monto, fecha, referencia)
            VALUES (%s, %s, %s, %s, %s)
        """, (tipo, concepto, monto, fecha, referencia))

        conn.commit()
        conn.close()

        flash("Movimiento registrado.", "ok")
        return redirect(url_for("ver_caja", mes=mes_movimiento))

    fecha_default = fecha_movimiento_default(request.args.get("mes"))
    return render_template(
        "movimiento_form.html",
        movimiento={"tipo": "egreso", "fecha": fecha_default},
    )

@app.before_request
def proteger_rutas():

    rutas_publicas = {
        "login",
        "solicitar_recuperacion_password",
        "restablecer_password",
        "logout",
        "static",
        "portal_jugador",
        "portal_actualizar_contacto",
        "portal_subir_comprobante",
        "portal_descargar_recibo",
    }

    if request.method == "POST" and request.endpoint != "static":
        if not csrf_valido():
            abort(400)

    if request.endpoint in rutas_publicas:
        return

    if "user_id" not in session:
        return redirect(url_for("login"))

    if (
        session.get("debe_cambiar_password")
        and request.endpoint not in {"cambiar_mi_password", "logout", "static"}
    ):
        flash("Cambiá tu contraseña para continuar.", "warning")
        return redirect(url_for("cambiar_mi_password"))

    try:
        g.mantenimiento = obtener_config_mantenimiento()
    except Exception:
        app.logger.exception("No se pudo consultar el modo mantenimiento.")
        g.mantenimiento = {
            "activo": False,
            "mensaje": MAINTENANCE_DEFAULT_MESSAGE,
            "actualizado_en": None,
            "actualizado_por": None,
        }

    if g.mantenimiento["activo"] and session.get("rol") != "admin":
        return render_template("mantenimiento.html", mantenimiento=g.mantenimiento), 503


@app.after_request
def auditar_acciones(response):
    if request.method != "POST":
        return response

    if request.endpoint in {"login", "static"}:
        return response

    accion_base, entidad = AUDIT_ENDPOINTS.get(
        request.endpoint,
        (request.endpoint or request.path, "sistema"),
    )
    resultado = "ok" if response.status_code < 400 else "error"

    registrar_auditoria(
        accion=f"{accion_base}_{resultado}",
        entidad=entidad,
        entidad_id=audit_entity_id(),
        detalle={
            "endpoint": request.endpoint,
            "path": request.path,
            "method": request.method,
            "status_code": response.status_code,
            "form": sanitized_audit_form(),
            "route_args": request.view_args or {},
        },
    )
    return response


def rol_requerido(*roles_permitidos):
    rol = session.get("rol")
    if rol == "admin":
        return None

    if rol not in roles_permitidos:
        flash("No tenés permiso para acceder a esa sección.", "error")
        return redirect(url_for("index"))
    return None


def tiene_rol(*roles):
    rol = session.get("rol")
    return rol == "admin" or rol in roles


def tiene_permiso(*permisos):
    if session.get("rol") == "admin":
        return True
    permisos_usuario = set(session.get("permisos") or permisos_default_rol(session.get("rol")))
    return any(permiso in permisos_usuario for permiso in permisos)


def permiso_requerido(*permisos):
    if tiene_permiso(*permisos):
        return None
    flash("No tenes permiso para acceder a esa seccion.", "error")
    return redirect(url_for("index"))


def puede_ver_bitacora_tipo(tipo):
    if tipo == "general":
        return tiene_permiso("jugadores_ver")
    if tipo == "finanzas":
        return tiene_permiso("cuotas_ver", "caja_ver")
    if tipo == "salud":
        return tiene_permiso("salud_ver")
    if tipo == "deportivo":
        return tiene_permiso("asistencia_ver", "calendario_ver", "tests_ver")
    return False


def puede_crear_bitacora_tipo(tipo):
    if tipo == "general":
        return tiene_permiso("jugadores_gestionar")
    if tipo == "finanzas":
        return tiene_permiso("cuotas_gestionar", "caja_gestionar")
    if tipo == "salud":
        return tiene_permiso("salud_gestionar")
    if tipo == "deportivo":
        return tiene_permiso("asistencia_gestionar", "calendario_gestionar", "tests_gestionar")
    return False


def tipos_bitacora_disponibles():
    return [
        {"clave": clave, "nombre": nombre}
        for clave, nombre in BITACORA_TIPOS.items()
        if puede_crear_bitacora_tipo(clave)
    ]


def filtrar_bitacora_visible(items):
    return [item for item in items if puede_ver_bitacora_tipo(item["tipo"])]


def validar_password_nueva(password, confirmacion):
    if not password or not confirmacion:
        return "La contraseña y la confirmación son obligatorias."
    if password != confirmacion:
        return "La confirmación no coincide."
    if len(password) < 8:
        return "La contraseña debe tener al menos 8 caracteres."
    return None

def normalizar_username(username):
    return (username or "").strip().lower()


def normalizar_email(email):
    return (email or "").strip().lower()


def smtp_configurado():
    return bool(SMTP_HOST and SMTP_FROM)


def enviar_email(destinatario, asunto, cuerpo):
    if not smtp_configurado():
        return False

    mensaje = EmailMessage()
    mensaje["From"] = SMTP_FROM
    mensaje["To"] = destinatario
    mensaje["Subject"] = asunto
    mensaje.set_content(cuerpo)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=15) as smtp:
        if SMTP_USE_TLS:
            smtp.starttls()
        if SMTP_USER:
            smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(mensaje)
    return True


def crear_token_recuperacion(conn, usuario_id):
    token = secrets.token_urlsafe(32)
    token_hash = generate_password_hash(token)
    conn.execute("""
        INSERT INTO password_reset_tokens (
            usuario_id, token_hash, creado_en, expira_en, usado
        )
        VALUES (
            %s,
            %s,
            CURRENT_TIMESTAMP,
            CURRENT_TIMESTAMP + (%s || ' minutes')::interval,
            0
        )
    """, (usuario_id, token_hash, str(PASSWORD_RESET_TOKEN_MINUTES)))
    return token


def buscar_token_recuperacion(conn, token):
    tokens = conn.execute("""
        SELECT
            t.*,
            u.username,
            u.email
        FROM password_reset_tokens t
        JOIN usuarios u ON u.id = t.usuario_id
        WHERE t.usado = 0
          AND t.expira_en > CURRENT_TIMESTAMP
        ORDER BY t.creado_en DESC
        LIMIT 100
    """).fetchall()

    for fila in tokens:
        if check_password_hash(fila["token_hash"], token):
            return fila
    return None


def login_bloqueado(conn, username, ip):
    intentos = conn.execute("""
        SELECT COUNT(*) AS total
        FROM login_attempts
        WHERE success = 0
          AND username = %s
          AND ip = %s
          AND fecha >= CURRENT_TIMESTAMP - (%s || ' minutes')::interval
    """, (normalizar_username(username), ip, str(LOGIN_ATTEMPT_WINDOW_MINUTES))).fetchone()
    return (intentos["total"] or 0) >= MAX_LOGIN_ATTEMPTS


def registrar_intento_login(conn, username, ip, success):
    normalized_username = normalizar_username(username)
    conn.execute("""
        INSERT INTO login_attempts (username, ip, success)
        VALUES (%s, %s, %s)
    """, (normalized_username, ip, 1 if success else 0))

    if success:
        conn.execute("""
            DELETE FROM login_attempts
            WHERE username = %s
              AND ip = %s
              AND success = 0
        """, (normalized_username, ip))
    else:
        conn.execute("""
            DELETE FROM login_attempts
            WHERE fecha < CURRENT_TIMESTAMP - INTERVAL '7 days'
        """)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        ip = audit_request_ip()

        conn = get_connection()
        if login_bloqueado(conn, username, ip):
            conn.close()
            registrar_auditoria(
                "login_bloqueado",
                "usuario",
                None,
                {"username": username, "ip": ip},
                username=username,
            )
            flash("Demasiados intentos fallidos. Proba nuevamente en unos minutos.", "error")
            return render_template("login.html")

        usuario = conn.execute("""
            SELECT u.*, r.permisos AS rol_permisos
            FROM usuarios u
            LEFT JOIN roles r ON r.nombre = u.rol
            WHERE lower(u.username) = %s
        """, (normalizar_username(username),)).fetchone()

        if usuario and check_password_hash(usuario["password"], password):
            registrar_intento_login(conn, username, ip, True)
            conn.execute("""
                UPDATE usuarios
                SET ultimo_login = CURRENT_TIMESTAMP
                WHERE id = %s
            """, (usuario["id"],))
            conn.commit()
            conn.close()

            session.clear()
            session.permanent = True
            csrf_token()
            session["user_id"] = usuario["id"]
            session["username"] = usuario["username"]
            session["rol"] = usuario["rol"]
            session["debe_cambiar_password"] = bool(usuario.get("debe_cambiar_password"))
            session["onboarding_visto"] = bool(usuario.get("onboarding_visto"))
            session["permisos"] = deserializar_permisos(
                usuario["rol_permisos"],
                usuario["rol"],
            )
            registrar_auditoria(
                "login_ok",
                "usuario",
                str(usuario["id"]),
                {"username": username},
                usuario_id=usuario["id"],
                username=usuario["username"],
                rol=usuario["rol"],
            )
            if session["debe_cambiar_password"]:
                flash("Por seguridad, cambiá tu contraseña para continuar.", "warning")
                return redirect(url_for("cambiar_mi_password"))
            return redirect(url_for("index"))
        else:
            registrar_intento_login(conn, username, ip, False)
            conn.commit()
            conn.close()
            registrar_auditoria(
                "login_error",
                "usuario",
                None,
                {"username": username},
                username=username,
            )
            flash("Usuario o contraseña incorrectos", "error")

    return render_template("login.html")


@app.route("/password/recuperar", methods=["GET", "POST"])
def solicitar_recuperacion_password():
    if request.method == "POST":
        email = normalizar_email(request.form.get("email", ""))

        if email:
            conn = get_connection()
            usuario = conn.execute("""
                SELECT id, username, email
                FROM usuarios
                WHERE lower(email) = %s
            """, (email,)).fetchone()

            if usuario:
                token = crear_token_recuperacion(conn, usuario["id"])
                reset_url = url_for("restablecer_password", token=token, _external=True)
                enviado = False
                try:
                    enviado = enviar_email(
                        usuario["email"],
                        "Recuperar clave - SIG Ruda Macho",
                        (
                            f"Hola {usuario['username']},\n\n"
                            "Recibimos un pedido para restablecer tu clave del SIG.\n"
                            f"El enlace vence en {PASSWORD_RESET_TOKEN_MINUTES} minutos:\n\n"
                            f"{reset_url}\n\n"
                            "Si no pediste este cambio, podés ignorar este mensaje."
                        ),
                    )
                except Exception:
                    app.logger.exception("No se pudo enviar email de recuperacion a %s.", email)
                conn.commit()
                conn.close()

                registrar_auditoria(
                    "password_reset_solicitado",
                    "usuario",
                    str(usuario["id"]),
                    {"email": email, "email_enviado": enviado},
                    username=usuario["username"],
                )
            else:
                conn.close()

        flash("Si el email está registrado, vas a recibir un enlace de recuperación.", "ok")
        return redirect(url_for("login"))

    return render_template("password_recuperar.html")


@app.route("/password/restablecer/<token>", methods=["GET", "POST"])
def restablecer_password(token):
    conn = get_connection()
    token_row = buscar_token_recuperacion(conn, token)
    if token_row is None:
        conn.close()
        flash("El enlace de recuperación no es válido o venció.", "error")
        return redirect(url_for("login"))

    usuario = {
        "id": token_row["usuario_id"],
        "username": token_row["username"],
        "email": token_row["email"],
    }

    if request.method == "POST":
        password_nueva = request.form.get("password_nueva", "")
        password_confirmacion = request.form.get("password_confirmacion", "")

        error = validar_password_nueva(password_nueva, password_confirmacion)
        if error:
            conn.close()
            flash(error, "error")
            return render_template("password_reset_form.html", token=token, usuario=usuario)

        conn.execute("""
            UPDATE usuarios
            SET password = %s,
                debe_cambiar_password = 0
            WHERE id = %s
        """, (generate_password_hash(password_nueva), usuario["id"]))
        conn.execute("""
            UPDATE password_reset_tokens
            SET usado = 1,
                usado_en = CURRENT_TIMESTAMP
            WHERE id = %s
        """, (token_row["id"],))
        conn.commit()
        conn.close()

        registrar_auditoria(
            "password_reset_ok",
            "usuario",
            str(usuario["id"]),
            {"origen": "recuperacion_email"},
            username=usuario["username"],
        )
        flash("Contraseña actualizada. Ya podés iniciar sesión.", "ok")
        return redirect(url_for("login"))

    conn.close()
    return render_template("password_reset_form.html", token=token, usuario=usuario)


def login_required():
    if "user_id" not in session:
        return redirect(url_for("login"))

@app.route("/logout")
def logout():
    registrar_auditoria(
        "logout",
        "usuario",
        str(session.get("user_id")) if session.get("user_id") else None,
        {"username": session.get("username")},
    )
    session.clear()
    return redirect(url_for("login"))


@app.route("/mi-cuenta/password", methods=["GET", "POST"])
def cambiar_mi_password():
    usuario_id = session.get("user_id")
    conn = get_connection()
    usuario = conn.execute("""
        SELECT id, username, password, rol, debe_cambiar_password
        FROM usuarios
        WHERE id = %s
    """, (usuario_id,)).fetchone()

    if usuario is None:
        conn.close()
        session.clear()
        flash("Tu usuario ya no existe. Iniciá sesión nuevamente.", "error")
        return redirect(url_for("login"))

    if request.method == "POST":
        password_actual = request.form.get("password_actual", "")
        password_nueva = request.form.get("password_nueva", "")
        password_confirmacion = request.form.get("password_confirmacion", "")

        error = validar_password_nueva(password_nueva, password_confirmacion)
        if error:
            conn.close()
            flash(error, "error")
            return render_template("password_form.html", usuario=usuario, modo="propio")

        if not usuario["debe_cambiar_password"] and not check_password_hash(usuario["password"], password_actual):
            conn.close()
            flash("La contraseña actual no es correcta.", "error")
            return render_template("password_form.html", usuario=usuario, modo="propio")

        conn.execute("""
            UPDATE usuarios
            SET password = %s,
                debe_cambiar_password = 0
            WHERE id = %s
        """, (generate_password_hash(password_nueva), usuario_id))
        conn.commit()
        conn.close()

        session["debe_cambiar_password"] = False
        flash("Contraseña actualizada correctamente.", "ok")
        return redirect(url_for("index"))

    conn.close()
    return render_template("password_form.html", usuario=usuario, modo="propio")


@app.route("/mi-cuenta/onboarding", methods=["POST"])
def descartar_onboarding():
    conn = get_connection()
    conn.execute("""
        UPDATE usuarios
        SET onboarding_visto = 1
        WHERE id = %s
    """, (session.get("user_id"),))
    conn.commit()
    conn.close()
    session["onboarding_visto"] = True
    return redirect(destino_interno(request.form.get("next")))


@app.route("/")
def index():
    mes = request.args.get("mes")  # formato YYYY-MM
    mes_actual = mes or datetime.now().strftime("%Y-%m")

    conn = get_connection()

    total_jugadores = conn.execute("""
        SELECT COUNT(*) AS total
        FROM jugadores
    """).fetchone()["total"]

    jugadores_con_deuda = conn.execute("""
        SELECT
            j.id,
            j.nombre,
            j.apellido,
            COALESCE(SUM(c.importe), 0) AS deuda,
            SUM(
                CASE
                    WHEN c.fecha_vencimiento IS NOT NULL
                     AND c.fecha_vencimiento <> ''
                     AND c.fecha_vencimiento::date < CURRENT_DATE
                    THEN 1
                    ELSE 0
                END
            ) AS cuotas_vencidas
        FROM jugadores j
        JOIN cuotas c ON j.id = c.jugador_id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
        GROUP BY j.id, j.nombre, j.apellido
        HAVING COALESCE(SUM(c.importe), 0) > 0
        ORDER BY deuda DESC, j.apellido, j.nombre
    """).fetchall()

    fichas_vencidas = conn.execute("""
        SELECT
            j.id,
            j.nombre,
            j.apellido,
            f.fecha_vencimiento
        FROM jugadores j
        JOIN fichas_medicas f ON j.id = f.jugador_id
        WHERE f.fecha_vencimiento IS NOT NULL
          AND f.fecha_vencimiento <> ''
          AND f.fecha_vencimiento::date < CURRENT_DATE
        ORDER BY f.fecha_vencimiento ASC, j.apellido, j.nombre
    """).fetchall()

    lesiones_activas = conn.execute("""
        SELECT
            l.id,
            j.id AS jugador_id,
            j.nombre,
            j.apellido,
            l.fecha_lesion,
            l.tipo_lesion,
            l.zona_cuerpo,
            l.estado
        FROM lesiones l
        JOIN jugadores j ON j.id = l.jugador_id
        WHERE l.estado IN ('Activa', 'En recuperación')
        ORDER BY
            CASE
                WHEN l.estado = 'Activa' THEN 0
                ELSE 1
            END,
            l.fecha_lesion DESC
    """).fetchall()

    total_recaudado_mes = conn.execute("""
        SELECT COALESCE(SUM(importe), 0) AS total
        FROM cuotas
        WHERE pagado = 1
           AND substring(fecha_pago from 1 for 7) = %s
    """, (mes_actual,)).fetchone()["total"]

    deuda_total = conn.execute("""
        SELECT COALESCE(SUM(importe), 0) AS total
        FROM cuotas
        WHERE pagado = 0
        AND COALESCE(importe, 0) > 0
    """).fetchone()["total"]

    cuotas_pagadas_mes = conn.execute("""
    SELECT COUNT(*) AS total
    FROM cuotas
    WHERE pagado = 1
      AND substring(fecha_pago from 1 for 7) = %s
    """, (mes_actual,)).fetchone()["total"]

    cuotas_pendientes = conn.execute("""
    SELECT COUNT(*) AS total
    FROM cuotas
    WHERE pagado = 0
      AND COALESCE(importe, 0) > 0
      AND substring(periodo from 1 for 7) = %s
    """, (mes_actual,)).fetchone()["total"]

    cuotas_pendientes_lista = conn.execute("""
    SELECT
        c.id,
        c.periodo,
        c.importe,
        c.fecha_vencimiento,
        j.id AS jugador_id,
        j.nombre,
        j.apellido,
        j.categoria
    FROM cuotas c
    JOIN jugadores j ON j.id = c.jugador_id
    WHERE c.pagado = 0
      AND COALESCE(c.importe, 0) > 0
    ORDER BY
        CASE
            WHEN c.fecha_vencimiento IS NOT NULL
             AND c.fecha_vencimiento <> ''
             AND c.fecha_vencimiento::date < CURRENT_DATE
            THEN 0
            ELSE 1
        END,
        c.fecha_vencimiento ASC,
        j.apellido,
        j.nombre
    LIMIT 20
    """).fetchall()

    comprobantes_pendientes_lista = conn.execute("""
        SELECT
            c.id,
            c.jugador_id,
            c.periodo,
            c.importe,
            c.comprobante_fecha,
            c.comprobante_usuario,
            j.nombre,
            j.apellido
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.comprobante_drive_file_id IS NOT NULL
          AND COALESCE(NULLIF(c.comprobante_estado, ''), 'sin_comprobante') IN ('pendiente', 'sin_comprobante')
        ORDER BY c.comprobante_fecha DESC NULLS LAST, c.id DESC
        LIMIT 10
    """).fetchall()

    comprobantes_pendientes_count = conn.execute("""
        SELECT COUNT(*) AS total
        FROM cuotas
        WHERE comprobante_drive_file_id IS NOT NULL
          AND COALESCE(NULLIF(comprobante_estado, ''), 'sin_comprobante') IN ('pendiente', 'sin_comprobante')
    """).fetchone()["total"]

    conn.close()

    return render_template(
    "dashboard.html",
    total_jugadores=total_jugadores,
    jugadores_con_deuda=jugadores_con_deuda,
    fichas_vencidas=fichas_vencidas,
    lesiones_activas=lesiones_activas,
    total_recaudado_mes=total_recaudado_mes,
    deuda_total=deuda_total,
    cuotas_pagadas_mes=cuotas_pagadas_mes,
    cuotas_pendientes=cuotas_pendientes,
    cuotas_pendientes_lista=cuotas_pendientes_lista,
    comprobantes_pendientes_count=comprobantes_pendientes_count,
    comprobantes_pendientes_lista=comprobantes_pendientes_lista,
    mes_actual=mes_actual,
    puede_ver_jugadores=tiene_permiso("jugadores_ver"),
    puede_ver_finanzas=tiene_permiso("cuotas_ver", "cuotas_gestionar"),
    puede_ver_salud=tiene_permiso("salud_ver")
)

@app.route("/jugadores/<int:jugador_id>/ficha-medica")
def ver_ficha_medica(jugador_id):
    check = permiso_requerido("salud_ver")
    if check:
        return check

    conn = get_connection()

    jugador = conn.execute(
        "SELECT * FROM jugadores WHERE id = %s",
        (jugador_id,)
    ).fetchone()

    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    ficha = conn.execute("""
        SELECT * FROM fichas_medicas
        WHERE jugador_id = %s
    """, (jugador_id,)).fetchone()

    conn.close()

    return render_template(
        "ficha_medica.html",
        jugador=jugador,
        ficha=ficha
    )


@app.route("/jugadores/<int:jugador_id>/ficha-medica/documento")
def descargar_ficha_medica_documento(jugador_id):
    check = permiso_requerido("salud_ver")
    if check:
        return check

    conn = get_connection()
    jugador = conn.execute(
        "SELECT * FROM jugadores WHERE id = %s",
        (jugador_id,)
    ).fetchone()

    ficha = conn.execute("""
        SELECT * FROM fichas_medicas
        WHERE jugador_id = %s
    """, (jugador_id,)).fetchone()
    conn.close()

    if jugador is None:
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    if not ficha or not ficha["documento_drive_file_id"]:
        flash("La ficha médica no tiene documento adjunto.", "error")
        return redirect(url_for("ver_ficha_medica", jugador_id=jugador_id))

    try:
        archivo = descargar_drive_file(ficha["documento_drive_file_id"])
    except RuntimeError as error:
        flash(str(error), "error")
        return redirect(url_for("ver_ficha_medica", jugador_id=jugador_id))
    except Exception as error:
        app.logger.exception("No se pudo descargar ficha médica del jugador %s.", jugador_id)
        flash(mensaje_error_drive(error, carpeta=DRIVE_FICHAS_MEDICAS_SUBFOLDER, accion="descargar la ficha médica"), "error")
        return redirect(url_for("ver_ficha_medica", jugador_id=jugador_id))

    registrar_auditoria(
        "descargar_ok",
        "ficha_medica_documento",
        str(jugador_id),
        {
            "archivo": ficha["documento_nombre"],
            "drive_file_id": ficha["documento_drive_file_id"],
        },
    )

    return send_file(
        archivo,
        mimetype=ficha["documento_mime_type"] or "application/octet-stream",
        as_attachment=False,
        download_name=ficha["documento_nombre"] or f"ficha_medica_{jugador_id}",
    )


@app.route("/jugadores/<int:jugador_id>/ficha-medica/editar", methods=["GET", "POST"])
def editar_ficha_medica(jugador_id):
    check = permiso_requerido("salud_gestionar")
    if check:
        return check

    conn = get_connection()

    jugador = conn.execute(
        "SELECT * FROM jugadores WHERE id = %s",
        (jugador_id,)
    ).fetchone()

    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    ficha = conn.execute("""
        SELECT * FROM fichas_medicas
        WHERE jugador_id = %s
    """, (jugador_id,)).fetchone()

    if request.method == "POST":
        presentada = 1 if request.form.get("presentada") == "on" else 0
        fecha_vencimiento = request.form.get("fecha_vencimiento", "").strip()
        apto_fisico = 1 if request.form.get("apto_fisico") == "on" else 0
        contacto_emergencia = request.form.get("contacto_emergencia", "").strip()
        telefono_emergencia = request.form.get("telefono_emergencia", "").strip()
        observaciones = request.form.get("observaciones", "").strip()
        procesar_ocr = request.form.get("procesar_ocr") == "on"
        archivo_ficha = request.files.get("ficha_archivo")
        documento_info = None
        documento_fecha = None
        documento_usuario = None
        ocr_texto = None
        ocr_fecha = None
        ocr_usuario = None

        ficha_form = {
            "presentada": presentada,
            "fecha_vencimiento": fecha_vencimiento,
            "apto_fisico": apto_fisico,
            "contacto_emergencia": contacto_emergencia,
            "telefono_emergencia": telefono_emergencia,
            "observaciones": observaciones,
            "documento_drive_file_id": ficha["documento_drive_file_id"] if ficha else None,
            "documento_nombre": ficha["documento_nombre"] if ficha else None,
            "documento_fecha": ficha["documento_fecha"] if ficha else None,
            "documento_usuario": ficha["documento_usuario"] if ficha else None,
            "ocr_texto": ficha["ocr_texto"] if ficha else None,
            "ocr_fecha": ficha["ocr_fecha"] if ficha else None,
            "ocr_usuario": ficha["ocr_usuario"] if ficha else None,
        }

        try:
            ficha_validada = validar_ficha_medica_upload(archivo_ficha)
        except ValueError as error:
            conn.close()
            flash(str(error), "error")
            return render_template(
                "ficha_medica_form.html",
                jugador=jugador,
                ficha=ficha_form,
            )

        if ficha_validada:
            try:
                documento_info = subir_ficha_medica_a_drive(ficha_validada, jugador, ficha)
            except RuntimeError as error:
                conn.close()
                flash(str(error), "error")
                return render_template(
                    "ficha_medica_form.html",
                    jugador=jugador,
                    ficha=ficha_form,
                )
            except Exception as error:
                conn.close()
                app.logger.exception("No se pudo subir ficha médica del jugador %s.", jugador_id)
                flash(mensaje_error_drive(error, carpeta=DRIVE_FICHAS_MEDICAS_SUBFOLDER, accion="guardar la ficha médica"), "error")
                return render_template(
                    "ficha_medica_form.html",
                    jugador=jugador,
                    ficha=ficha_form,
                )

            presentada = 1
            documento_fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            documento_usuario = session.get("username")

            if procesar_ocr:
                try:
                    ocr_texto = normalizar_ocr_texto(
                        extraer_texto_ocr_drive(
                            ficha_validada,
                            jugador,
                            documento_info.get("folder_id"),
                        )
                    )
                    if ocr_texto:
                        datos_ocr = datos_ficha_desde_ocr(ocr_texto)
                        if not fecha_vencimiento and datos_ocr.get("fecha_vencimiento"):
                            fecha_vencimiento = datos_ocr["fecha_vencimiento"]
                        if datos_ocr.get("apto_fisico") is not None:
                            apto_fisico = datos_ocr["apto_fisico"]
                        if not contacto_emergencia and datos_ocr.get("contacto_emergencia"):
                            contacto_emergencia = datos_ocr["contacto_emergencia"]
                        if not telefono_emergencia and datos_ocr.get("telefono_emergencia"):
                            telefono_emergencia = datos_ocr["telefono_emergencia"]
                        ocr_fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        ocr_usuario = session.get("username")
                    else:
                        flash("El documento se guardó, pero OCR no devolvió texto para completar campos.", "warning")
                except Exception:
                    app.logger.exception("No se pudo procesar OCR de ficha médica del jugador %s.", jugador_id)
                    flash(
                        "El documento se guardó, pero no se pudo procesar OCR. Revisá los datos manualmente.",
                        "warning",
                    )

        if ficha:
            conn.execute("""
                UPDATE fichas_medicas
                SET presentada = %s, fecha_vencimiento = %s, apto_fisico = %s,
                    contacto_emergencia = %s, telefono_emergencia = %s, observaciones = %s,
                    documento_drive_file_id = COALESCE(%s, documento_drive_file_id),
                    documento_nombre = COALESCE(%s, documento_nombre),
                    documento_mime_type = COALESCE(%s, documento_mime_type),
                    documento_tamano = COALESCE(%s, documento_tamano),
                    documento_fecha = COALESCE(%s, documento_fecha),
                    documento_usuario = COALESCE(%s, documento_usuario),
                    documento_web_url = COALESCE(%s, documento_web_url),
                    ocr_texto = COALESCE(%s, ocr_texto),
                    ocr_fecha = COALESCE(%s, ocr_fecha),
                    ocr_usuario = COALESCE(%s, ocr_usuario)
                WHERE jugador_id = %s
            """, (
                presentada, fecha_vencimiento, apto_fisico,
                contacto_emergencia, telefono_emergencia, observaciones,
                documento_info["file_id"] if documento_info else None,
                documento_info["nombre"] if documento_info else None,
                documento_info["mime_type"] if documento_info else None,
                documento_info["tamano"] if documento_info else None,
                documento_fecha,
                documento_usuario,
                documento_info["web_url"] if documento_info else None,
                ocr_texto,
                ocr_fecha,
                ocr_usuario,
                jugador_id
            ))
        else:
            conn.execute("""
                INSERT INTO fichas_medicas (
                    jugador_id, presentada, fecha_vencimiento, apto_fisico,
                    contacto_emergencia, telefono_emergencia, observaciones,
                    documento_drive_file_id, documento_nombre, documento_mime_type,
                    documento_tamano, documento_fecha, documento_usuario, documento_web_url,
                    ocr_texto, ocr_fecha, ocr_usuario
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                jugador_id, presentada, fecha_vencimiento, apto_fisico,
                contacto_emergencia, telefono_emergencia, observaciones,
                documento_info["file_id"] if documento_info else None,
                documento_info["nombre"] if documento_info else None,
                documento_info["mime_type"] if documento_info else None,
                documento_info["tamano"] if documento_info else None,
                documento_fecha,
                documento_usuario,
                documento_info["web_url"] if documento_info else None,
                ocr_texto,
                ocr_fecha,
                ocr_usuario,
            ))

        conn.commit()
        conn.close()

        flash("Ficha médica guardada correctamente.", "ok")
        if documento_info:
            flash("Documento de ficha médica guardado en Google Drive.", "ok")
        if ocr_texto:
            flash("OCR procesado y campos detectados aplicados.", "ok")
        return redirect(url_for("ver_ficha_medica", jugador_id=jugador_id))

    conn.close()
    return render_template(
        "ficha_medica_form.html",
        jugador=jugador,
        ficha=ficha
    )


@app.route("/fichas-medicas/batch", methods=["GET", "POST"])
def cargar_fichas_medicas_batch():
    check = permiso_requerido("salud_gestionar")
    if check:
        return check

    if request.method == "POST":
        archivos = [
            archivo
            for archivo in request.files.getlist("fichas_archivos")
            if archivo and archivo.filename
        ]
        if not archivos:
            flash("Seleccioná al menos una ficha médica para cargar.", "error")
            conn = get_connection()
            batches_recientes = obtener_fichas_medicas_batch_recientes(conn)
            conn.close()
            return render_template("fichas_medicas_batch.html", batches_recientes=batches_recientes)

        conn = get_connection()
        jugadores = obtener_jugadores_selector(conn)
        batch_id = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{secrets.token_urlsafe(6)}"
        cargadas = 0
        errores = 0

        for archivo in archivos:
            archivo_original = secure_filename(archivo.filename) or archivo.filename
            try:
                validado = validar_ficha_medica_upload(archivo)
                documento_info = subir_ficha_medica_batch_pendiente(validado, batch_id)
                ocr_texto = ""
                ocr_fecha = None
                ocr_usuario = None
                error = None

                datos_ocr = datos_ficha_desde_ocr(ocr_texto)
                jugador_sugerido, confianza, motivo = sugerir_jugador_ficha_ocr(archivo_original, jugadores)
                if confianza == "sin_coincidencia":
                    error = "OCR pendiente. Procesalo desde la revisión o asigná manualmente."

                conn.execute("""
                    INSERT INTO fichas_medicas_batch (
                        batch_id, estado, archivo_original, drive_file_id, drive_folder_id,
                        documento_nombre, documento_mime_type, documento_tamano, documento_web_url,
                        extension, ocr_texto, ocr_fecha, ocr_usuario, jugador_sugerido_id,
                        confianza, motivo, fecha_vencimiento_sugerida, apto_sugerido,
                        contacto_emergencia_sugerido, telefono_emergencia_sugerido,
                        error, creado_en, creado_por
                    )
                    VALUES (%s, 'pendiente', %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    batch_id,
                    archivo_original,
                    documento_info["file_id"],
                    documento_info["folder_id"],
                    documento_info["nombre"],
                    documento_info["mime_type"],
                    documento_info["tamano"],
                    documento_info["web_url"],
                    validado["ext"],
                    ocr_texto or None,
                    ocr_fecha,
                    ocr_usuario,
                    jugador_sugerido["id"] if jugador_sugerido else None,
                    confianza,
                    motivo,
                    datos_ocr.get("fecha_vencimiento") or None,
                    datos_ocr.get("apto_fisico"),
                    datos_ocr.get("contacto_emergencia") or None,
                    datos_ocr.get("telefono_emergencia") or None,
                    error,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    session.get("username"),
                ))
                conn.commit()
                cargadas += 1
            except ValueError as error:
                errores += 1
                flash(f"{archivo_original}: {error}", "error")
            except RuntimeError as error:
                errores += 1
                flash(str(error), "error")
            except Exception as error:
                conn.rollback()
                errores += 1
                app.logger.exception("No se pudo cargar ficha médica batch %s.", archivo_original)
                flash(mensaje_error_drive(error, carpeta=DRIVE_FICHAS_MEDICAS_SUBFOLDER, accion="guardar la ficha médica"), "error")

        conn.commit()

        if cargadas:
            conn.close()
            flash(f"Se cargaron {cargadas} ficha(s) para revisar.", "ok")
            flash("Para evitar timeouts, el OCR se procesa por archivo desde la pantalla de revisión.", "warning")
            if errores:
                flash(f"{errores} archivo(s) no pudieron cargarse.", "warning")
            return redirect(url_for("revisar_fichas_medicas_batch", batch_id=batch_id))

        flash("No se pudo cargar ninguna ficha médica.", "error")
        batches_recientes = obtener_fichas_medicas_batch_recientes(conn)
        conn.close()
        return render_template("fichas_medicas_batch.html", batches_recientes=batches_recientes)

    conn = get_connection()
    batches_recientes = obtener_fichas_medicas_batch_recientes(conn)
    conn.close()
    return render_template("fichas_medicas_batch.html", batches_recientes=batches_recientes)


@app.route("/fichas-medicas/batch/<batch_id>/revisar", methods=["GET", "POST"])
def revisar_fichas_medicas_batch(batch_id):
    check = permiso_requerido("salud_gestionar")
    if check:
        return check

    conn = get_connection()

    if request.method == "POST":
        item_ids = request.form.getlist("item_ids")
        procesadas = 0
        omitidas = 0
        errores = 0

        for item_id in item_ids:
            if request.form.get(f"procesar_{item_id}") != "on":
                omitidas += 1
                continue

            jugador_id = request.form.get(f"jugador_id_{item_id}", "").strip()
            fecha_vencimiento = request.form.get(f"fecha_vencimiento_{item_id}", "").strip()
            apto_fisico = 1 if request.form.get(f"apto_fisico_{item_id}") == "on" else 0
            contacto_emergencia = request.form.get(f"contacto_emergencia_{item_id}", "").strip()
            telefono_emergencia = request.form.get(f"telefono_emergencia_{item_id}", "").strip()
            observaciones = request.form.get(f"observaciones_{item_id}", "").strip()

            if not jugador_id:
                errores += 1
                flash(f"El archivo #{item_id} no tiene jugador asignado.", "error")
                continue

            if fecha_vencimiento and not validar_fecha_movimiento(fecha_vencimiento):
                errores += 1
                flash(f"El archivo #{item_id} tiene una fecha de vencimiento inválida.", "error")
                continue

            item = conn.execute("""
                SELECT *
                FROM fichas_medicas_batch
                WHERE id = %s AND batch_id = %s AND estado = 'pendiente'
            """, (item_id, batch_id)).fetchone()

            jugador = conn.execute("""
                SELECT *
                FROM jugadores
                WHERE id = %s
            """, (jugador_id,)).fetchone()

            if not item or not jugador:
                errores += 1
                flash(f"No se encontró el archivo pendiente #{item_id} o el jugador asignado.", "error")
                continue

            try:
                documento_info = mover_ficha_medica_batch_a_jugador(
                    item["drive_file_id"],
                    jugador,
                    item["extension"],
                    item["drive_folder_id"],
                )

                conn.execute("""
                    INSERT INTO fichas_medicas (
                        jugador_id, presentada, fecha_vencimiento, apto_fisico,
                        contacto_emergencia, telefono_emergencia, observaciones,
                        documento_drive_file_id, documento_nombre, documento_mime_type,
                        documento_tamano, documento_fecha, documento_usuario, documento_web_url,
                        ocr_texto, ocr_fecha, ocr_usuario
                    )
                    VALUES (%s, 1, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (jugador_id) DO UPDATE SET
                        presentada = 1,
                        fecha_vencimiento = COALESCE(NULLIF(EXCLUDED.fecha_vencimiento, ''), fichas_medicas.fecha_vencimiento),
                        apto_fisico = EXCLUDED.apto_fisico,
                        contacto_emergencia = COALESCE(NULLIF(EXCLUDED.contacto_emergencia, ''), fichas_medicas.contacto_emergencia),
                        telefono_emergencia = COALESCE(NULLIF(EXCLUDED.telefono_emergencia, ''), fichas_medicas.telefono_emergencia),
                        observaciones = COALESCE(NULLIF(EXCLUDED.observaciones, ''), fichas_medicas.observaciones),
                        documento_drive_file_id = EXCLUDED.documento_drive_file_id,
                        documento_nombre = EXCLUDED.documento_nombre,
                        documento_mime_type = EXCLUDED.documento_mime_type,
                        documento_tamano = EXCLUDED.documento_tamano,
                        documento_fecha = EXCLUDED.documento_fecha,
                        documento_usuario = EXCLUDED.documento_usuario,
                        documento_web_url = EXCLUDED.documento_web_url,
                        ocr_texto = COALESCE(EXCLUDED.ocr_texto, fichas_medicas.ocr_texto),
                        ocr_fecha = COALESCE(EXCLUDED.ocr_fecha, fichas_medicas.ocr_fecha),
                        ocr_usuario = COALESCE(EXCLUDED.ocr_usuario, fichas_medicas.ocr_usuario)
                """, (
                    jugador["id"],
                    fecha_vencimiento or None,
                    apto_fisico,
                    contacto_emergencia or None,
                    telefono_emergencia or None,
                    observaciones or None,
                    documento_info["file_id"],
                    documento_info["nombre"],
                    documento_info["mime_type"] or item["documento_mime_type"],
                    documento_info["tamano"] or item["documento_tamano"],
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    session.get("username"),
                    documento_info["web_url"],
                    item["ocr_texto"],
                    item["ocr_fecha"],
                    item["ocr_usuario"],
                ))

                conn.execute("""
                    UPDATE fichas_medicas_batch
                    SET estado = 'procesado',
                        jugador_id = %s,
                        fecha_vencimiento_sugerida = %s,
                        apto_sugerido = %s,
                        contacto_emergencia_sugerido = %s,
                        telefono_emergencia_sugerido = %s,
                        procesado_en = %s,
                        procesado_por = %s,
                        error = NULL
                    WHERE id = %s
                """, (
                    jugador["id"],
                    fecha_vencimiento or None,
                    apto_fisico,
                    contacto_emergencia or None,
                    telefono_emergencia or None,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    session.get("username"),
                    item["id"],
                ))
                conn.commit()
                procesadas += 1
            except Exception as error:
                conn.rollback()
                errores += 1
                app.logger.exception("No se pudo confirmar ficha médica batch item %s.", item_id)
                flash(mensaje_error_drive(error, carpeta=DRIVE_FICHAS_MEDICAS_SUBFOLDER, accion="asignar la ficha médica"), "error")

        if procesadas:
            flash(f"Se asignaron {procesadas} ficha(s) médica(s).", "ok")
        if omitidas:
            flash(f"{omitidas} archivo(s) quedaron pendientes.", "warning")
        if errores:
            flash(f"{errores} archivo(s) requieren revisión.", "error")
        return redirect(url_for("revisar_fichas_medicas_batch", batch_id=batch_id))

    items = conn.execute("""
        SELECT
            b.*,
            js.apellido AS sugerido_apellido,
            js.nombre AS sugerido_nombre,
            js.dni AS sugerido_dni,
            ja.apellido AS asignado_apellido,
            ja.nombre AS asignado_nombre
        FROM fichas_medicas_batch b
        LEFT JOIN jugadores js ON js.id = b.jugador_sugerido_id
        LEFT JOIN jugadores ja ON ja.id = b.jugador_id
        WHERE b.batch_id = %s
        ORDER BY b.id
    """, (batch_id,)).fetchall()

    jugadores = obtener_jugadores_selector(conn)
    conn.close()

    if not items:
        flash("No se encontró la tanda de fichas médicas.", "error")
        return redirect(url_for("cargar_fichas_medicas_batch"))

    return render_template(
        "fichas_medicas_batch_revision.html",
        batch_id=batch_id,
        items=items,
        jugadores=jugadores,
    )


@app.route("/fichas-medicas/batch/items/<int:item_id>/ocr", methods=["POST"])
def procesar_ficha_medica_batch_ocr(item_id):
    check = permiso_requerido("salud_gestionar")
    if check:
        return check

    conn = get_connection()
    item = conn.execute("""
        SELECT *
        FROM fichas_medicas_batch
        WHERE id = %s
    """, (item_id,)).fetchone()

    if not item:
        conn.close()
        flash("No se encontró el archivo pendiente.", "error")
        return redirect(url_for("cargar_fichas_medicas_batch"))

    if item["estado"] != "pendiente":
        conn.close()
        flash("La ficha ya fue procesada.", "warning")
        return redirect(url_for("revisar_fichas_medicas_batch", batch_id=item["batch_id"]))

    try:
        ocr_texto = procesar_ocr_ficha_medica_batch_item(conn, item)
        conn.commit()
        if ocr_texto:
            flash("OCR procesado para el archivo seleccionado.", "ok")
        else:
            flash("OCR no devolvió texto para este archivo. Podés asignarlo manualmente.", "warning")
    except Exception as error:
        conn.rollback()
        app.logger.exception("No se pudo procesar OCR batch item %s.", item_id)
        conn.execute("""
            UPDATE fichas_medicas_batch
            SET error = %s
            WHERE id = %s
        """, (f"OCR no disponible: {truncate_audit_value(error, 160)}", item_id))
        conn.commit()
        flash("No se pudo procesar OCR para este archivo. Podés cargar los datos manualmente.", "warning")
    finally:
        conn.close()

    return redirect(url_for("revisar_fichas_medicas_batch", batch_id=item["batch_id"]))


@app.route("/fichas-medicas/batch/items/<int:item_id>/documento")
def descargar_ficha_medica_batch_documento(item_id):
    check = permiso_requerido("salud_gestionar")
    if check:
        return check

    conn = get_connection()
    item = conn.execute("""
        SELECT *
        FROM fichas_medicas_batch
        WHERE id = %s
    """, (item_id,)).fetchone()
    conn.close()

    if not item or not item["drive_file_id"]:
        flash("No se encontró el documento pendiente.", "error")
        return redirect(url_for("cargar_fichas_medicas_batch"))

    try:
        archivo = descargar_drive_file(item["drive_file_id"])
    except Exception as error:
        app.logger.exception("No se pudo descargar ficha médica batch item %s.", item_id)
        flash(mensaje_error_drive(error, carpeta=DRIVE_FICHAS_MEDICAS_SUBFOLDER, accion="descargar la ficha médica"), "error")
        return redirect(url_for("revisar_fichas_medicas_batch", batch_id=item["batch_id"]))

    return send_file(
        archivo,
        mimetype=item["documento_mime_type"] or "application/octet-stream",
        as_attachment=False,
        download_name=item["archivo_original"] or item["documento_nombre"] or f"ficha_medica_batch_{item_id}",
    )


@app.route("/documentos")
def ver_documentos_vencidos():
    check = permiso_requerido("documentos_ver", "salud_ver")
    if check:
        return check

    conn = get_connection()
    manuales = conn.execute("""
        SELECT
            d.*,
            j.apellido,
            j.nombre,
            j.categoria
        FROM documentos_jugadores d
        JOIN jugadores j ON j.id = d.jugador_id
        WHERE d.fecha_vencimiento IS NOT NULL
          AND d.fecha_vencimiento <> ''
          AND d.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
          AND d.fecha_vencimiento::date <= CURRENT_DATE + INTERVAL '30 days'
        ORDER BY d.fecha_vencimiento ASC, j.apellido, j.nombre
    """).fetchall()

    fichas = conn.execute("""
        SELECT
            j.id AS jugador_id,
            j.apellido,
            j.nombre,
            j.categoria,
            f.fecha_vencimiento,
            f.presentada,
            CASE
                WHEN f.id IS NULL THEN 'faltante'
                WHEN f.fecha_vencimiento IS NULL OR f.fecha_vencimiento = '' THEN 'sin_vencimiento'
                WHEN f.fecha_vencimiento !~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' THEN 'sin_vencimiento'
                WHEN f.fecha_vencimiento::date < CURRENT_DATE THEN 'vencida'
                WHEN f.fecha_vencimiento::date <= CURRENT_DATE + INTERVAL '30 days' THEN 'por_vencer'
                ELSE 'vigente'
            END AS estado_documento
        FROM jugadores j
        LEFT JOIN fichas_medicas f ON f.jugador_id = j.id
        WHERE j.estado = 'Activo'
          AND (
              f.id IS NULL
              OR f.fecha_vencimiento IS NULL
              OR f.fecha_vencimiento = ''
              OR (
                  f.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                  AND f.fecha_vencimiento::date <= CURRENT_DATE + INTERVAL '30 days'
              )
          )
        ORDER BY estado_documento, f.fecha_vencimiento ASC NULLS FIRST, j.apellido, j.nombre
    """).fetchall()
    conn.close()

    return render_template("documentos.html", manuales=manuales, fichas=fichas)


@app.route("/jugadores/<int:jugador_id>/documentos/nuevo", methods=["GET", "POST"])
def nuevo_documento_jugador(jugador_id):
    check = permiso_requerido("documentos_gestionar", "salud_gestionar")
    if check:
        return check

    conn = get_connection()
    jugador = conn.execute("SELECT * FROM jugadores WHERE id = %s", (jugador_id,)).fetchone()
    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    if request.method == "POST":
        data = {
            "tipo": request.form.get("tipo", "").strip(),
            "nombre": request.form.get("nombre", "").strip(),
            "fecha_presentacion": request.form.get("fecha_presentacion", "").strip(),
            "fecha_vencimiento": request.form.get("fecha_vencimiento", "").strip(),
            "url": request.form.get("url", "").strip(),
            "observaciones": request.form.get("observaciones", "").strip(),
        }

        if not data["tipo"]:
            conn.close()
            flash("El tipo de documento es obligatorio.", "error")
            return render_template("documento_form.html", jugador=jugador, documento=data)

        if data["fecha_presentacion"] and not validar_fecha_movimiento(data["fecha_presentacion"]):
            conn.close()
            flash("La fecha de presentacion no es valida.", "error")
            return render_template("documento_form.html", jugador=jugador, documento=data)

        if data["fecha_vencimiento"] and not validar_fecha_movimiento(data["fecha_vencimiento"]):
            conn.close()
            flash("La fecha de vencimiento no es valida.", "error")
            return render_template("documento_form.html", jugador=jugador, documento=data)

        conn.execute("""
            INSERT INTO documentos_jugadores (
                jugador_id, tipo, nombre, fecha_presentacion, fecha_vencimiento,
                url, observaciones, creado_por
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            jugador_id,
            data["tipo"],
            data["nombre"],
            data["fecha_presentacion"] or None,
            data["fecha_vencimiento"] or None,
            data["url"],
            data["observaciones"],
            session.get("username"),
        ))
        conn.commit()
        conn.close()

        flash("Documento cargado correctamente.", "ok")
        return redirect(url_for("detalle_jugador", jugador_id=jugador_id))

    conn.close()
    return render_template("documento_form.html", jugador=jugador, documento={})


@app.route("/documentos/<int:documento_id>/eliminar", methods=["POST"])
def eliminar_documento_jugador(documento_id):
    check = permiso_requerido("documentos_gestionar", "salud_gestionar")
    if check:
        return check

    conn = get_connection()
    documento = conn.execute(
        "SELECT * FROM documentos_jugadores WHERE id = %s",
        (documento_id,),
    ).fetchone()
    if documento is None:
        conn.close()
        flash("Documento no encontrado.", "error")
        return redirect(url_for("ver_documentos_vencidos"))

    conn.execute("DELETE FROM documentos_jugadores WHERE id = %s", (documento_id,))
    conn.commit()
    conn.close()

    flash("Documento eliminado.", "ok")
    return redirect(url_for("detalle_jugador", jugador_id=documento["jugador_id"]))


@app.route("/jugadores")
def listar_jugadores():
    check = permiso_requerido("jugadores_ver")
    if check:
        return check

    busqueda = request.args.get("q", "").strip()
    conn = get_connection()

    if busqueda:
        terminos = [termino for termino in re.split(r"\s+", busqueda) if termino]
        condiciones = []
        parametros = []

        for termino in terminos:
            like = f"%{termino}%"
            condiciones.append("""
                (
                    nombre ILIKE %s
                    OR apellido ILIKE %s
                    OR dni ILIKE %s
                    OR categoria ILIKE %s
                    OR telefono ILIKE %s
                    OR email ILIKE %s
                    OR estado ILIKE %s
                    OR concat_ws(' ', nombre, apellido) ILIKE %s
                    OR concat_ws(' ', apellido, nombre) ILIKE %s
                )
            """)
            parametros.extend([like] * 9)

        jugadores = conn.execute("""
            SELECT * FROM jugadores
            WHERE """ + " AND ".join(condiciones) + """
            ORDER BY apellido, nombre
        """, parametros).fetchall()
    else:
        jugadores = conn.execute("""
            SELECT * FROM jugadores
            ORDER BY apellido, nombre
        """).fetchall()

    conn.close()
    return render_template("jugadores.html", jugadores=jugadores, busqueda=busqueda)


def obtener_madrinas_disponibles(conn):
    return conn.execute("""
        SELECT id, nombre, apellido, categoria
        FROM jugadores
        WHERE estado = 'Activo'
        ORDER BY apellido, nombre
    """).fetchall()


def aspirante_desde_formulario(aspirante=None):
    objetivo_raw = request.form.get(
        "entrenamientos_objetivo",
        str(aspirante["entrenamientos_objetivo"] if aspirante else ASPIRANTE_ENTRENAMIENTOS_OBJETIVO),
    )
    try:
        entrenamientos_objetivo = max(0, int(objetivo_raw))
    except (TypeError, ValueError):
        entrenamientos_objetivo = ASPIRANTE_ENTRENAMIENTOS_OBJETIVO

    madrina_raw = request.form.get("madrina_jugador_id", "").strip()
    try:
        madrina_jugador_id = int(madrina_raw) if madrina_raw else None
    except ValueError:
        madrina_jugador_id = None

    estado = request.form.get("estado", "").strip() or (
        aspirante["estado"] if aspirante else "Aspirante"
    )
    if estado not in ASPIRANTE_ESTADOS:
        estado = "Aspirante"

    return {
        "nombre": request.form.get("nombre", "").strip(),
        "apellido": request.form.get("apellido", "").strip(),
        "dni": request.form.get("dni", "").strip(),
        "fecha_nacimiento": request.form.get("fecha_nacimiento", "").strip(),
        "telefono": request.form.get("telefono", "").strip(),
        "email": request.form.get("email", "").strip(),
        "categoria": request.form.get("categoria", "").strip(),
        "fecha_postulacion": request.form.get("fecha_postulacion", "").strip(),
        "estado": estado,
        "madrina_jugador_id": madrina_jugador_id,
        "entrenamientos_objetivo": entrenamientos_objetivo,
        "observaciones": request.form.get("observaciones", "").strip(),
    }


def aspirante_con_progreso(aspirante):
    fila = dict(aspirante)
    objetivo = fila.get("entrenamientos_objetivo") or ASPIRANTE_ENTRENAMIENTOS_OBJETIVO
    presentes = fila.get("entrenamientos_realizados") or 0
    fila["progreso"] = min(100, round((presentes / objetivo) * 100)) if objetivo else 100
    fila["listo_para_ingresar"] = (
        fila.get("estado") == "Aspirante"
        and presentes >= objetivo
    )
    return fila


def buscar_aspirante(conn, aspirante_id):
    return conn.execute("""
        SELECT
            a.*,
            m.nombre AS madrina_nombre,
            m.apellido AS madrina_apellido,
            m.categoria AS madrina_categoria,
            COALESCE(stats.entrenamientos_realizados, 0) AS entrenamientos_realizados
        FROM aspirantes a
        LEFT JOIN jugadores m ON m.id = a.madrina_jugador_id
        LEFT JOIN (
            SELECT
                aa.aspirante_id,
                SUM(CASE WHEN aa.presente = 1 AND LOWER(e.tipo) = 'entrenamiento' THEN 1 ELSE 0 END) AS entrenamientos_realizados
            FROM aspirante_asistencias aa
            JOIN eventos_asistencia e ON e.id = aa.evento_id
            GROUP BY aa.aspirante_id
        ) stats ON stats.aspirante_id = a.id
        WHERE a.id = %s
    """, (aspirante_id,)).fetchone()


@app.route("/ahijadxs")
def listar_aspirantes():
    check = permiso_requerido("aspirantes_ver")
    if check:
        return check

    busqueda = request.args.get("q", "").strip()
    estado = request.args.get("estado", "Aspirante").strip()
    if estado not in ASPIRANTE_ESTADOS and estado != "todos":
        estado = "Aspirante"

    condiciones = []
    parametros = []
    if estado != "todos":
        condiciones.append("a.estado = %s")
        parametros.append(estado)

    if busqueda:
        terminos = [termino for termino in re.split(r"\s+", busqueda) if termino]
        for termino in terminos:
            like = f"%{termino}%"
            condiciones.append("""
                (
                    a.nombre ILIKE %s
                    OR a.apellido ILIKE %s
                    OR a.dni ILIKE %s
                    OR a.categoria ILIKE %s
                    OR concat_ws(' ', a.nombre, a.apellido) ILIKE %s
                    OR concat_ws(' ', a.apellido, a.nombre) ILIKE %s
                    OR m.nombre ILIKE %s
                    OR m.apellido ILIKE %s
                )
            """)
            parametros.extend([like] * 8)

    where_sql = "WHERE " + " AND ".join(condiciones) if condiciones else ""

    conn = get_connection()
    aspirantes = conn.execute(f"""
        SELECT
            a.*,
            m.nombre AS madrina_nombre,
            m.apellido AS madrina_apellido,
            COALESCE(stats.entrenamientos_realizados, 0) AS entrenamientos_realizados
        FROM aspirantes a
        LEFT JOIN jugadores m ON m.id = a.madrina_jugador_id
        LEFT JOIN (
            SELECT
                aa.aspirante_id,
                SUM(CASE WHEN aa.presente = 1 AND LOWER(e.tipo) = 'entrenamiento' THEN 1 ELSE 0 END) AS entrenamientos_realizados
            FROM aspirante_asistencias aa
            JOIN eventos_asistencia e ON e.id = aa.evento_id
            GROUP BY aa.aspirante_id
        ) stats ON stats.aspirante_id = a.id
        {where_sql}
        ORDER BY a.estado, a.apellido, a.nombre
    """, parametros).fetchall()
    conn.close()

    aspirantes = [aspirante_con_progreso(aspirante) for aspirante in aspirantes]

    return render_template(
        "aspirantes.html",
        aspirantes=aspirantes,
        busqueda=busqueda,
        estado=estado,
        estados=sorted(ASPIRANTE_ESTADOS),
    )


@app.route("/ahijadxs/nuevo", methods=["GET", "POST"])
def nuevo_aspirante():
    check = permiso_requerido("aspirantes_gestionar")
    if check:
        return check

    conn = get_connection()
    madrinas = obtener_madrinas_disponibles(conn)

    if request.method == "POST":
        data = aspirante_desde_formulario()
        if not data["fecha_postulacion"]:
            data["fecha_postulacion"] = datetime.now().strftime("%Y-%m-%d")

        if not data["nombre"] or not data["apellido"]:
            conn.close()
            flash("Nombre y apellido son obligatorios.", "error")
            return render_template("aspirante_form.html", aspirante=data, madrinas=madrinas, modo="nuevo")

        if data["dni"]:
            existente_jugador = conn.execute(
                "SELECT id FROM jugadores WHERE dni = %s",
                (data["dni"],),
            ).fetchone()
            existente_aspirante = conn.execute(
                "SELECT id FROM aspirantes WHERE dni = %s AND estado <> 'Baja'",
                (data["dni"],),
            ).fetchone()
            if existente_jugador or existente_aspirante:
                conn.close()
                flash("Ya existe un jugador o ahijadx activo con ese DNI.", "error")
                return render_template("aspirante_form.html", aspirante=data, madrinas=madrinas, modo="nuevo")

        conn.execute("""
            INSERT INTO aspirantes (
                nombre, apellido, dni, fecha_nacimiento, telefono, email, categoria,
                fecha_postulacion, estado, madrina_jugador_id, entrenamientos_objetivo,
                observaciones
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            data["nombre"], data["apellido"], data["dni"], data["fecha_nacimiento"],
            data["telefono"], data["email"], data["categoria"], data["fecha_postulacion"],
            data["estado"], data["madrina_jugador_id"], data["entrenamientos_objetivo"],
            data["observaciones"],
        ))
        conn.commit()
        conn.close()

        flash("Ahijadx cargado correctamente.", "ok")
        return redirect(url_for("listar_aspirantes"))

    conn.close()
    return render_template(
        "aspirante_form.html",
        aspirante={
            "fecha_postulacion": datetime.now().strftime("%Y-%m-%d"),
            "estado": "Aspirante",
            "entrenamientos_objetivo": ASPIRANTE_ENTRENAMIENTOS_OBJETIVO,
        },
        madrinas=madrinas,
        modo="nuevo",
    )


@app.route("/ahijadxs/<int:aspirante_id>")
def detalle_aspirante(aspirante_id):
    check = permiso_requerido("aspirantes_ver")
    if check:
        return check

    conn = get_connection()
    aspirante = buscar_aspirante(conn, aspirante_id)
    if aspirante is None:
        conn.close()
        flash("Ahijadx no encontrado.", "error")
        return redirect(url_for("listar_aspirantes"))

    asistencias = conn.execute("""
        SELECT aa.*, e.fecha, e.tipo, e.descripcion
        FROM aspirante_asistencias aa
        JOIN eventos_asistencia e ON e.id = aa.evento_id
        WHERE aa.aspirante_id = %s
        ORDER BY e.fecha DESC, e.id DESC
    """, (aspirante_id,)).fetchall()
    conn.close()

    return render_template(
        "aspirante_detalle.html",
        aspirante=aspirante_con_progreso(aspirante),
        asistencias=asistencias,
    )


@app.route("/ahijadxs/<int:aspirante_id>/editar", methods=["GET", "POST"])
def editar_aspirante(aspirante_id):
    check = permiso_requerido("aspirantes_gestionar")
    if check:
        return check

    conn = get_connection()
    aspirante = conn.execute("SELECT * FROM aspirantes WHERE id = %s", (aspirante_id,)).fetchone()
    if aspirante is None:
        conn.close()
        flash("Ahijadx no encontrado.", "error")
        return redirect(url_for("listar_aspirantes"))

    madrinas = obtener_madrinas_disponibles(conn)

    if request.method == "POST":
        data = aspirante_desde_formulario(aspirante)
        if not data["nombre"] or not data["apellido"]:
            conn.close()
            flash("Nombre y apellido son obligatorios.", "error")
            data["id"] = aspirante_id
            return render_template("aspirante_form.html", aspirante=data, madrinas=madrinas, modo="editar")

        if data["dni"]:
            jugador_actual_id = aspirante["jugador_id"] or 0
            existente_jugador = conn.execute(
                "SELECT id FROM jugadores WHERE dni = %s AND id <> %s",
                (data["dni"], jugador_actual_id),
            ).fetchone()
            existente_aspirante = conn.execute(
                "SELECT id FROM aspirantes WHERE dni = %s AND id <> %s AND estado <> 'Baja'",
                (data["dni"], aspirante_id),
            ).fetchone()
            if existente_jugador or existente_aspirante:
                conn.close()
                flash("Ya existe otro jugador o ahijadx activo con ese DNI.", "error")
                data["id"] = aspirante_id
                return render_template("aspirante_form.html", aspirante=data, madrinas=madrinas, modo="editar")

        conn.execute("""
            UPDATE aspirantes
            SET nombre = %s,
                apellido = %s,
                dni = %s,
                fecha_nacimiento = %s,
                telefono = %s,
                email = %s,
                categoria = %s,
                fecha_postulacion = %s,
                estado = %s,
                madrina_jugador_id = %s,
                entrenamientos_objetivo = %s,
                observaciones = %s
            WHERE id = %s
        """, (
            data["nombre"], data["apellido"], data["dni"], data["fecha_nacimiento"],
            data["telefono"], data["email"], data["categoria"], data["fecha_postulacion"],
            data["estado"], data["madrina_jugador_id"], data["entrenamientos_objetivo"],
            data["observaciones"], aspirante_id,
        ))
        conn.commit()
        conn.close()

        flash("Ahijadx actualizado correctamente.", "ok")
        return redirect(url_for("detalle_aspirante", aspirante_id=aspirante_id))

    conn.close()
    return render_template("aspirante_form.html", aspirante=aspirante, madrinas=madrinas, modo="editar")


@app.route("/ahijadxs/<int:aspirante_id>/convertir", methods=["POST"])
def convertir_aspirante(aspirante_id):
    check = permiso_requerido("aspirantes_gestionar")
    if check:
        return check

    conn = get_connection()
    aspirante = buscar_aspirante(conn, aspirante_id)
    if aspirante is None:
        conn.close()
        flash("Ahijadx no encontrado.", "error")
        return redirect(url_for("listar_aspirantes"))

    aspirante = aspirante_con_progreso(aspirante)
    if aspirante["estado"] != "Aspirante":
        conn.close()
        flash("Solo se pueden ingresar ahijadxs en seguimiento.", "error")
        return redirect(url_for("detalle_aspirante", aspirante_id=aspirante_id))

    if not aspirante["listo_para_ingresar"]:
        conn.close()
        flash("El ahijadx todavia no alcanzo la cantidad requerida de entrenamientos.", "error")
        return redirect(url_for("detalle_aspirante", aspirante_id=aspirante_id))

    if aspirante["dni"]:
        existente = conn.execute(
            "SELECT id FROM jugadores WHERE dni = %s",
            (aspirante["dni"],),
        ).fetchone()
        if existente:
            conn.close()
            flash("Ya existe un jugador con ese DNI.", "error")
            return redirect(url_for("detalle_aspirante", aspirante_id=aspirante_id))

    fecha_ingreso = request.form.get("fecha_ingreso", "").strip() or datetime.now().strftime("%Y-%m-%d")
    jugador_creado = conn.execute("""
        INSERT INTO jugadores (
            nombre, apellido, dni, fecha_nacimiento, telefono, email, categoria,
            fecha_ingreso, estado, observaciones
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'Activo', %s)
        RETURNING id
    """, (
        aspirante["nombre"], aspirante["apellido"], aspirante["dni"], aspirante["fecha_nacimiento"],
        aspirante["telefono"], aspirante["email"], aspirante["categoria"], fecha_ingreso,
        aspirante["observaciones"],
    )).fetchone()
    jugador_id = jugador_creado["id"]

    conn.execute("""
        UPDATE aspirantes
        SET estado = 'Ingresado',
            fecha_ingreso_club = %s,
            jugador_id = %s
        WHERE id = %s
    """, (fecha_ingreso, jugador_id, aspirante_id))
    conn.commit()
    conn.close()

    flash("Ahijadx ingresado como jugador activo.", "ok")
    return redirect(url_for("detalle_jugador", jugador_id=jugador_id))


@app.route("/ahijadxs/<int:aspirante_id>/eliminar", methods=["POST"])
def eliminar_aspirante(aspirante_id):
    check = permiso_requerido("aspirantes_gestionar")
    if check:
        return check

    conn = get_connection()
    conn.execute("""
        UPDATE aspirantes
        SET estado = 'Baja'
        WHERE id = %s
    """, (aspirante_id,))
    conn.commit()
    conn.close()

    flash("Ahijadx dado de baja.", "ok")
    return redirect(url_for("listar_aspirantes"))


@app.route("/jugadores/nuevo", methods=["GET", "POST"])
def nuevo_jugador():
    check = permiso_requerido("jugadores_gestionar")
    if check:
        return check

    if request.method == "POST":
        data = {
            "nombre": request.form.get("nombre", "").strip(),
            "apellido": request.form.get("apellido", "").strip(),
            "dni": request.form.get("dni", "").strip(),
            "fecha_nacimiento": request.form.get("fecha_nacimiento", "").strip(),
            "telefono": request.form.get("telefono", "").strip(),
            "email": request.form.get("email", "").strip(),
            "categoria": request.form.get("categoria", "").strip(),
            "fecha_ingreso": request.form.get("fecha_ingreso", "").strip(),
            "estado": request.form.get("estado", "").strip() or "Activo",
            "contacto_tutor": request.form.get("contacto_tutor", "").strip(),
            "parentesco_tutor": request.form.get("parentesco_tutor", "").strip(),
            "telefono_tutor": request.form.get("telefono_tutor", "").strip(),
            "email_tutor": request.form.get("email_tutor", "").strip(),
            "direccion": request.form.get("direccion", "").strip(),
            "obra_social": request.form.get("obra_social", "").strip(),
            "numero_socio": request.form.get("numero_socio", "").strip(),
            "documentos": request.form.get("documentos", "").strip(),
            "observaciones": request.form.get("observaciones", "").strip(),
        }
        beca_data, beca_error = datos_beca_form()
        data.update(beca_data or {
            "beca_activa": 1 if request.form.get("beca_activa") == "on" else 0,
            "beca_porcentaje": request.form.get("beca_porcentaje", "").strip(),
            "beca_desde": request.form.get("beca_desde", "").strip(),
            "beca_hasta": request.form.get("beca_hasta", "").strip(),
            "beca_motivo": request.form.get("beca_motivo", "").strip(),
        })

        if not data["nombre"] or not data["apellido"]:
            flash("Nombre y apellido son obligatorios.", "error")
            return render_template("jugador_form.html", jugador=data, modo="nuevo")
        if beca_error:
            flash(beca_error, "error")
            return render_template("jugador_form.html", jugador=data, modo="nuevo")

        conn = get_connection()
        creado = conn.execute("""
            INSERT INTO jugadores
            (
                nombre, apellido, dni, fecha_nacimiento, telefono, email, categoria,
                fecha_ingreso, estado, contacto_tutor, parentesco_tutor, telefono_tutor,
                email_tutor, direccion, obra_social, numero_socio, documentos, observaciones,
                beca_activa, beca_porcentaje, beca_desde, beca_hasta, beca_motivo
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id
        """, (
            data["nombre"], data["apellido"], data["dni"], data["fecha_nacimiento"],
            data["telefono"], data["email"], data["categoria"], data["fecha_ingreso"],
            data["estado"], data["contacto_tutor"], data["parentesco_tutor"],
            data["telefono_tutor"], data["email_tutor"], data["direccion"],
            data["obra_social"], data["numero_socio"], data["documentos"], data["observaciones"],
            data["beca_activa"], data["beca_porcentaje"], data["beca_desde"],
            data["beca_hasta"], data["beca_motivo"]
        )).fetchone()

        if data["beca_activa"]:
            registrar_historial_beca(
                conn,
                creado["id"],
                data,
                "alta",
                {"origen": "alta_jugador"},
            )
        conn.commit()
        conn.close()

        flash("Jugador cargado correctamente.", "ok")
        return redirect(url_for("listar_jugadores"))

    return render_template("jugador_form.html", jugador=None, modo="nuevo")


@app.route("/jugadores/importar", methods=["GET", "POST"])
def importar_jugadores():
    check = permiso_requerido("jugadores_gestionar")
    if check:
        return check

    resultado = None

    if request.method == "POST":
        archivo = request.files.get("archivo")

        if not archivo or not archivo.filename:
            flash("Debés seleccionar un archivo Excel.", "error")
            return render_template("importar_jugadores.html", resultado=resultado)

        if not archivo.filename.lower().endswith(".xlsx"):
            flash("El archivo debe ser .xlsx.", "error")
            return render_template("importar_jugadores.html", resultado=resultado)

        try:
            wb = load_workbook(archivo, read_only=True, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
        except Exception:
            flash("No se pudo leer el archivo Excel.", "error")
            return render_template("importar_jugadores.html", resultado=resultado)

        if not rows:
            flash("El archivo está vacío.", "error")
            return render_template("importar_jugadores.html", resultado=resultado)

        headers = [normalizar_header_excel(valor) for valor in rows[0]]
        creados = 0
        omitidos = 0
        errores = []

        conn = get_connection()

        for numero_fila, row in enumerate(rows[1:], start=2):
            if not any(limpiar_valor_excel(valor) for valor in row):
                continue

            data = mapear_fila_jugador(headers, row)

            if not data["nombre"] or not data["apellido"]:
                omitidos += 1
                errores.append(f"Fila {numero_fila}: falta nombre o apellido.")
                continue

            if data["dni"]:
                existente = conn.execute(
                    "SELECT id FROM jugadores WHERE dni = %s",
                    (data["dni"],)
                ).fetchone()
                if existente:
                    omitidos += 1
                    errores.append(f"Fila {numero_fila}: DNI ya existente ({data['dni']}).")
                    continue

            conn.execute("""
                INSERT INTO jugadores (
                    nombre, apellido, dni, fecha_nacimiento, telefono, email, categoria,
                    fecha_ingreso, estado, contacto_tutor, parentesco_tutor, telefono_tutor,
                    email_tutor, direccion, obra_social, numero_socio, documentos, observaciones
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                data["nombre"], data["apellido"], data["dni"], data["fecha_nacimiento"],
                data["telefono"], data["email"], data["categoria"], data["fecha_ingreso"],
                data["estado"], data["contacto_tutor"], data["parentesco_tutor"],
                data["telefono_tutor"], data["email_tutor"], data["direccion"],
                data["obra_social"], data["numero_socio"], data["documentos"],
                data["observaciones"],
            ))
            creados += 1

        conn.commit()
        conn.close()

        resultado = {
            "creados": creados,
            "omitidos": omitidos,
            "errores": errores[:30],
        }

        registrar_auditoria(
            "importar_ok",
            "jugadores",
            None,
            {
                "archivo": archivo.filename,
                "creados": creados,
                "omitidos": omitidos,
                "errores": len(errores),
            },
        )

        flash(f"Importación terminada. Creados: {creados}. Omitidos: {omitidos}.", "ok")

    return render_template("importar_jugadores.html", resultado=resultado)


@app.route("/jugadores/acciones-masivas", methods=["POST"])
def acciones_masivas_jugadores():
    check = permiso_requerido("jugadores_gestionar")
    if check:
        return check

    jugador_ids = request.form.getlist("jugador_ids")
    accion = request.form.get("accion", "").strip()
    nuevo_estado = request.form.get("estado", "").strip()
    nueva_categoria = request.form.get("categoria", "").strip()

    if not jugador_ids:
        flash("Seleccioná al menos un jugador.", "error")
        return redirect(url_for("listar_jugadores"))

    ids = []
    for valor in jugador_ids:
        try:
            ids.append(int(valor))
        except ValueError:
            continue

    if not ids:
        flash("La selección no es válida.", "error")
        return redirect(url_for("listar_jugadores"))

    conn = get_connection()

    if accion == "estado":
        if nuevo_estado not in {"Activo", "Inactivo", "Suspendido", "Baja"}:
            conn.close()
            flash("El estado seleccionado no es válido.", "error")
            return redirect(url_for("listar_jugadores"))

        conn.execute(
            "UPDATE jugadores SET estado = %s WHERE id = ANY(%s)",
            (nuevo_estado, ids)
        )
        detalle = {"accion": "estado", "estado": nuevo_estado, "cantidad": len(ids)}
        mensaje = f"Estado actualizado para {len(ids)} jugador(es)."
    elif accion == "categoria":
        if not nueva_categoria:
            conn.close()
            flash("Ingresá una categoría.", "error")
            return redirect(url_for("listar_jugadores"))

        conn.execute(
            "UPDATE jugadores SET categoria = %s WHERE id = ANY(%s)",
            (nueva_categoria, ids)
        )
        detalle = {"accion": "categoria", "categoria": nueva_categoria, "cantidad": len(ids)}
        mensaje = f"Categoría actualizada para {len(ids)} jugador(es)."
    else:
        conn.close()
        flash("La acción masiva no es válida.", "error")
        return redirect(url_for("listar_jugadores"))

    conn.commit()
    conn.close()

    registrar_auditoria("accion_masiva", "jugadores", None, detalle)

    flash(mensaje, "ok")
    return redirect(url_for("listar_jugadores"))


@app.route("/jugadores/<int:jugador_id>/editar", methods=["GET", "POST"])
def editar_jugador(jugador_id):
    check = permiso_requerido("jugadores_gestionar")
    if check:
        return check

    conn = get_connection()
    jugador = conn.execute("SELECT * FROM jugadores WHERE id = %s", (jugador_id,)).fetchone()

    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    if request.method == "POST":
        data = {
            "nombre": request.form.get("nombre", "").strip(),
            "apellido": request.form.get("apellido", "").strip(),
            "dni": request.form.get("dni", "").strip(),
            "fecha_nacimiento": request.form.get("fecha_nacimiento", "").strip(),
            "telefono": request.form.get("telefono", "").strip(),
            "email": request.form.get("email", "").strip(),
            "categoria": request.form.get("categoria", "").strip(),
            "fecha_ingreso": request.form.get("fecha_ingreso", "").strip(),
            "estado": request.form.get("estado", "").strip() or "Activo",
            "contacto_tutor": request.form.get("contacto_tutor", "").strip(),
            "parentesco_tutor": request.form.get("parentesco_tutor", "").strip(),
            "telefono_tutor": request.form.get("telefono_tutor", "").strip(),
            "email_tutor": request.form.get("email_tutor", "").strip(),
            "direccion": request.form.get("direccion", "").strip(),
            "obra_social": request.form.get("obra_social", "").strip(),
            "numero_socio": request.form.get("numero_socio", "").strip(),
            "documentos": request.form.get("documentos", "").strip(),
            "observaciones": request.form.get("observaciones", "").strip(),
        }
        beca_data, beca_error = datos_beca_form()
        data.update(beca_data or {
            "beca_activa": 1 if request.form.get("beca_activa") == "on" else 0,
            "beca_porcentaje": request.form.get("beca_porcentaje", "").strip(),
            "beca_desde": request.form.get("beca_desde", "").strip(),
            "beca_hasta": request.form.get("beca_hasta", "").strip(),
            "beca_motivo": request.form.get("beca_motivo", "").strip(),
        })

        if not data["nombre"] or not data["apellido"]:
            conn.close()
            flash("Nombre y apellido son obligatorios.", "error")
            jugador_dict = dict(data)
            jugador_dict["id"] = jugador_id
            return render_template("jugador_form.html", jugador=jugador_dict, modo="editar")
        if beca_error:
            conn.close()
            flash(beca_error, "error")
            jugador_dict = dict(data)
            jugador_dict["id"] = jugador_id
            return render_template("jugador_form.html", jugador=jugador_dict, modo="editar")

        registro_beca = beca_modificada(jugador, data)
        conn.execute("""
            UPDATE jugadores
            SET nombre = %s, apellido = %s, dni = %s, fecha_nacimiento = %s, telefono = %s,
                email = %s, categoria = %s, fecha_ingreso = %s, estado = %s,
                contacto_tutor = %s, parentesco_tutor = %s, telefono_tutor = %s,
                email_tutor = %s, direccion = %s, obra_social = %s, numero_socio = %s,
                documentos = %s, observaciones = %s,
                beca_activa = %s, beca_porcentaje = %s, beca_desde = %s,
                beca_hasta = %s, beca_motivo = %s
            WHERE id = %s
        """, (
            data["nombre"], data["apellido"], data["dni"], data["fecha_nacimiento"],
            data["telefono"], data["email"], data["categoria"], data["fecha_ingreso"],
            data["estado"], data["contacto_tutor"], data["parentesco_tutor"],
            data["telefono_tutor"], data["email_tutor"], data["direccion"],
            data["obra_social"], data["numero_socio"], data["documentos"],
            data["observaciones"], data["beca_activa"], data["beca_porcentaje"],
            data["beca_desde"], data["beca_hasta"], data["beca_motivo"], jugador_id
        ))
        if registro_beca:
            registrar_historial_beca(
                conn,
                jugador_id,
                data,
                "actualizacion",
                {"anterior": snapshot_beca(jugador), "nuevo": snapshot_beca(data)},
            )
        conn.commit()
        conn.close()

        flash("Jugador actualizado correctamente.", "ok")
        return redirect(url_for("listar_jugadores"))

    conn.close()
    return render_template("jugador_form.html", jugador=jugador, modo="editar")


@app.route("/jugadores/<int:jugador_id>/eliminar", methods=["POST"])
def eliminar_jugador(jugador_id):
    check = permiso_requerido("jugadores_eliminar")
    if check:
        return check

    conn = get_connection()
    conn.execute("DELETE FROM jugadores WHERE id = %s", (jugador_id,))
    conn.commit()
    conn.close()
    flash("Jugador eliminado.", "ok")
    return redirect(url_for("listar_jugadores"))

@app.route("/jugadores/<int:jugador_id>/cuotas")
def ver_cuotas(jugador_id):
    check = permiso_requerido("cuotas_ver", "cuotas_gestionar")
    if check:
        return check

    conn = get_connection()

    jugador = conn.execute(
        "SELECT * FROM jugadores WHERE id = %s",
        (jugador_id,)
    ).fetchone()

    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    cuotas = conn.execute("""
        SELECT * FROM cuotas
        WHERE jugador_id = %s
        ORDER BY periodo DESC
    """, (jugador_id,)).fetchall()

    deuda = conn.execute("""
        SELECT SUM(importe) AS total
        FROM cuotas
        WHERE jugador_id = %s AND pagado = 0 AND COALESCE(importe, 0) > 0
    """, (jugador_id,)).fetchone()["total"] or 0

    becas_historial = conn.execute("""
        SELECT *
        FROM becas_historial
        WHERE jugador_id = %s
        ORDER BY creado_en DESC, id DESC
        LIMIT 10
    """, (jugador_id,)).fetchall()

    planes_pago = conn.execute("""
        SELECT *
        FROM planes_pago
        WHERE jugador_id = %s
        ORDER BY
            CASE WHEN estado = 'Activo' THEN 0 ELSE 1 END,
            fecha_inicio DESC,
            id DESC
    """, (jugador_id,)).fetchall()

    conn.close()

    return render_template(
        "cuotas.html",
        jugador=jugador,
        cuotas=cuotas,
        deuda=deuda,
        becas_historial=becas_historial,
        planes_pago=planes_pago,
    )


@app.route("/jugadores/<int:jugador_id>/cuotas/recalcular-becas", methods=["POST"])
def recalcular_becas_jugador(jugador_id):
    check = permiso_requerido("cuotas_gestionar")
    if check:
        return check

    periodo_desde = validar_mes_beca(request.form.get("periodo_desde", ""))
    periodo_hasta = validar_mes_beca(request.form.get("periodo_hasta", ""))
    if periodo_desde is None or periodo_hasta is None:
        flash("Los periodos de recalculo deben tener formato YYYY-MM.", "error")
        return redirect(url_for("ver_cuotas", jugador_id=jugador_id))
    if periodo_desde and periodo_hasta and periodo_hasta < periodo_desde:
        flash("El periodo hasta no puede ser anterior al periodo desde.", "error")
        return redirect(url_for("ver_cuotas", jugador_id=jugador_id))

    conn = get_connection()
    jugador = conn.execute(
        "SELECT * FROM jugadores WHERE id = %s",
        (jugador_id,),
    ).fetchone()
    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    resultado = recalcular_cuotas_becadas(conn, jugador, periodo_desde or "", periodo_hasta or "")
    registrar_historial_beca(
        conn,
        jugador_id,
        jugador,
        "recalculo_cuotas",
        {
            "periodo_desde": periodo_desde or None,
            "periodo_hasta": periodo_hasta or None,
            **resultado,
        },
    )
    conn.commit()
    conn.close()

    flash(
        "Recalculo terminado. "
        f"Revisadas: {resultado['revisadas']}. "
        f"Actualizadas: {resultado['actualizadas']}. "
        f"Becas totales: {resultado['becas_totales']}. "
        f"Becas parciales: {resultado['becas_parciales']}.",
        "ok",
    )
    return redirect(url_for("ver_cuotas", jugador_id=jugador_id))


@app.route("/planes-pago")
def listar_planes_pago():
    check = permiso_requerido("planes_pago_ver")
    if check:
        return check

    estado = request.args.get("estado", "Activo").strip()
    if estado not in {"Activo", "Cumplido", "Caido", "todos"}:
        estado = "Activo"

    condiciones = []
    parametros = []
    if estado != "todos":
        condiciones.append("p.estado = %s")
        parametros.append(estado)

    where_sql = "WHERE " + " AND ".join(condiciones) if condiciones else ""

    conn = get_connection()
    planes = conn.execute(f"""
        SELECT
            p.*,
            j.apellido,
            j.nombre,
            j.categoria,
            COALESCE(deuda.deuda, 0) AS deuda_actual
        FROM planes_pago p
        JOIN jugadores j ON j.id = p.jugador_id
        LEFT JOIN (
            SELECT jugador_id, SUM(importe) AS deuda
            FROM cuotas
            WHERE pagado = 0
              AND COALESCE(importe, 0) > 0
            GROUP BY jugador_id
        ) deuda ON deuda.jugador_id = j.id
        {where_sql}
        ORDER BY
            CASE WHEN p.estado = 'Activo' THEN 0 ELSE 1 END,
            p.fecha_inicio DESC,
            j.apellido,
            j.nombre
    """, parametros).fetchall()
    conn.close()

    return render_template("planes_pago.html", planes=planes, estado=estado)


@app.route("/jugadores/<int:jugador_id>/planes/nuevo", methods=["GET", "POST"])
def nuevo_plan_pago(jugador_id):
    check = permiso_requerido("planes_pago_gestionar")
    if check:
        return check

    conn = get_connection()
    jugador = conn.execute("SELECT * FROM jugadores WHERE id = %s", (jugador_id,)).fetchone()
    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    deuda = conn.execute("""
        SELECT COALESCE(SUM(importe), 0) AS total
        FROM cuotas
        WHERE jugador_id = %s
          AND pagado = 0
          AND COALESCE(importe, 0) > 0
    """, (jugador_id,)).fetchone()["total"]

    if request.method == "POST":
        fecha_inicio = request.form.get("fecha_inicio", "").strip() or datetime.now().strftime("%Y-%m-%d")
        descripcion = request.form.get("descripcion", "").strip()
        observaciones = request.form.get("observaciones", "").strip()

        try:
            monto_total = float(request.form.get("monto_total", "0") or 0)
            cantidad_cuotas = max(1, int(request.form.get("cantidad_cuotas", "1") or 1))
        except ValueError:
            conn.close()
            flash("El monto y la cantidad de cuotas deben ser numericos.", "error")
            return render_template("plan_pago_form.html", jugador=jugador, deuda=deuda)

        if monto_total <= 0:
            conn.close()
            flash("El monto total del plan debe ser mayor a cero.", "error")
            return render_template("plan_pago_form.html", jugador=jugador, deuda=deuda)

        monto_cuota = round(monto_total / cantidad_cuotas, 2)
        conn.execute("""
            INSERT INTO planes_pago (
                jugador_id, fecha_inicio, monto_total, cantidad_cuotas, monto_cuota,
                estado, descripcion, observaciones, creado_por
            )
            VALUES (%s, %s, %s, %s, %s, 'Activo', %s, %s, %s)
        """, (
            jugador_id,
            fecha_inicio,
            monto_total,
            cantidad_cuotas,
            monto_cuota,
            descripcion,
            observaciones,
            session.get("username"),
        ))
        conn.commit()
        conn.close()

        flash("Plan de pago creado correctamente.", "ok")
        return redirect(url_for("ver_cuotas", jugador_id=jugador_id))

    conn.close()
    return render_template("plan_pago_form.html", jugador=jugador, deuda=deuda)


@app.route("/planes-pago/<int:plan_id>/actualizar", methods=["POST"])
def actualizar_plan_pago(plan_id):
    check = permiso_requerido("planes_pago_gestionar")
    if check:
        return check

    estado = request.form.get("estado", "").strip()
    if estado not in {"Activo", "Cumplido", "Caido"}:
        flash("Estado de plan no valido.", "error")
        return redirect(url_for("listar_planes_pago"))

    conn = get_connection()
    plan = conn.execute("SELECT * FROM planes_pago WHERE id = %s", (plan_id,)).fetchone()
    if plan is None:
        conn.close()
        flash("Plan de pago no encontrado.", "error")
        return redirect(url_for("listar_planes_pago"))

    cerrado_en = datetime.now().strftime("%Y-%m-%d") if estado != "Activo" else None
    cerrado_por = session.get("username") if estado != "Activo" else None
    conn.execute("""
        UPDATE planes_pago
        SET estado = %s,
            cerrado_en = %s,
            cerrado_por = %s
        WHERE id = %s
    """, (estado, cerrado_en, cerrado_por, plan_id))
    conn.commit()
    conn.close()

    flash("Plan de pago actualizado.", "ok")
    return redirect(request.form.get("next") or url_for("ver_cuotas", jugador_id=plan["jugador_id"]))

@app.route("/jugadores/<int:jugador_id>/cuotas/nueva", methods=["GET", "POST"])
def nueva_cuota(jugador_id):
    check = permiso_requerido("cuotas_gestionar")
    if check:
        return check

    if request.method == "POST":
        periodo = request.form.get("periodo", "").strip()
        importe = request.form.get("importe", "").strip()
        fecha_vencimiento = request.form.get("fecha_vencimiento", "").strip()

        if not periodo or not importe:
            flash("Período e importe son obligatorios.", "error")
            return render_template("cuota_form.html", jugador_id=jugador_id)

        try:
            importe_valor = float(importe)
        except ValueError:
            flash("El importe debe ser numérico.", "error")
            return render_template(
                "cuota_form.html",
                jugador_id=jugador_id,
                periodo=periodo,
                importe=importe,
                fecha_vencimiento=fecha_vencimiento
            )

        conn = get_connection()
        jugador = conn.execute(
            "SELECT * FROM jugadores WHERE id = %s",
            (jugador_id,)
        ).fetchone()

        if jugador is None:
            conn.close()
            flash("Jugador no encontrado.", "error")
            return redirect(url_for("listar_jugadores"))

        existente = conn.execute("""
            SELECT id
            FROM cuotas
            WHERE jugador_id = %s AND periodo = %s
        """, (jugador_id, periodo)).fetchone()

        if existente:
            conn.close()
            flash("Ese jugador ya tiene una cuota cargada para ese período.", "error")
            return render_template(
                "cuota_form.html",
                jugador_id=jugador_id,
                periodo=periodo,
                importe=importe,
                fecha_vencimiento=fecha_vencimiento
            )

        cuota_calculada = calcular_importe_con_beca(jugador, periodo, importe_valor)
        pagado_inicial = 1 if cuota_calculada["beca_total"] else 0
        fecha_pago_inicial = datetime.now().strftime("%Y-%m-%d") if pagado_inicial else None
        metodo_inicial = "Beca" if pagado_inicial else None
        referencia_inicial = (
            f"Beca total {cuota_calculada['beca_porcentaje']:g}%"
            if pagado_inicial else None
        )

        conn.execute("""
            INSERT INTO cuotas (
                jugador_id, periodo, importe, pagado, fecha_pago, fecha_vencimiento,
                importe_original, descuento_beca, beca_porcentaje, beca_motivo,
                becada, metodo_pago, referencia_pago
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            jugador_id,
            periodo,
            cuota_calculada["importe"],
            pagado_inicial,
            fecha_pago_inicial,
            fecha_vencimiento or None,
            cuota_calculada["importe_original"],
            cuota_calculada["descuento_beca"],
            cuota_calculada["beca_porcentaje"],
            cuota_calculada["beca_motivo"],
            cuota_calculada["becada"],
            metodo_inicial,
            referencia_inicial,
        ))

        conn.commit()
        conn.close()

        if cuota_calculada["beca_total"]:
            flash("Cuota cargada como beca total.", "ok")
        elif cuota_calculada["becada"]:
            flash("Cuota cargada con beca parcial aplicada.", "ok")
        else:
            flash("Cuota cargada correctamente.", "ok")
        return redirect(url_for("ver_cuotas", jugador_id=jugador_id))

    return render_template("cuota_form.html", jugador_id=jugador_id)

@app.route("/cuotas/<int:cuota_id>/pagar", methods=["GET", "POST"])
def pagar_cuota(cuota_id):
    check = permiso_requerido("cuotas_gestionar")
    if check:
        return check

    conn = get_connection()

    cuota = conn.execute("""
        SELECT
            c.*,
            j.nombre,
            j.apellido
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.id = %s
    """, (cuota_id,)).fetchone()

    if cuota is None:
        conn.close()
        flash("Cuota no encontrada.", "error")
        return redirect(url_for("listar_jugadores"))

    if cuota["pagado"]:
        conn.close()
        flash("La cuota ya estaba marcada como pagada.", "error")
        return redirect(url_for("ver_cuotas", jugador_id=cuota["jugador_id"]))

    if request.method == "POST":
        metodo_pago = request.form.get("metodo_pago", "").strip()
        referencia_pago = request.form.get("referencia_pago", "").strip()
        comprobante_pago = request.files.get("comprobante_pago")

        if not metodo_pago:
            conn.close()
            flash("Debe seleccionar un método de pago.", "error")
            return render_template("pagar_cuota.html", cuota=cuota)

        try:
            comprobante_info = subir_comprobante_a_drive(comprobante_pago, cuota)
        except (RuntimeError, ValueError) as error:
            conn.close()
            flash(str(error), "error")
            return render_template("pagar_cuota.html", cuota=cuota)
        except Exception as error:
            app.logger.exception("No se pudo subir comprobante de cuota %s al registrar pago.", cuota_id)
            conn.close()
            flash(f"{mensaje_error_drive(error)} La cuota no fue marcada como pagada.", "error")
            return render_template("pagar_cuota.html", cuota=cuota)

        nuevo_numero = cuota["numero_recibo"] or siguiente_numero_recibo(conn)
        comprobante_fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if comprobante_info else None
        comprobante_usuario = session.get("username") if comprobante_info else None

        conn.execute("""
            UPDATE cuotas
            SET pagado = 1,
                fecha_pago = CURRENT_DATE,
                numero_recibo = %s,
                metodo_pago = %s,
                referencia_pago = %s,
                comprobante_drive_file_id = COALESCE(%s, comprobante_drive_file_id),
                comprobante_nombre = COALESCE(%s, comprobante_nombre),
                comprobante_mime_type = COALESCE(%s, comprobante_mime_type),
                comprobante_tamano = COALESCE(%s, comprobante_tamano),
                comprobante_fecha = COALESCE(%s, comprobante_fecha),
                comprobante_usuario = COALESCE(%s, comprobante_usuario),
                comprobante_web_url = COALESCE(%s, comprobante_web_url),
                comprobante_estado = CASE
                    WHEN %s IS NOT NULL THEN 'aceptado'
                    ELSE comprobante_estado
                END,
                comprobante_revisado_en = CASE
                    WHEN %s IS NOT NULL THEN %s
                    ELSE comprobante_revisado_en
                END,
                comprobante_revisado_por = CASE
                    WHEN %s IS NOT NULL THEN %s
                    ELSE comprobante_revisado_por
                END,
                comprobante_observaciones = CASE
                    WHEN %s IS NOT NULL THEN NULL
                    ELSE comprobante_observaciones
                END
            WHERE id = %s
        """, (
            nuevo_numero,
            metodo_pago,
            referencia_pago,
            comprobante_info["file_id"] if comprobante_info else None,
            comprobante_info["nombre"] if comprobante_info else None,
            comprobante_info["mime_type"] if comprobante_info else None,
            comprobante_info["tamano"] if comprobante_info else None,
            comprobante_fecha,
            comprobante_usuario,
            comprobante_info["web_url"] if comprobante_info else None,
            comprobante_info["file_id"] if comprobante_info else None,
            comprobante_info["file_id"] if comprobante_info else None,
            comprobante_fecha,
            comprobante_info["file_id"] if comprobante_info else None,
            comprobante_usuario,
            comprobante_info["file_id"] if comprobante_info else None,
            cuota_id,
        ))

        conn.execute("""
            INSERT INTO movimientos (tipo, concepto, monto, fecha, referencia)
            VALUES (%s, %s, %s, CURRENT_DATE, %s)
        """, (
            "ingreso",
            f"Cuota {cuota['periodo']} - {cuota['apellido']}, {cuota['nombre']}",
            cuota["importe"],
            f"Cuota Social ({metodo_pago})"
        ))

        conn.commit()
        conn.close()

        generar_recibo_pdf(cuota_id)

        if comprobante_info:
            flash("Cuota marcada como pagada, registrada en caja y comprobante guardado en Drive.", "ok")
        else:
            flash("Cuota marcada como pagada y registrada en caja.", "ok")
        return redirect(url_for("ver_cuotas", jugador_id=cuota["jugador_id"]))

    conn.close()
    return render_template("pagar_cuota.html", cuota=cuota)

@app.route("/cuotas/<int:cuota_id>/comprobante/subir", methods=["GET", "POST"])
def subir_comprobante_cuota(cuota_id):
    check = permiso_requerido("cuotas_gestionar")
    if check:
        return check

    conn = get_connection()
    cuota = conn.execute("""
        SELECT
            c.*,
            j.nombre,
            j.apellido
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.id = %s
    """, (cuota_id,)).fetchone()

    if cuota is None:
        conn.close()
        flash("Cuota no encontrada.", "error")
        return redirect(url_for("listar_jugadores"))

    if not cuota["pagado"]:
        conn.close()
        flash("Solo se pueden adjuntar comprobantes en cuotas pagadas.", "error")
        return redirect(url_for("ver_cuotas", jugador_id=cuota["jugador_id"]))

    if request.method == "POST":
        comprobante_pago = request.files.get("comprobante_pago")
        try:
            comprobante_info = subir_comprobante_a_drive(comprobante_pago, cuota)
        except (RuntimeError, ValueError) as error:
            conn.close()
            flash(str(error), "error")
            return render_template("comprobante_form.html", cuota=cuota)
        except Exception as error:
            app.logger.exception("No se pudo subir comprobante de cuota %s.", cuota_id)
            conn.close()
            flash(mensaje_error_drive(error), "error")
            return render_template("comprobante_form.html", cuota=cuota)

        if not comprobante_info:
            conn.close()
            flash("Debe seleccionar un archivo para adjuntar.", "error")
            return render_template("comprobante_form.html", cuota=cuota)

        conn.execute("""
            UPDATE cuotas
            SET comprobante_drive_file_id = %s,
                comprobante_nombre = %s,
                comprobante_mime_type = %s,
                comprobante_tamano = %s,
                comprobante_fecha = %s,
                comprobante_usuario = %s,
                comprobante_web_url = %s,
                comprobante_estado = 'aceptado',
                comprobante_revisado_en = %s,
                comprobante_revisado_por = %s,
                comprobante_observaciones = NULL
            WHERE id = %s
        """, (
            comprobante_info["file_id"],
            comprobante_info["nombre"],
            comprobante_info["mime_type"],
            comprobante_info["tamano"],
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            session.get("username"),
            comprobante_info["web_url"],
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            session.get("username"),
            cuota_id,
        ))

        conn.commit()
        conn.close()

        flash("Comprobante guardado en Google Drive.", "ok")
        return redirect(url_for("ver_cuotas", jugador_id=cuota["jugador_id"]))

    conn.close()
    return render_template("comprobante_form.html", cuota=cuota)


@app.route("/cuotas/comprobantes")
def listar_comprobantes_revision():
    check = permiso_requerido("cuotas_gestionar")
    if check:
        return check

    estado = request.args.get("estado", "pendiente").strip() or "pendiente"
    if estado not in COMPROBANTE_ESTADOS and estado != "todos":
        estado = "pendiente"

    filtros = ["c.comprobante_drive_file_id IS NOT NULL"]
    params = []
    if estado != "todos":
        if estado == "pendiente":
            filtros.append("COALESCE(NULLIF(c.comprobante_estado, ''), 'sin_comprobante') IN ('pendiente', 'sin_comprobante')")
        else:
            filtros.append("COALESCE(NULLIF(c.comprobante_estado, ''), 'sin_comprobante') = %s")
            params.append(estado)

    where_sql = " AND ".join(filtros)

    conn = get_connection()
    comprobantes = conn.execute(f"""
        SELECT
            c.*,
            CASE
                WHEN c.comprobante_drive_file_id IS NOT NULL
                 AND COALESCE(NULLIF(c.comprobante_estado, ''), 'sin_comprobante') = 'sin_comprobante'
                THEN 'pendiente'
                ELSE COALESCE(NULLIF(c.comprobante_estado, ''), 'sin_comprobante')
            END AS comprobante_estado_resuelto,
            j.apellido,
            j.nombre
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE {where_sql}
        ORDER BY
            CASE COALESCE(NULLIF(c.comprobante_estado, ''), 'pendiente')
                WHEN 'pendiente' THEN 0
                WHEN 'sin_comprobante' THEN 0
                WHEN 'rechazado' THEN 1
                WHEN 'aceptado' THEN 2
                ELSE 3
            END,
            c.comprobante_fecha DESC NULLS LAST,
            c.id DESC
        LIMIT 150
    """, tuple(params)).fetchall()

    resumen = conn.execute("""
        SELECT
            COUNT(*) AS total,
            COUNT(*) FILTER (
                WHERE COALESCE(NULLIF(comprobante_estado, ''), 'sin_comprobante') IN ('pendiente', 'sin_comprobante')
                  AND comprobante_drive_file_id IS NOT NULL
            ) AS pendientes,
            COUNT(*) FILTER (
                WHERE COALESCE(NULLIF(comprobante_estado, ''), 'pendiente') = 'aceptado'
            ) AS aceptados,
            COUNT(*) FILTER (
                WHERE COALESCE(NULLIF(comprobante_estado, ''), 'pendiente') = 'rechazado'
            ) AS rechazados
        FROM cuotas
        WHERE comprobante_drive_file_id IS NOT NULL
    """).fetchone()
    conn.close()

    return render_template(
        "comprobantes_revision.html",
        comprobantes=comprobantes,
        estado=estado,
        resumen=resumen,
    )


@app.route("/cuotas/<int:cuota_id>/comprobante/revisar", methods=["POST"])
def revisar_comprobante_cuota(cuota_id):
    check = permiso_requerido("cuotas_gestionar")
    if check:
        return check

    accion = request.form.get("accion", "").strip()
    observaciones = request.form.get("observaciones", "").strip()
    next_url = destino_interno(request.form.get("next"), fallback="listar_comprobantes_revision")

    if accion not in {"aceptar", "rechazar"}:
        flash("La accion de revision no es valida.", "error")
        return redirect(next_url)

    conn = get_connection()
    cuota = conn.execute("""
        SELECT
            c.*,
            j.apellido,
            j.nombre
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.id = %s
        FOR UPDATE
    """, (cuota_id,)).fetchone()

    if cuota is None:
        conn.close()
        flash("Cuota no encontrada.", "error")
        return redirect(url_for("listar_jugadores"))

    if not cuota["comprobante_drive_file_id"]:
        conn.close()
        flash("La cuota no tiene comprobante adjunto.", "error")
        return redirect(next_url)

    revisado_en = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    revisado_por = session.get("username")

    if accion == "rechazar":
        conn.execute("""
            UPDATE cuotas
            SET comprobante_estado = 'rechazado',
                comprobante_revisado_en = %s,
                comprobante_revisado_por = %s,
                comprobante_observaciones = %s
            WHERE id = %s
        """, (revisado_en, revisado_por, observaciones or None, cuota_id))
        conn.commit()
        conn.close()
        flash("Comprobante rechazado.", "ok")
        return redirect(next_url)

    numero_recibo = cuota["numero_recibo"] or siguiente_numero_recibo(conn)
    fecha_pago = cuota["fecha_pago"] or datetime.now().strftime("%Y-%m-%d")
    metodo_pago = cuota["metodo_pago"] or "Comprobante portal"
    referencia_pago = cuota["referencia_pago"] or "Comprobante validado"

    conn.execute("""
        UPDATE cuotas
        SET pagado = 1,
            fecha_pago = %s,
            numero_recibo = %s,
            metodo_pago = %s,
            referencia_pago = %s,
            comprobante_estado = 'aceptado',
            comprobante_revisado_en = %s,
            comprobante_revisado_por = %s,
            comprobante_observaciones = %s
        WHERE id = %s
    """, (
        fecha_pago,
        numero_recibo,
        metodo_pago,
        referencia_pago,
        revisado_en,
        revisado_por,
        observaciones or None,
        cuota_id,
    ))

    try:
        importe = float(cuota["importe"] or 0)
    except (TypeError, ValueError):
        importe = 0

    if not cuota["pagado"] and importe > 0:
        conn.execute("""
            INSERT INTO movimientos (tipo, concepto, monto, fecha, referencia)
            VALUES (%s, %s, %s, %s, %s)
        """, (
            "ingreso",
            f"Cuota {cuota['periodo']} - {cuota['apellido']}, {cuota['nombre']}",
            cuota["importe"],
            fecha_pago,
            f"Cuota Social ({metodo_pago})",
        ))

    conn.commit()
    conn.close()

    generar_recibo_pdf(cuota_id)
    flash("Comprobante aceptado. Cuota marcada como pagada y recibo generado.", "ok")
    return redirect(next_url)


@app.route("/cuotas/<int:cuota_id>/comprobante")
def descargar_comprobante_cuota(cuota_id):
    check = permiso_requerido("cuotas_ver", "cuotas_gestionar")
    if check:
        return check

    conn = get_connection()
    cuota = conn.execute("""
        SELECT
            c.*,
            j.nombre,
            j.apellido
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.id = %s
    """, (cuota_id,)).fetchone()
    conn.close()

    if cuota is None:
        flash("Cuota no encontrada.", "error")
        return redirect(url_for("listar_jugadores"))

    if not cuota["comprobante_drive_file_id"]:
        flash("La cuota no tiene comprobante adjunto.", "error")
        return redirect(url_for("ver_cuotas", jugador_id=cuota["jugador_id"]))

    try:
        archivo = descargar_drive_file(cuota["comprobante_drive_file_id"])
    except RuntimeError as error:
        flash(str(error), "error")
        return redirect(url_for("ver_cuotas", jugador_id=cuota["jugador_id"]))
    except Exception as error:
        app.logger.exception("No se pudo descargar comprobante de cuota %s.", cuota_id)
        flash(mensaje_error_drive(error), "error")
        return redirect(url_for("ver_cuotas", jugador_id=cuota["jugador_id"]))

    registrar_auditoria(
        "descargar_ok",
        "comprobante_cuota",
        str(cuota_id),
        {
            "archivo": cuota["comprobante_nombre"],
            "drive_file_id": cuota["comprobante_drive_file_id"],
        },
    )

    return send_file(
        archivo,
        mimetype=cuota["comprobante_mime_type"] or "application/octet-stream",
        as_attachment=False,
        download_name=cuota["comprobante_nombre"] or f"comprobante_cuota_{cuota_id}",
    )


@app.route("/cuotas/conciliacion", methods=["GET", "POST"])
def conciliar_pagos():
    check = permiso_requerido("cuotas_gestionar")
    if check:
        return check

    resultado = None
    matches = []
    matches_json = ""

    if request.method == "POST":
        accion = request.form.get("accion", "preview")
        if accion == "aplicar":
            try:
                matches = json.loads(request.form.get("matches_json", "[]"))
            except ValueError:
                matches = []

            conn = get_connection()
            resultado = aplicar_matches_conciliacion(conn, matches)
            conn.commit()
            conn.close()
            registrar_auditoria("conciliar_ok", "cuotas", None, resultado)
            flash(
                f"Conciliacion aplicada. Pagos registrados: {resultado['aplicados']}. "
                f"Omitidos: {resultado['omitidos']}.",
                "ok",
            )
            return redirect(url_for("conciliar_pagos"))

        archivo = request.files.get("archivo")
        if not archivo or not archivo.filename:
            flash("Selecciona un archivo CSV o Excel.", "error")
            return render_template(
                "conciliacion.html",
                resultado=resultado,
                matches=matches,
                matches_json=matches_json,
            )

        try:
            filas = leer_csv_conciliacion(archivo)
        except ValueError as error:
            flash(str(error), "error")
            return render_template(
                "conciliacion.html",
                resultado=resultado,
                matches=matches,
                matches_json=matches_json,
            )

        conn = get_connection()
        matches = [buscar_match_conciliacion(conn, fila) for fila in filas]
        conn.close()

        resultado = {
            "filas": len(matches),
            "matches": sum(1 for item in matches if item["estado"] == "match"),
            "sin_match": sum(1 for item in matches if item["estado"] == "sin_match"),
            "multiples": sum(1 for item in matches if item["estado"] == "multiple"),
            "errores": sum(1 for item in matches if item["estado"] == "error"),
        }
        matches_json = json.dumps(matches, ensure_ascii=False, default=str)

    return render_template(
        "conciliacion.html",
        resultado=resultado,
        matches=matches,
        matches_json=matches_json,
    )


@app.route("/jugadores/<int:jugador_id>/lesiones")
def ver_lesiones(jugador_id):
    check = permiso_requerido("salud_ver")
    if check:
        return check

    conn = get_connection()

    jugador = conn.execute(
        "SELECT * FROM jugadores WHERE id = %s",
        (jugador_id,)
    ).fetchone()

    lesiones = conn.execute("""
        SELECT * FROM lesiones
        WHERE jugador_id = %s
        ORDER BY fecha_lesion DESC, id DESC
    """, (jugador_id,)).fetchall()

    conn.close()

    return render_template(
        "lesiones.html",
        jugador=jugador,
        lesiones=lesiones
    )

@app.route("/jugadores/<int:jugador_id>/lesiones/nueva", methods=["GET", "POST"])
def nueva_lesion(jugador_id):
    check = permiso_requerido("salud_gestionar")
    if check:
        return check

    conn = get_connection()

    jugador = conn.execute(
        "SELECT * FROM jugadores WHERE id = %s",
        (jugador_id,)
    ).fetchone()

    if request.method == "POST":
        fecha_lesion = request.form.get("fecha_lesion", "").strip()
        tipo_lesion = request.form.get("tipo_lesion", "").strip()
        zona_cuerpo = request.form.get("zona_cuerpo", "").strip()
        diagnostico = request.form.get("diagnostico", "").strip()
        tratamiento = request.form.get("tratamiento", "").strip()
        estado = request.form.get("estado", "").strip() or "Activa"
        fecha_alta = request.form.get("fecha_alta", "").strip()
        observaciones = request.form.get("observaciones", "").strip()

        conn.execute("""
            INSERT INTO lesiones (
                jugador_id, fecha_lesion, tipo_lesion, zona_cuerpo,
                diagnostico, tratamiento, estado, fecha_alta, observaciones
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            jugador_id, fecha_lesion, tipo_lesion, zona_cuerpo,
            diagnostico, tratamiento, estado, fecha_alta, observaciones
        ))

        conn.commit()
        conn.close()

        flash("Lesión cargada correctamente.", "ok")
        return redirect(url_for("ver_lesiones", jugador_id=jugador_id))

    conn.close()
    return render_template("lesion_form.html", jugador=jugador)

@app.route("/lesiones/<int:lesion_id>/editar", methods=["GET", "POST"])
def editar_lesion(lesion_id):
    check = permiso_requerido("salud_gestionar")
    if check:
        return check
    conn = get_connection()

    lesion = conn.execute("""
        SELECT * FROM lesiones
        WHERE id = %s
    """, (lesion_id,)).fetchone()

    if lesion is None:
        conn.close()
        flash("Lesión no encontrada.", "error")
        return redirect(url_for("listar_jugadores"))

    jugador = conn.execute("""
        SELECT * FROM jugadores
        WHERE id = %s
    """, (lesion["jugador_id"],)).fetchone()

    if request.method == "POST":
        fecha_lesion = request.form.get("fecha_lesion", "").strip()
        tipo_lesion = request.form.get("tipo_lesion", "").strip()
        zona_cuerpo = request.form.get("zona_cuerpo", "").strip()
        diagnostico = request.form.get("diagnostico", "").strip()
        tratamiento = request.form.get("tratamiento", "").strip()
        estado = request.form.get("estado", "").strip() or "Activa"
        fecha_alta = request.form.get("fecha_alta", "").strip()
        observaciones = request.form.get("observaciones", "").strip()

        conn.execute("""
            UPDATE lesiones
            SET fecha_lesion = %s, tipo_lesion = %s, zona_cuerpo = %s,
                diagnostico = %s, tratamiento = %s, estado = %s,
                fecha_alta = %s, observaciones = %s
            WHERE id = %s
        """, (
            fecha_lesion, tipo_lesion, zona_cuerpo,
            diagnostico, tratamiento, estado,
            fecha_alta, observaciones, lesion_id
        ))

        conn.commit()
        conn.close()

        flash("Lesión actualizada correctamente.", "ok")
        return redirect(url_for("ver_lesiones", jugador_id=jugador["id"]))

    conn.close()
    return render_template("lesion_form.html", jugador=jugador, lesion=lesion)

@app.route("/lesiones/<int:lesion_id>/eliminar", methods=["POST"])
def eliminar_lesion(lesion_id):
    check = permiso_requerido("salud_gestionar")
    if check:
        return check
    conn = get_connection()

    lesion = conn.execute("""
        SELECT jugador_id FROM lesiones
        WHERE id = %s
    """, (lesion_id,)).fetchone()

    if lesion is None:
        conn.close()
        flash("Lesión no encontrada.", "error")
        return redirect(url_for("listar_jugadores"))

    conn.execute("DELETE FROM lesiones WHERE id = %s", (lesion_id,))
    conn.commit()
    conn.close()

    flash("Lesión eliminada.", "ok")
    return redirect(url_for("ver_lesiones", jugador_id=lesion["jugador_id"]))

@app.route("/cuotas/generar", methods=["GET", "POST"])
def generar_cuotas():
    check = permiso_requerido("cuotas_gestionar")
    if check:
        return check
    if request.method == "POST":
        periodo = request.form.get("periodo", "").strip()
        importe = request.form.get("importe", "").strip()
        fecha_vencimiento = request.form.get("fecha_vencimiento", "").strip()
        categoria = request.form.get("categoria", "").strip()

        if not periodo or not importe:
            flash("Período e importe son obligatorios.", "error")
            return render_template("generar_cuotas.html")

        try:
            importe_valor = float(importe)
        except ValueError:
            flash("El importe debe ser numérico.", "error")
            return render_template(
                "generar_cuotas.html",
                periodo=periodo,
                importe=importe,
                fecha_vencimiento=fecha_vencimiento,
                categoria=categoria
            )

        conn = get_connection()

        if categoria:
            jugadores = conn.execute("""
                SELECT *
                FROM jugadores
                WHERE estado = 'Activo' AND categoria = %s
                ORDER BY apellido, nombre
            """, (categoria,)).fetchall()
        else:
            jugadores = conn.execute("""
                SELECT *
                FROM jugadores
                WHERE estado = 'Activo'
                ORDER BY apellido, nombre
            """).fetchall()

        creadas = 0
        omitidas = 0
        becadas_totales = 0
        becadas_parciales = 0

        for jugador in jugadores:
            existente = conn.execute("""
                SELECT id
                FROM cuotas
                WHERE jugador_id = %s AND periodo = %s
            """, (jugador["id"], periodo)).fetchone()

            if existente:
                omitidas += 1
                continue

            cuota_calculada = calcular_importe_con_beca(jugador, periodo, importe_valor)
            pagado_inicial = 1 if cuota_calculada["beca_total"] else 0
            fecha_pago_inicial = datetime.now().strftime("%Y-%m-%d") if pagado_inicial else None
            metodo_inicial = "Beca" if pagado_inicial else None
            referencia_inicial = (
                f"Beca total {cuota_calculada['beca_porcentaje']:g}%"
                if pagado_inicial else None
            )

            conn.execute("""
                INSERT INTO cuotas (
                    jugador_id, periodo, importe, pagado, fecha_pago, fecha_vencimiento,
                    importe_original, descuento_beca, beca_porcentaje, beca_motivo,
                    becada, metodo_pago, referencia_pago
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                jugador["id"],
                periodo,
                cuota_calculada["importe"],
                pagado_inicial,
                fecha_pago_inicial,
                fecha_vencimiento or None,
                cuota_calculada["importe_original"],
                cuota_calculada["descuento_beca"],
                cuota_calculada["beca_porcentaje"],
                cuota_calculada["beca_motivo"],
                cuota_calculada["becada"],
                metodo_inicial,
                referencia_inicial,
            ))
            creadas += 1
            if cuota_calculada["beca_total"]:
                becadas_totales += 1
            elif cuota_calculada["becada"]:
                becadas_parciales += 1

        conn.commit()
        conn.close()

        flash(
            f"Generación terminada. Cuotas creadas: {creadas}. Becas totales: {becadas_totales}. "
            f"Becas parciales: {becadas_parciales}. Ya existentes: {omitidas}.",
            "ok",
        )
        return redirect(url_for("index"))

    return render_template("generar_cuotas.html")


@app.route("/jugadores/<int:jugador_id>/timeline")
def timeline_jugador(jugador_id):
    check = permiso_requerido("jugadores_ver")
    if check:
        return check

    conn = get_connection()
    jugador = conn.execute("SELECT * FROM jugadores WHERE id = %s", (jugador_id,)).fetchone()
    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    eventos = []

    if jugador["fecha_ingreso"]:
        eventos.append({
            "fecha": jugador["fecha_ingreso"],
            "tipo": "Alta",
            "titulo": "Ingreso al club",
            "detalle": jugador["estado"],
        })

    if tiene_permiso("cuotas_ver"):
        cuotas = conn.execute("""
            SELECT *
            FROM cuotas
            WHERE jugador_id = %s
            ORDER BY periodo DESC, id DESC
            LIMIT 80
        """, (jugador_id,)).fetchall()
        for cuota in cuotas:
            eventos.append({
                "fecha": cuota["fecha_pago"] or cuota["fecha_vencimiento"] or f"{cuota['periodo']}-01",
                "tipo": "Cuota",
                "titulo": f"Cuota {cuota['periodo']}",
                "detalle": (
                    f"{'Pagada' if cuota['pagado'] else 'Pendiente'} - "
                    f"{formato_moneda(cuota['importe'])}"
                    + (f" - Beca {cuota['beca_porcentaje']}%" if cuota.get("becada") else "")
                ),
            })

        becas = conn.execute("""
            SELECT *
            FROM becas_historial
            WHERE jugador_id = %s
            ORDER BY creado_en DESC
            LIMIT 40
        """, (jugador_id,)).fetchall()
        for beca in becas:
            eventos.append({
                "fecha": str(beca["creado_en"])[:10],
                "tipo": "Beca",
                "titulo": beca["accion"],
                "detalle": f"{beca['beca_porcentaje'] or 0}% - {beca['beca_motivo'] or '-'}",
            })

    if tiene_permiso("salud_ver"):
        ficha = conn.execute("SELECT * FROM fichas_medicas WHERE jugador_id = %s", (jugador_id,)).fetchone()
        if ficha:
            eventos.append({
                "fecha": ficha["fecha_vencimiento"] or str(ficha["id"]),
                "tipo": "Ficha medica",
                "titulo": "Ficha medica",
                "detalle": f"Vencimiento: {ficha['fecha_vencimiento'] or '-'}",
            })

        lesiones = conn.execute("""
            SELECT *
            FROM lesiones
            WHERE jugador_id = %s
            ORDER BY fecha_lesion DESC, id DESC
            LIMIT 50
        """, (jugador_id,)).fetchall()
        for lesion in lesiones:
            eventos.append({
                "fecha": lesion["fecha_lesion"] or lesion["fecha_alta"] or "",
                "tipo": "Lesion",
                "titulo": lesion["tipo_lesion"] or "Lesion",
                "detalle": f"{lesion['estado']} - {lesion['zona_cuerpo'] or '-'}",
            })

        documentos = conn.execute("""
            SELECT *
            FROM documentos_jugadores
            WHERE jugador_id = %s
            ORDER BY COALESCE(fecha_presentacion, fecha_vencimiento) DESC, id DESC
            LIMIT 50
        """, (jugador_id,)).fetchall()
        for documento in documentos:
            eventos.append({
                "fecha": documento["fecha_presentacion"] or documento["fecha_vencimiento"] or "",
                "tipo": "Documento",
                "titulo": documento["tipo"],
                "detalle": f"Vence: {documento['fecha_vencimiento'] or '-'}",
            })

    if tiene_permiso("asistencia_ver"):
        asistencias = conn.execute("""
            SELECT a.*, e.fecha, e.tipo, e.descripcion
            FROM asistencias a
            JOIN eventos_asistencia e ON e.id = a.evento_id
            WHERE a.jugador_id = %s
            ORDER BY e.fecha DESC, e.id DESC
            LIMIT 80
        """, (jugador_id,)).fetchall()
        for asistencia in asistencias:
            eventos.append({
                "fecha": asistencia["fecha"],
                "tipo": "Asistencia",
                "titulo": asistencia["tipo"],
                "detalle": asistencia["estado_asistencia"] or ("presente" if asistencia["presente"] else "ausente"),
            })

    bitacora = conn.execute("""
        SELECT *
        FROM jugador_bitacora
        WHERE jugador_id = %s
        ORDER BY creado_en DESC, id DESC
        LIMIT 80
    """, (jugador_id,)).fetchall()
    for nota in filtrar_bitacora_visible(bitacora):
        eventos.append({
            "fecha": str(nota["creado_en"])[:10],
            "tipo": f"Bitacora {BITACORA_TIPOS.get(nota['tipo'], nota['tipo'])}",
            "titulo": nota["creado_por"] or "SIG",
            "detalle": nota["nota"],
        })

    conn.close()
    eventos.sort(key=lambda item: item["fecha"] or "", reverse=True)

    return render_template("timeline_jugador.html", jugador=jugador, eventos=eventos)


@app.route("/jugadores/<int:jugador_id>")
def detalle_jugador(jugador_id):
    check = permiso_requerido("jugadores_ver")
    if check:
        return check

    puede_ver_finanzas = tiene_permiso("cuotas_ver")
    puede_ver_salud = tiene_permiso("salud_ver")
    puede_ver_asistencia = tiene_permiso("asistencia_ver")
    puede_ver_tests = tiene_permiso("tests_ver")
    puede_gestionar_portal = tiene_permiso("portal_jugador_gestionar")

    conn = get_connection()

    jugador = conn.execute("""
        SELECT * FROM jugadores
        WHERE id = %s
    """, (jugador_id,)).fetchone()

    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    deuda = conn.execute("""
        SELECT COALESCE(SUM(importe), 0) AS total
        FROM cuotas
        WHERE jugador_id = %s
          AND pagado = 0
          AND COALESCE(importe, 0) > 0
    """, (jugador_id,)).fetchone()["total"]

    resumen_cuotas = conn.execute("""
        SELECT
            COUNT(*) AS total,
            COUNT(CASE WHEN pagado = 1 THEN 1 END) AS pagadas,
            COUNT(CASE WHEN pagado = 0 AND COALESCE(importe, 0) > 0 THEN 1 END) AS pendientes,
            COALESCE(SUM(CASE WHEN pagado = 1 THEN importe ELSE 0 END), 0) AS total_pagado,
            COALESCE(SUM(CASE WHEN pagado = 0 AND COALESCE(importe, 0) > 0 THEN importe ELSE 0 END), 0) AS total_pendiente
        FROM cuotas
        WHERE jugador_id = %s
    """, (jugador_id,)).fetchone()

    ultimas_cuotas = conn.execute("""
        SELECT *
        FROM cuotas
        WHERE jugador_id = %s
        ORDER BY periodo DESC, id DESC
        LIMIT 5
    """, (jugador_id,)).fetchall()

    ficha = conn.execute("""
        SELECT *
        FROM fichas_medicas
        WHERE jugador_id = %s
    """, (jugador_id,)).fetchone()

    lesiones = conn.execute("""
        SELECT *
        FROM lesiones
        WHERE jugador_id = %s
        ORDER BY
            CASE
                WHEN estado = 'Activa' THEN 0
                WHEN estado = 'En recuperación' THEN 1
                ELSE 2
            END,
            fecha_lesion DESC,
            id DESC
        LIMIT 5
    """, (jugador_id,)).fetchall()

    ficha_vencida = False
    if ficha and validar_fecha_movimiento(ficha["fecha_vencimiento"]):
        check_vencimiento = conn.execute("""
            SELECT CASE
                WHEN %s::date < CURRENT_DATE THEN 1
                ELSE 0
            END AS vencida
        """, (ficha["fecha_vencimiento"],)).fetchone()
        ficha_vencida = bool(check_vencimiento["vencida"])

    lesiones_activas_count = conn.execute("""
        SELECT COUNT(*) AS total
        FROM lesiones
        WHERE jugador_id = %s
          AND estado IN ('Activa', 'En recuperación')
    """, (jugador_id,)).fetchone()["total"]

    cuotas_pendientes_count = conn.execute("""
        SELECT COUNT(*) AS total
        FROM cuotas
        WHERE jugador_id = %s
          AND pagado = 0
          AND COALESCE(importe, 0) > 0
    """, (jugador_id,)).fetchone()["total"]

    asistencia_resumen = conn.execute("""
        SELECT
            COUNT(*) AS total,
            COALESCE(SUM(CASE WHEN a.presente = 1 THEN 1 ELSE 0 END), 0) AS presentes,
            COALESCE(SUM(CASE WHEN a.presente = 0 THEN 1 ELSE 0 END), 0) AS ausentes
        FROM asistencias a
        JOIN eventos_asistencia e ON e.id = a.evento_id
        WHERE a.jugador_id = %s
    """, (jugador_id,)).fetchone()

    ultimas_asistencias = conn.execute("""
        SELECT
            e.fecha,
            e.tipo,
            e.descripcion,
            a.presente,
            a.estado_asistencia,
            a.observaciones
        FROM asistencias a
        JOIN eventos_asistencia e ON e.id = a.evento_id
        WHERE a.jugador_id = %s
        ORDER BY e.fecha DESC, e.id DESC
        LIMIT 8
    """, (jugador_id,)).fetchall()

    documentos_manual = conn.execute("""
        SELECT *
        FROM documentos_jugadores
        WHERE jugador_id = %s
        ORDER BY COALESCE(fecha_vencimiento, fecha_presentacion) DESC NULLS LAST, id DESC
    """, (jugador_id,)).fetchall()

    planes_pago = conn.execute("""
        SELECT *
        FROM planes_pago
        WHERE jugador_id = %s
        ORDER BY
            CASE WHEN estado = 'Activo' THEN 0 ELSE 1 END,
            fecha_inicio DESC,
            id DESC
        LIMIT 5
    """, (jugador_id,)).fetchall()

    bitacora_raw = conn.execute("""
        SELECT *
        FROM jugador_bitacora
        WHERE jugador_id = %s
        ORDER BY creado_en DESC, id DESC
        LIMIT 40
    """, (jugador_id,)).fetchall()

    tests_recientes = []
    if puede_ver_tests:
        tests_recientes = conn.execute("""
            SELECT
                r.*,
                t.nombre AS test_nombre,
                t.unidad
            FROM test_resultados r
            JOIN test_tipos t ON t.id = r.test_id
            WHERE r.jugador_id = %s
            ORDER BY r.fecha DESC, r.id DESC
            LIMIT 8
        """, (jugador_id,)).fetchall()

    conn.close()

    asistencia_total = asistencia_resumen["total"] or 0
    asistencia_presentes = asistencia_resumen["presentes"] or 0
    asistencia_porcentaje = (
        round((asistencia_presentes / asistencia_total) * 100, 1)
        if asistencia_total else 0
    )

    documentos = []
    if jugador["documentos"]:
        documentos = [
            linea.strip()
            for linea in jugador["documentos"].splitlines()
            if linea.strip()
        ]

    portal_url = None
    if jugador.get("portal_activo") and jugador.get("portal_token"):
        portal_url = url_for("portal_jugador", token=jugador["portal_token"], _external=True)

    bitacora = filtrar_bitacora_visible(bitacora_raw)

    return render_template(
        "jugador_detalle.html",
        jugador=jugador,
        deuda=deuda,
        resumen_cuotas=resumen_cuotas,
        ultimas_cuotas=ultimas_cuotas,
        ficha=ficha,
        lesiones=lesiones,
        ficha_vencida=ficha_vencida,
        lesiones_activas_count=lesiones_activas_count,
        cuotas_pendientes_count=cuotas_pendientes_count,
        asistencia_resumen=asistencia_resumen,
        asistencia_porcentaje=asistencia_porcentaje,
        ultimas_asistencias=ultimas_asistencias,
        documentos=documentos,
        documentos_manual=documentos_manual,
        planes_pago=planes_pago,
        bitacora=bitacora,
        bitacora_tipos=tipos_bitacora_disponibles(),
        bitacora_labels=BITACORA_TIPOS,
        portal_url=portal_url,
        puede_ver_finanzas=puede_ver_finanzas,
        puede_ver_salud=puede_ver_salud,
        puede_ver_asistencia=puede_ver_asistencia,
        puede_ver_tests=puede_ver_tests,
        tests_recientes=tests_recientes,
        puede_gestionar_portal=puede_gestionar_portal
        )


@app.route("/jugadores/<int:jugador_id>/bitacora", methods=["POST"])
def nueva_bitacora_jugador(jugador_id):
    tipos_disponibles = {item["clave"] for item in tipos_bitacora_disponibles()}
    if not tipos_disponibles:
        flash("No tenes permiso para agregar notas en la bitacora.", "error")
        return redirect(url_for("detalle_jugador", jugador_id=jugador_id))

    tipo = request.form.get("tipo", "general").strip()
    nota = request.form.get("nota", "").strip()

    if tipo not in tipos_disponibles:
        flash("No tenes permiso para ese tipo de nota.", "error")
        return redirect(url_for("detalle_jugador", jugador_id=jugador_id))

    if not nota:
        flash("La nota de bitacora no puede estar vacia.", "error")
        return redirect(url_for("detalle_jugador", jugador_id=jugador_id))

    conn = get_connection()
    jugador = conn.execute("SELECT id FROM jugadores WHERE id = %s", (jugador_id,)).fetchone()
    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    conn.execute("""
        INSERT INTO jugador_bitacora (jugador_id, tipo, nota, creado_por)
        VALUES (%s, %s, %s, %s)
    """, (jugador_id, tipo, nota, session.get("username")))
    conn.commit()
    conn.close()

    flash("Nota agregada a la bitacora.", "ok")
    return redirect(url_for("detalle_jugador", jugador_id=jugador_id))


@app.route("/jugadores/<int:jugador_id>/portal/generar", methods=["POST"])
def generar_portal_jugador(jugador_id):
    check = permiso_requerido("portal_jugador_gestionar")
    if check:
        return check

    conn = get_connection()
    jugador = conn.execute("SELECT * FROM jugadores WHERE id = %s", (jugador_id,)).fetchone()
    if jugador is None:
        conn.close()
        flash("Jugador no encontrado.", "error")
        return redirect(url_for("listar_jugadores"))

    token = jugador.get("portal_token") or generar_portal_token()
    conn.execute("""
        UPDATE jugadores
        SET portal_token = %s,
            portal_activo = 1,
            portal_actualizado_en = %s
        WHERE id = %s
    """, (token, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), jugador_id))
    conn.commit()
    conn.close()

    flash("Portal externo habilitado para el jugador.", "ok")
    return redirect(url_for("detalle_jugador", jugador_id=jugador_id))


@app.route("/jugadores/<int:jugador_id>/portal/desactivar", methods=["POST"])
def desactivar_portal_jugador(jugador_id):
    check = permiso_requerido("portal_jugador_gestionar")
    if check:
        return check

    conn = get_connection()
    conn.execute("""
        UPDATE jugadores
        SET portal_activo = 0,
            portal_actualizado_en = %s
        WHERE id = %s
    """, (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), jugador_id))
    conn.commit()
    conn.close()

    flash("Portal externo desactivado.", "ok")
    return redirect(url_for("detalle_jugador", jugador_id=jugador_id))


@app.route("/portal/<token>")
def portal_jugador(token):
    conn = get_connection()
    jugador = conn.execute("""
        SELECT *
        FROM jugadores
        WHERE portal_token = %s
          AND COALESCE(portal_activo, 0) = 1
    """, (token,)).fetchone()

    if jugador is None:
        conn.close()
        abort(404)

    cuotas = conn.execute("""
        SELECT id, periodo, importe, pagado, fecha_vencimiento, fecha_pago,
               metodo_pago, becada, beca_porcentaje, descuento_beca,
               comprobante_drive_file_id, comprobante_nombre, comprobante_fecha,
               comprobante_estado, comprobante_observaciones
        FROM cuotas
        WHERE jugador_id = %s
        ORDER BY periodo DESC, id DESC
        LIMIT 24
    """, (jugador["id"],)).fetchall()

    deuda = conn.execute("""
        SELECT COALESCE(SUM(importe), 0) AS total
        FROM cuotas
        WHERE jugador_id = %s
          AND pagado = 0
          AND COALESCE(importe, 0) > 0
    """, (jugador["id"],)).fetchone()["total"]

    ficha = conn.execute("""
        SELECT presentada, fecha_vencimiento, apto_fisico
        FROM fichas_medicas
        WHERE jugador_id = %s
    """, (jugador["id"],)).fetchone()

    documentos = conn.execute("""
        SELECT tipo, nombre, fecha_vencimiento, url
        FROM documentos_jugadores
        WHERE jugador_id = %s
        ORDER BY fecha_vencimiento DESC NULLS LAST, id DESC
    """, (jugador["id"],)).fetchall()
    conn.close()

    return render_template(
        "portal_jugador.html",
        jugador=jugador,
        cuotas=cuotas,
        deuda=deuda,
        ficha=ficha,
        documentos=documentos,
        token=token,
    )


@app.route("/portal/<token>/contacto", methods=["POST"])
def portal_actualizar_contacto(token):
    conn = get_connection()
    jugador = conn.execute("""
        SELECT *
        FROM jugadores
        WHERE portal_token = %s
          AND COALESCE(portal_activo, 0) = 1
    """, (token,)).fetchone()
    if jugador is None:
        conn.close()
        abort(404)

    data = {
        "telefono": request.form.get("telefono", "").strip(),
        "email": request.form.get("email", "").strip(),
        "telefono_tutor": request.form.get("telefono_tutor", "").strip(),
        "email_tutor": request.form.get("email_tutor", "").strip(),
    }
    conn.execute("""
        UPDATE jugadores
        SET telefono = %s,
            email = %s,
            telefono_tutor = %s,
            email_tutor = %s
        WHERE id = %s
    """, (
        data["telefono"],
        data["email"],
        data["telefono_tutor"],
        data["email_tutor"],
        jugador["id"],
    ))
    conn.commit()
    conn.close()

    registrar_auditoria(
        "actualizar_contacto",
        "portal_jugador",
        str(jugador["id"]),
        {"campos": list(data.keys())},
        username="portal",
        rol="portal",
    )
    flash("Datos de contacto actualizados.", "ok")
    return redirect(url_for("portal_jugador", token=token))


@app.route("/portal/<token>/cuotas/<int:cuota_id>/comprobante", methods=["POST"])
def portal_subir_comprobante(token, cuota_id):
    conn = get_connection()
    cuota = conn.execute("""
        SELECT
            c.*,
            j.nombre,
            j.apellido
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.id = %s
          AND j.portal_token = %s
          AND COALESCE(j.portal_activo, 0) = 1
    """, (cuota_id, token)).fetchone()
    if cuota is None:
        conn.close()
        abort(404)

    if cuota["pagado"]:
        conn.close()
        flash("La cuota ya figura pagada.", "error")
        return redirect(url_for("portal_jugador", token=token))

    comprobante_pago = request.files.get("comprobante_pago")
    referencia_pago = request.form.get("referencia_pago", "").strip()
    try:
        comprobante_info = subir_comprobante_a_drive(comprobante_pago, cuota)
    except (RuntimeError, ValueError) as error:
        conn.close()
        flash(str(error), "error")
        return redirect(url_for("portal_jugador", token=token))
    except Exception as error:
        app.logger.exception("No se pudo subir comprobante desde portal para cuota %s.", cuota_id)
        conn.close()
        flash(mensaje_error_drive(error), "error")
        return redirect(url_for("portal_jugador", token=token))

    comprobante_fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute("""
        UPDATE cuotas
        SET referencia_pago = COALESCE(NULLIF(%s, ''), referencia_pago),
            comprobante_drive_file_id = %s,
            comprobante_nombre = %s,
            comprobante_mime_type = %s,
            comprobante_tamano = %s,
            comprobante_fecha = %s,
            comprobante_usuario = 'portal',
            comprobante_web_url = %s,
            comprobante_estado = 'pendiente',
            comprobante_revisado_en = NULL,
            comprobante_revisado_por = NULL,
            comprobante_observaciones = NULL
        WHERE id = %s
    """, (
        referencia_pago,
        comprobante_info["file_id"],
        comprobante_info["nombre"],
        comprobante_info["mime_type"],
        comprobante_info["tamano"],
        comprobante_fecha,
        comprobante_info["web_url"],
        cuota_id,
    ))
    conn.commit()
    conn.close()

    registrar_auditoria(
        "subir_comprobante",
        "portal_jugador",
        str(cuota["jugador_id"]),
        {"cuota_id": cuota_id, "archivo": comprobante_info["nombre"]},
        username="portal",
        rol="portal",
    )
    flash("Comprobante recibido. La cuota queda pendiente de validacion interna.", "ok")
    return redirect(url_for("portal_jugador", token=token))


@app.route("/portal/<token>/cuotas/<int:cuota_id>/recibo")
def portal_descargar_recibo(token, cuota_id):
    conn = get_connection()
    cuota = conn.execute("""
        SELECT c.*
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.id = %s
          AND j.portal_token = %s
          AND COALESCE(j.portal_activo, 0) = 1
    """, (cuota_id, token)).fetchone()
    conn.close()

    if cuota is None:
        abort(404)
    if not cuota["pagado"]:
        flash("El recibo esta disponible cuando la cuota figura pagada.", "error")
        return redirect(url_for("portal_jugador", token=token))

    archivo = BASE_DIR / "recibos" / f"recibo_cuota_{cuota_id}.pdf"
    if not archivo.exists():
        archivo = generar_recibo_pdf(cuota_id)
    if archivo is None or not archivo.exists():
        flash("No se pudo generar el recibo.", "error")
        return redirect(url_for("portal_jugador", token=token))

    return send_file(
        archivo,
        as_attachment=True,
        download_name=f"recibo_cuota_{cuota_id}.pdf",
    )


@app.route("/calendario")
def ver_calendario():
    check = permiso_requerido("calendario_ver")
    if check:
        return check

    mes = normalizar_mes(request.args.get("mes"), datetime.now().strftime("%Y-%m"))
    calendario = obtener_calendario(mes)
    mes_actual = datetime.strptime(mes, "%Y-%m")

    if mes_actual.month == 1:
        mes_anterior = mes_actual.replace(year=mes_actual.year - 1, month=12)
    else:
        mes_anterior = mes_actual.replace(month=mes_actual.month - 1)

    if mes_actual.month == 12:
        mes_siguiente = mes_actual.replace(year=mes_actual.year + 1, month=1)
    else:
        mes_siguiente = mes_actual.replace(month=mes_actual.month + 1)

    return render_template(
        "calendario.html",
        calendario=calendario,
        mes=mes,
        mes_anterior=mes_anterior.strftime("%Y-%m"),
        mes_siguiente=mes_siguiente.strftime("%Y-%m"),
    )


@app.route("/calendario/nuevo", methods=["GET", "POST"])
def nuevo_evento_calendario():
    check = permiso_requerido("calendario_gestionar")
    if check:
        return check

    if request.method == "POST":
        data = {
            "fecha": request.form.get("fecha", "").strip(),
            "tipo": request.form.get("tipo", "").strip(),
            "titulo": request.form.get("titulo", "").strip(),
            "descripcion": request.form.get("descripcion", "").strip(),
            "ubicacion": request.form.get("ubicacion", "").strip(),
            "categoria": request.form.get("categoria", "").strip(),
        }

        if not data["fecha"] or not data["tipo"] or not data["titulo"]:
            flash("Fecha, tipo y título son obligatorios.", "error")
            return render_template("calendario_evento_form.html", evento=data)

        try:
            datetime.strptime(data["fecha"], "%Y-%m-%d")
        except ValueError:
            flash("La fecha del evento no es válida.", "error")
            return render_template("calendario_evento_form.html", evento=data)

        conn = get_connection()
        conn.execute("""
            INSERT INTO calendario_eventos (
                fecha, tipo, titulo, descripcion, ubicacion, categoria
            )
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (
            data["fecha"],
            data["tipo"],
            data["titulo"],
            data["descripcion"],
            data["ubicacion"],
            data["categoria"],
        ))
        conn.commit()
        conn.close()

        flash("Evento agregado al calendario.", "ok")
        return redirect(url_for("ver_calendario", mes=data["fecha"][:7]))

    return render_template("calendario_evento_form.html", evento=None)


@app.route("/alertas")
def ver_alertas():
    check = permiso_requerido("alertas_finanzas", "alertas_salud")
    if check:
        return check

    puede_ver_finanzas = tiene_permiso("alertas_finanzas")
    puede_ver_salud = tiene_permiso("alertas_salud")
    alertas = obtener_alertas()
    alertas = filtrar_alertas_por_permisos(
        alertas,
        puede_ver_finanzas=puede_ver_finanzas,
        puede_ver_salud=puede_ver_salud,
    )
    return render_template(
        "alertas.html",
        alertas=alertas,
        puede_ver_finanzas=puede_ver_finanzas,
        puede_ver_salud=puede_ver_salud,
    )


@app.route("/salud")
def panel_salud():
    check = permiso_requerido("salud_ver")
    if check:
        return check

    return render_template("salud_panel.html", panel=obtener_panel_salud())


@app.route("/reportes")
def ver_reportes():
    check = permiso_requerido("reportes_ver")
    if check:
        return check

    filtros = filtros_reportes()
    reportes = obtener_reportes(filtros["desde"], filtros["hasta"])

    return render_template(
        "reportes.html",
        filtros=filtros,
        reportes=reportes,
    )


def estilizar_hoja_reporte(ws, header_row=1):
    encabezado = PatternFill("solid", fgColor="1F2937")
    thin = Side(style="thin", color="D1D5DB")

    for cell in ws[header_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = encabezado
        cell.alignment = Alignment(horizontal="center")

    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
            cell.alignment = Alignment(vertical="top")

    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 42)

    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = ws.dimensions


def agregar_hoja_reporte(wb, titulo, encabezados, filas):
    ws = wb.create_sheet(titulo)
    ws.append(encabezados)
    for fila in filas:
        ws.append(fila)
    estilizar_hoja_reporte(ws)
    return ws


def aplicar_formato_columnas(ws, formatos):
    for columna, formato in formatos.items():
        for cell in ws[columna][1:]:
            if isinstance(cell.value, (int, float)):
                cell.number_format = formato


@app.route("/reportes/exportar")
def exportar_reportes():
    check = permiso_requerido("reportes_ver")
    if check:
        return check

    filtros = filtros_reportes()
    reportes = obtener_reportes(filtros["desde"], filtros["hasta"])

    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    ws.append(["Reporte", "Valor"])
    ws.append(["Periodo", f"{filtros['desde']} a {filtros['hasta']}"])
    ws.append(["Ingresos", reportes["resumen"]["ingresos"]])
    ws.append(["Egresos", reportes["resumen"]["egresos"]])
    ws.append(["Resultado", reportes["resumen"]["resultado"]])
    ws.append(["Cuotas cobradas", reportes["resumen"]["cuotas_cobradas"]])
    ws.append(["Cuotas pagadas", reportes["resumen"]["cuotas_pagadas"]])
    ws.append(["Deuda total pendiente", reportes["resumen"]["deuda"]])
    ws.append(["Deuda vencida", reportes["resumen"]["deuda_vencida"]])
    ws.append(["Total bonificado por becas", reportes["resumen"]["total_bonificado_becas"]])
    ws.append(["Cuotas becadas", reportes["resumen"]["cuotas_becadas"]])
    ws.append(["Becas totales", reportes["resumen"]["becas_totales"]])
    ws.append(["Becas parciales", reportes["resumen"]["becas_parciales"]])
    ws.append(["Jugadores activos", reportes["resumen"]["jugadores_activos"]])
    ws.append(["Asistencia promedio", f"{reportes['resumen']['asistencia_porcentaje']}%"])
    estilizar_hoja_reporte(ws)
    for row in (3, 4, 5, 6, 8, 9):
        ws.cell(row=row, column=2).number_format = '$ #,##0'
    for row in (7, 10):
        ws.cell(row=row, column=2).number_format = '#,##0'

    ws_mensual = agregar_hoja_reporte(
        wb,
        "Mensual",
        [
            "Mes", "Ingresos", "Egresos", "Resultado", "Movimientos",
            "Cuotas emitidas", "Cuotas pagadas", "Cuotas pendientes",
            "Total emitido", "Total cobrado", "Cuotas becadas", "Becas totales",
            "Becas parciales", "Total bonificado"
        ],
        [
            [
                fila["mes"],
                fila["ingresos"],
                fila["egresos"],
                fila["resultado"],
                fila["movimientos"],
                fila["cuotas_emitidas"],
                fila["cuotas_pagadas"],
                fila["cuotas_pendientes"],
                fila["total_emitido"],
                fila["total_cobrado"],
                fila["cuotas_becadas"],
                fila["becas_totales"],
                fila["becas_parciales"],
                fila["total_bonificado"],
            ]
            for fila in reportes["mensual"]
        ],
    )
    aplicar_formato_columnas(ws_mensual, {
        "B": '$ #,##0',
        "C": '$ #,##0',
        "D": '$ #,##0',
        "E": '#,##0',
        "F": '#,##0',
        "G": '#,##0',
        "H": '#,##0',
        "I": '$ #,##0',
        "J": '$ #,##0',
        "K": '#,##0',
        "L": '#,##0',
        "M": '#,##0',
        "N": '$ #,##0',
    })

    ws_becas = agregar_hoja_reporte(
        wb,
        "Becas",
        [
            "Apellido", "Nombre", "Categoria", "Porcentaje", "Desde", "Hasta",
            "Cuotas becadas", "Total bonificado", "Motivo"
        ],
        [
            [
                fila["apellido"],
                fila["nombre"],
                fila["categoria"],
                fila["beca_porcentaje"],
                fila["beca_desde"],
                fila["beca_hasta"],
                fila["cuotas_becadas"],
                fila["total_bonificado"],
                fila["beca_motivo"],
            ]
            for fila in reportes["becas_jugadores"]
        ],
    )
    aplicar_formato_columnas(ws_becas, {
        "D": '0.00',
        "G": '#,##0',
        "H": '$ #,##0',
    })

    ws_deuda = agregar_hoja_reporte(
        wb,
        "Deuda categoria",
        ["Categoria", "Jugadores", "Cuotas pendientes", "Deuda", "Deuda vencida"],
        [
            [
                fila["categoria"],
                fila["jugadores"],
                fila["cuotas_pendientes"],
                fila["deuda"],
                fila["deuda_vencida"],
            ]
            for fila in reportes["deuda_por_categoria"]
        ],
    )
    aplicar_formato_columnas(ws_deuda, {
        "B": '#,##0',
        "C": '#,##0',
        "D": '$ #,##0',
        "E": '$ #,##0',
    })

    ws_egresos = agregar_hoja_reporte(
        wb,
        "Egresos concepto",
        ["Concepto", "Cantidad", "Total"],
        [
            [fila["concepto"], fila["cantidad"], fila["total"]]
            for fila in reportes["egresos_por_concepto"]
        ],
    )
    aplicar_formato_columnas(ws_egresos, {
        "B": '#,##0',
        "C": '$ #,##0',
    })

    ws_morosos = agregar_hoja_reporte(
        wb,
        "Morosos",
        [
            "Apellido", "Nombre", "Categoria", "Telefono", "Email",
            "Cuotas pendientes", "Cuotas vencidas", "Deuda", "Primer vencimiento"
        ],
        [
            [
                fila["apellido"],
                fila["nombre"],
                fila["categoria"],
                fila["telefono"],
                fila["email"],
                fila["cuotas_pendientes"],
                fila["cuotas_vencidas"],
                fila["deuda"],
                fila["primer_vencimiento"],
            ]
            for fila in reportes["morosos_recurrentes"]
        ],
    )
    aplicar_formato_columnas(ws_morosos, {
        "F": '#,##0',
        "G": '#,##0',
        "H": '$ #,##0',
    })

    ws_asistencia = agregar_hoja_reporte(
        wb,
        "Asistencia",
        ["Categoria", "Eventos", "Registros", "Presentes", "Ausentes", "Porcentaje"],
        [
            [
                fila["categoria"],
                fila["eventos"],
                fila["registros"],
                fila["presentes"],
                fila["ausentes"],
                fila["porcentaje"],
            ]
            for fila in reportes["asistencia_por_categoria"]
        ],
    )
    aplicar_formato_columnas(ws_asistencia, {
        "B": '#,##0',
        "C": '#,##0',
        "D": '#,##0',
        "E": '#,##0',
        "F": '0.0',
    })

    export_dir = BASE_DIR / "exports"
    export_dir.mkdir(exist_ok=True)

    archivo = export_dir / f"reportes_{filtros['desde']}_a_{filtros['hasta']}.xlsx"
    wb.save(archivo)

    registrar_auditoria(
        "exportar_ok",
        "reportes",
        f"{filtros['desde']}:{filtros['hasta']}",
        {"formato": "xlsx"},
    )

    return send_file(
        archivo,
        as_attachment=True,
        download_name=f"reportes_{filtros['desde']}_a_{filtros['hasta']}.xlsx"
    )


@app.route("/exportar/datos")
def exportar_datos_integral():
    check = permiso_requerido("reportes_ver")
    if check:
        return check

    conn = get_connection()

    jugadores = conn.execute("""
        SELECT
            id, apellido, nombre, dni, categoria, estado, fecha_ingreso,
            telefono, email, contacto_tutor, telefono_tutor, email_tutor,
            beca_activa, beca_porcentaje, beca_desde, beca_hasta, beca_motivo
        FROM jugadores
        ORDER BY apellido, nombre
    """).fetchall()

    deudores = conn.execute("""
        SELECT
            j.id,
            j.apellido,
            j.nombre,
            j.dni,
            j.categoria,
            j.telefono,
            j.email,
            COUNT(c.id) AS cuotas_pendientes,
            SUM(
                CASE
                    WHEN c.fecha_vencimiento IS NOT NULL
                     AND c.fecha_vencimiento <> ''
                     AND c.fecha_vencimiento::text ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                     AND c.fecha_vencimiento::date < CURRENT_DATE
                    THEN 1 ELSE 0
                END
            ) AS cuotas_vencidas,
            COALESCE(SUM(c.importe), 0) AS deuda,
            MIN(c.fecha_vencimiento) AS primer_vencimiento
        FROM jugadores j
        JOIN cuotas c ON c.jugador_id = j.id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
        GROUP BY j.id, j.apellido, j.nombre, j.dni, j.categoria, j.telefono, j.email
        HAVING COALESCE(SUM(c.importe), 0) > 0
        ORDER BY deuda DESC, j.apellido, j.nombre
    """).fetchall()

    asistencia = conn.execute("""
        SELECT
            j.apellido,
            j.nombre,
            j.categoria,
            COUNT(a.id) AS registros,
            SUM(CASE WHEN COALESCE(a.presente, 0) = 1 THEN 1 ELSE 0 END) AS presentes,
            SUM(CASE WHEN COALESCE(a.presente, 0) = 0 THEN 1 ELSE 0 END) AS ausentes,
            ROUND(
                CASE WHEN COUNT(a.id) = 0 THEN 0
                ELSE (SUM(CASE WHEN COALESCE(a.presente, 0) = 1 THEN 1 ELSE 0 END)::numeric / COUNT(a.id)) * 100
                END,
                1
            ) AS porcentaje
        FROM jugadores j
        LEFT JOIN asistencias a ON a.jugador_id = j.id
        GROUP BY j.id, j.apellido, j.nombre, j.categoria
        ORDER BY j.apellido, j.nombre
    """).fetchall()

    becas = conn.execute("""
        SELECT
            j.apellido,
            j.nombre,
            j.categoria,
            j.beca_activa,
            j.beca_porcentaje,
            j.beca_desde,
            j.beca_hasta,
            j.beca_motivo,
            COUNT(c.id) FILTER (WHERE COALESCE(c.becada, 0) = 1) AS cuotas_becadas,
            COALESCE(SUM(c.descuento_beca), 0) AS total_bonificado
        FROM jugadores j
        LEFT JOIN cuotas c ON c.jugador_id = j.id
        WHERE COALESCE(j.beca_activa, 0) = 1
           OR COALESCE(c.becada, 0) = 1
        GROUP BY
            j.id, j.apellido, j.nombre, j.categoria, j.beca_activa,
            j.beca_porcentaje, j.beca_desde, j.beca_hasta, j.beca_motivo
        ORDER BY j.apellido, j.nombre
    """).fetchall()

    fichas = conn.execute("""
        SELECT
            j.apellido,
            j.nombre,
            j.categoria,
            f.presentada,
            f.apto_fisico,
            f.fecha_vencimiento,
            f.contacto_emergencia,
            f.telefono_emergencia,
            f.documento_nombre,
            f.ocr_fecha,
            f.observaciones
        FROM jugadores j
        LEFT JOIN fichas_medicas f ON f.jugador_id = j.id
        ORDER BY j.apellido, j.nombre
    """).fetchall()

    auditoria = []
    if tiene_permiso("auditoria_ver"):
        auditoria = conn.execute("""
            SELECT fecha, username, rol, accion, entidad, entidad_id, ip
            FROM auditoria
            ORDER BY fecha DESC, id DESC
            LIMIT 2000
        """).fetchall()

    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Jugadores"
    ws.append([
        "ID", "Apellido", "Nombre", "DNI", "Categoria", "Estado", "Ingreso",
        "Telefono", "Email", "Contacto tutor", "Telefono tutor", "Email tutor",
        "Beca activa", "Beca %", "Beca desde", "Beca hasta", "Motivo beca"
    ])
    for jugador in jugadores:
        ws.append([
            jugador["id"],
            jugador["apellido"],
            jugador["nombre"],
            jugador["dni"],
            jugador["categoria"],
            jugador["estado"],
            jugador["fecha_ingreso"],
            jugador["telefono"],
            jugador["email"],
            jugador["contacto_tutor"],
            jugador["telefono_tutor"],
            jugador["email_tutor"],
            "Si" if jugador["beca_activa"] else "No",
            jugador["beca_porcentaje"],
            jugador["beca_desde"],
            jugador["beca_hasta"],
            jugador["beca_motivo"],
        ])
    estilizar_hoja_reporte(ws)

    ws_deudores = agregar_hoja_reporte(
        wb,
        "Deudores",
        [
            "ID", "Apellido", "Nombre", "DNI", "Categoria", "Telefono", "Email",
            "Cuotas pendientes", "Cuotas vencidas", "Deuda", "Primer vencimiento"
        ],
        [
            [
                fila["id"],
                fila["apellido"],
                fila["nombre"],
                fila["dni"],
                fila["categoria"],
                fila["telefono"],
                fila["email"],
                fila["cuotas_pendientes"],
                fila["cuotas_vencidas"],
                fila["deuda"],
                fila["primer_vencimiento"],
            ]
            for fila in deudores
        ],
    )
    aplicar_formato_columnas(ws_deudores, {"H": '#,##0', "I": '#,##0', "J": '$ #,##0'})

    ws_asistencia = agregar_hoja_reporte(
        wb,
        "Asistencia",
        ["Apellido", "Nombre", "Categoria", "Registros", "Presentes", "Ausentes", "Porcentaje"],
        [
            [
                fila["apellido"],
                fila["nombre"],
                fila["categoria"],
                fila["registros"],
                fila["presentes"],
                fila["ausentes"],
                fila["porcentaje"],
            ]
            for fila in asistencia
        ],
    )
    aplicar_formato_columnas(ws_asistencia, {"D": '#,##0', "E": '#,##0', "F": '#,##0', "G": '0.0'})

    ws_becas = agregar_hoja_reporte(
        wb,
        "Becas",
        [
            "Apellido", "Nombre", "Categoria", "Activa", "Porcentaje", "Desde",
            "Hasta", "Motivo", "Cuotas becadas", "Total bonificado"
        ],
        [
            [
                fila["apellido"],
                fila["nombre"],
                fila["categoria"],
                "Si" if fila["beca_activa"] else "No",
                fila["beca_porcentaje"],
                fila["beca_desde"],
                fila["beca_hasta"],
                fila["beca_motivo"],
                fila["cuotas_becadas"],
                fila["total_bonificado"],
            ]
            for fila in becas
        ],
    )
    aplicar_formato_columnas(ws_becas, {"E": '0.00', "I": '#,##0', "J": '$ #,##0'})

    agregar_hoja_reporte(
        wb,
        "Fichas medicas",
        [
            "Apellido", "Nombre", "Categoria", "Presentada", "Apto fisico",
            "Vencimiento", "Contacto emergencia", "Telefono emergencia",
            "Documento", "OCR fecha", "Observaciones"
        ],
        [
            [
                fila["apellido"],
                fila["nombre"],
                fila["categoria"],
                "Si" if fila["presentada"] else "No",
                "Si" if fila["apto_fisico"] else "No",
                fila["fecha_vencimiento"],
                fila["contacto_emergencia"],
                fila["telefono_emergencia"],
                fila["documento_nombre"],
                fila["ocr_fecha"],
                fila["observaciones"],
            ]
            for fila in fichas
        ],
    )

    if auditoria:
        agregar_hoja_reporte(
            wb,
            "Auditoria",
            ["Fecha", "Usuario", "Rol", "Accion", "Entidad", "Entidad ID", "IP"],
            [
                [
                    fila["fecha"],
                    fila["username"],
                    fila["rol"],
                    fila["accion"],
                    fila["entidad"],
                    fila["entidad_id"],
                    fila["ip"],
                ]
                for fila in auditoria
            ],
        )

    export_dir = BASE_DIR / "exports"
    export_dir.mkdir(exist_ok=True)
    fecha = datetime.now().strftime("%Y%m%d_%H%M")
    archivo = export_dir / f"sig_export_integral_{fecha}.xlsx"
    wb.save(archivo)

    registrar_auditoria(
        "exportar_ok",
        "datos_integrales",
        None,
        {
            "formato": "xlsx",
            "jugadores": len(jugadores),
            "deudores": len(deudores),
            "incluye_auditoria": bool(auditoria),
        },
    )

    return send_file(
        archivo,
        as_attachment=True,
        download_name=f"sig_export_integral_{fecha}.xlsx",
    )


@app.route("/comunicaciones")
def ver_comunicaciones():
    check = permiso_requerido("comunicaciones_ver")
    if check:
        return check

    template_default = (
        "Hola {nombre}, te escribimos de Ruda Macho Rugby Club. "
        "Registramos {cuotas_pendientes} cuota(s) pendiente(s) por {deuda}. "
        "Te pedimos regularizar la situación o avisarnos si ya realizaste el pago. Gracias."
    )
    template = request.args.get("mensaje", template_default).strip() or template_default
    morosos = obtener_morosos_para_comunicacion()

    comunicaciones = []
    for jugador in morosos:
        mensaje = mensaje_moroso(template, jugador)
        telefono = jugador["telefono_tutor"] or jugador["telefono"]
        telefono_whatsapp = normalizar_telefono_whatsapp(telefono)
        if telefono_whatsapp:
            whatsapp_url = f"https://wa.me/{telefono_whatsapp}?text={quote(mensaje)}"
        else:
            whatsapp_url = f"https://wa.me/?text={quote(mensaje)}"

        comunicaciones.append({
            "jugador": jugador,
            "mensaje": mensaje,
            "telefono_whatsapp": telefono_whatsapp,
            "whatsapp_url": whatsapp_url,
        })

    return render_template(
        "comunicaciones.html",
        template=template,
        comunicaciones=comunicaciones,
    )


@app.route("/notificaciones")
def ver_notificaciones():
    check = permiso_requerido("comunicaciones_ver")
    if check:
        return check

    datos = obtener_notificaciones_operativas()

    cuotas_vencidas = []
    for cuota in datos["cuotas_vencidas"]:
        mensaje = (
            f"Hola {cuota['nombre']}, registramos pendiente la cuota {cuota['periodo']} "
            f"por {formato_moneda(cuota['importe'])}, vencida el {cuota['fecha_vencimiento']}. "
            "Por favor avisanos si ya fue abonada. Gracias."
        )
        cuotas_vencidas.append({
            "item": cuota,
            "mensaje": mensaje,
            "whatsapp_url": whatsapp_mensaje(cuota["telefono_tutor"] or cuota["telefono"], mensaje),
        })

    cuotas_por_vencer = []
    for cuota in datos["cuotas_por_vencer"]:
        mensaje = (
            f"Hola {cuota['nombre']}, te recordamos que la cuota {cuota['periodo']} "
            f"por {formato_moneda(cuota['importe'])} vence el {cuota['fecha_vencimiento']}. Gracias."
        )
        cuotas_por_vencer.append({
            "item": cuota,
            "mensaje": mensaje,
            "whatsapp_url": whatsapp_mensaje(cuota["telefono_tutor"] or cuota["telefono"], mensaje),
        })

    fichas = []
    for ficha in datos["fichas"]:
        if ficha["estado_documento"] == "vencida":
            texto_estado = f"esta vencida desde el {ficha['fecha_vencimiento']}"
        elif ficha["estado_documento"] == "por_vencer":
            texto_estado = f"vence el {ficha['fecha_vencimiento']}"
        else:
            texto_estado = "figura pendiente de carga"
        mensaje = (
            f"Hola {ficha['nombre']}, la ficha medica {texto_estado}. "
            "Cuando puedas, acercanos la actualizacion. Gracias."
        )
        fichas.append({
            "item": ficha,
            "mensaje": mensaje,
            "whatsapp_url": whatsapp_mensaje(ficha["telefono_tutor"] or ficha["telefono"], mensaje),
        })

    asistencia_baja = []
    for jugador in datos["asistencia_baja"]:
        registros = jugador["registros"] or 0
        presentes = jugador["presentes"] or 0
        porcentaje = round((presentes / registros) * 100, 1) if registros else 0
        mensaje = (
            f"Hola {jugador['nombre']}, notamos baja asistencia en los ultimos entrenamientos "
            f"({porcentaje}%). Queremos saber si esta todo bien y como podemos acompañar."
        )
        asistencia_baja.append({
            "item": jugador,
            "porcentaje": porcentaje,
            "mensaje": mensaje,
            "whatsapp_url": whatsapp_mensaje(jugador["telefono_tutor"] or jugador["telefono"], mensaje),
        })

    return render_template(
        "notificaciones.html",
        cuotas_vencidas=cuotas_vencidas,
        cuotas_por_vencer=cuotas_por_vencer,
        fichas=fichas,
        asistencia_baja=asistencia_baja,
        comprobantes_pendientes=datos["comprobantes_pendientes"],
        ahijadxs_objetivo=datos["ahijadxs_objetivo"],
    )

@app.route("/exportar/morosos")
def exportar_morosos():
    check = permiso_requerido("comunicaciones_ver")
    if check:
        return check

    conn = get_connection()

    morosos = conn.execute("""
        SELECT
            j.apellido,
            j.nombre,
            j.dni,
            j.categoria,
            j.telefono,
            j.email,
            COALESCE(SUM(c.importe), 0) AS deuda,
            SUM(
                CASE
                    WHEN c.fecha_vencimiento IS NOT NULL
                     AND c.fecha_vencimiento <> ''
                     AND c.fecha_vencimiento::date < CURRENT_DATE
                    THEN 1
                    ELSE 0
                END
            ) AS cuotas_vencidas
        FROM jugadores j
        JOIN cuotas c ON j.id = c.jugador_id
        WHERE c.pagado = 0
          AND COALESCE(c.importe, 0) > 0
        GROUP BY j.id, j.apellido, j.nombre, j.dni, j.categoria, j.telefono, j.email
        HAVING COALESCE(SUM(c.importe), 0) > 0
        ORDER BY deuda DESC, j.apellido, j.nombre
    """).fetchall()

    conn.close()

    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")

    writer.writerow([
        "Apellido",
        "Nombre",
        "DNI",
        "Categoria",
        "Telefono",
        "Email",
        "Deuda",
        "Cuotas vencidas"
    ])

    for jugador in morosos:
        writer.writerow([
            jugador["apellido"],
            jugador["nombre"],
            jugador["dni"],
            jugador["categoria"],
            jugador["telefono"],
            jugador["email"],
            jugador["deuda"],
            jugador["cuotas_vencidas"]
        ])

    contenido = output.getvalue()
    output.close()

    registrar_auditoria(
        "exportar_ok",
        "morosos",
        None,
        {"formato": "csv", "cantidad_registros": len(morosos)},
    )

    return Response(
        contenido,
        mimetype="text/csv",
        headers={
            "Content-Disposition": "attachment; filename=morosos.csv"
        }
    )

@app.route("/backup")
def backup_db():
    check = permiso_requerido("backup_ver")
    if check:
        return check

    registrar_auditoria(
        "backup_info",
        "backup",
        None,
        {"mensaje": "Consulta de instrucciones de backup Cloud SQL"},
    )
    flash("En la versión con Cloud SQL, el backup se gestiona desde Google Cloud SQL. Usá los backups automáticos o exportaciones desde la consola de Google Cloud.", "ok")
    return render_template(
        "sistema_admin.html",
        estado=obtener_estado_sistema_admin(),
        solo_backup=True,
    )


@app.route("/admin/sistema")
def panel_sistema_admin():
    check = rol_requerido("admin")
    if check:
        return check

    return render_template(
        "sistema_admin.html",
        estado=obtener_estado_sistema_admin(),
        solo_backup=False,
    )


@app.route("/admin/mantenimiento", methods=["GET", "POST"])
def configurar_mantenimiento():
    check = rol_requerido("admin")
    if check:
        return check

    conn = get_connection()

    if request.method == "POST":
        activo = request.form.get("activo") == "on"
        mensaje = request.form.get("mensaje", "").strip() or MAINTENANCE_DEFAULT_MESSAGE

        guardar_app_setting(conn, "maintenance_mode", "true" if activo else "false", session.get("username"))
        guardar_app_setting(conn, "maintenance_message", mensaje, session.get("username"))
        conn.commit()

        g.mantenimiento = obtener_config_mantenimiento(conn)
        conn.close()

        if activo:
            flash("Modo mantenimiento activado. Solo admin puede usar el sistema.", "ok")
        else:
            flash("Modo mantenimiento desactivado. El sistema vuelve a estar disponible.", "ok")
        return redirect(url_for("configurar_mantenimiento"))

    mantenimiento = obtener_config_mantenimiento(conn)
    conn.close()

    return render_template(
        "mantenimiento_admin.html",
        mantenimiento=mantenimiento,
        mensaje_default=MAINTENANCE_DEFAULT_MESSAGE,
    )


@app.route("/auditoria")
def ver_auditoria():
    check = permiso_requerido("auditoria_ver")
    if check:
        return check

    filtros = {
        "q": request.args.get("q", "").strip(),
        "accion": request.args.get("accion", "").strip(),
        "entidad": request.args.get("entidad", "").strip(),
        "usuario": request.args.get("usuario", "").strip(),
        "desde": request.args.get("desde", "").strip(),
        "hasta": request.args.get("hasta", "").strip(),
    }

    condiciones = []
    params = []

    if filtros["q"]:
        patron = f"%{filtros['q']}%"
        condiciones.append("""
            (
                accion ILIKE %s OR entidad ILIKE %s OR entidad_id ILIKE %s
                OR detalle ILIKE %s OR ip ILIKE %s OR username ILIKE %s
            )
        """)
        params.extend([patron, patron, patron, patron, patron, patron])

    if filtros["accion"]:
        condiciones.append("accion = %s")
        params.append(filtros["accion"])

    if filtros["entidad"]:
        condiciones.append("entidad = %s")
        params.append(filtros["entidad"])

    if filtros["usuario"]:
        condiciones.append("username ILIKE %s")
        params.append(f"%{filtros['usuario']}%")

    if validar_fecha_movimiento(filtros["desde"]):
        condiciones.append("fecha::date >= %s::date")
        params.append(filtros["desde"])
    else:
        filtros["desde"] = ""

    if validar_fecha_movimiento(filtros["hasta"]):
        condiciones.append("fecha::date <= %s::date")
        params.append(filtros["hasta"])
    else:
        filtros["hasta"] = ""

    where_sql = ""
    if condiciones:
        where_sql = "WHERE " + " AND ".join(condiciones)

    conn = get_connection()
    registros = conn.execute(f"""
        SELECT *
        FROM auditoria
        {where_sql}
        ORDER BY fecha DESC, id DESC
        LIMIT 300
    """, tuple(params)).fetchall()

    acciones = conn.execute("""
        SELECT DISTINCT accion
        FROM auditoria
        ORDER BY accion
    """).fetchall()

    entidades = conn.execute("""
        SELECT DISTINCT entidad
        FROM auditoria
        WHERE entidad IS NOT NULL
        ORDER BY entidad
    """).fetchall()

    conn.close()

    return render_template(
        "auditoria.html",
        registros=registros,
        filtros=filtros,
        acciones=[row["accion"] for row in acciones],
        entidades=[row["entidad"] for row in entidades],
    )


@app.route("/auditoria/exportar")
def exportar_auditoria():
    check = permiso_requerido("auditoria_ver")
    if check:
        return check

    filtros = {
        "q": request.args.get("q", "").strip(),
        "accion": request.args.get("accion", "").strip(),
        "entidad": request.args.get("entidad", "").strip(),
        "usuario": request.args.get("usuario", "").strip(),
        "desde": request.args.get("desde", "").strip(),
        "hasta": request.args.get("hasta", "").strip(),
    }

    condiciones = []
    params = []
    if filtros["q"]:
        patron = f"%{filtros['q']}%"
        condiciones.append("""
            (
                accion ILIKE %s OR entidad ILIKE %s OR entidad_id ILIKE %s
                OR detalle ILIKE %s OR ip ILIKE %s OR username ILIKE %s
            )
        """)
        params.extend([patron, patron, patron, patron, patron, patron])
    if filtros["accion"]:
        condiciones.append("accion = %s")
        params.append(filtros["accion"])
    if filtros["entidad"]:
        condiciones.append("entidad = %s")
        params.append(filtros["entidad"])
    if filtros["usuario"]:
        condiciones.append("username ILIKE %s")
        params.append(f"%{filtros['usuario']}%")
    if validar_fecha_movimiento(filtros["desde"]):
        condiciones.append("fecha::date >= %s::date")
        params.append(filtros["desde"])
    if validar_fecha_movimiento(filtros["hasta"]):
        condiciones.append("fecha::date <= %s::date")
        params.append(filtros["hasta"])

    where_sql = "WHERE " + " AND ".join(condiciones) if condiciones else ""
    conn = get_connection()
    registros = conn.execute(f"""
        SELECT *
        FROM auditoria
        {where_sql}
        ORDER BY fecha DESC, id DESC
        LIMIT 5000
    """, tuple(params)).fetchall()
    conn.close()

    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow(["Fecha", "Usuario", "Rol", "Accion", "Entidad", "Entidad ID", "IP", "Detalle"])
    for registro in registros:
        writer.writerow([
            registro["fecha"],
            registro["username"],
            registro["rol"],
            registro["accion"],
            registro["entidad"],
            registro["entidad_id"],
            registro["ip"],
            registro["detalle"],
        ])

    contenido = output.getvalue()
    output.close()
    registrar_auditoria("exportar_ok", "auditoria", None, {"cantidad_registros": len(registros)})

    return Response(
        contenido,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=auditoria.csv"},
    )


@app.route("/seguridad")
def ver_seguridad():
    check = permiso_requerido("seguridad_ver")
    if check:
        return check

    conn = get_connection()
    login_resumen = conn.execute("""
        SELECT
            COUNT(*) AS intentos,
            SUM(CASE WHEN success = 1 THEN 1 ELSE 0 END) AS exitosos,
            SUM(CASE WHEN success = 0 THEN 1 ELSE 0 END) AS fallidos
        FROM login_attempts
        WHERE fecha >= CURRENT_TIMESTAMP - INTERVAL '7 days'
    """).fetchone()

    intentos_recientes = conn.execute("""
        SELECT *
        FROM login_attempts
        ORDER BY fecha DESC, id DESC
        LIMIT 80
    """).fetchall()

    ips_fallidas = conn.execute("""
        SELECT ip, username, COUNT(*) AS fallidos, MAX(fecha) AS ultimo
        FROM login_attempts
        WHERE success = 0
          AND fecha >= CURRENT_TIMESTAMP - INTERVAL '24 hours'
        GROUP BY ip, username
        HAVING COUNT(*) >= 2
        ORDER BY fallidos DESC, ultimo DESC
        LIMIT 20
    """).fetchall()

    acciones_sensibles = conn.execute("""
        SELECT *
        FROM auditoria
        WHERE accion ILIKE %s
           OR accion ILIKE %s
           OR accion ILIKE %s
           OR accion ILIKE %s
           OR accion ILIKE %s
           OR accion ILIKE %s
           OR accion ILIKE %s
        ORDER BY fecha DESC, id DESC
        LIMIT 80
    """, (
        "%eliminar%",
        "%password%",
        "%login%",
        "%portal%",
        "%mantenimiento%",
        "%rol%",
        "%usuario%",
    )).fetchall()
    conn.close()

    return render_template(
        "seguridad.html",
        login_resumen=login_resumen,
        intentos_recientes=intentos_recientes,
        ips_fallidas=ips_fallidas,
        acciones_sensibles=acciones_sensibles,
        max_login_attempts=MAX_LOGIN_ATTEMPTS,
        login_window=LOGIN_ATTEMPT_WINDOW_MINUTES,
    )


def obtener_roles(conn):
    return conn.execute("""
        SELECT id, nombre, descripcion, sistema
        FROM roles
        ORDER BY sistema DESC, nombre ASC
    """).fetchall()


def permisos_desde_formulario():
    return normalizar_permisos(request.form.getlist("permisos"))


@app.route("/roles")
def listar_roles():
    check = rol_requerido("admin")
    if check:
        return check

    conn = get_connection()
    roles = conn.execute("""
        SELECT
            r.id,
            r.nombre,
            r.descripcion,
            r.permisos,
            r.sistema,
            COUNT(u.id) AS usuarios
        FROM roles r
        LEFT JOIN usuarios u ON u.rol = r.nombre
        GROUP BY r.id, r.nombre, r.descripcion, r.permisos, r.sistema
        ORDER BY r.sistema DESC, r.nombre ASC
    """).fetchall()
    conn.close()

    roles = [dict(rol) for rol in roles]
    for rol in roles:
        rol["cantidad_permisos"] = len(deserializar_permisos(rol["permisos"], rol["nombre"]))

    return render_template("roles.html", roles=roles)


@app.route("/roles/nuevo", methods=["GET", "POST"])
def nuevo_rol():
    check = rol_requerido("admin")
    if check:
        return check

    if request.method == "POST":
        nombre = request.form.get("nombre", "").strip().lower()
        descripcion = request.form.get("descripcion", "").strip()
        permisos = permisos_desde_formulario()

        if not nombre:
            flash("El nombre del rol es obligatorio.", "error")
            return render_template(
                "rol_form.html",
                rol={"nombre": nombre, "descripcion": descripcion},
                grupos=grupos_permisos(),
                permisos_seleccionados=permisos,
                modo="nuevo",
            )

        if not re.match(r"^[a-z0-9_ -]+$", nombre):
            flash("El rol solo puede usar letras, números, espacios, guiones o guiones bajos.", "error")
            return render_template(
                "rol_form.html",
                rol={"nombre": nombre, "descripcion": descripcion},
                grupos=grupos_permisos(),
                permisos_seleccionados=permisos,
                modo="nuevo",
            )

        conn = get_connection()
        existente = conn.execute("""
            SELECT id
            FROM roles
            WHERE nombre = %s
        """, (nombre,)).fetchone()

        if existente:
            conn.close()
            flash("Ya existe un rol con ese nombre.", "error")
            return render_template(
                "rol_form.html",
                rol={"nombre": nombre, "descripcion": descripcion},
                grupos=grupos_permisos(),
                permisos_seleccionados=permisos,
                modo="nuevo",
            )

        conn.execute("""
            INSERT INTO roles (nombre, descripcion, permisos, sistema)
            VALUES (%s, %s, %s, 0)
        """, (nombre, descripcion, serializar_permisos(permisos)))
        conn.commit()
        conn.close()

        flash("Rol creado correctamente.", "ok")
        return redirect(url_for("listar_roles"))

    return render_template(
        "rol_form.html",
        rol=None,
        grupos=grupos_permisos(),
        permisos_seleccionados=[],
        modo="nuevo",
    )


@app.route("/roles/<int:rol_id>/editar", methods=["GET", "POST"])
def editar_rol(rol_id):
    check = rol_requerido("admin")
    if check:
        return check

    conn = get_connection()
    rol = conn.execute("""
        SELECT *
        FROM roles
        WHERE id = %s
    """, (rol_id,)).fetchone()

    if rol is None:
        conn.close()
        flash("Rol no encontrado.", "error")
        return redirect(url_for("listar_roles"))

    if rol["sistema"]:
        conn.close()
        flash("Los roles base no se editan. Creá un rol personalizado si necesitás otra combinación.", "error")
        return redirect(url_for("listar_roles"))

    if request.method == "POST":
        descripcion = request.form.get("descripcion", "").strip()
        permisos = permisos_desde_formulario()
        conn.execute("""
            UPDATE roles
            SET descripcion = %s,
                permisos = %s
            WHERE id = %s
        """, (descripcion, serializar_permisos(permisos), rol_id))
        conn.commit()
        conn.close()
        flash("Rol actualizado correctamente. Los usuarios con este rol deberán volver a iniciar sesión para tomar los nuevos permisos.", "ok")
        return redirect(url_for("listar_roles"))

    permisos_actuales = deserializar_permisos(rol["permisos"], rol["nombre"])
    conn.close()
    return render_template(
        "rol_form.html",
        rol=rol,
        grupos=grupos_permisos(),
        permisos_seleccionados=permisos_actuales,
        modo="editar",
    )


@app.route("/roles/<int:rol_id>/eliminar", methods=["POST"])
def eliminar_rol(rol_id):
    check = rol_requerido("admin")
    if check:
        return check

    conn = get_connection()
    rol = conn.execute("""
        SELECT *
        FROM roles
        WHERE id = %s
    """, (rol_id,)).fetchone()

    if rol is None:
        conn.close()
        flash("Rol no encontrado.", "error")
        return redirect(url_for("listar_roles"))

    if rol["sistema"]:
        conn.close()
        flash("No se pueden eliminar los roles base del sistema.", "error")
        return redirect(url_for("listar_roles"))

    usuarios = conn.execute("""
        SELECT COUNT(*) AS total
        FROM usuarios
        WHERE rol = %s
    """, (rol["nombre"],)).fetchone()["total"]

    if usuarios:
        conn.close()
        flash("No se puede eliminar un rol asignado a usuarios.", "error")
        return redirect(url_for("listar_roles"))

    conn.execute("DELETE FROM roles WHERE id = %s", (rol_id,))
    conn.commit()
    conn.close()
    flash("Rol eliminado correctamente.", "ok")
    return redirect(url_for("listar_roles"))


@app.route("/usuarios/nuevo", methods=["GET", "POST"])
def nuevo_usuario():
    check = rol_requerido("admin")
    if check:
        return check

    conn = get_connection()
    roles = obtener_roles(conn)

    if request.method == "POST":
        username = normalizar_username(request.form.get("username", ""))
        email = normalizar_email(request.form.get("email", ""))
        password = request.form.get("password", "")
        rol = request.form.get("rol", "tesorero")

        rol_existe = conn.execute("""
            SELECT id
            FROM roles
            WHERE nombre = %s
        """, (rol,)).fetchone()

        if not rol_existe:
            conn.close()
            flash("Rol invalido.", "error")
            return render_template("usuario_form.html", username=username, email=email, rol=rol, roles=roles)

        if not username or not password:
            conn.close()
            flash("Usuario y clave son obligatorios.", "error")
            return render_template("usuario_form.html", username=username, email=email, rol=rol, roles=roles)

        password_hash = generate_password_hash(password)

        existente = conn.execute("""
            SELECT id FROM usuarios
            WHERE lower(username) = %s
               OR (email IS NOT NULL AND email <> '' AND lower(email) = %s)
        """, (username, email)).fetchone()

        if existente:
            conn.close()
            flash("Ese usuario o email ya existe.", "error")
            return render_template("usuario_form.html", username=username, email=email, rol=rol, roles=roles)

        conn.execute("""
            INSERT INTO usuarios (username, email, password, rol, debe_cambiar_password, onboarding_visto)
            VALUES (%s, %s, %s, %s, 1, 0)
        """, (username, email or None, password_hash, rol))

        conn.commit()
        conn.close()

        flash(f"Usuario creado correctamente con rol {rol}.", "ok")
        return redirect(url_for("listar_usuarios"))

    conn.close()
    return render_template("usuario_form.html", roles=roles)


@app.route("/usuarios")
def listar_usuarios():
    check = rol_requerido("admin")
    if check:
        return check

    conn = get_connection()

    usuarios = conn.execute("""
        SELECT id, username, email, rol, debe_cambiar_password, onboarding_visto, ultimo_login
        FROM usuarios
        ORDER BY username
    """).fetchall()

    conn.close()

    return render_template("usuarios.html", usuarios=usuarios)


@app.route("/usuarios/<int:usuario_id>/editar", methods=["GET", "POST"])
def editar_usuario(usuario_id):
    check = rol_requerido("admin")
    if check:
        return check

    conn = get_connection()
    usuario = conn.execute("""
        SELECT id, username, email, rol
        FROM usuarios
        WHERE id = %s
    """, (usuario_id,)).fetchone()

    if usuario is None:
        conn.close()
        flash("Usuario no encontrado.", "error")
        return redirect(url_for("listar_usuarios"))

    roles = obtener_roles(conn)

    if request.method == "POST":
        username = normalizar_username(request.form.get("username", ""))
        email = normalizar_email(request.form.get("email", ""))
        rol = request.form.get("rol", "").strip()
        usuario_form = {
            **usuario,
            "username": username,
            "email": email,
            "rol": rol,
        }

        if not username:
            conn.close()
            flash("El usuario es obligatorio.", "error")
            return render_template("usuario_form.html", usuario=usuario_form, rol=rol, roles=roles, modo="editar")

        rol_existe = conn.execute("""
            SELECT id
            FROM roles
            WHERE nombre = %s
        """, (rol,)).fetchone()

        if not rol_existe:
            conn.close()
            flash("Rol invalido.", "error")
            return render_template("usuario_form.html", usuario=usuario_form, rol=rol, roles=roles, modo="editar")

        email_param = email or None
        if email_param:
            duplicado = conn.execute("""
                SELECT id
                FROM usuarios
                WHERE id <> %s
                  AND (
                      lower(username) = %s
                      OR lower(email) = %s
                  )
                LIMIT 1
            """, (usuario_id, username, email_param)).fetchone()
        else:
            duplicado = conn.execute("""
                SELECT id
                FROM usuarios
                WHERE id <> %s
                  AND lower(username) = %s
                LIMIT 1
            """, (usuario_id, username)).fetchone()

        if duplicado:
            conn.close()
            flash("Ese usuario o email ya existe.", "error")
            return render_template("usuario_form.html", usuario=usuario_form, rol=rol, roles=roles, modo="editar")

        conn.execute("""
            UPDATE usuarios
            SET username = %s,
                email = %s,
                rol = %s
            WHERE id = %s
        """, (username, email_param, rol, usuario_id))
        permisos_actualizados = cargar_permisos_rol(conn, rol)
        conn.commit()
        conn.close()

        if session.get("user_id") == usuario_id:
            session["username"] = username
            session["rol"] = rol
            session["permisos"] = permisos_actualizados
            flash("Tu usuario fue actualizado.", "ok")
        else:
            flash("Usuario actualizado correctamente. Si estaba conectado, debera volver a iniciar sesion para tomar los cambios.", "ok")
        return redirect(url_for("listar_usuarios"))

    conn.close()
    return render_template("usuario_form.html", usuario=usuario, rol=usuario["rol"], roles=roles, modo="editar")


@app.route("/usuarios/<int:usuario_id>/password", methods=["GET", "POST"])
def resetear_password_usuario(usuario_id):
    check = rol_requerido("admin")
    if check:
        return check

    if session.get("user_id") == usuario_id:
        return redirect(url_for("cambiar_mi_password"))

    conn = get_connection()
    usuario = conn.execute("""
        SELECT id, username, email, rol
        FROM usuarios
        WHERE id = %s
    """, (usuario_id,)).fetchone()

    if usuario is None:
        conn.close()
        flash("Usuario no encontrado.", "error")
        return redirect(url_for("listar_usuarios"))

    if request.method == "POST":
        password_nueva = request.form.get("password_nueva", "")
        password_confirmacion = request.form.get("password_confirmacion", "")

        error = validar_password_nueva(password_nueva, password_confirmacion)
        if error:
            conn.close()
            flash(error, "error")
            return render_template("password_form.html", usuario=usuario, modo="admin")

        conn.execute("""
            UPDATE usuarios
            SET password = %s,
                debe_cambiar_password = 1
            WHERE id = %s
        """, (generate_password_hash(password_nueva), usuario_id))
        conn.commit()
        conn.close()

        flash(f"Clave de {usuario['username']} actualizada. Debera cambiarla al ingresar.", "ok")
        return redirect(url_for("listar_usuarios"))

    conn.close()
    return render_template("password_form.html", usuario=usuario, modo="admin")

@app.route("/usuarios/<int:usuario_id>/eliminar", methods=["POST"])
def eliminar_usuario(usuario_id):
    check = rol_requerido("admin")
    if check:
        return check

    if session.get("user_id") == usuario_id:
        flash("No podés eliminar tu propio usuario mientras estás logueado.", "error")
        return redirect(url_for("listar_usuarios"))

    conn = get_connection()
    conn.execute("DELETE FROM password_reset_tokens WHERE usuario_id = %s", (usuario_id,))
    conn.execute("DELETE FROM usuarios WHERE id = %s", (usuario_id,))
    conn.commit()
    conn.close()

    flash("Usuario eliminado correctamente.", "ok")
    return redirect(url_for("listar_usuarios"))

@app.route("/cuotas/<int:cuota_id>/eliminar", methods=["POST"])
def eliminar_cuota(cuota_id):
    check = permiso_requerido("cuotas_gestionar")
    if check:
        return check

    conn = get_connection()

    cuota = conn.execute("""
        SELECT jugador_id
        FROM cuotas
        WHERE id = %s
    """, (cuota_id,)).fetchone()

    if cuota is None:
        conn.close()
        flash("Cuota no encontrada.", "error")
        return redirect(url_for("listar_jugadores"))

    conn.execute("DELETE FROM cuotas WHERE id = %s", (cuota_id,))
    conn.commit()
    conn.close()

    flash("Cuota eliminada correctamente.", "ok")
    return redirect(url_for("ver_cuotas", jugador_id=cuota["jugador_id"]))

def generar_recibo_pdf(cuota_id):
    conn = get_connection()

    datos = conn.execute("""
        SELECT
            c.id AS cuota_id,
            c.periodo,
            c.importe,
            c.fecha_pago,
            c.fecha_vencimiento,
            c.numero_recibo,
            c.metodo_pago,
            c.referencia_pago,
            c.importe_original,
            c.descuento_beca,
            c.beca_porcentaje,
            c.becada,
            j.nombre,
            j.apellido,
            j.dni,
            j.categoria
        FROM cuotas c
        JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.id = %s
    """, (cuota_id,)).fetchone()

    conn.close()

    if datos is None:
        return None

    recibos_dir = BASE_DIR / "recibos"
    recibos_dir.mkdir(exist_ok=True)

    archivo = recibos_dir / f"recibo_cuota_{cuota_id}.pdf"

    pdf = canvas.Canvas(str(archivo), pagesize=A4)
    width, height = A4

    # Encabezado con logo
    logo_path = BASE_DIR / "static" / "img" / "logo.png"

    if logo_path.exists():
        pdf.drawImage(
            ImageReader(str(logo_path)),
            25 * mm,
            height - 35 * mm,
            width=22 * mm,
            height=22 * mm,
            preserveAspectRatio=True,
            mask="auto"
        )
        texto_x = 52 * mm
    else:
        texto_x = 25 * mm

    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(texto_x, height - 22 * mm, "Ruda Macho Rugby Club")

    pdf.setFont("Helvetica", 10)
    pdf.drawString(texto_x, height - 29 * mm, "Recibo interno no válido como factura")

    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawRightString(
        185 * mm,
        height - 22 * mm,
        f"N° {datos['numero_recibo'] or datos['cuota_id']}"
    )

    pdf.setFont("Helvetica", 10)
    pdf.drawRightString(
        185 * mm,
        height - 29 * mm,
        f"Emitido: {datetime.now().strftime('%d/%m/%Y')}"
    )

    pdf.line(25 * mm, height - 42 * mm, 185 * mm, height - 42 * mm)

    # Título
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(25 * mm, height - 58 * mm, "RECIBO DE CUOTA")

    # Cuerpo
    pdf.setFont("Helvetica", 12)
    y = height - 78 * mm

    importe_formateado = f"${int(float(datos['importe'])):,}".replace(",", ".")

    pdf.drawString(25 * mm, y, f"Jugador: {datos['apellido']}, {datos['nombre']}")
    y -= 10 * mm

    pdf.drawString(25 * mm, y, f"DNI: {datos['dni'] or '-'}")
    y -= 10 * mm

    pdf.drawString(25 * mm, y, f"Categoría: {datos['categoria'] or '-'}")
    y -= 10 * mm

    pdf.drawString(25 * mm, y, f"Periodo abonado: {datos['periodo']}")
    y -= 10 * mm

    pdf.drawString(25 * mm, y, f"Importe abonado: {importe_formateado}")
    y -= 10 * mm

    if datos["becada"]:
        original = f"${int(float(datos['importe_original'] or datos['importe'])):,}".replace(",", ".")
        descuento = f"${int(float(datos['descuento_beca'] or 0)):,}".replace(",", ".")
        porcentaje_beca_pdf = float(datos["beca_porcentaje"] or 0)
        pdf.drawString(
            25 * mm,
            y,
            f"Beca aplicada: {porcentaje_beca_pdf:g}% - Original {original} - Descuento {descuento}"
        )
        y -= 10 * mm

    pdf.drawString(25 * mm, y, f"Fecha de pago: {datos['fecha_pago'] or '-'}")
    y -= 10 * mm

    pdf.drawString(25 * mm, y, f"Vencimiento original: {datos['fecha_vencimiento'] or '-'}")
    y -= 10 * mm

    pdf.drawString(25 * mm, y, f"Método de pago: {datos['metodo_pago'] or '-'}")
    y -= 10 * mm

    pdf.drawString(25 * mm, y, f"Referencia: {datos['referencia_pago'] or '-'}")
    y -= 25 * mm

    # Firma
    pdf.line(25 * mm, y, 90 * mm, y)
    pdf.drawString(25 * mm, y - 7 * mm, "Firma / aclaración")

    # Pie
    pdf.setFont("Helvetica-Oblique", 9)
    pdf.drawString(25 * mm, 20 * mm, "Ruda Macho Rugby Club - Sistema Integral de Gestion")

    pdf.save()

    return archivo

@app.route("/cuotas/<int:cuota_id>/recibo")
def descargar_recibo(cuota_id):
    check = permiso_requerido("cuotas_ver")
    if check:
        return check

    archivo = BASE_DIR / "recibos" / f"recibo_cuota_{cuota_id}.pdf"

    if not archivo.exists():
        archivo = generar_recibo_pdf(cuota_id)

    if archivo is None or not archivo.exists():
        flash("No se pudo generar el recibo.", "error")
        return redirect(url_for("listar_jugadores"))

    registrar_auditoria(
        "descargar_ok",
        "recibo",
        str(cuota_id),
        {"archivo": archivo.name},
    )

    return send_file(
        archivo,
        as_attachment=True,
        download_name=f"recibo_cuota_{cuota_id}.pdf"
    )

@app.route("/caja/exportar")
def exportar_caja():
    check = permiso_requerido("caja_ver")
    if check:
        return check

    mes = request.args.get("mes")
    mes_actual = mes or datetime.now().strftime("%Y-%m")

    conn = get_connection()

    movimientos = conn.execute("""
        SELECT *
        FROM movimientos
        WHERE substring(fecha from 1 for 7) = %s
          AND COALESCE(anulado, 0) = 0
        ORDER BY fecha ASC, id ASC
    """, (mes_actual,)).fetchall()

    ingresos_mes = conn.execute("""
        SELECT COALESCE(SUM(monto), 0) AS total
        FROM movimientos
        WHERE tipo = 'ingreso'
          AND COALESCE(anulado, 0) = 0
          AND substring(fecha from 1 for 7) = %s
    """, (mes_actual,)).fetchone()["total"]

    egresos_mes = conn.execute("""
        SELECT COALESCE(SUM(monto), 0) AS total
        FROM movimientos
        WHERE tipo = 'egreso'
          AND COALESCE(anulado, 0) = 0
          AND substring(fecha from 1 for 7) = %s
    """, (mes_actual,)).fetchone()["total"]

    conn.close()

    resultado_mes = ingresos_mes - egresos_mes

    wb = Workbook()
    ws = wb.active
    ws.title = "Caja"

    ws["A1"] = "Ruda Macho Rugby Club"
    ws["A2"] = f"Balance de caja - {mes_actual}"

    ws["A4"] = "Ingresos del mes"
    ws["B4"] = ingresos_mes
    ws["A5"] = "Egresos del mes"
    ws["B5"] = egresos_mes
    ws["A6"] = "Resultado del mes"
    ws["B6"] = resultado_mes

    encabezados = ["Fecha", "Tipo", "Concepto", "Referencia", "Monto"]
    for col, encabezado in enumerate(encabezados, start=1):
        celda = ws.cell(row=8, column=col)
        celda.value = encabezado
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F2937")
        celda.alignment = Alignment(horizontal="center")

    fila = 9
    for m in movimientos:
        ws.cell(row=fila, column=1).value = m["fecha"]
        ws.cell(row=fila, column=2).value = m["tipo"]
        ws.cell(row=fila, column=3).value = m["concepto"]
        ws.cell(row=fila, column=4).value = m["referencia"]
        ws.cell(row=fila, column=5).value = m["monto"]

        if m["tipo"] == "ingreso":
            ws.cell(row=fila, column=2).font = Font(color="166534", bold=True)
            ws.cell(row=fila, column=5).font = Font(color="166534", bold=True)
        else:
            ws.cell(row=fila, column=2).font = Font(color="DC2626", bold=True)
            ws.cell(row=fila, column=5).font = Font(color="DC2626", bold=True)

        fila += 1

    for row in range(4, 7):
        ws.cell(row=row, column=1).font = Font(bold=True)
        ws.cell(row=row, column=2).number_format = '$ #,##0'

    for row in range(9, fila):
        ws.cell(row=row, column=5).number_format = '$ #,##0'

    thin = Side(style="thin", color="D1D5DB")
    for row in ws.iter_rows(min_row=8, max_row=max(fila - 1, 8), min_col=1, max_col=5):
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"].font = Font(size=12)

    widths = {
        "A": 14,
        "B": 14,
        "C": 40,
        "D": 32,
        "E": 14,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "A9"

    export_dir = BASE_DIR / "exports"
    export_dir.mkdir(exist_ok=True)

    archivo = export_dir / f"caja_{mes_actual}.xlsx"
    wb.save(archivo)

    registrar_auditoria(
        "exportar_ok",
        "caja",
        mes_actual,
        {"formato": "xlsx", "cantidad_registros": len(movimientos)},
    )

    return send_file(
        archivo,
        as_attachment=True,
        download_name=f"caja_{mes_actual}.xlsx"
    )

@app.route("/movimientos/<int:movimiento_id>/editar", methods=["GET", "POST"])
def editar_movimiento(movimiento_id):
    check = permiso_requerido("caja_gestionar")
    if check:
        return check

    conn = get_connection()

    movimiento = conn.execute("""
        SELECT *
        FROM movimientos
        WHERE id = %s
    """, (movimiento_id,)).fetchone()

    if movimiento is None:
        conn.close()
        flash("Movimiento no encontrado.", "error")
        return redirect(url_for("ver_caja"))

    if movimiento["anulado"]:
        conn.close()
        flash("No se puede editar un movimiento anulado.", "error")
        return redirect(url_for("ver_caja", mes=movimiento["fecha"][:7]))

    if request.method == "POST":
        tipo = request.form.get("tipo")
        concepto = request.form.get("concepto")
        monto = request.form.get("monto")
        fecha = validar_fecha_movimiento(request.form.get("fecha", "").strip())
        referencia = request.form.get("referencia")

        if not fecha:
            flash("La fecha del movimiento no es válida.", "error")
            movimiento_form = dict(movimiento)
            movimiento_form.update({
                "tipo": tipo,
                "concepto": concepto,
                "monto": monto,
                "fecha": request.form.get("fecha", "").strip(),
                "referencia": referencia,
            })
            conn.close()
            return render_template("movimiento_form_editar.html", movimiento=movimiento_form)

        mes_destino = fecha[:7]
        if mes_esta_cerrado(mes_destino):
            flash("No se puede mover un movimiento a un mes cerrado.", "error")
            movimiento_form = dict(movimiento)
            movimiento_form.update({
                "tipo": tipo,
                "concepto": concepto,
                "monto": monto,
                "fecha": fecha,
                "referencia": referencia,
            })
            conn.close()
            return render_template("movimiento_form_editar.html", movimiento=movimiento_form)

        conn.execute("""
            UPDATE movimientos
            SET tipo = %s, concepto = %s, monto = %s, fecha = %s, referencia = %s
            WHERE id = %s
        """, (tipo, concepto, monto, fecha, referencia, movimiento_id))

        conn.commit()
        conn.close()

        flash("Movimiento actualizado.", "ok")
        return redirect(url_for("ver_caja", mes=mes_destino))

    mes_movimiento = movimiento["fecha"][:7]

    if mes_esta_cerrado(mes_movimiento):
        conn.close()
        flash("No se puede editar un movimiento de un mes cerrado.", "error")
        return redirect(url_for("ver_caja", mes=mes_movimiento))

    conn.close()
    return render_template("movimiento_form_editar.html", movimiento=movimiento)

@app.route("/movimientos/<int:movimiento_id>/eliminar", methods=["POST"])
def eliminar_movimiento(movimiento_id):
    check = permiso_requerido("caja_gestionar")
    if check:
        return check

    conn = get_connection()

    movimiento = conn.execute("""
        SELECT *
        FROM movimientos
        WHERE id = %s
    """, (movimiento_id,)).fetchone()

    if movimiento is None:
        conn.close()
        flash("Movimiento no encontrado.", "error")
        return redirect(url_for("ver_caja"))

    if movimiento["anulado"]:
        conn.close()
        flash("Ese movimiento ya estaba anulado.", "error")
        return redirect(url_for("ver_caja", mes=movimiento["fecha"][:7]))

    mes_movimiento = movimiento["fecha"][:7]
    if mes_esta_cerrado(mes_movimiento):
        conn.close()
        flash("No se puede anular un movimiento de un mes cerrado.", "error")
        return redirect(url_for("ver_caja", mes=mes_movimiento))

    motivo = request.form.get("motivo_anulacion", "").strip()
    if not motivo:
        conn.close()
        flash("Debe indicar un motivo para anular el movimiento.", "error")
        return redirect(url_for("ver_caja", mes=mes_movimiento))

    conn.execute("""
        UPDATE movimientos
        SET anulado = 1,
            fecha_anulacion = CURRENT_TIMESTAMP,
            usuario_anulacion = %s,
            motivo_anulacion = %s
        WHERE id = %s
    """, (session.get("username"), motivo, movimiento_id))
    conn.commit()
    conn.close()

    registrar_auditoria(
        "anular_detalle",
        "movimiento_caja",
        str(movimiento_id),
        {
            "fecha": movimiento["fecha"],
            "tipo": movimiento["tipo"],
            "concepto": movimiento["concepto"],
            "monto": movimiento["monto"],
            "referencia": movimiento["referencia"],
            "motivo_anulacion": motivo,
            "usuario_anulacion": session.get("username"),
        },
    )

    flash("Movimiento anulado correctamente.", "ok")
    return redirect(url_for("ver_caja", mes=mes_movimiento))

def mes_esta_cerrado(mes):
    conn = get_connection()
    cierre = conn.execute("""
        SELECT id
        FROM cierres_mensuales
        WHERE mes = %s
    """, (mes,)).fetchone()
    conn.close()
    return cierre is not None

@app.route("/caja/cerrar", methods=["POST"])
def cerrar_mes():
    check = permiso_requerido("caja_gestionar")
    if check:
        return check

    mes = request.form.get("mes", "").strip()

    if not mes:
        flash("Debe indicar un mes para cerrar.", "error")
        return redirect(url_for("ver_caja"))

    conn = get_connection()

    cierre_existente = conn.execute("""
        SELECT id
        FROM cierres_mensuales
        WHERE mes = %s
    """, (mes,)).fetchone()

    if cierre_existente:
        conn.close()
        flash("Ese mes ya está cerrado.", "error")
        return redirect(url_for("ver_caja", mes=mes))

    ingresos = conn.execute("""
        SELECT COALESCE(SUM(monto), 0) AS total
        FROM movimientos
        WHERE tipo = 'ingreso'
          AND COALESCE(anulado, 0) = 0
          AND substring(fecha from 1 for 7) = %s
    """, (mes,)).fetchone()["total"]

    egresos = conn.execute("""
        SELECT COALESCE(SUM(monto), 0) AS total
        FROM movimientos
        WHERE tipo = 'egreso'
          AND COALESCE(anulado, 0) = 0
          AND substring(fecha from 1 for 7) = %s
    """, (mes,)).fetchone()["total"]

    resultado = ingresos - egresos

    conn.execute("""
        INSERT INTO cierres_mensuales (
            mes, ingresos, egresos, resultado, fecha_cierre, usuario
        )
        VALUES (%s, %s, %s, %s, CURRENT_TIMESTAMP, %s)
    """, (
        mes,
        ingresos,
        egresos,
        resultado,
        session.get("username")
    ))

    conn.commit()
    conn.close()

    flash(f"Mes {mes} cerrado correctamente.", "ok")
    return redirect(url_for("ver_caja", mes=mes))


@app.route("/tests")
def listar_tests():
    check = permiso_requerido("tests_ver")
    if check:
        return check

    conn = get_connection()

    tests = conn.execute("""
        SELECT
            t.*,
            COUNT(r.id) AS mediciones,
            COUNT(DISTINCT r.jugador_id) AS jugadores_medidos,
            MAX(r.fecha) AS ultima_fecha
        FROM test_tipos t
        LEFT JOIN test_resultados r ON r.test_id = t.id
        GROUP BY t.id
        ORDER BY t.activo DESC, t.nombre
    """).fetchall()

    recientes = conn.execute("""
        SELECT
            r.*,
            t.nombre AS test_nombre,
            t.unidad,
            j.apellido,
            j.nombre,
            j.categoria
        FROM test_resultados r
        JOIN test_tipos t ON t.id = r.test_id
        JOIN jugadores j ON j.id = r.jugador_id
        ORDER BY r.fecha DESC, r.id DESC
        LIMIT 20
    """).fetchall()

    conn.close()

    return render_template("tests.html", tests=tests, recientes=recientes)


@app.route("/tests/nuevo", methods=["GET", "POST"])
def nuevo_test_tipo():
    check = permiso_requerido("tests_gestionar")
    if check:
        return check

    if request.method == "POST":
        nombre = request.form.get("nombre", "").strip()
        descripcion = request.form.get("descripcion", "").strip()
        unidad = request.form.get("unidad", "").strip()
        puntaje_min = parsear_puntaje_test(request.form.get("puntaje_min"))
        puntaje_max = parsear_puntaje_test(request.form.get("puntaje_max"))
        mayor_es_mejor = 1 if request.form.get("mayor_es_mejor") else 0
        activo = 1 if request.form.get("activo") else 0

        if not nombre:
            flash("El nombre del test es obligatorio.", "error")
            return render_template("test_form.html", test=request.form)

        if puntaje_min is not None and puntaje_max is not None and puntaje_min > puntaje_max:
            flash("El puntaje minimo no puede ser mayor al maximo.", "error")
            return render_template("test_form.html", test=request.form)

        conn = get_connection()
        try:
            conn.execute("""
                INSERT INTO test_tipos (
                    nombre, descripcion, unidad, puntaje_min, puntaje_max,
                    mayor_es_mejor, activo, creado_por
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                nombre,
                descripcion,
                unidad,
                puntaje_min,
                puntaje_max,
                mayor_es_mejor,
                activo,
                session.get("username"),
            ))
            conn.commit()
        except Exception:
            conn.rollback()
            conn.close()
            flash("No se pudo crear el test. Revisá que el nombre no esté repetido.", "error")
            return render_template("test_form.html", test=request.form)

        conn.close()
        flash("Test creado correctamente.", "ok")
        return redirect(url_for("listar_tests"))

    return render_template("test_form.html", test={"mayor_es_mejor": 1, "activo": 1})


@app.route("/tests/<int:test_id>/cargar", methods=["GET", "POST"])
def cargar_test_resultados(test_id):
    check = permiso_requerido("tests_gestionar")
    if check:
        return check

    conn = get_connection()

    test = conn.execute("""
        SELECT *
        FROM test_tipos
        WHERE id = %s
    """, (test_id,)).fetchone()

    if test is None:
        conn.close()
        flash("Test no encontrado.", "error")
        return redirect(url_for("listar_tests"))

    jugadores = conn.execute("""
        SELECT id, apellido, nombre, dni, categoria, estado
        FROM jugadores
        WHERE COALESCE(estado, 'Activo') <> 'Baja'
        ORDER BY categoria, apellido, nombre
    """).fetchall()

    if request.method == "POST":
        fecha = normalizar_fecha_test(request.form.get("fecha"))
        cargados = 0
        omitidos = 0

        for jugador in jugadores:
            puntaje = parsear_puntaje_test(request.form.get(f"puntaje_{jugador['id']}"))
            observaciones = request.form.get(f"obs_{jugador['id']}", "").strip()

            if puntaje is None:
                omitidos += 1
                continue

            conn.execute("""
                INSERT INTO test_resultados (
                    test_id, jugador_id, fecha, puntaje, observaciones, creado_por
                )
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (
                test_id,
                jugador["id"],
                fecha,
                puntaje,
                observaciones,
                session.get("username"),
            ))
            cargados += 1

        conn.commit()
        conn.close()

        if cargados:
            flash(f"Se cargaron {cargados} mediciones.", "ok")
            if omitidos:
                flash(f"{omitidos} jugadores quedaron sin puntaje en esta carga.", "info")
            return redirect(url_for("graficos_tests", test_id=test_id))

        flash("No se cargaron puntajes. Completá al menos una medición.", "error")
        return redirect(url_for("cargar_test_resultados", test_id=test_id))

    conn.close()
    return render_template(
        "test_carga.html",
        test=test,
        jugadores=jugadores,
        fecha_hoy=datetime.now().strftime("%Y-%m-%d"),
    )


@app.route("/tests/importar", methods=["GET", "POST"])
def importar_test_resultados():
    check = permiso_requerido("tests_gestionar")
    if check:
        return check

    conn = get_connection()
    tests = obtener_test_tipos(conn, solo_activos=True)
    batches_recientes = obtener_test_importaciones_batch_recientes(conn)

    if request.method == "POST":
        archivo = request.files.get("archivo")
        test_id_form = request.form.get("test_id", "").strip()
        test_id_fijo = int(test_id_form) if test_id_form.isdigit() else None

        if not archivo or not archivo.filename:
            flash("Seleccioná un archivo Excel para importar.", "error")
            conn.close()
            return render_template("test_importar.html", tests=tests, batches_recientes=batches_recientes)

        try:
            wb = load_workbook(archivo, read_only=True, data_only=True)
            ws = wb.active
            filas = list(ws.iter_rows(values_only=True))
        except Exception:
            conn.close()
            flash("No se pudo leer el Excel. Verificá el archivo.", "error")
            return render_template("test_importar.html", tests=tests, batches_recientes=batches_recientes)

        if not filas:
            conn.close()
            flash("El archivo no tiene filas para importar.", "error")
            return render_template("test_importar.html", tests=tests, batches_recientes=batches_recientes)

        headers = [normalizar_header_excel(valor) for valor in filas[0]]
        tests_por_nombre = {
            normalizar_header_excel(test["nombre"]): test["id"]
            for test in tests
        }

        jugadores = obtener_jugadores_selector(conn)
        batch_id = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{secrets.token_urlsafe(6)}"
        pendientes = 0

        for numero_fila, row in enumerate(filas[1:], start=2):
            datos = {
                headers[index]: row[index] if index < len(row) else None
                for index in range(len(headers))
            }
            nombre_completo = str(
                datos.get("jugador") or
                datos.get("nombre_completo") or
                datos.get("jugador_nombre") or
                ""
            ).strip()
            nombre = str(datos.get("nombre") or "").strip()
            apellido = str(datos.get("apellido") or "").strip()
            jugador_sugerido, confianza, motivo = sugerir_jugador_por_nombre_test(datos, jugadores)
            puntaje = parsear_puntaje_test(
                datos.get("puntaje") or datos.get("score") or datos.get("valor") or datos.get("resultado")
            )
            fecha = normalizar_fecha_test(datos.get("fecha"))
            observaciones = str(datos.get("observaciones") or datos.get("obs") or "").strip()

            test_id = test_id_fijo
            nombre_test = datos.get("test") or datos.get("test_nombre") or datos.get("nombre_test")
            if not test_id:
                test_id = tests_por_nombre.get(normalizar_header_excel(nombre_test))

            errores_fila = []
            if not jugador_sugerido:
                errores_fila.append("Revisar jugador")
            if not test_id:
                errores_fila.append("Revisar test")
            if puntaje is None:
                errores_fila.append("Revisar puntaje")

            error = ", ".join(errores_fila) if errores_fila else None
            conn.execute("""
                INSERT INTO test_importaciones_batch (
                    batch_id, estado, fila, test_id, test_nombre,
                    jugador_sugerido_id, confianza, motivo,
                    nombre_excel, apellido_excel, nombre_completo_excel,
                    fecha, puntaje, observaciones, error, creado_por
                )
                VALUES (%s, 'pendiente', %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                batch_id,
                numero_fila,
                test_id,
                str(nombre_test or "").strip() or None,
                jugador_sugerido["id"] if jugador_sugerido else None,
                confianza,
                motivo,
                nombre or None,
                apellido or None,
                nombre_completo or None,
                fecha,
                puntaje,
                observaciones or None,
                error,
                session.get("username"),
            ))
            pendientes += 1

        conn.commit()
        conn.close()

        if pendientes:
            flash(f"Se prepararon {pendientes} mediciones para revisar antes de confirmar.", "ok")
            return redirect(url_for("revisar_test_importacion", batch_id=batch_id))

        flash("El Excel no tenia filas para revisar.", "error")
        return redirect(url_for("importar_test_resultados"))
    conn.close()
    return render_template("test_importar.html", tests=tests, batches_recientes=batches_recientes)


@app.route("/tests/importar/<batch_id>/revisar", methods=["GET", "POST"])
def revisar_test_importacion(batch_id):
    check = permiso_requerido("tests_gestionar")
    if check:
        return check

    conn = get_connection()
    tests = obtener_test_tipos(conn, solo_activos=True)
    jugadores = obtener_jugadores_selector(conn)

    if request.method == "POST":
        item_ids = request.form.getlist("item_ids")
        procesadas = 0
        omitidas = 0
        errores = 0

        for item_id in item_ids:
            if request.form.get(f"procesar_{item_id}") != "on":
                omitidas += 1
                continue

            jugador_id = request.form.get(f"jugador_id_{item_id}", "").strip()
            test_id = request.form.get(f"test_id_{item_id}", "").strip()
            fecha = normalizar_fecha_test(request.form.get(f"fecha_{item_id}"))
            puntaje = parsear_puntaje_test(request.form.get(f"puntaje_{item_id}"))
            observaciones = request.form.get(f"observaciones_{item_id}", "").strip()

            if not jugador_id or not jugador_id.isdigit():
                errores += 1
                flash(f"La fila #{item_id} no tiene jugador asignado.", "error")
                continue
            if not test_id or not test_id.isdigit():
                errores += 1
                flash(f"La fila #{item_id} no tiene test asignado.", "error")
                continue
            if puntaje is None:
                errores += 1
                flash(f"La fila #{item_id} no tiene puntaje valido.", "error")
                continue

            item = conn.execute("""
                SELECT *
                FROM test_importaciones_batch
                WHERE id = %s AND batch_id = %s AND estado = 'pendiente'
            """, (item_id, batch_id)).fetchone()

            jugador = conn.execute("""
                SELECT id
                FROM jugadores
                WHERE id = %s
            """, (jugador_id,)).fetchone()

            test = conn.execute("""
                SELECT id
                FROM test_tipos
                WHERE id = %s
            """, (test_id,)).fetchone()

            if not item or not jugador or not test:
                errores += 1
                flash(f"No se encontro la fila pendiente #{item_id}, el jugador o el test.", "error")
                continue

            conn.execute("""
                INSERT INTO test_resultados (
                    test_id, jugador_id, fecha, puntaje, observaciones, creado_por
                )
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (
                int(test_id),
                int(jugador_id),
                fecha,
                puntaje,
                observaciones or item["observaciones"],
                session.get("username"),
            ))

            conn.execute("""
                UPDATE test_importaciones_batch
                SET estado = 'procesado',
                    test_id = %s,
                    jugador_id = %s,
                    fecha = %s,
                    puntaje = %s,
                    observaciones = %s,
                    procesado_en = CURRENT_TIMESTAMP,
                    procesado_por = %s,
                    error = NULL
                WHERE id = %s
            """, (
                int(test_id),
                int(jugador_id),
                fecha,
                puntaje,
                observaciones or item["observaciones"],
                session.get("username"),
                item["id"],
            ))
            procesadas += 1

        conn.commit()
        conn.close()

        if procesadas:
            flash(f"Se confirmaron {procesadas} mediciones.", "ok")
        if omitidas:
            flash(f"{omitidas} mediciones quedaron pendientes.", "warning")
        if errores:
            flash(f"{errores} mediciones requieren revision.", "error")
        return redirect(url_for("revisar_test_importacion", batch_id=batch_id))

    items = conn.execute("""
        SELECT
            b.*,
            js.apellido AS sugerido_apellido,
            js.nombre AS sugerido_nombre,
            js.dni AS sugerido_dni,
            ja.apellido AS asignado_apellido,
            ja.nombre AS asignado_nombre,
            t.nombre AS test_nombre_actual
        FROM test_importaciones_batch b
        LEFT JOIN jugadores js ON js.id = b.jugador_sugerido_id
        LEFT JOIN jugadores ja ON ja.id = b.jugador_id
        LEFT JOIN test_tipos t ON t.id = b.test_id
        WHERE b.batch_id = %s
        ORDER BY b.id
    """, (batch_id,)).fetchall()

    conn.close()

    if not items:
        flash("No se encontro la tanda de importacion.", "error")
        return redirect(url_for("importar_test_resultados"))

    return render_template(
        "test_importar_revision.html",
        batch_id=batch_id,
        items=items,
        tests=tests,
        jugadores=jugadores,
    )


@app.route("/tests/graficos")
def graficos_tests():
    check = permiso_requerido("tests_ver")
    if check:
        return check

    conn = get_connection()

    tests = obtener_test_tipos(conn, solo_activos=True)
    categorias = conn.execute("""
        SELECT DISTINCT categoria
        FROM jugadores
        WHERE categoria IS NOT NULL AND TRIM(categoria) <> ''
        ORDER BY categoria
    """).fetchall()
    jugadores = conn.execute("""
        SELECT id, apellido, nombre, categoria
        FROM jugadores
        WHERE COALESCE(estado, 'Activo') <> 'Baja'
        ORDER BY categoria, apellido, nombre
    """).fetchall()

    test_id_raw = request.args.get("test_id", "").strip()
    test_id = int(test_id_raw) if test_id_raw.isdigit() else None
    if not test_id and tests:
        test_id = tests[0]["id"]

    categoria = request.args.get("categoria", "").strip()
    desde = validar_fecha_movimiento(request.args.get("desde", "").strip())
    hasta = validar_fecha_movimiento(request.args.get("hasta", "").strip())
    jugadores_seleccionados = [
        int(valor)
        for valor in request.args.getlist("jugadores")
        if str(valor).isdigit()
    ]

    resultados = []
    if test_id:
        filtros = ["r.test_id = %s"]
        params = [test_id]

        if categoria:
            filtros.append("j.categoria = %s")
            params.append(categoria)
        if desde:
            filtros.append("r.fecha >= %s")
            params.append(desde)
        if hasta:
            filtros.append("r.fecha <= %s")
            params.append(hasta)
        if jugadores_seleccionados:
            filtros.append("j.id = ANY(%s)")
            params.append(jugadores_seleccionados)

        where = " AND ".join(filtros)
        resultados = conn.execute(f"""
            SELECT
                r.*,
                j.apellido,
                j.nombre,
                j.categoria
            FROM test_resultados r
            JOIN jugadores j ON j.id = r.jugador_id
            WHERE {where}
            ORDER BY j.apellido, j.nombre, r.fecha
        """, params).fetchall()

    grafico = construir_grafico_tests(resultados)
    test_actual = next((test for test in tests if test["id"] == test_id), None)

    conn.close()
    return render_template(
        "tests_graficos.html",
        tests=tests,
        test_actual=test_actual,
        categorias=categorias,
        jugadores=jugadores,
        jugadores_seleccionados=jugadores_seleccionados,
        categoria=categoria,
        desde=desde or "",
        hasta=hasta or "",
        grafico=grafico,
        resultados=resultados,
        test_id=test_id,
    )


@app.route("/asistencia")
def listar_eventos_asistencia():
    check = permiso_requerido("asistencia_ver")
    if check:
        return check

    conn = get_connection()

    eventos = conn.execute("""
        SELECT *
        FROM eventos_asistencia
        ORDER BY fecha DESC, id DESC
    """).fetchall()

    conn.close()

    return render_template("asistencia_eventos.html", eventos=eventos)


@app.route("/asistencia/nuevo", methods=["GET", "POST"])
def nuevo_evento_asistencia():
    check = permiso_requerido("asistencia_gestionar")
    if check:
        return check

    if request.method == "POST":
        fecha = request.form.get("fecha", "").strip()
        tipo = request.form.get("tipo", "").strip()
        descripcion = request.form.get("descripcion", "").strip()

        if not fecha or not tipo:
            flash("Fecha y tipo son obligatorios.", "error")
            return render_template("asistencia_evento_form.html")

        conn = get_connection()
        conn.execute("""
            INSERT INTO eventos_asistencia (fecha, tipo, descripcion)
            VALUES (%s, %s, %s)
        """, (fecha, tipo, descripcion))
        conn.commit()
        conn.close()

        flash("Evento de asistencia creado.", "ok")
        return redirect(url_for("listar_eventos_asistencia"))

    return render_template("asistencia_evento_form.html")


@app.route("/asistencia/<int:evento_id>", methods=["GET", "POST"])
def tomar_asistencia(evento_id):
    check = permiso_requerido("asistencia_ver")
    if check:
        return check

    conn = get_connection()

    evento = conn.execute("""
        SELECT *
        FROM eventos_asistencia
        WHERE id = %s
    """, (evento_id,)).fetchone()

    if evento is None:
        conn.close()
        flash("Evento no encontrado.", "error")
        return redirect(url_for("listar_eventos_asistencia"))

    jugadores = conn.execute("""
        SELECT *
        FROM jugadores
        WHERE estado = 'Activo'
        ORDER BY apellido, nombre
    """).fetchall()

    aspirantes = conn.execute("""
        SELECT *
        FROM aspirantes
        WHERE estado = 'Aspirante'
        ORDER BY apellido, nombre
    """).fetchall()

    if request.method == "POST":
        check = permiso_requerido("asistencia_gestionar")
        if check:
            conn.close()
            return check

        for jugador in jugadores:
            estado_asistencia = request.form.get(
                f"estado_jugador_{jugador['id']}",
                "ausente",
            )
            if estado_asistencia not in {"ausente", "a_tiempo", "tarde"}:
                estado_asistencia = "ausente"

            presente = 1 if estado_asistencia in {"a_tiempo", "tarde"} else 0
            observaciones = request.form.get(f"obs_jugador_{jugador['id']}", "").strip()

            conn.execute("""
                INSERT INTO asistencias (
                    evento_id, jugador_id, presente, estado_asistencia, observaciones
                )
                VALUES (%s, %s, %s, %s, %s)
                ON CONFLICT(evento_id, jugador_id)
                DO UPDATE SET
                    presente = excluded.presente,
                    estado_asistencia = excluded.estado_asistencia,
                    observaciones = excluded.observaciones
            """, (
                evento_id,
                jugador["id"],
                presente,
                estado_asistencia,
                observaciones,
            ))

        for aspirante in aspirantes:
            estado_asistencia = request.form.get(
                f"estado_aspirante_{aspirante['id']}",
                "ausente",
            )
            if estado_asistencia not in {"ausente", "a_tiempo", "tarde"}:
                estado_asistencia = "ausente"

            presente = 1 if estado_asistencia in {"a_tiempo", "tarde"} else 0
            observaciones = request.form.get(f"obs_aspirante_{aspirante['id']}", "").strip()

            conn.execute("""
                INSERT INTO aspirante_asistencias (
                    evento_id, aspirante_id, presente, estado_asistencia, observaciones
                )
                VALUES (%s, %s, %s, %s, %s)
                ON CONFLICT(evento_id, aspirante_id)
                DO UPDATE SET
                    presente = excluded.presente,
                    estado_asistencia = excluded.estado_asistencia,
                    observaciones = excluded.observaciones
            """, (
                evento_id,
                aspirante["id"],
                presente,
                estado_asistencia,
                observaciones,
            ))

        conn.commit()
        conn.close()

        flash("Asistencia guardada.", "ok")
        return redirect(url_for("listar_eventos_asistencia"))

    asistencias = conn.execute("""
        SELECT *
        FROM asistencias
        WHERE evento_id = %s
    """, (evento_id,)).fetchall()

    aspirante_asistencias = conn.execute("""
        SELECT *
        FROM aspirante_asistencias
        WHERE evento_id = %s
    """, (evento_id,)).fetchall()

    conn.close()

    asistencias_por_jugador = {
        a["jugador_id"]: a for a in asistencias
    }
    asistencias_por_aspirante = {
        a["aspirante_id"]: a for a in aspirante_asistencias
    }

    participantes = []
    for jugador in jugadores:
        item = dict(jugador)
        item["tipo"] = "jugador"
        item["tipo_label"] = "Jugador"
        item["form_key"] = f"jugador_{jugador['id']}"
        item["asistencia"] = asistencias_por_jugador.get(jugador["id"])
        participantes.append(item)

    for aspirante in aspirantes:
        item = dict(aspirante)
        item["tipo"] = "aspirante"
        item["tipo_label"] = "Ahijadx"
        item["form_key"] = f"aspirante_{aspirante['id']}"
        item["asistencia"] = asistencias_por_aspirante.get(aspirante["id"])
        participantes.append(item)

    return render_template(
        "tomar_asistencia.html",
        evento=evento,
        participantes=participantes
    )

if os.environ.get("INIT_DB", "true").lower() in {"1", "true", "yes", "on"}:
    init_db()

if __name__ == "__main__":
    app.run(debug=True)
