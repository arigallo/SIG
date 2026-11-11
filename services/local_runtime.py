"""Adaptadores exclusivos del iniciador local. Producción no importa este módulo."""
import io
import hashlib
import json
import os
from pathlib import Path
import re
import secrets
import socket
import threading
from datetime import datetime, timezone


class LocalRequest:
    def __init__(self, result):
        self.result = result

    def execute(self):
        return self.result


class LocalDrive:
    """Conserva la interfaz usada por SIG, con archivos y metadatos en disco."""
    def __init__(self, directory):
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)
        self.lock = threading.RLock()
        if not self.path("local-root", ".json").exists():
            self.save({"id": "local-root", "name": "Archivos locales", "mimeType": "application/vnd.google-apps.folder", "parents": []})

    def path(self, identifier, extension):
        if not re.fullmatch(r"local-[a-z0-9]+", identifier or ""):
            raise ValueError("Identificador de archivo local inválido")
        return self.directory / (identifier + extension)

    def metadata(self, identifier):
        return json.loads(self.path(identifier, ".json").read_text(encoding="utf-8"))

    def save(self, metadata):
        target = self.path(metadata["id"], ".json")
        temporary = target.with_suffix("." + secrets.token_hex(8) + ".tmp")
        temporary.write_text(json.dumps(metadata, ensure_ascii=False), encoding="utf-8")
        temporary.replace(target)

    def files(self):
        return self

    def list(self, q="", **kwargs):
        def value(pattern):
            match = re.search(pattern, q)
            return match.group(1).replace("\\'", "'").replace("\\\\", "\\") if match else None
        name = value(r"name = '((?:\\.|[^'])*)'")
        parent = value(r"'((?:\\.|[^'])*)' in parents")
        mime = value(r"mimeType = '([^']*)'")
        with self.lock:
            files = [json.loads(path.read_text(encoding="utf-8")) for path in self.directory.glob("*.json")]
        return LocalRequest({"files": [item for item in files
            if (name is None or item["name"] == name)
            and (parent is None or parent in item.get("parents", []))
            and (mime is None or item["mimeType"] == mime)
            and not item.get("trashed")]})

    def create(self, body, media_body=None, **kwargs):
        with self.lock:
            identifier = "local-" + secrets.token_hex(16)
            metadata = {"id": identifier, "mimeType": "application/octet-stream", **body,
                        "webViewLink": f"/__local/files/{identifier}"}
            if body.get("mimeType") == "application/vnd.google-apps.document":
                raise RuntimeError("El OCR de Google no se ejecuta en local. Cargá los datos de la ficha manualmente.")
            self.write_media(metadata, media_body)
            self.save(metadata)
            return LocalRequest(metadata)

    def write_media(self, metadata, media):
        if media is not None:
            stream = media.stream()
            stream.seek(0)
            content = stream.read()
            self.path(metadata["id"], ".bin").write_bytes(content)
            metadata.update(size=len(content), mimeType=media.mimetype())

    def update(self, fileId, body=None, media_body=None, addParents=None, removeParents=None, **kwargs):
        with self.lock:
            metadata = self.metadata(fileId)
            metadata.update(body or {})
            parents = set(metadata.get("parents", []))
            parents.update(filter(None, (addParents or "").split(",")))
            parents.difference_update(filter(None, (removeParents or "").split(",")))
            metadata["parents"] = sorted(parents)
            self.write_media(metadata, media_body)
            self.save(metadata)
            return LocalRequest(metadata)

    def get(self, fileId, **kwargs):
        return LocalRequest(self.metadata(fileId))

    def get_media(self, fileId, **kwargs):
        return LocalRequest(self.path(fileId, ".bin").read_bytes())

    def delete(self, fileId, **kwargs):
        with self.lock:
            self.path(fileId, ".json").unlink(missing_ok=True)
            self.path(fileId, ".bin").unlink(missing_ok=True)
        return LocalRequest({})


def block_external_connections(db_port):
    """Las conexiones Python salientes solo pueden ir al PostgreSQL de prueba."""
    original_connect = socket.socket.connect
    original_connect_ex = socket.socket.connect_ex

    def validate(address):
        if not isinstance(address, tuple) or address[0] not in {"127.0.0.1", "::1", "localhost"} or address[1] != db_port:
            raise ConnectionRefusedError("SIG local: la conexión externa está bloqueada")

    def connect(sock, address):
        validate(address)
        return original_connect(sock, address)

    def connect_ex(sock, address):
        validate(address)
        return original_connect_ex(sock, address)

    socket.socket.connect = connect
    socket.socket.connect_ex = connect_ex


def configure(sig, state, config):
    from flask import abort, jsonify, render_template, request, send_file, session, url_for, Response

    if getattr(sig, "_local_runtime", False):
        return
    sig._local_runtime = True
    state = Path(state)
    outbox = state / "outbox"
    outbox.mkdir(parents=True, exist_ok=True)
    drive = LocalDrive(state / "files" / "drive")

    def capture(channel, payload, content=None):
        identifier = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S") + "-" + secrets.token_hex(8)
        item = {"id": identifier, "canal": channel, "fecha": datetime.now(timezone.utc).isoformat(), **payload}
        (outbox / (identifier + ".json")).write_text(json.dumps(item, ensure_ascii=False, default=str), encoding="utf-8")
        if content is not None:
            (outbox / (identifier + ".eml")).write_bytes(content)
        return identifier

    class LocalSMTP:
        def __init__(self, *args, **kwargs):
            pass

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

        def starttls(self):
            pass

        def login(self, *args):
            pass

        def send_message(self, message, **kwargs):
            capture("email", {"destinatario": str(message["To"]), "asunto": str(message["Subject"]),
                              "texto": message.get_body(preferencelist=("plain",)).get_content()}, message.as_bytes())
            return {}

    def local_whatsapp(telefono, payload_builder, *, tipo="general", entidad="sistema", entidad_id=None, jugador_id=None, mensaje_log=None):
        variants = sig.variantes_telefono_whatsapp(telefono)
        if not variants:
            return False, None, "sin_telefono", None
        destination = variants[0]
        payload = payload_builder(destination)
        identifier = "local-" + capture("whatsapp", {"destinatario": destination, "payload": payload, "texto": mensaje_log})
        response = {"local": True, "messages": [{"id": identifier}]}
        sig.registrar_whatsapp_envio(telefono=telefono, tipo=tipo, entidad=entidad, entidad_id=entidad_id,
            jugador_id=jugador_id, mensaje=mensaje_log, payload=payload, respuesta=response, estado="enviado", meta_message_id=identifier)
        sig.registrar_whatsapp_mensaje(telefono=destination, jugador_id=jugador_id, direccion="out",
            tipo=payload.get("type", "text"), texto=mensaje_log or json.dumps(payload, ensure_ascii=False),
            meta_message_id=identifier, estado="enviado", payload=payload, respuesta=response)
        return True, destination, None, None

    def local_push(subscription, payload):
        capture("push", {"payload": payload})
        return True, None

    def unavailable(*args, **kwargs):
        raise RuntimeError("Esta integración externa no está disponible en SIG local.")

    # Configuración inmutable de la conexión usada por todas las rutas.
    sig.get_db_connect_args = lambda: ((), {
        "host": "127.0.0.1", "port": config["db_port"], "dbname": "sig_local",
        "user": "sig_local_owner", "password": config["db_password"], "sslmode": "disable",
        "row_factory": sig.dict_row, "connect_timeout": 5,
    })
    sig.app.config.update(SESSION_COOKIE_NAME="sig_local_session", SESSION_COOKIE_SECURE=False,
                          TEMPLATES_AUTO_RELOAD=True)
    sig.app.jinja_env.auto_reload = True
    sig.drive_service = lambda: drive
    sig.descargar_drive_file = lambda identifier: io.BytesIO(drive.get_media(identifier).execute())
    sig.DRIVE_COMPROBANTES_FOLDER_ID = sig.DRIVE_FICHAS_MEDICAS_FOLDER_ID = sig.DRIVE_SECRETARIA_FOLDER_ID = "local-root"
    sig.DRIVE_SHARED_DRIVE_ID = None
    sig.DRIVE_COMPROBANTES_SUBFOLDER = sig.DRIVE_FICHAS_MEDICAS_SUBFOLDER = sig.DRIVE_SECRETARIA_SUBFOLDER = ""
    sig.cloud_sql_admin_service = sig.secret_manager_service = unavailable
    sig.google_auth_default = None
    sig.SMTP_HOST, sig.SMTP_FROM, sig.SMTP_USER, sig.SMTP_PASSWORD = "local", "sig@example.invalid", "", ""
    sig.SMTP_USE_TLS = False
    sig.smtplib.SMTP = LocalSMTP
    sig.whatsapp_api_disponible = lambda: True
    sig._enviar_whatsapp_payload = local_whatsapp
    sig.enviar_push_subscription = local_push

    def interface_revision():
        root = Path(sig.__file__).resolve().parent
        files = sorted((root / "templates").glob("*.html")) + sorted((root / "static").glob("*.css")) + sorted((root / "static").glob("*.js"))
        signature = "|".join(f"{path.relative_to(root)}:{path.stat().st_mtime_ns}:{path.stat().st_size}" for path in files)
        return hashlib.sha256(signature.encode()).hexdigest()[:10]

    def local_static_asset(filename):
        path = (sig.BASE_DIR / "static" / filename).resolve()
        if not path.is_relative_to((sig.BASE_DIR / "static").resolve()):
            abort(404)
        version = hashlib.sha256(path.read_bytes()).hexdigest()[:16] if path.is_file() else "missing"
        return url_for("static", filename=filename, v=version)

    sig.static_asset = local_static_asset

    @sig.app.after_request
    def local_no_cache(response):
        response.headers["Cache-Control"] = "no-store, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        if response.mimetype == "text/html":
            response.headers["X-SIG-Local-Revision"] = interface_revision()
        return response

    def local_service_worker():
        # Keep local pages online-only, replacing an older cache-first worker.
        script = """self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', event => event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(key => key.startsWith('sig-')).map(key => caches.delete(key))))
        .then(() => self.clients.claim())
));
// No fetch handler: local requests always use the network.
"""
        return Response(script, mimetype="application/javascript", headers={"Service-Worker-Allowed": "/"})

    sig.app.view_functions["pwa_service_worker"] = local_service_worker

    @sig.app.context_processor
    def local_context():
        return {"sig_local_mode": True, "local_interface_revision": interface_revision()}

    @sig.app.route("/__local/health")
    def local_health():
        conn = sig.get_connection()
        try:
            conn.execute("SELECT 1")
        finally:
            conn.close()
        return jsonify(local=True, pid=os.getpid(), workspace=str(Path(sig.__file__).resolve().parent))

    original_protection = sig.proteger_rutas

    def protect_local_routes():
        if request.endpoint == "local_health":
            return None
        return original_protection()

    sig.app.before_request_funcs[None] = [protect_local_routes if fn == original_protection else fn
                                         for fn in sig.app.before_request_funcs.get(None, [])]

    @sig.app.route("/__local/outbox")
    def local_outbox():
        items = [json.loads(path.read_text(encoding="utf-8")) for path in sorted(outbox.glob("*.json"), reverse=True)[:100]]
        return render_template("local_outbox.html", items=items)

    @sig.app.route("/__local/files/<identifier>")
    def local_file(identifier):
        try:
            metadata = drive.metadata(identifier)
            content = drive.path(identifier, ".bin")
            if not content.exists():
                abort(404)
        except (ValueError, FileNotFoundError):
            abort(404)
        return send_file(content, mimetype=metadata["mimeType"], download_name=metadata["name"], as_attachment=True)

    @sig.app.route("/__local/outbox/<identifier>.eml")
    def local_email(identifier):
        if not re.fullmatch(r"[0-9]{8}T[0-9]{6}-[a-f0-9]{16}", identifier):
            abort(404)
        return send_file(outbox / (identifier + ".eml"), as_attachment=True)

    block_external_connections(config["db_port"])
