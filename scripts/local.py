"""SIG completo con PostgreSQL portátil y datos persistentes aislados.

Ejecutar con .codex-venv/Scripts/python.exe scripts/local.py start
"""
import argparse
import json
import os
from pathlib import Path
import secrets
import subprocess
import sys
import tempfile
import time
from urllib.request import urlopen
import zipfile

ROOT = Path(__file__).resolve().parents[1]
STATE = ROOT / ".local"
PG = STATE / "postgres" / "pgsql" / "bin"
DATA = STATE / "postgres-data"
VERSION = "17.11-5"
DOWNLOAD = f"https://get.enterprisedb.com/postgresql/postgresql-{VERSION}-windows-x64-binaries.zip"
sys.path.insert(0, str(ROOT))


def settings():
    STATE.mkdir(exist_ok=True)
    path = STATE / "config.json"
    if not path.exists():
        path.write_text(json.dumps({
            "db_password": secrets.token_urlsafe(32),
            "secret_key": secrets.token_urlsafe(48),
            "admin_password": "Local-" + secrets.token_urlsafe(12),
            "db_port": 55432,
            "web_port": 5080,
        }, indent=2), encoding="utf-8")
    config = json.loads(path.read_text(encoding="utf-8"))
    for key in ("db_port", "web_port"):
        if not isinstance(config[key], int) or not 1024 <= config[key] <= 65535:
            raise RuntimeError(f"Puerto local inválido: {key}")
    return config


def pg_command(*args, **kwargs):
    # pg_ctl crea procesos que pueden heredar pipes en Windows: usar un archivo
    # evita esperar indefinidamente a que el servidor cierre stdout/stderr.
    with tempfile.TemporaryFile() as output:
        result = subprocess.run([str(PG / (args[0] + ".exe")), *map(str, args[1:])],
                                cwd=ROOT, stdout=output, stderr=subprocess.STDOUT, **kwargs)
        output.seek(0)
        message = output.read().decode("utf-8", errors="replace")
    if result.returncode:
        raise RuntimeError(message)
    return message.strip()


def install_postgres():
    if (PG / "pg_ctl.exe").exists():
        return
    if sys.platform != "win32":
        raise RuntimeError("Este iniciador portátil está preparado para Windows.")
    archive = STATE / "downloads" / f"postgresql-{VERSION}-windows-x64-binaries.zip"
    archive.parent.mkdir(exist_ok=True)
    if not archive.exists():
        print("Descargando PostgreSQL portátil desde EDB...", flush=True)
        temporary = archive.with_suffix(".partial")
        with urlopen(DOWNLOAD, timeout=60) as response, temporary.open("wb") as output:
            while chunk := response.read(1024 * 1024):
                output.write(chunk)
        temporary.replace(archive)
    destination = STATE / "postgres"
    destination.mkdir(exist_ok=True)
    print("Preparando PostgreSQL portátil...", flush=True)
    with zipfile.ZipFile(archive) as source:
        for member in source.infolist():
            if not member.filename.startswith(("pgsql/bin/", "pgsql/lib/", "pgsql/share/")):
                continue
            target = (destination / member.filename).resolve()
            if not target.is_relative_to(destination.resolve()):
                raise RuntimeError("Ruta inválida en el archivo de PostgreSQL")
            source.extract(member, destination)


def db_connection(config, database="sig_local", **kwargs):
    import psycopg
    return psycopg.connect(host="127.0.0.1", port=config["db_port"],
                           dbname=database, user="sig_local_owner", password=config["db_password"],
                           connect_timeout=5, **kwargs)


def ensure_database(config):
    install_postgres()
    if not (DATA / "PG_VERSION").exists():
        password_file = STATE / "initdb-password.tmp"
        password_file.write_text(config["db_password"], encoding="utf-8")
        try:
            pg_command("initdb", "-D", DATA, "-U", "sig_local_owner", "--pwfile", password_file,
                       "--auth=scram-sha-256", "--encoding=UTF8", "--locale=C")
        finally:
            password_file.unlink(missing_ok=True)
    status = subprocess.run([str(PG / "pg_ctl.exe"), "-D", str(DATA), "status"],
                            capture_output=True)
    if status.returncode:
        pg_command("pg_ctl", "-D", DATA, "-l", STATE / "postgres.log", "-o",
                   f'-h 127.0.0.1 -p {config["db_port"]}', "-w", "start")
    # Impide usar por error otro servidor o un proxy que escuche en ese puerto.
    with db_connection(config, "postgres", autocommit=True) as conn:
        actual = Path(conn.execute("SHOW data_directory").fetchone()[0]).resolve()
        if actual != DATA.resolve():
            raise RuntimeError("El servidor no pertenece a .local/postgres-data")
        if not conn.execute("SELECT 1 FROM pg_database WHERE datname = 'sig_local'").fetchone():
            conn.execute("CREATE DATABASE sig_local")


def prepare_environment(config):
    # Ignora credenciales heredadas de una terminal de producción.
    prefixes = ("DB_", "POSTGRES_", "CLOUD_SQL_", "SMTP_", "WHATSAPP_", "GOOGLE_",
                "FACTURA_EMAIL", "PWA_VAPID_", "AUTOMATION_", "ADMIN_", "FORCE_ADMIN_")
    for key in list(os.environ):
        if key.startswith(prefixes) or key in {"DATABASE_URL", "INSTANCE_CONNECTION_NAME", "K_SERVICE"}:
            os.environ.pop(key)
    os.environ.update({
        "INIT_DB": "false", "SECRET_KEY": config["secret_key"],
        "ADMIN_PASSWORD": config["admin_password"], "SESSION_COOKIE_SECURE": "false",
        "DB_HOST": "127.0.0.1", "DB_PORT": str(config["db_port"]), "DB_NAME": "sig_local",
        "DB_USER": "sig_local_owner", "DB_PASS": config["db_password"], "DB_SSLMODE": "disable",
        "SIG_DATA_DIR": str(STATE / "files"), "FLASK_DEBUG": "false",
    })


def load_application(config):
    prepare_environment(config)
    import app as sig
    from services.local_runtime import configure
    configure(sig, STATE, config)
    sig.init_db()
    seed(sig, config)
    seed_extra_roles(sig, config)
    from scripts.local_scenarios import seed_review_scenarios, seed_financial_review
    seed_review_scenarios(sig)
    seed_financial_review(sig)
    return sig


def seed(sig, config):
    """Datos iniciales una sola vez; no repone ni sobrescribe cambios posteriores."""
    conn = sig.get_connection()
    try:
        conn.execute("CREATE TABLE IF NOT EXISTS sig_local_seed (version INTEGER PRIMARY KEY)")
        if conn.execute("SELECT 1 FROM sig_local_seed WHERE version = 1").fetchone():
            conn.close()
            return
        conn.execute("UPDATE usuarios SET debe_cambiar_password = 0, onboarding_visto = 1 WHERE username = 'admin'")
        for role in ("tesorero", "medico", "entrenador"):
            conn.execute("""INSERT INTO usuarios (username, password, rol, debe_cambiar_password, onboarding_visto)
                VALUES (%s, %s, %s, 0, 1) ON CONFLICT (username) DO NOTHING""",
                         (role, sig.generate_password_hash(config["admin_password"]), role))
        today = sig.ahora_sig().date()
        categoria = sig.planteles_disponibles(conn)[0]
        for index, (nombre, apellido) in enumerate((("Alex", "Prueba"), ("Sam", "Ejemplo"), ("Juli", "Local")), 1):
            jugador = conn.execute("""INSERT INTO jugadores
                (nombre, apellido, dni, email, telefono, categoria, fecha_ingreso, estado,
                 tipo_miembro, cobra_cuota, portal_activo, portal_token)
                VALUES (%s, %s, %s, %s, %s, %s, %s, 'Activo', 'Jugador', 1, 1, %s) RETURNING id""",
                (nombre, apellido, f"9900000{index}", f"jugador{index}@example.invalid", f"110000000{index}",
                 categoria, today.isoformat(), secrets.token_urlsafe(24))).fetchone()
            conn.execute("""INSERT INTO cuotas (jugador_id, periodo, importe, pagado, fecha_vencimiento)
                VALUES (%s, %s, 20000, 0, %s)""", (jugador["id"], today.strftime("%Y-%m"), today.isoformat()))
            if index == 1:
                conn.execute("""INSERT INTO jugador_bitacora (jugador_id, tipo, nota, creado_por, seguimiento_fecha)
                    VALUES (%s, 'finanzas', 'Avisó que paga la semana próxima. Nota de ejemplo editable en el entorno local.', 'tesorero', %s)""",
                    (jugador["id"], today + sig.timedelta(days=7)))
        conn.execute("""INSERT INTO aspirantes (nombre, apellido, categoria, estado, fecha_postulacion, etapa_seguimiento)
            VALUES ('Luz', 'Prueba local', %s, 'Aspirante', %s, 'pendiente_contacto')""", (categoria, today.isoformat()))
        conn.execute("""INSERT INTO calendario_eventos (fecha, tipo, titulo, categoria, publicar_portal)
            VALUES (%s, 'Entrenamiento', 'Entrenamiento de prueba local', %s, 1)""",
                     ((today + sig.timedelta(days=2)).isoformat(), categoria))
        conn.execute("INSERT INTO sig_local_seed VALUES (1)")
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def seed_extra_roles(sig, config):
    conn = sig.get_connection()
    try:
        if conn.execute("SELECT 1 FROM sig_local_seed WHERE version = 2").fetchone():
            return
        for name, permissions in {
            "madrinas": ["aspirantes_ver", "aspirantes_gestionar", "comunicaciones_ver"],
            "secretaria": ["secretaria_ver", "secretaria_gestionar", "tareas_ver", "tareas_gestionar"],
        }.items():
            conn.execute("""INSERT INTO roles (nombre, descripcion, permisos, sistema)
                VALUES (%s, 'Rol de prueba local', %s, 0) ON CONFLICT (nombre) DO NOTHING""",
                         (name, sig.serializar_permisos(permissions)))
            conn.execute("""INSERT INTO usuarios (username, password, rol, debe_cambiar_password, onboarding_visto)
                VALUES (%s, %s, %s, 0, 1) ON CONFLICT (username) DO NOTHING""",
                         (name, sig.generate_password_hash(config["admin_password"]), name))
        conn.execute("INSERT INTO sig_local_seed VALUES (2)")
        conn.commit()
    finally:
        conn.close()


def server_pid(config):
    path = STATE / "web.pid"
    if not path.exists():
        return None
    try:
        with urlopen(f'http://127.0.0.1:{config["web_port"]}/__local/health', timeout=2) as response:
            health = json.load(response)
        if health.get("local") and health.get("workspace") == str(ROOT):
            # El ejecutable de un venv de Windows puede lanzar otro python.exe.
            pid = int(health["pid"])
            path.write_text(str(pid))
            return pid
        return None
    except (ValueError, OSError):
        return None


def start(config):
    ensure_database(config)
    pid = server_pid(config)
    if not pid:
        with (STATE / "web.log").open("ab") as log:
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            process = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), "run"],
                                       cwd=ROOT, stdout=log, stderr=log, creationflags=creationflags)
        (STATE / "web.pid").write_text(str(process.pid))
        for _ in range(60):
            if process.poll() is not None:
                raise RuntimeError("El servidor no inició. Revisá .local/web.log")
            if server_pid(config):
                break
            time.sleep(0.5)
        else:
            raise RuntimeError("El servidor sigue iniciando. Revisá .local/web.log")
    print(f'SIG local: http://127.0.0.1:{config["web_port"]}')
    print("Usuarios: admin, tesorero, medico, entrenador, madrinas, secretaria")
    print(f'Contraseña local inicial: {config["admin_password"]}')
    print("Datos y registros: .local/ (se conservan al reiniciar)")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("start", "run", "stop", "status", "check", "backup"), default="start", nargs="?")
    args = parser.parse_args()
    config = settings()
    if args.action == "start":
        start(config)
    elif args.action == "run":
        sig = load_application(config)
        (STATE / "web.pid").write_text(str(os.getpid()))
        print("SIG LOCAL: guardado real; servicios externos simulados.", flush=True)
        sig.app.run(host="127.0.0.1", port=config["web_port"], debug=False, use_reloader=False)
    elif args.action == "stop":
        pid = server_pid(config)
        if pid:
            result = subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
            if result.returncode:
                raise RuntimeError(result.stdout.decode(errors="replace") + result.stderr.decode(errors="replace"))
        if (PG / "pg_ctl.exe").exists() and (DATA / "postmaster.pid").exists():
            pg_command("pg_ctl", "-D", DATA, "-m", "fast", "-w", "stop")
        print("SIG local detenido. Los datos se conservan.")
    elif args.action == "status":
        print("Activo" if server_pid(config) else "Detenido")
    elif args.action == "check":
        ensure_database(config)
        sig = load_application(config)
        from scripts.local_check import check
        check(sig, config)
    elif args.action == "backup":
        ensure_database(config)
        folder = STATE / "backups"
        folder.mkdir(exist_ok=True)
        target = folder / (time.strftime("sig-local-%Y%m%d-%H%M%S") + ".dump")
        pg_command("pg_dump", "-h", "127.0.0.1", "-p", config["db_port"], "-U", "sig_local_owner",
                   "-d", "sig_local", "-Fc", "-f", target,
                   env={**os.environ, "PGPASSWORD": config["db_password"]})
        print(target)


if __name__ == "__main__":
    main()
