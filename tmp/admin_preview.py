import os
import sys
from pathlib import Path

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "local-preview-only")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app as sig
from flask import render_template, session, request, redirect


def seed_preview_session():
    if request.method != "GET":
        return "Vista de diseño: los cambios están deshabilitados.", 405
    destinations = {"/": "/preview/admin", "/ahijadxs": "/preview/ahijadxs", "/ahijadxs/nuevo": "/preview/ahijadxs/nuevo", "/ahijadxs/1": "/preview/ahijadxs/1", "/ahijadxs/2": "/preview/ahijadxs/1", "/ahijadxs/1/editar": "/preview/ahijadxs/nuevo", "/ahijadxs/2/editar": "/preview/ahijadxs/nuevo"}
    if request.path in destinations:
        return redirect(destinations[request.path])
    if not (request.path.startswith("/preview/") or request.path.startswith("/static/") or request.path == "/postulate"):
        return "Sección fuera de esta vista de diseño. Volvé al panel de demostración.", 404
    session["user_id"] = 1
    session["username"] = ""
    if request.args.get("rol") in {"admin", "madrinas"}:
        session["preview_rol"] = request.args["rol"]
    session["rol"] = session.get("preview_rol", "admin")
    session["permisos"] = ["aspirantes_ver", "aspirantes_gestionar"] if session["rol"] == "madrinas" else []
    session["onboarding_visto"] = True
    session["debe_cambiar_password"] = False


sig.app.before_request_funcs[None] = [seed_preview_session]
def deny_database():
    raise RuntimeError("La vista previa no permite conexiones a la base de datos")
sig.get_connection = deny_database
sig.obtener_config_mantenimiento = lambda: {"activo": False}

@sig.app.context_processor
def preview_context():
    return {"preview_mode": True}

DEMO_ASPIRANTE = {"id": 1, "nombre": "Alex", "apellido": "Pérez", "estado": "Aspirante", "categoria": "Plantel superior", "etapa_seguimiento": "contactado", "proxima_accion_fecha": "2026-09-18", "entrenamientos_realizados": 2, "entrenamientos_objetivo": 8, "progreso": 25, "listo_para_ingresar": False, "telefono": "11 2345 6789", "experiencia_previa": "Sin experiencia previa", "disponibilidad": "Martes y jueves por la tarde", "fecha_postulacion": "2026-09-15"}

@sig.app.route("/preview/ahijadxs/1")
def preview_detalle():
    return render_template("aspirante_detalle.html", aspirante=DEMO_ASPIRANTE, etapas=sig.ASPIRANTE_ETAPAS, asistencias=[], seguimientos=[{"tipo": "whatsapp", "fecha": "2026-09-16 10:30:00", "detalle": "Se compartieron los horarios de entrenamiento. Confirmar disponibilidad el viernes.", "creado_por": "Madrina demo", "proxima_accion_fecha": "2026-09-18"}])

@sig.app.route("/preview/ahijadxs/nuevo")
def preview_form():
    return render_template("aspirante_form.html", aspirante=DEMO_ASPIRANTE, madrinas=[{"id": 3, "nombre": "Sol", "apellido": "García", "categoria": "Plantel superior"}], modo="editar")


@sig.app.route("/preview/admin")
def preview_admin():
    return render_template(
        "dashboard.html",
        total_jugadores=86,
        jugadores_con_deuda=[{"id": 1, "nombre": "Alex", "apellido": "Pérez", "deuda": 35000, "cuotas_vencidas": 2}],
        fichas_vencidas=[{"id": 2, "nombre": "Sam", "apellido": "López", "fecha_vencimiento": "2026-09-10"}],
        lesiones_activas=[{"id": 1, "jugador_id": 3, "nombre": "Martín", "apellido": "Díaz", "tipo_lesion": "Contractura", "zona_cuerpo": "Pierna", "estado": "Activa"}],
        total_recaudado_mes=825000,
        deuda_total=186000,
        deuda_vencida_total=92000,
        cuotas_pagadas_mes=61,
        cuotas_pendientes=12,
        cobranza_ratio=83.6,
        fichas_por_vencer_count=4,
        cuotas_pendientes_lista=[],
        comprobantes_pendientes_count=3,
        comprobantes_pendientes_lista=[],
        mes_actual="2026-09",
        resumen_notificaciones={"cuotas_vencidas": 5, "fichas": 1, "comprobantes": 3, "cambios_portal": 2, "asistencia_baja": 4},
        sistema_resumen={"db_ok": True, "integraciones": {"smtp_ok": True}},
        puede_ver_jugadores=True,
        puede_ver_finanzas=True,
        puede_ver_salud=True,
    )


@sig.app.route("/preview/ahijadxs")
def preview_ahijadxs():
    aspirantes = [
        {"id": 1, "apellido": "Pérez", "nombre": "Alex", "dni": "", "categoria": "Plantel superior", "madrina_apellido": None, "madrina_nombre": None, "entrenamientos_realizados": 0, "entrenamientos_objetivo": 8, "progreso": 0, "estado": "Aspirante", "listo_para_ingresar": False, "etapa_seguimiento": "pendiente_contacto", "proxima_accion_fecha": "2026-09-15"},
        {"id": 2, "apellido": "López", "nombre": "Sam", "dni": "", "categoria": "Plantel superior", "madrina_apellido": "García", "madrina_nombre": "Sol", "entrenamientos_realizados": 3, "entrenamientos_objetivo": 8, "progreso": 38, "estado": "Aspirante", "listo_para_ingresar": False, "etapa_seguimiento": "confirmo_asistencia", "proxima_accion_fecha": "2026-09-17"},
    ]
    return render_template("aspirantes.html", aspirantes=aspirantes, busqueda="", estado="Aspirante", estados=sorted(sig.ASPIRANTE_ESTADOS), etapas=sig.ASPIRANTE_ETAPAS, etapa="todas", asignacion="todas", resumen_aspirantes={"en_seguimiento": 12, "sin_asignar": 3, "pendientes_contacto": 4, "acciones_vencidas": 2})


if __name__ == "__main__":
    sig.app.run(host="127.0.0.1", port=5072, debug=False, use_reloader=False)
