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
    destinations = {"/": "/preview/admin", "/operacion": "/preview/operacion", "/finanzas/ventas": "/preview/ventas", "/ahijadxs": "/preview/ahijadxs", "/ahijadxs/nuevo": "/preview/ahijadxs/nuevo", "/ahijadxs/1": "/preview/ahijadxs/1", "/ahijadxs/2": "/preview/ahijadxs/1", "/ahijadxs/1/editar": "/preview/ahijadxs/nuevo", "/ahijadxs/2/editar": "/preview/ahijadxs/nuevo"}
    if request.path in destinations:
        return redirect(destinations[request.path])
    if not (request.path.startswith("/preview/") or request.path.startswith("/static/") or request.path == "/postulate"):
        return "Sección fuera de esta vista de diseño. Volvé al panel de demostración.", 404
    session["user_id"] = 1
    session["username"] = ""
    if request.args.get("rol") in {"admin", "madrinas", "tesorero", "medico", "entrenador", "secretaria"}:
        session["preview_rol"] = request.args["rol"]
    session["rol"] = session.get("preview_rol", "admin")
    preview_permissions = {
        "admin": [],
        "madrinas": ["aspirantes_ver", "aspirantes_gestionar", "comunicaciones_ver"],
        "tesorero": sig.ROLE_PRESETS["tesorero"],
        "medico": sig.ROLE_PRESETS["medico"],
        "entrenador": sig.ROLE_PRESETS["entrenador"],
        "secretaria": ["secretaria_ver", "secretaria_gestionar", "tareas_ver", "tareas_gestionar"],
    }
    session["permisos"] = preview_permissions[session["rol"]]
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
    puede_ver_finanzas = sig.tiene_permiso("cuotas_ver", "cuotas_gestionar")
    puede_ver_salud = sig.tiene_permiso("salud_ver")
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
        resumen_notificaciones={"cuotas_vencidas": 5, "fichas": 1, "comprobantes": 3, "cambios_portal": 2, "asistencia_baja": 4, "whatsapp": 2, "ahijadxs": 3},
        sistema_resumen={"db_ok": True, "integraciones": {"smtp_ok": True}} if session.get("rol") == "admin" else None,
        puede_ver_jugadores=sig.tiene_permiso("jugadores_ver"),
        puede_ver_finanzas=puede_ver_finanzas,
        puede_ver_salud=puede_ver_salud,
    )


@sig.app.route("/preview/operacion")
def preview_operacion():
    if not sig.puede_ver_operacion():
        return redirect("/preview/admin")
    modulos = sig.ordenar_modulos_tareas()
    tareas_demo = [
        {"id": 1, "titulo": "Revisar documentación pendiente", "descripcion": "Validar el archivo recibido.", "modulo": modulos[0], "prioridad": "media", "responsable": "Equipo demo", "fecha_vencimiento": "2026-09-18", "estado": "pendiente", "apellido": "Pérez", "nombre": "Alex", "categoria": "Plantel superior"}
    ] if modulos else []
    return render_template(
        "operacion.html",
        revision={"whatsapp": 2, "comprobantes": 3, "cambios_portal": 2, "cuotas": 5, "fichas": 1, "asistencia_baja": 4, "secretaria": 2, "ahijadxs": 3, "proximos_eventos": [{"id": 1, "fecha": "2026-09-18", "hora_inicio": "20:30", "tipo": "Entrenamiento", "descripcion": "Entrenamiento plantel superior"}], "tareas_vencidas": 1},
        tareas=tareas_demo,
        estado="pendiente",
        puede_gestionar_tareas=sig.puede_gestionar_tareas_sig(),
        puede_ver_tareas=bool(modulos) and sig.tiene_permiso("tareas_ver", "tareas_gestionar"),
        modulos_tareas=modulos,
    )


@sig.app.route("/preview/ventas")
def preview_ventas():
    if not sig.tiene_permiso("ventas_ver", "ventas_gestionar"):
        session["preview_rol"] = "tesorero"
        session["rol"] = "tesorero"
        session["permisos"] = sig.ROLE_PRESETS["tesorero"]
    productos = [
        {"id": 1, "nombre": "Remera oficial", "descripcion": "Modelo 2026", "precio": 25000, "stock": 8, "activo": 1},
        {"id": 2, "nombre": "Gorra del club", "descripcion": "Negra y verde", "precio": 12000, "stock": 4, "activo": 1},
    ]
    ventas = [
        {"id": 21, "fecha": "2026-09-16", "comprador": "Juan Pérez", "detalle": "Remera oficial x1", "notas": "", "medio_pago": "Transferencia", "total": 25000, "estado": "confirmada", "comprobante_drive_file_id": "demo-drive", "comprobante_fecha": "2026-09-16 15:40", "comprobante_usuario": "tesoreria"},
        {"id": 20, "fecha": "2026-09-15", "comprador": "María López", "detalle": "Gorra del club x1", "notas": "Entrega en entrenamiento", "medio_pago": "Efectivo", "total": 12000, "estado": "confirmada", "comprobante_drive_file_id": None, "comprobante_fecha": None, "comprobante_usuario": None},
    ]
    return render_template(
        "ventas.html",
        productos=productos,
        ventas=ventas,
        resumen={"operaciones": 2, "total": 37000},
        mes="2026-09",
        fecha_hoy="2026-09-16",
    )


@sig.app.route("/preview/ahijadxs")
def preview_ahijadxs():
    aspirantes = [
        {"id": 1, "apellido": "Pérez", "nombre": "Alex", "dni": "", "categoria": "Plantel superior", "madrina_apellido": None, "madrina_nombre": None, "entrenamientos_realizados": 0, "entrenamientos_objetivo": 8, "progreso": 0, "estado": "Aspirante", "listo_para_ingresar": False, "etapa_seguimiento": "pendiente_contacto", "proxima_accion_fecha": "2026-09-15"},
        {"id": 2, "apellido": "López", "nombre": "Sam", "dni": "", "categoria": "Plantel superior", "madrina_apellido": "García", "madrina_nombre": "Sol", "entrenamientos_realizados": 3, "entrenamientos_objetivo": 8, "progreso": 38, "estado": "Aspirante", "listo_para_ingresar": False, "etapa_seguimiento": "confirmo_asistencia", "proxima_accion_fecha": "2026-09-17"},
    ]
    return render_template("aspirantes.html", aspirantes=aspirantes, busqueda="", estado="Aspirante", estados=sorted(sig.ASPIRANTE_ESTADOS), etapas=sig.ASPIRANTE_ETAPAS, etapa="todas", asignacion="todas", resumen_aspirantes={"en_seguimiento": 12, "sin_asignar": 3, "pendientes_contacto": 4, "acciones_vencidas": 2})


if __name__ == "__main__":
    sig.app.run(host="127.0.0.1", port=int(os.environ.get("PREVIEW_PORT", "5073")), debug=False, use_reloader=False)
