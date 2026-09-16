import os
import sys
from pathlib import Path

os.environ.setdefault("SECRET_KEY", "local-preview-only")
os.environ.setdefault("SESSION_COOKIE_SECURE", "false")
os.environ.setdefault("INIT_DB", "false")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from flask import render_template

import app as sig


def portal_preview(token):
    jugador = {
        "id": 1,
        "nombre": "Emmanuel Lucas",
        "apellido": "Cristaldo",
        "categoria": "Plantel Superior",
        "estado": "Activo",
        "numero_socio": "00032",
        "beca_activa": 0,
        "beca_porcentaje": 0,
        "dni": "",
        "fecha_nacimiento": "",
        "telefono": "",
        "email": "",
        "direccion": "",
        "obra_social": "",
        "numero_afiliado_obra_social": "",
        "contacto_tutor": "",
        "parentesco_tutor": "",
        "telefono_tutor": "",
        "email_tutor": "",
        "fecha_ingreso": "2024-03-01",
        "portal_onboarding_visto": 1,
        "portal_accesos_rapidos": "entrenamientos,pagos,salud,calendario",
        "posicion": "Fullback",
        "numero_camiseta": "15",
        "objetivo_temporada": "Mejorar la asistencia y llegar en mi mejor nivel a las finales.",
    }
    eventos = [
        {
            "id": 1,
            "tipo": "Entrenamiento",
            "fecha": "2026-09-15",
            "hora_inicio": "20:30",
            "titulo": "Entrenamiento",
            "descripcion": "Preparación semanal del Plantel Superior",
            "ubicacion": "Predio Ruda Macho",
            "convocatoria_texto": "Convocatoria abierta",
            "convocatoria_cierre": "2026-09-15 18:00",
            "minuta_post_evento": "",
            "asistencia_evento_id": 1,
            "es_partido": False,
            "confirmacion_portal": None,
        },
        {
            "id": 2,
            "tipo": "Partido",
            "fecha": "2026-09-19",
            "hora_inicio": "15:30",
            "titulo": "Ruda Macho Rugby vs. Hand Off Rugby",
            "descripcion": "Fecha del torneo",
            "ubicacion": "Cancha principal",
            "convocatoria_texto": "A confirmar",
            "convocatoria_cierre": "2026-09-18 20:00",
            "minuta_post_evento": "",
            "asistencia_evento_id": 2,
            "es_partido": True,
            "confirmacion_portal": {"estado": "confirmado"},
        },
    ]
    return render_template(
        "portal_jugador.html",
        jugador=jugador,
        cuotas=[],
        deuda=35000,
        deuda_cuotas=35000,
        ficha={"presentada": True, "fecha_vencimiento": "2026-12-31", "apto_fisico": True},
        documentos=[{"tipo": "Ficha médica", "nombre": "Apto 2026", "fecha_vencimiento": "2026-12-31", "url": ""}],
        documentos_por_vencer=0,
        accesos_disponibles={
            "entrenamientos": {"label": "Entrenamientos", "href": "#equipo", "hint": "Agenda y confirmaciones"},
            "pagos": {"label": "Pagos", "href": "#pagos", "hint": "Cuenta y comprobantes"},
            "salud": {"label": "Salud", "href": "#salud", "hint": "Apto y disponibilidad"},
            "calendario": {"label": "Calendario", "href": "#equipo", "hint": "Próximos eventos"},
            "documentos": {"label": "Documentos", "href": "#documentos", "hint": "Ficha y certificados"},
            "perfil": {"label": "Perfil", "href": "#perfil", "hint": "Datos y preferencias"},
        },
        accesos_guardados=["entrenamientos", "pagos", "salud", "calendario"],
        accesos_rapidos=[
            {"clave": "entrenamientos", "label": "Entrenamientos", "href": "#equipo", "hint": "Agenda y confirmaciones"},
            {"clave": "pagos", "label": "Pagos", "href": "#pagos", "hint": "Cuenta y comprobantes"},
            {"clave": "salud", "label": "Salud", "href": "#salud", "hint": "Apto y disponibilidad"},
            {"clave": "calendario", "label": "Calendario", "href": "#equipo", "hint": "Próximos eventos"},
        ],
        temporada_resumen={"anio": 2026, "eventos": 24, "asistencia_porcentaje": 86, "partidos": 14, "entrenamientos": 10},
        historial_asistencia=[],
        lesiones_activas_portal=[],
        portal_alertas=[{"nivel": "danger", "titulo": "Tenés deuda pendiente", "detalle": "Actualmente tenés $35.000 pendientes de pago."}],
        ficha_portal={"nivel": "success", "label": "Apto físico vigente"},
        planes_pago=[],
        gastos_compartidos=[],
        gastos_pendientes=[],
        gastos_pagados=[],
        gasto_pendiente_total=0,
        tests_recientes=[],
        resumen_tests=[],
        portal_tests=[],
        portal_test_id=None,
        portal_test_desde="",
        portal_test_hasta="",
        portal_test_actual=None,
        portal_test_grafico={"series": []},
        portal_test_comparativo={},
        portal_test_resultados=[],
        eventos_deportivos=eventos,
        calendario_ics_url="#calendario",
        calendario_webcal_url="#iphone",
        calendario_google_url="#google",
        calendario_android_url="#android",
        cuenta_corriente=[{"fecha_pago": None, "fecha": "2026-09-01", "concepto": "Cuota septiembre", "importe": 35000, "estado": "pendiente"}],
        comunicaciones_portal=[],
        token=token,
    )


def bienestar_preview(token, evento_id=None):
    return render_template(
        "portal_bienestar.html",
        jugador={"nombre": "Emmanuel Lucas", "apellido": "Cristaldo"},
        evento={"titulo": "Entrenamiento", "fecha": "2026-09-15", "hora_inicio": "20:30"},
        estado="confirmado",
        actual={},
        horas_opciones=sig.BIENESTAR_HORAS_OPCIONES,
        dolor_zonas=sig.BIENESTAR_DOLOR_ZONAS,
        token=token,
    )


sig.app.view_functions["portal_jugador"] = portal_preview
sig.app.view_functions["portal_bienestar_asistencia"] = bienestar_preview
sig.portal_tiene_notificaciones_activas = lambda _token: True

if __name__ == "__main__":
    sig.app.run(host="127.0.0.1", port=5059, debug=False)
