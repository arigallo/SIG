import os
import unittest
from unittest.mock import patch

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")

import app


class PlanPagosTests(unittest.TestCase):
    def test_solicitud_masiva_bloqueada_en_ambos_canales(self):
        for vista in (app.enviar_email_comunicacion_morosos_lote, app.enviar_whatsapp_comunicacion_morosos_lote):
            with self.subTest(vista=vista.__name__), app.app.test_request_context(method="POST", data={"plantilla": "solicitar_comprobante"}), patch.object(app, "permiso_requerido", return_value=None), patch.object(app, "obtener_morosos_para_comunicacion") as buscar:
                self.assertEqual(vista().status_code, 302)
                buscar.assert_not_called()

    def test_solicitud_desde_finanzas_no_requiere_deuda(self):
        from unittest.mock import MagicMock
        conn = MagicMock()
        conn.execute.return_value.fetchone.return_value = {"id": 7, "nombre": "Alex"}
        with app.app.test_request_context(method="POST"), patch.object(app, "permiso_requerido", return_value=None), patch.object(app, "get_connection", return_value=conn), patch.object(app, "enviar_whatsapp_recordatorio_cuota_template", return_value=(True, "5491112345678", None, None)) as enviar, patch.object(app, "registrar_auditoria"):
            respuesta = app.solicitar_comprobante_jugador(7)
        self.assertEqual(respuesta.status_code, 302)
        self.assertEqual(enviar.call_args.kwargs["plantilla"], "solicitar_comprobante")

    def test_solicitud_comprobante_usa_nombre_y_plantilla_correcta(self):
        jugador = {
            "id": 7, "nombre": "Alex", "apellido": "Ejemplo",
            "telefono": "5491112345678", "deuda": 45000,
            "cuotas_pendientes": 3, "cuotas_vencidas": 2,
            "primer_vencimiento": "2026-09-01",
        }
        with patch.object(app, "enviar_whatsapp_template_meta") as enviar:
            app.enviar_whatsapp_recordatorio_cuota_template(jugador, plantilla="solicitar_comprobante")
        self.assertEqual(enviar.call_args.args[1], "solicitud_comprobante_pago_v2")
        self.assertEqual(enviar.call_args.args[2], "es_AR")
        self.assertEqual(len(enviar.call_args.args[3]), 1)
        self.assertIn("Alex", enviar.call_args.args[3][0])
        self.assertIn("https://sig.rudamachorugby.com/portal", enviar.call_args.kwargs["mensaje_log"])
        self.assertNotIn("{nombre}", enviar.call_args.kwargs["mensaje_log"])

    def test_meta_recibe_cantidad_y_monto_sin_duplicar_pesos(self):
        jugador = {
            "id": 7, "nombre": "Ana", "apellido": "Pérez",
            "telefono": "5491112345678", "telefono_tutor": "",
            "deuda": 45000, "cuotas_pendientes": 3,
            "cuotas_vencidas": 2, "primer_vencimiento": "2026-09-01",
        }
        with patch.object(app, "enviar_whatsapp_template_meta") as enviar:
            app.enviar_whatsapp_recordatorio_cuota_template(jugador, plantilla="plan_pagos")
        self.assertEqual(enviar.call_args.args[1:], (
            "recordatorio_cuotas_pendientes", "es_AR", ["3", "45.000"],
        ))
        texto = enviar.call_args.kwargs["mensaje_log"]
        self.assertIn("3 cuotas pendientes", texto)
        self.assertIn("$45.000", texto)
        self.assertIn("https://wa.me/5491161885141?text=Hola%20Ari%21", texto)

    def test_recordatorio_anterior_conserva_una_variable(self):
        with patch.object(app, "enviar_whatsapp_template_meta") as enviar:
            app.enviar_whatsapp_recordatorio_cuota_template({"nombre": "Ana"})
        self.assertEqual(enviar.call_args.args[1], app.WHATSAPP_TEMPLATE_CUOTA)
        self.assertEqual(len(enviar.call_args.args[3]), 1)
