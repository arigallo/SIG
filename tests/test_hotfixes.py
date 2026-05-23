import os
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

os.environ["INIT_DB"] = "false"

from flask import render_template
from openpyxl import Workbook, load_workbook

import app


class HotfixTests(unittest.TestCase):
    def test_timezone_aware_datetime_is_serialized_before_saving(self):
        wb = Workbook()
        ws = wb.active

        app.append_fila_reporte(
            ws,
            [datetime(2026, 5, 20, 14, 0, 51, tzinfo=timezone.utc)],
        )

        self.assertEqual(ws["A1"].value, "2026-05-20 14:00:51")

        fd, path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        try:
            wb.save(path)
            reloaded = load_workbook(path)
            self.assertEqual(reloaded.active["A1"].value, "2026-05-20 14:00:51")
        finally:
            if os.path.exists(path):
                os.remove(path)

    def test_usuarios_template_accepts_string_ultimo_login(self):
        usuario = {
            "id": 1,
            "username": "admin",
            "email": "admin@example.com",
            "rol": "admin",
            "debe_cambiar_password": 0,
            "onboarding_visto": 1,
            "ultimo_login": "2026-05-20 12:34:56+00",
        }

        with app.app.test_request_context("/usuarios"):
            html = render_template("usuarios.html", usuarios=[usuario])

        self.assertIn("2026-05-20 12:34", html)

    def test_no_known_unsafe_fecha_vencimiento_casts(self):
        source = Path("app.py").read_text(encoding="utf-8")
        unsafe_patterns = [
            "fecha_vencimiento <> ''\n                     AND fecha_vencimiento::date",
            "c.fecha_vencimiento <> ''\n                     AND c.fecha_vencimiento::date",
            "f.fecha_vencimiento <> ''\n          AND f.fecha_vencimiento::date",
            "c.fecha_vencimiento <> ''\n             AND c.fecha_vencimiento::date",
        ]

        for pattern in unsafe_patterns:
            with self.subTest(pattern=pattern):
                self.assertNotIn(pattern, source)

    def test_asistencia_nuevo_redirects_to_calendario_form(self):
        with patch.object(app, "obtener_config_mantenimiento", return_value={"activo": False}):
            client = app.app.test_client()
            with client.session_transaction() as session:
                session["user_id"] = 1
                session["rol"] = "entrenador"
                session["permisos"] = ["asistencia_gestionar"]

            response = client.get("/asistencia/nuevo")

        self.assertEqual(response.status_code, 302)
        self.assertIn("/calendario/nuevo?origen=asistencia", response.headers["Location"])

    def test_asistencia_listing_orders_upcoming_events_first(self):
        source = Path("app.py").read_text(encoding="utf-8")

        self.assertIn("WHEN e.fecha >= CURRENT_DATE::text THEN 0", source)
        self.assertIn("WHEN e.fecha >= CURRENT_DATE::text THEN e.fecha", source)
        self.assertIn("WHEN e.fecha < CURRENT_DATE::text THEN e.fecha", source)

    def test_whatsapp_unsupported_message_shows_meta_reason(self):
        resumen = app.resumir_contenido_whatsapp(
            "unsupported",
            {
                "type": "unsupported",
                "errors": [
                    {
                        "code": 131051,
                        "title": "Unsupported message type",
                        "details": "Message type is not currently supported",
                    }
                ],
            },
        )

        self.assertEqual(resumen, "[Mensaje no compatible] Unsupported message type")

    def test_whatsapp_interactive_and_contact_messages_are_readable(self):
        interactive = app.resumir_contenido_whatsapp(
            "interactive",
            {"interactive": {"button_reply": {"id": "ok", "title": "Confirmar"}}},
        )
        contacto = app.resumir_contenido_whatsapp(
            "contacts",
            {"contacts": [{"name": {"formatted_name": "Ariel Gallo"}}]},
        )

        self.assertEqual(interactive, "Confirmar")
        self.assertEqual(contacto, "[Contacto] Ariel Gallo")

    def test_whatsapp_inbox_email_notification_uses_configured_recipients(self):
        enviados = []

        def fake_enviar_email(destinatario, asunto, cuerpo):
            enviados.append((destinatario, asunto, cuerpo))
            return True, None

        with app.app.test_request_context("/webhooks/whatsapp", base_url="https://sig.example.test"):
            with patch.object(app, "WHATSAPP_INBOX_NOTIFY_EMAILS", ["avisos@example.com"]):
                with patch.object(app, "suprimir_email_whatsapp_por_presencia", return_value=False):
                    with patch.object(app, "enviar_email", side_effect=fake_enviar_email):
                        enviado = app.enviar_notificacion_whatsapp_inbox_email(
                            mensaje={"tipo": "text", "texto": "Hola"},
                            telefono="5491112345678",
                            jugador={"apellido": "Gallo", "nombre": "Ariel"},
                        )

        self.assertTrue(enviado)
        self.assertEqual(enviados[0][0], "avisos@example.com")
        self.assertIn("Nueva respuesta WhatsApp - Gallo, Ariel", enviados[0][1])
        self.assertIn("Mensaje: Hola", enviados[0][2])
        self.assertIn("/comunicacion/whatsapp?telefono=5491112345678", enviados[0][2])

    def test_whatsapp_inbox_email_notification_falls_back_to_smtp_from(self):
        with patch.object(app, "WHATSAPP_INBOX_NOTIFY_EMAILS", []):
            with patch.object(app, "SMTP_FROM", "tesoreria@example.com"):
                self.assertEqual(
                    app.destinatarios_notificacion_whatsapp_inbox(),
                    ["tesoreria@example.com"],
                )

    def test_whatsapp_email_notification_is_suppressed_by_presence(self):
        with patch.object(app, "suprimir_email_whatsapp_por_presencia", return_value=True):
            with patch.object(app, "enviar_email") as enviar_email:
                enviado = app.enviar_notificacion_whatsapp_inbox_email(
                    mensaje={"tipo": "text", "texto": "Hola"},
                    telefono="5491112345678",
                    jugador=None,
                )

        self.assertFalse(enviado)
        enviar_email.assert_not_called()

    def test_presence_heartbeat_requires_session_and_records_user(self):
        with patch.object(app, "obtener_config_mantenimiento", return_value={"activo": False}):
            with patch.object(app, "registrar_presencia_usuario", return_value=True) as registrar:
                client = app.app.test_client()
                with client.session_transaction() as session:
                    session["user_id"] = 1
                    session["username"] = "arielgallo"
                    session["rol"] = "admin"
                    session["permisos"] = []
                    session["_csrf_token"] = "token"

                response = client.post(
                    "/presencia/heartbeat",
                    headers={"X-CSRF-Token": "token"},
                )

        self.assertEqual(response.status_code, 200)
        registrar.assert_called_once_with("arielgallo")

    def test_whatsapp_status_endpoint_returns_json(self):
        with patch.object(app, "obtener_config_mantenimiento", return_value={"activo": False}):
            with patch.object(app, "permiso_requerido", return_value=None):
                with patch.object(app, "obtener_estado_whatsapp_inbox", return_value={"sin_leer": 2, "ultimo_id": 10}):
                    client = app.app.test_client()
                    with client.session_transaction() as session:
                        session["user_id"] = 1
                        session["username"] = "arielgallo"
                        session["rol"] = "admin"
                        session["permisos"] = ["comunicaciones_ver"]

                    response = client.get("/comunicacion/whatsapp/estado")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json(), {"sin_leer": 2, "ultimo_id": 10})

    def test_notificaciones_template_highlights_sig_user_events(self):
        with app.app.test_request_context("/notificaciones"):
            with patch.object(app, "obtener_contador_notificaciones", return_value=3):
                with patch.object(app, "obtener_contador_whatsapp_inbox", return_value=1):
                    html = render_template(
                        "notificaciones.html",
                        cuotas_vencidas=[],
                        cuotas_por_vencer=[],
                        fichas=[],
                        asistencia_baja=[],
                        comprobantes_pendientes=[{
                            "apellido": "Gallo",
                            "nombre": "Ariel",
                            "periodo": "2026-05",
                            "importe": 1000,
                            "comprobante_fecha": "2026-05-23",
                            "comprobante_usuario": "portal",
                            "comprobante_nombre": "ticket.pdf",
                            "jugador_id": 1,
                            "_notificacion_tipo": "comprobante",
                            "_notificacion_id": "1:2026-05-23",
                            "_notificacion_key": "comprobante:1:2026-05-23",
                        }],
                        whatsapp_conversaciones=[{
                            "apellido": None,
                            "nombre": None,
                            "jugador_id": None,
                            "telefono": "5491111111111",
                            "categoria": "",
                            "texto": "Hola",
                            "tipo": "text",
                            "sin_leer": 1,
                            "creado_en": "2026-05-23 10:00",
                            "_notificacion_tipo": "whatsapp",
                            "_notificacion_id": "5491111111111:10",
                            "_notificacion_key": "whatsapp:5491111111111:10",
                        }],
                        secretaria_documentos=[],
                        ahijadxs_objetivo=[],
                        cambios_portal=[{
                            "apellido": "Perez",
                            "nombre": "Juan",
                            "detalle_resumen": "Cambio telefono",
                            "fecha": "2026-05-23 09:00",
                            "jugador_id": 2,
                            "_notificacion_tipo": "cambio_portal",
                            "_notificacion_id": "20",
                            "_notificacion_key": "cambio_portal:20",
                        }],
                        whatsapp_api_activa=False,
                    )

        self.assertIn("Hay un nuevo mensaje de WhatsApp", html)
        self.assertIn("Sin vincular", html)
        self.assertIn("Perez, Juan modific", html)
        self.assertIn("Hay comprobantes por verificar", html)
        self.assertIn("Ignorar", html)

    def test_dismiss_notification_records_user_setting(self):
        with patch.object(app, "obtener_config_mantenimiento", return_value={"activo": False}):
            with patch.object(app, "descartar_notificacion_usuario", return_value=True) as descartar:
                with patch.object(app, "registrar_auditoria") as auditar:
                    client = app.app.test_client()
                    with client.session_transaction() as session:
                        session["user_id"] = 1
                        session["username"] = "arielgallo"
                        session["rol"] = "admin"
                        session["permisos"] = ["comunicaciones_ver"]
                        session["_csrf_token"] = "token"

                    response = client.post(
                        "/notificaciones/descartar",
                        data={
                            "_csrf_token": "token",
                            "tipo": "comprobante",
                            "entidad_id": "1:2026-05-23",
                        },
                    )

        self.assertEqual(response.status_code, 302)
        self.assertIn("/notificaciones", response.headers["Location"])
        descartar.assert_called_once_with("comprobante", "1:2026-05-23")
        auditar.assert_any_call(
            "descartar",
            "notificacion",
            "comprobante:1:2026-05-23",
            {"tipo": "comprobante", "entidad_id": "1:2026-05-23"},
        )


if __name__ == "__main__":
    unittest.main()
