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


if __name__ == "__main__":
    unittest.main()
