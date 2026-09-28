import os
import unittest
from unittest.mock import MagicMock, patch

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")

import app


class AvisosPostulacionTests(unittest.TestCase):
    def test_only_new_public_application_creates_and_sends_notice(self):
        conn = MagicMock()

        def execute(sql, params=None):
            result = MagicMock()
            if "SELECT id FROM aspirantes" in sql:
                result.fetchone.return_value = None
            elif "INSERT INTO aspirantes" in sql:
                result.fetchone.return_value = {"id": 42}
            elif "INSERT INTO avisos_postulacion" in sql:
                result.fetchone.return_value = {"id": 7}
            return result

        conn.execute.side_effect = execute
        destinatario = {"id": 3, "username": "madrina", "email": "madrina@example.org"}
        datos = {
            "nombre": "Ana", "apellido": "Prueba", "fecha_nacimiento": "2000-01-01",
            "telefono": "1123456789", "consentimiento_contacto": "on",
        }
        with app.app.test_request_context("/postulate", method="POST", data=datos), patch.object(
            app, "get_connection", return_value=conn
        ), patch.object(app, "consumir_limite_publico", return_value=True), patch.object(
            app, "destinatario_avisos_postulacion", return_value=destinatario
        ), patch.object(app, "enviar_aviso_postulacion") as enviar:
            response = app.postulacion_aspirante_publica()

        self.assertEqual(response.status_code, 303)
        self.assertTrue(any("INSERT INTO avisos_postulacion" in call.args[0] for call in conn.execute.call_args_list))
        enviar.assert_called_once()
        self.assertEqual(enviar.call_args.args[:2], (7, destinatario))

    def test_email_failure_does_not_prevent_push_or_in_app_notice(self):
        conn = MagicMock()
        destinatario = {"id": 3, "email": "madrina@example.org"}
        aspirante = {"nombre": "Ana", "apellido": "Prueba", "telefono": "1123456789", "email": ""}
        with app.app.test_request_context("/postulate"), patch.object(
            app, "enviar_email", return_value=(False, "smtp")
        ), patch.object(app, "get_connection", return_value=conn), patch.object(
            app, "enviar_push_por_actor"
        ) as push:
            app.enviar_aviso_postulacion(7, destinatario, aspirante)
        self.assertEqual(conn.execute.call_args.args[1], ("fallido", "smtp", 7))
        push.assert_called_once()


if __name__ == "__main__":
    unittest.main()
