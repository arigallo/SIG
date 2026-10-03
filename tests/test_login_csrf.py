import os
import unittest
from unittest.mock import MagicMock, patch

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")

import app
from werkzeug.security import generate_password_hash


class LoginCsrfTests(unittest.TestCase):
    def setUp(self):
        self.client = app.app.test_client()
        self.notices = patch.object(app, "obtener_avisos_login_publicos", return_value=[])
        self.notices.start()
        self.addCleanup(self.notices.stop)

    def test_login_is_not_cached_and_has_autofill_fields(self):
        response = self.client.get("/login")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["Cache-Control"], "no-store")
        self.assertIn(b'autocomplete="current-password"', response.data)
        self.assertIn(b'data-csrf-url="/login/csrf"', response.data)

    def test_refresh_returns_current_session_token_without_rotating_it(self):
        first = self.client.get("/login/csrf")
        self.assertEqual(first.status_code, 200)
        self.assertEqual(first.headers["Cache-Control"], "no-store")
        token = first.json["token"]
        with self.client.session_transaction() as session:
            self.assertEqual(session["_csrf_token"], token)
        self.assertEqual(self.client.get("/login/csrf").json["token"], token)

    def test_stale_or_missing_token_rejects_credentials_with_recoverable_notice(self):
        with self.client.session_transaction() as session:
            session["_csrf_token"] = "current"
        for token in ("stale", ""):
            with self.subTest(token=token), patch.object(app, "get_connection") as connection:
                response = self.client.post("/login", data={
                    "username": "admin", "password": "secret", "_csrf_token": token,
                })
                self.assertEqual(response.status_code, 303)
                connection.assert_not_called()
                page = self.client.get(response.location)
                self.assertIn("La sesión de ingreso cambió o venció".encode(), page.data)
                self.assertIn(b'entry-action-admin" open', page.data)
                with self.client.session_transaction() as session:
                    self.assertNotIn("user_id", session)

    def test_refresh_after_session_loss_allows_valid_autofilled_credentials(self):
        self.client.get("/login")
        with self.client.session_transaction() as session:
            session.clear()
        token = self.client.get("/login/csrf").json["token"]
        password = "Autofill & + ñ password"
        connection = MagicMock()
        connection.execute.return_value.fetchone.return_value = {
            "id": 1, "username": "admin", "rol": "admin",
            "password": generate_password_hash(password), "rol_permisos": "[]",
            "debe_cambiar_password": False, "onboarding_visto": True,
        }
        with patch.object(app, "get_connection", return_value=connection), patch.object(
            app, "login_bloqueado", return_value=False
        ), patch.object(app, "registrar_intento_login"), patch.object(app, "registrar_auditoria"):
            response = self.client.post("/login", data={
                "username": "admin", "password": password, "_csrf_token": token,
            })
        self.assertEqual(response.status_code, 302)
        with self.client.session_transaction() as session:
            self.assertEqual(session["user_id"], 1)
            self.assertNotEqual(session["_csrf_token"], token)

    def test_other_post_routes_still_reject_invalid_csrf(self):
        with patch.object(app, "registrar_auditoria"):
            self.assertEqual(self.client.post("/admin/simular-rol").status_code, 400)


if __name__ == "__main__":
    unittest.main()
