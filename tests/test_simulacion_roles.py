import os
import unittest
from unittest.mock import patch

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")

import app


class _Rows:
    def fetchall(self):
        return [{"nombre": "entrenador", "descripcion": "Entrenamiento", "permisos": '["calendario_ver"]'}]


class _Connection:
    def execute(self, sql, params=None):
        return _Rows()

    def close(self):
        pass


class SimulacionRolesTests(unittest.TestCase):
    def test_admin_can_simulate_read_only_and_return(self):
        client = app.app.test_client()
        with client.session_transaction() as session:
            session.update({
                "user_id": 1,
                "username": "admin",
                "rol": "admin",
                "permisos": ["roles_gestionar"],
                "_csrf_token": "test-token",
            })
        with patch.object(app, "get_connection", return_value=_Connection()), patch.object(
            app, "obtener_config_mantenimiento", return_value={"activo": False}
        ):
            response = client.post("/admin/simular-rol", data={"rol": "entrenador", "_csrf_token": "test-token"})
            self.assertEqual(response.status_code, 302)
            with client.session_transaction() as session:
                self.assertEqual(session["rol"], "entrenador")
                self.assertEqual(session["permisos"], ["calendario_ver"])

            response = client.post("/admin/versiones", data={"_csrf_token": "test-token"})
            self.assertEqual(response.status_code, 403)

            response = client.post("/admin/simular-rol/salir", data={"_csrf_token": "test-token"})
            self.assertEqual(response.status_code, 302)
            with client.session_transaction() as session:
                self.assertEqual(session["rol"], "admin")
                self.assertEqual(session["permisos"], ["roles_gestionar"])
                self.assertNotIn("simulacion_rol_original", session)


if __name__ == "__main__":
    unittest.main()
