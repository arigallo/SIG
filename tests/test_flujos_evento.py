import os
import unittest
from unittest.mock import patch

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")

import app


class _Result:
    def __init__(self, row=None, rows=None):
        self.row = row
        self.rows = rows or []

    def fetchone(self):
        return self.row

    def fetchall(self):
        return self.rows


class _Connection:
    def __init__(self):
        self.params = []
        self.closed = False

    def execute(self, sql, params=None):
        self.params.append(params)
        if "FROM calendario_eventos WHERE id" in sql:
            return _Result(row={"id": 11, "titulo": "Partido", "fecha": "2026-10-01", "asistencia_evento_id": 22})
        if "FROM portal_asistencia_confirmaciones" in sql:
            return _Result(rows=[{"estado": "confirmado", "total": 4}, {"estado": "no_asiste", "total": 2}])
        if "FROM asistencias" in sql:
            return _Result(row={"presentes": 3, "ausentes": 1})
        if "FROM gastos_compartidos" in sql:
            return _Result(rows=[{"id": 30, "titulo": "Tercer tiempo"}])
        raise AssertionError(sql)

    def close(self):
        self.closed = True


class FlujoEventoTests(unittest.TestCase):
    def test_event_hub_uses_linked_attendance_and_expenses(self):
        conn = _Connection()
        with app.app.test_request_context("/calendario/11"), patch.object(app, "permiso_requerido", return_value=None), patch.object(
            app, "tiene_permiso", return_value=True
        ), patch.object(app, "get_connection", return_value=conn), patch.object(app, "render_template", side_effect=lambda _name, **kwargs: kwargs):
            data = app.detalle_evento_calendario(11)
        self.assertEqual(data["confirmaciones"], {"confirmado": 4, "dudoso": 0, "no_asiste": 2})
        self.assertEqual(data["asistencia_real"]["presentes"], 3)
        self.assertEqual(data["gastos"][0]["id"], 30)
        self.assertEqual(conn.params, [(11,), (22,), (22,), (11,)])
        self.assertTrue(conn.closed)


if __name__ == "__main__":
    unittest.main()
