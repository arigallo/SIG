import unittest
from unittest.mock import patch

from services import calendario


class CalendarioCoreTests(unittest.TestCase):
    def setUp(self):
        self.data = {"fecha": "2026-10-01", "tipo": "Entrenamiento", "titulo": "Entrenamiento"}

    def test_single_event_does_not_duplicate_existing_event(self):
        with patch.object(calendario, "existe_evento_calendario", return_value={"id": 10}), patch.object(
            calendario, "crear_evento_calendario_desde_data"
        ) as crear:
            creados, omitidos = calendario.crear_eventos_calendario(None, self.data)
        self.assertEqual(creados, [])
        self.assertEqual(omitidos, ["2026-10-01"])
        crear.assert_not_called()

    def test_monthly_batch_skips_only_existing_dates(self):
        with patch.object(calendario, "existe_evento_calendario", side_effect=[{"id": 10}, None]), patch.object(
            calendario, "crear_evento_calendario_desde_data", return_value=(11, 22)
        ) as crear:
            creados, omitidos = calendario.crear_eventos_calendario(
                None, self.data, fechas_recurrentes=["2026-10-01", "2026-10-03"], crear_recurrentes=True
            )
        self.assertEqual(omitidos, ["2026-10-01"])
        self.assertEqual(creados, [{"id": 11, "fecha": "2026-10-03", "asistencia_evento_id": 22}])
        self.assertEqual(crear.call_count, 1)

    def test_creation_serializes_same_event_key_before_duplicate_check(self):
        class Connection:
            def __init__(self):
                self.locked = False

            def execute(self, sql, params):
                self.locked = "pg_advisory_xact_lock" in sql

        conn = Connection()
        with patch.object(calendario, "existe_evento_calendario", side_effect=lambda _conn, _data: {"id": 1} if conn.locked else None):
            creados, omitidos = calendario.crear_eventos_calendario(conn, self.data)
        self.assertEqual(creados, [])
        self.assertEqual(omitidos, ["2026-10-01"])


if __name__ == "__main__":
    unittest.main()
