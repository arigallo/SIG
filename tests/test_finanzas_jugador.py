import os
import unittest
from datetime import date
from unittest.mock import MagicMock, patch

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")
import app


class FinanzasJugadorTests(unittest.TestCase):
    def test_puntualidad_pago_contempla_limite_y_fechas_faltantes(self):
        for pago, vencimiento, esperado in [
            ("2026-11-01", "2026-11-02", "Abonado a tiempo"),
            ("2026-11-02", "2026-11-02", "Abonado a tiempo"),
            ("2026-11-03", "2026-11-02", "Abonado fuera de término"),
            (None, "2026-11-02", "Sin fechas para comparar"),
            ("2026-11-03", None, "Sin fechas para comparar"),
            (date(2026, 11, 2), date(2026, 11, 2), "Abonado a tiempo"),
        ]:
            with self.subTest(pago=pago, vencimiento=vencimiento):
                self.assertEqual(app.puntualidad_pago_gasto({
                    "estado": "pagado", "fecha_pago": pago, "fecha_vencimiento": vencimiento,
                }), esperado)
        self.assertIsNone(app.puntualidad_pago_gasto({"estado": "pendiente"}))

    def test_alta_activa_portal_en_la_misma_transaccion(self):
        conn = MagicMock()
        with patch.object(app, "generar_portal_token", return_value="token-unico"):
            app.activar_portal_nuevo_jugador(conn, 17)
        sql, params = conn.execute.call_args.args
        self.assertIn("portal_activo = 1", sql)
        self.assertEqual(params[0], "token-unico")
        self.assertEqual(params[2], 17)
        conn.commit.assert_not_called()

    def test_finanzas_incluye_historial_y_solo_suma_saldos_pendientes(self):
        conn = MagicMock()
        conn.execute.side_effect = [
            MagicMock(fetchone=lambda: {"id": 7}),
            MagicMock(fetchall=lambda: []),
            MagicMock(fetchone=lambda: {"total": 100}),
            MagicMock(fetchall=lambda: []),
            MagicMock(fetchall=lambda: []),
            MagicMock(fetchall=lambda: [
                {"estado": "pagado", "importe": 50, "fecha_pago": "2026-11-03", "fecha_vencimiento": "2026-11-02"},
                {"estado": "pendiente", "importe": 25},
                {"estado": "exento", "importe": 30},
            ]),
        ]
        with app.app.test_request_context(), patch.object(app, "permiso_requerido", return_value=None), patch.object(app, "get_connection", return_value=conn), patch.object(app, "render_template") as render:
            app.ver_cuotas(7)
        datos = render.call_args.kwargs
        self.assertEqual(datos["deuda_finanzas"], 125)
        self.assertEqual(len(datos["gastos_compartidos"]), 3)
        self.assertEqual(datos["gastos_compartidos"][0]["puntualidad_pago"], "Abonado fuera de término")
