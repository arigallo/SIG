import os
import unittest
from unittest.mock import MagicMock, patch

from werkzeug.exceptions import Forbidden

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")
import app


class NotasFinanzasTests(unittest.TestCase):
    def test_crear_con_seguimiento_conserva_autor_y_no_cambia_saldos(self):
        conn = MagicMock()
        conn.execute.return_value.fetchone.return_value = {"id": 7}
        with app.app.test_request_context(method="POST", data={
            "tipo": "finanzas", "nota": " Paga el viernes ",
            "seguimiento_fecha": "2026-11-20",
        }), patch.object(app, "tipos_bitacora_disponibles", return_value=[{"clave": "finanzas"}]), patch.object(app, "get_connection", return_value=conn):
            app.session["username"] = "tesoreria"
            response = app.nueva_bitacora_jugador(7)
        self.assertTrue(response.location.endswith("#notas-finanzas"))
        self.assertEqual(conn.execute.call_count, 2)
        sql, params = conn.execute.call_args.args
        self.assertIn("INSERT INTO jugador_bitacora", sql)
        self.assertEqual(params, (7, "finanzas", "Paga el viernes", "tesoreria", "2026-11-20"))
        conn.commit.assert_called_once()

    def test_rechaza_fecha_invalida_texto_vacio_y_excesivo(self):
        for nota, fecha in [("Pago pendiente", "2026-02-30"), (" ", ""), ("x" * 3001, "")]:
            with self.subTest(nota=nota[:20], fecha=fecha), app.app.test_request_context(method="POST", data={
                "tipo": "finanzas", "nota": nota, "seguimiento_fecha": fecha,
            }), patch.object(app, "tipos_bitacora_disponibles", return_value=[{"clave": "finanzas"}]), patch.object(app, "get_connection") as connection:
                app.nueva_bitacora_jugador(7)
                connection.assert_not_called()

    def test_no_puede_crear_finanzas_con_permiso_general(self):
        with app.app.test_request_context(method="POST", data={"tipo": "finanzas", "nota": "Privado"}), patch.object(app, "tipos_bitacora_disponibles", return_value=[{"clave": "general"}]), patch.object(app, "get_connection") as connection:
            app.nueva_bitacora_jugador(7)
            connection.assert_not_called()

    def test_resolver_exige_permiso_financiero(self):
        with app.app.test_request_context(method="POST"), patch.object(app, "puede_crear_bitacora_tipo", return_value=False), patch.object(app, "get_connection") as connection:
            with self.assertRaises(Forbidden):
                app.resolver_nota_finanzas(7, 12)
            connection.assert_not_called()

    def test_resolver_limita_jugador_tipo_y_preserva_primera_resolucion(self):
        conn = MagicMock()
        with app.app.test_request_context(method="POST"), patch.object(app, "puede_crear_bitacora_tipo", return_value=True), patch.object(app, "get_connection", return_value=conn):
            app.session["username"] = "tesoreria"
            app.resolver_nota_finanzas(7, 12)
        sql, params = conn.execute.call_args.args
        self.assertEqual(params, ("tesoreria", 12, 7))
        self.assertIn("jugador_id = %s", sql)
        self.assertIn("tipo = 'finanzas'", sql)
        self.assertIn("resuelto_en IS NULL", sql)
        self.assertNotIn("DELETE", sql)
        conn.commit.assert_called_once()

    def test_plantilla_compila(self):
        app.app.jinja_env.get_template("jugador_detalle.html")
