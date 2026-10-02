import os
import unittest
from unittest.mock import patch

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")
import app


class ModificarMontosTests(unittest.TestCase):
    def test_preserva_exencion_social_con_plan_de_version_actual(self):
        resultado = app.calcular_monto_cuota_modificada({"becada": 1, "beca_porcentaje": 25, "plan_pago_monto": 1500, "plan_pago_detalle": "Plan #1; Cuota social cubierta por beca"}, 10000)
        self.assertEqual(resultado["importe"], 1500)
        self.assertEqual(resultado["descuento_beca"], 10000)

    def test_preserva_beca_y_plan(self):
        resultado = app.calcular_monto_cuota_modificada(
            {"beca_porcentaje": 25, "plan_pago_monto": 1500}, 10000)
        self.assertEqual(resultado, {"importe": 9000, "importe_original": 11500, "descuento_beca": 2500})

    def test_rechaza_importes_y_periodos_invalidos_sin_abrir_db(self):
        for importe, desde, hasta in [("nan", "2026-01", "2026-02"), ("inf", "2026-01", "2026-02"), ("-1", "2026-01", "2026-02"), ("100", "2026-03", "2026-02"), ("100", "2026-13", "2026-14")]:
            with self.subTest(importe=importe, desde=desde):
                with app.app.test_request_context("/cuotas/modificar-montos", method="POST", data={"importe": importe, "periodo_desde": desde, "periodo_hasta": hasta}):
                    with patch.object(app, "permiso_requerido", return_value=None), patch.object(app, "get_connection") as conexion:
                        self.assertEqual(app.modificar_montos_cuotas()[1], 400)
                        conexion.assert_not_called()

    def test_aplica_solo_seleccionadas_y_audita(self):
        from unittest.mock import MagicMock
        conn = MagicMock()
        conn.execute.return_value.fetchall.return_value = [
            {"id": 1, "importe": 100, "beca_porcentaje": 50, "plan_pago_monto": 20},
            {"id": 2, "importe": 100},
        ]
        with app.app.test_request_context("/cuotas/modificar-montos", method="POST", data={"importe": "200", "periodo_desde": "2026-01", "periodo_hasta": "2026-02", "categoria": "Primera", "accion": "aplicar", "cuota_id": "1"}):
            with patch.object(app, "permiso_requerido", return_value=None), patch.object(app, "get_connection", return_value=conn), patch.object(app, "registrar_auditoria") as audit:
                self.assertEqual(app.modificar_montos_cuotas().status_code, 302)
                audit.assert_called_once()
        consulta = conn.execute.call_args_list[0].args[0]
        for condicion in ["c.pagado = 0", "COALESCE(c.anulada, 0) = 0", "'pendiente', 'aceptado'", "j.categoria = %s", "FOR UPDATE OF c"]:
            self.assertIn(condicion, consulta)
        self.assertEqual(conn.execute.call_args_list[1].args[1], (120, 220, 100, 1))
        self.assertEqual(conn.execute.call_count, 2)
        conn.commit.assert_called_once()
        conn.close.assert_called_once()


if __name__ == "__main__":
    unittest.main()
