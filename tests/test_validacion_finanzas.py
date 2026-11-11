import os
import unittest

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")

import app
from services.finanzas import beca_vigente, porcentaje_beca


class ValidacionFinanzasTests(unittest.TestCase):
    def test_beca_rechaza_porcentajes_no_finitos(self):
        for valor in ("nan", "NaN", "inf", "-inf", "1e309"):
            with self.subTest(valor=valor):
                self.assertIsNone(porcentaje_beca(valor))
                self.assertFalse(beca_vigente(
                    {"beca_activa": 1, "beca_porcentaje": valor}, "2026-10"
                ))

    def test_plan_rechaza_montos_no_finitos_aun_con_monto_opcional(self):
        for valor in ("nan", "NaN", "inf", "-inf", "1e309"):
            for obligatorio in (True, False):
                with self.subTest(valor=valor, obligatorio=obligatorio):
                    with app.app.test_request_context(
                        "/", method="POST", data={"monto_total": valor}
                    ):
                        with self.assertRaisesRegex(ValueError, "finito"):
                            app.datos_plan_pago_form(require_monto=obligatorio)

    def test_conserva_porcentajes_y_planes_validos(self):
        for valor in (0, 25.5, 100):
            with self.subTest(valor=valor):
                self.assertEqual(porcentaje_beca(valor), valor)
        with app.app.test_request_context(
            "/", method="POST",
            data={"monto_total": "1200.50", "cantidad_cuotas": "2"},
        ):
            plan = app.datos_plan_pago_form()
        self.assertEqual(plan["monto_total"], 1200.5)
        self.assertEqual(plan["monto_cuota"], 600.25)
        with app.app.test_request_context("/", method="POST", data={"monto_total": "0"}):
            self.assertEqual(app.datos_plan_pago_form(require_monto=False)["monto_total"], 0)
