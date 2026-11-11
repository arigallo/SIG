import os
import unittest
from datetime import datetime, timezone
from unittest.mock import MagicMock, patch

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")
import app
from services.madrinas import estados_rapidos, validar_cambios


class MadrinasRapidoTests(unittest.TestCase):
    def setUp(self):
        self.aspirante = {"id": 1, "estado": "Aspirante", "etapa_seguimiento": "pendiente_contacto",
                          "madrina_jugador_id": 2, "contacto_estado": None, "viene_estado": None}
        self.conn = MagicMock()

        def execute(sql, params=()):
            result = MagicMock()
            if "SELECT * FROM aspirantes" in sql:
                result.fetchone.return_value = self.aspirante
            elif "FROM jugadores" in sql:
                result.fetchone.return_value = {"id": 3, "nombre": "Sol", "apellido": "Prueba"} if params[0] == 3 else None
            elif "RETURNING fecha" in sql:
                result.fetchone.return_value = {"fecha": datetime.now(timezone.utc), "creado_por": "Mathias"}
            return result
        self.conn.execute.side_effect = execute

    def save(self, payload, permiso=None):
        with app.app.test_request_context("/ahijadxs/1/rapido", method="POST", json=payload), \
             patch.object(app, "permiso_requerido", return_value=permiso), \
             patch.object(app, "get_connection", return_value=self.conn):
            app.session["username"] = "Mathias"
            response = app.guardar_aspirante_rapido(1)
            if isinstance(response, tuple):
                return response[0].get_json(), response[1]
            return response.get_json(), response.status_code

    def test_legacy_contact_does_not_invent_confirmation(self):
        self.assertEqual(estados_rapidos({"etapa_seguimiento": "contactado"})["viene_estado"], "pendiente")
        self.assertEqual(estados_rapidos({"etapa_seguimiento": "confirmo_asistencia"})["viene_estado"], "si")
        self.assertEqual(estados_rapidos({"etapa_seguimiento": "no_interesado"})["contacto_estado"], "pendiente")

    def test_contact_save_preserves_confirmation_and_logs_actor(self):
        self.aspirante["viene_estado"] = "no"
        data, code = self.save({"cambios": {"contacto_estado": "contactado"}, "anteriores": {"contacto_estado": "pendiente"}})
        self.assertEqual(code, 200)
        self.assertEqual(data["valores"]["viene_estado"], "no")
        self.assertEqual(data["autor"], "Mathias")
        inserts = [c for c in self.conn.execute.call_args_list if "INSERT INTO aspirante_seguimientos" in c.args[0]]
        self.assertEqual(inserts[0].args[1][-1], "Mathias")
        self.conn.commit.assert_called_once()
        self.conn.close.assert_called_once()

    def test_stale_update_is_rejected_without_writing(self):
        self.aspirante["contacto_estado"] = "contactado"
        data, code = self.save({"cambios": {"contacto_estado": "contactado"}, "anteriores": {"contacto_estado": "pendiente"}})
        self.assertEqual(code, 409)
        self.conn.commit.assert_not_called()
        self.assertEqual(self.conn.execute.call_count, 1)

    def test_invalid_fields_and_values_rejected(self):
        for cambios in ({"estado": "Ingresado"}, {"viene_estado": []}, {"contacto_estado": "cualquier"},
                        {"madrina_jugador_id": True}, {"madrina_jugador_id": -1}):
            with self.subTest(cambios=cambios):
                with self.assertRaises(ValueError):
                    validar_cambios({"cambios": cambios, "anteriores": dict.fromkeys(cambios)})

    def test_notes_append_and_do_not_change_status(self):
        data, code = self.save({"nota": "Trabaja hasta las 21"})
        self.assertEqual(code, 200)
        self.assertFalse(any("UPDATE aspirantes" in c.args[0] for c in self.conn.execute.call_args_list))
        self.assertEqual(data["detalle"], "Trabaja hasta las 21")

    def test_inactive_madrina_cannot_be_assigned(self):
        _, code = self.save({"cambios": {"madrina_jugador_id": 999}, "anteriores": {"madrina_jugador_id": 2}})
        self.assertEqual(code, 400)
        self.conn.commit.assert_not_called()

    def test_assignment_logs_name_and_can_be_cleared(self):
        data, code = self.save({"cambios": {"madrina_jugador_id": 3}, "anteriores": {"madrina_jugador_id": 2}})
        self.assertEqual(code, 200)
        self.assertEqual(data["detalle"], "Responsable: Prueba, Sol")
        data, code = self.save({"cambios": {"madrina_jugador_id": None}, "anteriores": {"madrina_jugador_id": 2}})
        self.assertEqual(code, 200)
        self.assertEqual(data["valores"]["madrina_jugador_id"], None)

    def test_closed_and_missing_records_cannot_be_updated(self):
        self.aspirante["estado"] = "Ingresado"
        _, code = self.save({"nota": "Prueba"})
        self.assertEqual(code, 409)
        self.aspirante = None
        _, code = self.save({"nota": "Prueba"})
        self.assertEqual(code, 404)
        self.conn.commit.assert_not_called()

    def test_database_failure_rolls_back_and_closes(self):
        self.conn.execute.side_effect = RuntimeError("Database unavailable")
        with self.assertRaises(RuntimeError):
            self.save({"nota": "Prueba"})
        self.conn.rollback.assert_called_once()
        self.conn.close.assert_called_once()

    def test_permission_failure_does_not_open_connection(self):
        with app.app.test_request_context("/ahijadxs/1/rapido", method="POST", json={"nota": "Prueba"}), \
             patch.object(app, "permiso_requerido", return_value=("Forbidden", 403)), \
             patch.object(app, "get_connection") as connect:
            self.assertEqual(app.guardar_aspirante_rapido(1)[1], 403)
            connect.assert_not_called()

    def test_whatsapp_link_is_encoded_and_missing_phone_has_no_link(self):
        self.assertIsNone(app.preparar_aspirante_rapido({"nombre": "Alex", "telefono": ""})["whatsapp_url"])
        url = app.preparar_aspirante_rapido({"nombre": "Ana & Sol", "telefono": "+54 9 11 2345 6789"})["whatsapp_url"]
        self.assertTrue(url.startswith("https://wa.me/5491123456789?text="))
        self.assertIn("%26", url)


if __name__ == "__main__":
    unittest.main()
