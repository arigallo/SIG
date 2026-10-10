import csv
import io
import json
import os
import unittest
from unittest.mock import MagicMock, patch

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")

from werkzeug.datastructures import MultiDict
import app


class FormulariosTests(unittest.TestCase):
    def connection(self, tipo="texto_corto", requerida=True, estado="publicada", total=0):
        conn = MagicMock()
        encuesta = {"id": 1, "titulo": "Inscripción", "estado": estado,
                    "fecha_inicio": None, "fecha_fin": None, "descripcion": "",
                    "publico_objetivo": "", "token": "a" * 24}
        pregunta = {"id": 2, "texto": "Tu respuesta", "tipo": tipo,
                    "requerida": int(requerida), "opciones": '["A", "B"]', "orden": 1}
        def execute(sql, params=None):
            result = MagicMock()
            if "COUNT(*)" in sql:
                result.fetchone.return_value = {"total": total}
            elif "SELECT" in sql and "FROM encuestas_satisfaccion" in sql:
                result.fetchone.return_value = encuesta
            elif "SELECT" in sql and "FROM encuesta_satisfaccion_preguntas" in sql:
                result.fetchall.return_value = [pregunta]
            elif "INSERT INTO encuesta_satisfaccion_respuestas" in sql:
                result.fetchone.return_value = {"id": 3}
            else:
                result.fetchall.return_value = []
            return result
        conn.execute.side_effect = execute
        return conn

    def submit(self, tipo, values, requerida=True):
        conn = self.connection(tipo, requerida)
        with app.app.test_request_context("/encuestas/" + "a" * 24, method="POST", data=MultiDict(values)):
            with patch.object(app, "get_connection", return_value=conn), \
                 patch.object(app, "consumir_limite_publico", return_value=True), \
                 patch.object(app, "registrar_auditoria"), \
                 patch.object(app, "render_template", return_value="invalid"):
                response = app.responder_encuesta_satisfaccion("a" * 24)
        return response, conn

    def test_all_field_types_can_be_created(self):
        for tipo in app.ENCUESTA_TIPOS_PREGUNTA:
            with self.subTest(tipo=tipo):
                preguntas, error = app.normalizar_preguntas_encuesta(json.dumps([
                    {"texto": "Tu respuesta", "tipo": tipo, "opciones": ["A", "B", "A"]}
                ]))
                self.assertIsNone(error)
                if tipo in {"opcion_unica", "desplegable", "opcion_multiple"}:
                    self.assertEqual(preguntas[0]["opciones"], ["A", "B"])

    def test_invalid_options_are_rejected_for_all_choice_types(self):
        for tipo in ("opcion_unica", "desplegable", "opcion_multiple"):
            _, error = app.normalizar_preguntas_encuesta(json.dumps([
                {"texto": "Tu respuesta", "tipo": tipo, "opciones": ["A"]}
            ]))
            self.assertIsNotNone(error)

    def test_multiple_selections_are_saved_without_duplicates(self):
        response, conn = self.submit("opcion_multiple", [("pregunta_2", "A"), ("pregunta_2", "A")])
        self.assertEqual(response.status_code, 302)
        inserts = [call for call in conn.execute.call_args_list if "INSERT INTO encuesta_satisfaccion_respuesta_items" in call.args[0]]
        self.assertEqual(json.loads(inserts[0].args[1][2]), ["A"])
        conn.commit.assert_called_once()

    def test_invalid_field_answers_are_rejected_before_saving(self):
        for tipo, valor in (("email", "sin-correo"), ("fecha", "2026-02-30"),
                            ("numero", "NaN"), ("numero", "Infinity"),
                            ("desplegable", "C"), ("opcion_multiple", "C"),
                            ("texto_corto", "x" * 301), ("texto_largo", "x" * 4001)):
            with self.subTest(tipo=tipo, valor=valor[:20]):
                response, conn = self.submit(tipo, [("pregunta_2", valor)])
                self.assertEqual(response[1], 400)
                conn.commit.assert_not_called()

    def test_required_multiple_selection_and_optional_empty(self):
        response, conn = self.submit("opcion_multiple", [])
        self.assertEqual(response[1], 400)
        conn.commit.assert_not_called()
        response, conn = self.submit("opcion_multiple", [], requerida=False)
        self.assertEqual(response.status_code, 302)
        conn.commit.assert_called_once()

    def test_valid_field_answers_are_saved(self):
        for tipo, valor in (("email", "ana@example.com"), ("fecha", "2026-10-09"),
                            ("numero", "-12.5"), ("numero", "1e3"), ("numero", ".5"),
                            ("desplegable", "B"), ("texto_corto", "Ana")):
            with self.subTest(tipo=tipo):
                response, conn = self.submit(tipo, [("pregunta_2", valor)])
                self.assertEqual(response.status_code, 302)
                conn.commit.assert_called_once()

    def test_editing_published_or_answered_forms_is_blocked(self):
        for estado, total in (("publicada", 0), ("borrador", 1)):
            conn = self.connection(estado=estado, total=total)
            with app.app.test_request_context("/admin/encuestas/1/editar"):
                with patch.object(app, "permiso_requerido", return_value=None), patch.object(app, "get_connection", return_value=conn):
                    response = app.nueva_encuesta_satisfaccion(1)
            self.assertEqual(response.status_code, 302)
            self.assertFalse(any("DELETE" in call.args[0] for call in conn.execute.call_args_list))

    def test_editing_keeps_existing_form_and_token(self):
        conn = self.connection(estado="borrador")
        data = {"titulo": "Inscripción", "estado": "borrador", "preguntas_json": json.dumps([
            {"texto": "Nueva pregunta", "tipo": "email", "requerida": True}
        ])}
        with app.app.test_request_context("/admin/encuestas/1/editar", method="POST", data=data):
            with patch.object(app, "permiso_requerido", return_value=None), \
                 patch.object(app, "get_connection", return_value=conn), patch.object(app, "registrar_auditoria"):
                response = app.nueva_encuesta_satisfaccion(1)
        self.assertEqual(response.status_code, 302)
        sql = [call.args[0] for call in conn.execute.call_args_list]
        self.assertTrue(any("FOR UPDATE" in statement for statement in sql))
        updates = [statement for statement in sql if "UPDATE encuestas_satisfaccion" in statement]
        self.assertEqual(len(updates), 1)
        self.assertNotIn("token =", updates[0])
        self.assertFalse(any("INSERT INTO encuestas_satisfaccion" in statement for statement in sql))
        conn.commit.assert_called_once()

    def test_public_template_renders_all_field_types_and_escapes_values(self):
        preguntas = [{"id": index, "texto": "<script>pregunta</script>", "tipo": tipo,
                      "requerida": True, "opciones_lista": ["<script>A</script>", "B"]}
                     for index, tipo in enumerate(app.ENCUESTA_TIPOS_PREGUNTA, 1)]
        with app.app.test_request_context("/encuestas/" + "a" * 24):
            rendered = app.app.jinja_env.get_template("encuesta_satisfaccion_publica.html").render(
                encuesta={"titulo": "Formulario", "descripcion": ""}, preguntas=preguntas,
                disponible=True, data={"respuestas": {}, "nombre": "", "contacto": ""},
                static_asset=lambda value: value, csrf_token=lambda: "test")
        self.assertIn('type="checkbox"', rendered)
        self.assertIn('type="email"', rendered)
        self.assertIn('type="date"', rendered)
        self.assertIn('type="number"', rendered)
        self.assertIn("&lt;script&gt;", rendered)
        self.assertNotIn("<script>pregunta</script>", rendered)

    def test_csv_keeps_answers_together_and_escapes_formulas(self):
        conn = MagicMock()
        results = [
            {"id": 1},
            [{"id": 2, "texto": "=Título", "tipo": "opcion_multiple"}],
            [{"id": 3, "creado_en": "2026-10-09", "nombre": "=CMD()", "contacto": None}],
            [{"respuesta_id": 3, "pregunta_id": 2, "valor_texto": '["A", "B"]', "valor_numero": None}],
        ]
        def execute(sql, params=None):
            result = MagicMock()
            value = results.pop(0)
            result.fetchone.return_value = value
            result.fetchall.return_value = value
            return result
        conn.execute.side_effect = execute
        with app.app.test_request_context("/admin/encuestas/1/exportar"):
            with patch.object(app, "permiso_requerido", return_value=None), patch.object(app, "get_connection", return_value=conn):
                response = app.exportar_formulario(1)
        rows = list(csv.reader(io.StringIO(response.get_data(as_text=True).lstrip("\ufeff"))))
        self.assertEqual(rows[0][-1], "'=Título")
        self.assertEqual(rows[1][2], "'=CMD()")
        self.assertEqual(rows[1][-1], "A; B")
        conn.close.assert_called_once()

    def test_export_requires_permission(self):
        with app.app.test_request_context("/admin/encuestas/1/exportar"):
            with patch.object(app, "permiso_requerido", return_value="denied"), patch.object(app, "get_connection") as connection:
                self.assertEqual(app.exportar_formulario(1), "denied")
                connection.assert_not_called()
