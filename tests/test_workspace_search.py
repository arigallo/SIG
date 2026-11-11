import os
import unittest
from unittest.mock import MagicMock, patch
from werkzeug.exceptions import Forbidden

os.environ['INIT_DB'] = 'false'
os.environ.setdefault('SECRET_KEY', 'test-only-secret-key')
import app


class WorkspaceSearchTests(unittest.TestCase):
    def test_busqueda_exige_permiso_y_no_consulta_si_no_lo_tiene(self):
        with app.app.test_request_context('/api/buscar-jugadores?q=Sam'), patch.object(app, 'tiene_permiso', return_value=False), patch.object(app, 'get_connection') as connection:
            with self.assertRaises(Forbidden):
                app.buscar_jugadores_global()
            connection.assert_not_called()

    def test_busqueda_medica_no_expone_finanzas_ni_dni(self):
        conn = MagicMock()
        conn.execute.return_value.fetchall.return_value = [dict(id=7, nombre='Sam', apellido='Ejemplo', numero_socio='00007', categoria='Superior')]
        with app.app.test_request_context('/api/buscar-jugadores?q=Sam%20Ejemplo'), patch.object(app, 'tiene_permiso', side_effect=lambda *permissions: 'jugadores_ver' in permissions), patch.object(app, 'get_connection', return_value=conn):
            response = app.buscar_jugadores_global()
        item = response.json['resultados'][0]
        self.assertIn('perfil', item)
        self.assertNotIn('finanzas', item)
        self.assertNotIn('notas', item)
        self.assertNotIn('dni', item)
        self.assertEqual(response.headers['Cache-Control'], 'no-store')

    def test_busqueda_financiera_ofrece_accesos_y_trata_comodines_literalmente(self):
        conn = MagicMock()
        conn.execute.return_value.fetchall.return_value = [dict(id=7, nombre='Sam', apellido='Ejemplo', numero_socio='00007', categoria=None)]
        with app.app.test_request_context('/api/buscar-jugadores?q=25%25'), patch.object(app, 'tiene_permiso', return_value=True), patch.object(app, 'get_connection', return_value=conn):
            item = app.buscar_jugadores_global().json['resultados'][0]
        self.assertTrue(item['notas'].endswith('#notas-finanzas'))
        self.assertEqual(conn.execute.call_args.args[1], ['%25\\%%'])

    def test_busqueda_corta_no_enumera_jugadores(self):
        with app.app.test_request_context('/api/buscar-jugadores?q=a'), patch.object(app, 'tiene_permiso', return_value=True), patch.object(app, 'get_connection') as connection:
            self.assertEqual(app.buscar_jugadores_global().json, {'resultados': []})
            connection.assert_not_called()

    def test_filtros_se_recuerdan_y_limpiar_restaura_defaults(self):
        with app.app.test_request_context('/jugadores?q=Sam&pagina=4&ignorar=externo'):
            defaults = dict(q='', pagina='1')
            values = app.filtros_workspace('jugadores', defaults)
            self.assertEqual(values, dict(q='Sam', pagina='4'))
            saved = dict(app.session)
        with app.app.test_request_context('/jugadores'):
            app.session.update(saved)
            self.assertEqual(app.filtros_workspace('jugadores', defaults), values)
        with app.app.test_request_context('/jugadores?limpiar=1'):
            app.session.update(saved)
            self.assertEqual(app.filtros_workspace('jugadores', defaults), defaults)
