import os
import unittest
import json
from datetime import date
from unittest.mock import MagicMock, patch

os.environ['INIT_DB'] = 'false'
os.environ.setdefault('SECRET_KEY', 'test-only-secret-key')
import app
from services.presupuesto import numero, obtener_indicadores, obtener_padron, simular, validar_parametros


class PresupuestoTests(unittest.TestCase):
    def setUp(self):
        self.indicadores = {'base_cuotas': 1000000, 'tasa': 80}
        self.items = [{'activo': 1, 'tipo': 'egreso', 'monto': 550000,
                       'mes_inicio': '2026-01', 'mes_fin': None}]

    def test_margen_y_cuota_para_objetivo(self):
        r = simular(self.items, self.indicadores, {'objetivo': '900000', 'sueldo_incluido': 'si'}, 2027)
        bajo, base, alto = r['escenarios']
        self.assertAlmostEqual(base['aumento_maximo'], 250000)
        self.assertAlmostEqual(bajo['aumento_maximo'], 100000)
        self.assertFalse(base['viable'])
        self.assertAlmostEqual(base['aumento_cuota_adicional'], 12.5)
        r = simular(self.items, self.indicadores,
                    {'objetivo': '900000', 'aumento_cuota': '12.5', 'sueldo_incluido': 'si'}, 2027)
        self.assertTrue(r['base_esperada']['viable'])

    def test_cuotas_reemplazan_ingresos_para_no_duplicarlos(self):
        items = self.items + [{'activo': 1, 'tipo': 'ingreso', 'monto': 800000}]
        base = simular(items, self.indicadores, {'sueldo_incluido': 'si'}, 2027)['base_esperada']
        self.assertEqual(base['filas'][0]['ingresos'], 800000)

    def test_reserva_se_cumple_cada_mes_no_solo_en_diciembre(self):
        items = self.items + [{'activo': 1, 'tipo': 'egreso', 'monto': 300000,
                              'mes_inicio': '2027-01', 'mes_fin': '2027-01'}]
        base = simular(items, self.indicadores, {'sueldo_incluido': 'si'}, 2027)['base_esperada']
        self.assertTrue(base['deficit_base'])
        self.assertEqual(base['aumento_maximo'], 0)
        self.assertGreater(base['saldo_final'], 0)
        self.assertFalse(base['viable'])

    def test_sueldo_no_incluido_cargas_y_saldo_inicial(self):
        base = simular([], self.indicadores, {'sueldo_incluido': 'no', 'cargas': '20',
                       'saldo': '120000', 'reserva': '120000'}, 2027)['base_esperada']
        self.assertEqual(base['filas'][0]['gastos'], 660000)
        self.assertEqual(base['aumento_maximo'], 116666.66)

    def test_presupuesto_no_aplica_aumento_de_cuotas(self):
        base = simular(self.items, self.indicadores,
                       {'fuente': 'presupuesto', 'aumento_cuota': '100'}, 2027)['base_esperada']
        self.assertEqual(base['filas'][0]['ingresos'], 0)

    def test_sin_datos_y_entradas_invalidas(self):
        for valor in ['NaN', 'inf', '-1', 'basura', '1e1000']:
            self.assertEqual(numero(valor, 12), 12)
        base = simular([], {'base_cuotas': 0, 'tasa': 0}, {}, 2027)['base_esperada']
        self.assertEqual(base['aumento_maximo'], 0)

    def test_indicadores_incluyen_meses_vacios_y_tasa_ponderada(self):
        conn = MagicMock()
        conn.execute.return_value.fetchall.side_effect = [
            [{'mes': '2026-09', 'cantidad': 2, 'emitido': 100, 'cobrado': 50, 'en_termino': 25}],
            [{'id': 1, 'deuda': 50, 'vencido': 30}]]
        conn.execute.return_value.fetchone.return_value = {'saldo': 20}
        r = obtener_indicadores(conn, date(2026, 11, 14), ['2026-09', '2026-10'])
        self.assertEqual(r['base_cuotas'], 50)
        self.assertEqual(r['tasa'], 50)
        self.assertEqual(r['historial'][1]['cantidad'], 0)
        self.assertEqual(r['vencido'], 30)
        for call in conn.execute.call_args_list[:2]:
            self.assertIn('anulada', call.args[0])
            self.assertIn('incobrable', call.args[0])

    def test_ruta_renderiza_y_limita_horizonte(self):
        conn = MagicMock()
        conn.execute.return_value.fetchall.return_value = []
        conn.execute.return_value.fetchone.return_value = {'saldo': 0}
        presupuesto = {'items': [], 'proyeccion': [], 'resumen': {
            'ingresos_presupuestados': 0, 'egresos_presupuestados': 0,
            'resultado_presupuestado': 0, 'resultado_real': 0}}
        with app.app.test_request_context('/presupuesto?vista=planificar&meses=99'), \
             patch.object(app, 'permiso_requerido', return_value=None), \
             patch.object(app, 'obtener_presupuesto_mensual', return_value=presupuesto) as obtener, \
             patch.object(app, 'get_connection', return_value=conn):
            rendered = app.ver_presupuesto()
        self.assertIn('¿Alcanza para lo que queremos hacer?', rendered)
        self.assertIn('550000', rendered)
        self.assertEqual(obtener.call_args.args[1], 24)
        conn.close.assert_called_once()

    def test_cobro_diferido_no_es_dinero_disponible_en_enero(self):
        r = simular([], {'base_cuotas': 100, 'tasa': 100},
                    {'sueldo': 0, 'objetivo': 0, 'demora': 50}, 2027)['base_esperada']
        self.assertEqual(r['filas'][0]['cobro'], 50)
        self.assertEqual(r['filas'][1]['cobro'], 100)
        self.assertEqual(r['ingresos_total'], 1150)
        self.assertEqual(r['cuotas_por_cobrar_diciembre'], 50)

    def test_socios_becas_bajas_y_aumento_desde_mes_elegido(self):
        indicadores = {**self.indicadores, 'tasa': 100,
                       'padron': [{'socios': 10, 'descuento': 25}] * 12}
        r = simular([], indicadores, {'fuente': 'socios', 'cuota_unitaria': 100,
                    'sueldo': 0, 'objetivo': 0, 'socios_2': 8,
                    'aumento_cuota': 20, 'mes_cuota': 2}, 2027)['base_esperada']
        self.assertEqual(r['filas'][0]['emision'], 750)
        self.assertEqual(r['filas'][1]['emision'], 720)
        self.assertEqual(r['filas'][2]['emision'], 900)

    def test_aumento_sueldo_desde_julio_y_cuota_necesaria_se_puede_reaplicar(self):
        args = {'sueldo_incluido': 'si', 'objetivo': 1000000, 'mes_sueldo': 7,
                'mes_cuota': 7, 'demora': 20, 'saldo': 200000}
        r = simular(self.items, self.indicadores, args, 2027)['base_esperada']
        self.assertEqual(r['filas'][5]['adicional'], 0)
        self.assertEqual(r['filas'][6]['adicional'], 450000)
        self.assertTrue(r['cuota_calculable'])
        args['aumento_cuota'] = r['aumento_cuota_adicional']
        self.assertTrue(simular(self.items, self.indicadores, args, 2027)['base_esperada']['viable'])

    def test_aumento_tardio_no_resuelve_un_faltante_previo(self):
        r = simular(self.items, {'base_cuotas': 100000, 'tasa': 100},
                    {'mes_cuota': 7, 'mes_sueldo': 7, 'sueldo_incluido': 'si'}, 2027)['base_esperada']
        self.assertFalse(r['cuota_calculable'])
        self.assertEqual(r['primer_deficit'], '2027-01')
        self.assertEqual(r['aumento_maximo'], 0)

    def test_gastos_extraordinarios_ajuste_y_aguinaldo(self):
        items = [{'activo': 1, 'tipo': 'egreso', 'monto': 100}]
        r = simular(items, {'base_cuotas': 0, 'tasa': 0},
                    {'sueldo': 100, 'objetivo': 200, 'mes_sueldo': 6,
                     'aguinaldo': 'si', 'extra_6': 1000,
                     'aumento_gastos': 20, 'mes_gastos': 6}, 2027)['base_esperada']
        self.assertEqual(r['filas'][4]['gastos'], 200)
        self.assertEqual(r['filas'][5]['gastos'], 1270)
        self.assertEqual(r['filas'][5]['adicional'], 150)
        self.assertEqual(r['filas'][6]['adicional'], 100)

    def test_recuperacion_separada_acotada_al_stock_y_al_anio(self):
        r = simular([], {'base_cuotas': 100, 'tasa': 100, 'vencido': 1000},
                    {'sueldo': 0, 'objetivo': 0, 'recuperacion': 60,
                     'mes_recuperacion': 11, 'plazo_recuperacion': 6}, 2027)
        self.assertEqual(sum(f['recuperado'] for f in r['base_esperada']['filas']), 200)
        self.assertEqual(sum(f['recuperado'] for f in r['escenarios'][0]['filas']), 100)
        self.assertEqual(r['base_esperada']['filas'][10]['recuperado'], 100)
        self.assertEqual(r['base_esperada']['deuda_a_recuperar_despues'], 400)
        self.assertEqual(r['base_esperada']['filas'][10]['cobro'], 100)

    def test_padron_respeta_fin_beca_tipo_miembro_y_exclusion_de_no_aportantes(self):
        conn = MagicMock()
        conn.execute.return_value.fetchall.return_value = [
            {'tipo_miembro': 'Miembro Adherente'},
            {'tipo_miembro': '', 'beca_activa': 1, 'beca_porcentaje': 100,
             'beca_desde': '2027-01', 'beca_hasta': '2027-06'}]
        padron = obtener_padron(conn, 2027)
        self.assertEqual(padron[0]['socios'], 2)
        self.assertEqual(padron[0]['descuento'], 75)
        self.assertEqual(padron[6]['descuento'], 25)
        sql = conn.execute.call_args.args[0]
        self.assertIn("estado = 'Activo'", sql)
        self.assertIn('cobra_cuota', sql)

    def test_ajustes_vacios_siguen_campos_generales_y_no_se_congelan_al_recalcular(self):
        r = simular([], self.indicadores, {'cuotas': 100, 'emision_2': 200}, 2027)
        self.assertNotIn('emision_1', r['formulario'])
        args = {**r['formulario'], 'cuotas': '300'}
        segundo = simular([], self.indicadores, args, 2027)
        self.assertEqual(segundo['configuracion'][0]['emision'], 300)
        self.assertEqual(segundo['configuracion'][1]['emision'], 200)

    def test_validacion_rechaza_fechas_porcentajes_y_montos_invalidos(self):
        for args in [{'anio': 2027.5}, {'mes_cuota': 13}, {'cobranza_2': 101},
                     {'socios_1': 1.5}, {'extra_1': -1}, {'saldo': 'NaN'},
                     {'fuente': 'inventada'}, {'cuotas': 'texto'}]:
            with self.subTest(args=args), self.assertRaises(ValueError):
                validar_parametros(args)
        validar_parametros({'saldo': '-1', 'socios_1': '0', 'extra_1': ''})

    def test_guardar_conserva_resultados_servidor_y_no_modifica_cuotas(self):
        conn = MagicMock()
        conn.execute.return_value.fetchone.return_value = {'id': 9}
        indicadores = {**self.indicadores, 'padron': [{'socios': 0, 'descuento': 0}] * 12}
        with app.app.test_request_context(method='POST', data={
                'nombre': 'Plan 2027', 'objetivo': '800000', 'resultado_falso': '999999'}), \
             patch.object(app, 'permiso_requerido', return_value=None), \
             patch.object(app, 'obtener_presupuesto_mensual', return_value={'items': self.items}), \
             patch.object(app, 'obtener_indicadores', return_value=indicadores), \
             patch.object(app, 'obtener_padron', return_value=indicadores['padron']), \
             patch.object(app, 'get_connection', return_value=conn), \
             patch.object(app, 'registrar_auditoria'):
            app.session['username'] = 'tesorero'
            response = app.guardar_simulacion_presupuesto()
        self.assertIn('plan_id=9', response.location)
        sql, params = conn.execute.call_args.args
        self.assertIn('INSERT INTO presupuesto_simulaciones', sql)
        datos = json.loads(params[2])
        self.assertNotIn('resultado_falso', datos['formulario'])
        self.assertEqual(len(datos['escenarios']), 3)
        self.assertEqual(datos['conceptos'][0]['monto'], 550000)
        self.assertEqual(params[3], 'tesorero')
        conn.commit.assert_called_once()
        conn.close.assert_called_once()

    def test_guardar_sin_permiso_o_con_valor_invalido_no_accede_base(self):
        with app.app.test_request_context(method='POST'), \
             patch.object(app, 'permiso_requerido', return_value=('Prohibido', 403)), \
             patch.object(app, 'get_connection') as connection:
            self.assertEqual(app.guardar_simulacion_presupuesto()[1], 403)
            connection.assert_not_called()
        with app.app.test_request_context(method='POST', data={'nombre': 'Plan', 'mes_cuota': '20'}), \
             patch.object(app, 'permiso_requerido', return_value=None), \
             patch.object(app, 'get_connection') as connection:
            self.assertEqual(app.guardar_simulacion_presupuesto().status_code, 302)
            connection.assert_not_called()

    def test_vistas_separan_la_caja_de_la_planificacion_y_conservan_los_parametros(self):
        presupuesto = {'items': [], 'proyeccion': [], 'resumen': {
            'ingresos_presupuestados': 0, 'egresos_presupuestados': 0,
            'resultado_presupuestado': 0, 'resultado_real': 100,
            'ingresos_reales': 300, 'egresos_reales': 200}}
        for vista in ['resumen', 'gastos', 'planificar', 'guardados']:
            conn = MagicMock()
            conn.execute.return_value.fetchall.return_value = []
            conn.execute.return_value.fetchone.return_value = {'saldo': 100}
            with self.subTest(vista=vista), \
                 app.app.test_request_context(f'/presupuesto?vista={vista}&objetivo=700000'), \
                 patch.object(app, 'permiso_requerido', return_value=None), \
                 patch.object(app, 'obtener_presupuesto_mensual', return_value=presupuesto), \
                 patch.object(app, 'get_connection', return_value=conn):
                rendered = app.ver_presupuesto()
            self.assertEqual('id="datos-plan"' in rendered, vista == 'planificar')
            self.assertEqual('id="mes-resumen"' in rendered, vista == 'resumen')
            self.assertEqual('id="versiones-plan"' in rendered, vista == 'guardados')
            self.assertIn('objetivo=700000', rendered)
            if vista == 'planificar':
                self.assertIn('Primero hay que completar los gastos', rendered)
                self.assertIn('Ver cómo quedaría el año', rendered)
                self.assertIn('value="700000', rendered)


if __name__ == '__main__':
    unittest.main()
