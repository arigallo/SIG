import os
import unittest
from datetime import date
from unittest.mock import MagicMock

os.environ['INIT_DB'] = 'false'
os.environ.setdefault('SECRET_KEY', 'test-only-secret-key')
from services.workspace import preparar_gestion, agrupar_pendientes_inicio, obtener_inicio, ordenar_notas_financieras, paginacion


class WorkspaceTests(unittest.TestCase):
    def test_seguimiento_sin_deuda_sigue_visible_y_no_inventa_conceptos(self):
        item = preparar_gestion(self.item(origen='seguimiento', importe=0,
            nota='Confirmar acuerdo', seguimiento_fecha='2026-11-10'), '2026-11-10')
        self.assertEqual(item['carril'], 'seguimientos')
        result = agrupar_pendientes_inicio([item])
        self.assertEqual(result[0]['conceptos'], 0)
        self.assertEqual(result[0]['importe'], 0)

    def test_paginacion_acota_paginas_inexistentes_y_lista_vacia(self):
        self.assertEqual(paginacion(61, '999')['pagina'], 3)
        self.assertEqual(paginacion(61, '-2')['offset'], 0)
        self.assertEqual(paginacion(61, 'texto')['pagina'], 1)
        self.assertEqual(paginacion(0, 9)['paginas'], 1)

    def item(self, **updates):
        return {'id': 7, 'jugador_id': 2, 'origen': 'cuota', 'importe': 20000,
                'fecha_vencimiento': '2026-10-10', **updates}

    def test_acuerdo_futuro_impide_reclamo_prematuro(self):
        item = preparar_gestion(self.item(nota='Paga cuando cobre', seguimiento_fecha='2026-11-20'), '2026-11-10')
        self.assertEqual(item['situacion'], 'Acuerdo vigente')
        self.assertEqual(item['proxima_accion'], 'Esperar hasta 2026-11-20')
        self.assertEqual(agrupar_pendientes_inicio([item]), [])

    def test_comprobante_pendiente_tiene_prioridad_sobre_acuerdo(self):
        item = preparar_gestion(self.item(nota='Paga luego', seguimiento_fecha='2026-11-20',
                                         comprobante_estado='pendiente', archivo='local-file'), '2026-11-10')
        self.assertEqual(item['carril'], 'revision')
        self.assertEqual(item['proxima_accion'], 'Revisar comprobante')

    def test_comprobante_legacy_sin_estado_sigue_en_revision(self):
        item = preparar_gestion(self.item(archivo='local-pdf', comprobante_estado='sin_comprobante'), '2026-11-10')
        self.assertEqual(item['carril'], 'revision')

    def test_no_se_promete_revision_sin_archivo_y_seguimiento_incluye_hoy(self):
        item = preparar_gestion(self.item(nota='Confirmar', seguimiento_fecha='2026-11-10',
                                         comprobante_estado='pendiente'), '2026-11-10')
        self.assertEqual(item['carril'], 'cobros')
        self.assertEqual(item['situacion'], 'Seguimiento pendiente')

    def test_acuerdo_resuelto_no_pospone_una_deuda_vencida(self):
        item = preparar_gestion(self.item(nota='Acuerdo anterior', resuelto_en='2026-11-09',
                                         seguimiento_fecha='2026-12-20'), '2026-11-10')
        self.assertEqual(item['situacion'], 'Saldo vencido')

    def test_inicio_agrupa_persona_sin_repetirla_y_conserva_importes(self):
        items = [preparar_gestion(self.item(id=offset, importe=1000 * offset), '2026-11-10') for offset in (1, 2, 3)]
        result = agrupar_pendientes_inicio(items)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]['conceptos'], 3)
        self.assertEqual(result[0]['importe'], 6000)

    def test_inicio_sin_permisos_no_consulta_otros_modulos(self):
        conn = MagicMock()
        data = obtener_inicio(conn, lambda *permissions: False, 'madrina', set(), date(2026, 11, 10))
        conn.execute.assert_not_called()
        self.assertTrue(all(not items for items in data.values()))

    def test_tareas_se_limitan_al_usuario_y_modulos_permitidos(self):
        conn = MagicMock()
        conn.execute.return_value.fetchall.return_value = []
        obtener_inicio(conn, lambda *permissions: 'tareas_ver' in permissions, 'medico', {'salud'}, date(2026, 11, 10))
        sql, params = conn.execute.call_args.args
        self.assertIn('modulo = ANY', sql)
        self.assertIn('responsable', sql)
        self.assertEqual(params, (['salud'], 'medico'))

    def test_historial_solo_finanzas_y_resueltos_al_final(self):
        notes = [dict(id=1, tipo='finanzas', resuelto_en='2026-11-01'),
                 dict(id=2, tipo='general'), dict(id=3, tipo='finanzas', seguimiento_fecha='2026-11-10')]
        self.assertEqual([note['id'] for note in ordenar_notas_financieras(notes)], [3, 1])
