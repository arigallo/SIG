import os
import unittest
from unittest.mock import MagicMock, patch
os.environ['INIT_DB']='false'
os.environ.setdefault('SECRET_KEY','test-only-secret-key')
import app
from flask import render_template, session


class ComprobantesCompartidosTests(unittest.TestCase):
    def test_filtros_y_contadores_incluyen_compartidos(self):
        for estado, esperados in [('pendiente',1),('todos',2),('aceptado',1)]:
            conn=MagicMock()
            items=[{'id':4,'estado':'pendiente','comprobante_estado_resuelto':'pendiente'}, {'id':5,'estado':'pagado','comprobante_estado_resuelto':'aceptado'}]
            conn.execute.side_effect=[MagicMock(fetchall=lambda:[]),MagicMock(fetchone=lambda:dict(total=1,pendientes=1,aceptados=0,rechazados=0)),MagicMock(fetchall=lambda:items)]
            with self.subTest(estado=estado), app.app.test_request_context('/?estado='+estado), patch.object(app,'permiso_requerido',return_value=None),patch.object(app,'get_connection',return_value=conn),patch.object(app,'render_template') as render:
                app.listar_comprobantes_revision()
                datos=render.call_args.kwargs
                self.assertEqual(len(datos['comprobantes']),esperados)
                self.assertEqual(datos['resumen']['total'],3)
                self.assertEqual(datos['resumen']['pendientes'],2)
                self.assertEqual(datos['resumen']['aceptados'],1)

    def test_compartido_tiene_archivo_y_accion_propios(self):
        item=dict(id=4,jugador_id=None,nombre='Invitado',apellido='Demo',periodo='Tercer tiempo',importe=100,comprobante_estado_resuelto='pendiente',tipo_comprobante='compartido',pagado=False)
        with app.app.test_request_context('/comprobantes?estado=pendiente'):
            session.update(user_id=1,rol='admin')
            html=render_template('comprobantes_revision.html',comprobantes=[item],estado='pendiente',resumen=dict(total=1,pendientes=1))
        self.assertIn('/gastos-compartidos/items/4/estado',html)
        self.assertIn('value="aprobar"',html)
        self.assertIn('name="fecha_pago"',html)
        self.assertNotIn('/cuotas/4/comprobante',html)
