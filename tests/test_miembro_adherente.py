import os
import unittest
from unittest.mock import patch
os.environ['INIT_DB']='false'
os.environ.setdefault('SECRET_KEY','test-only-secret-key')
import app
from services import finanzas


class MiembroAdherenteTests(unittest.TestCase):
    def test_nombre_anterior_se_normaliza(self):
        self.assertEqual(app.normalizar_tipo_miembro('Colaborador'),'Miembro Adherente')
        self.assertIn('Miembro Adherente',app.TIPOS_MIEMBRO)
        self.assertNotIn('Colaborador',app.TIPOS_MIEMBRO)

    def test_mitad_solo_para_adherentes(self):
        for tipo,importe in [('Miembro Adherente',5000),('Colaborador',5000),('Jugador',10000),('Socio activo',10000)]:
            with self.subTest(tipo=tipo):
                self.assertEqual(finanzas.importe_social_segun_miembro({'tipo_miembro':tipo},10000),importe)

    def test_beca_sobre_mitad_y_plan_completo_sin_reducir_de_nuevo(self):
        jugador={'id':1,'tipo_miembro':'Miembro Adherente'}
        base=finanzas.importe_social_segun_miembro(jugador,10000)
        with patch.object(finanzas,'adicional_plan_pago_para_periodo',return_value={'monto':2000,'detalle':'Plan','plan_id':1}):
            cuota=finanzas.calcular_importe_cuota_mensual(None,jugador,'2026-11',base)
            self.assertEqual(cuota['importe'],7000)
            # Recalcular con el importe original no vuelve a aplicar el 50%.
            self.assertEqual(finanzas.calcular_importe_cuota_mensual(None,jugador,'2026-11',base)['importe'],7000)
        becado={**jugador,'beca_activa':1,'beca_porcentaje':20}
        cuota=finanzas.calcular_importe_con_beca(becado,'2026-11',base)
        self.assertEqual(cuota['importe'],4000)
