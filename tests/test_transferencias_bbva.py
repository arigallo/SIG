import unittest
from email.message import EmailMessage
from decimal import Decimal
from services import transferencias_bbva as bbva
from unittest.mock import MagicMock


def aviso(remitente='avisos@bbva.com.ar', monto='40.000,00'):
    m=EmailMessage()
    m['From']=remitente
    m['Subject']='Mensajes y Avisos por e-mail - AVISO TRANSFERENCIA INMEDIATA ACREDITADA'
    m['Message-ID']='<ejemplo-bbva@test.invalid>'
    m.set_content(f'<p>Hola TITULAR DESTINO, Fecha 02/11/2026 N&uacute;mero cuenta CA-ARS ***123/4 CBU origen 0000000000000000000001 CUIT origen 20123456786 Importe ARS {monto}</p>', subtype='html')
    return m.as_bytes()


class TransferenciasBBVATests(unittest.TestCase):
    def test_parser_html_y_entidades(self):
        d=bbva.parsear_aviso(aviso())
        self.assertEqual(d['importe'],Decimal('40000.00'))
        self.assertEqual(str(d['fecha']),'2026-11-02')
        self.assertEqual(d['cuit'],'20123456786')
        self.assertNotIn('nombre',d)
        self.assertEqual(d['mensaje_clave'],bbva.parsear_aviso(aviso())['mensaje_clave'])

    def test_ignora_otro_remitente_y_rechaza_importe_cero(self):
        self.assertIsNone(bbva.parsear_aviso(aviso('otro@example.com')))
        with self.assertRaises(ValueError): bbva.parsear_aviso(aviso(monto='0,00'))

    def test_coincidencia_unica_y_ambigua(self):
        datos=bbva.parsear_aviso(aviso())
        jugadores=[{'id':1,'dni':'12.345.678'}]
        self.assertEqual(bbva.sugerir_jugador(datos,jugadores,[])[0],1)
        self.assertIsNone(bbva.sugerir_jugador(datos,jugadores*2,[])[0])
        self.assertEqual(bbva.sugerir_jugador(datos,[],[{'cuit':datos['cuit'],'cbu':datos['cbu'],'jugador_id':9}])[0],9)

    def test_lectura_no_marca_leidos_y_duplicados_no_suman(self):
        conn=MagicMock()
        conn.execute.return_value.fetchone.side_effect=[{'id':1},None]
        imap=MagicMock()
        client=imap.IMAP4_SSL.return_value
        client.select.return_value=('OK',[])
        client.search.return_value=('OK',[b'1 2'])
        client.fetch.return_value=('OK',[(b'data',aviso())])
        config={'user':'tesoreria@rudamachorugby.com','use_ssl':True,'host':'imap.gmail.com','port':993,'password':'fake','folder':'INBOX','search_days':30}
        r=bbva.sincronizar(conn,[config],imap)
        self.assertEqual(r['nuevas'],1)
        client.select.assert_called_once_with('INBOX',readonly=True)
        self.assertEqual(client.fetch.call_args.args[1],'(BODY.PEEK[])')
        self.assertIn('ON CONFLICT',conn.execute.call_args.args[0])
        client.logout.assert_called_once()
