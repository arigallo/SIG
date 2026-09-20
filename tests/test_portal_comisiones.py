import os
import unittest
from unittest.mock import patch
os.environ['INIT_DB'] = 'false'
os.environ.setdefault('SECRET_KEY', 'test-only-secret-key')
import app

class PortalComisionesTests(unittest.TestCase):
    def test_public_page_without_private_data(self):
        with patch.object(app, 'get_connection', side_effect=AssertionError('Unexpected database access')):
            with app.app.test_client() as client:
                response = client.get('/portal/comisiones')
                self.assertEqual(response.status_code, 200)
                html = response.get_data(as_text=True)
                for value in ['Comisión Directiva', 'Sub-comisión de Disciplina', 'Sub-comisión de Madrinas', 'Sub-comisión de Comunicación', 'Fabian Luciano Fretti', 'Eduardo Del Valle', 'Mathias Quiroz Mendoza', 'Federico Ferraro', 'pendientes de validación']:
                    self.assertIn(value, html)
                self.assertIn('href="/portal"', html)
                self.assertEqual(client.head('/portal/comisiones').status_code, 200)

if __name__ == '__main__':
    unittest.main()
