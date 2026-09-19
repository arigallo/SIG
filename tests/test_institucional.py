import os
import unittest
from unittest.mock import patch

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")

import app


class InstitucionalTests(unittest.TestCase):
    def test_public_home_does_not_query_private_data(self):
        with patch.object(app, "get_connection", side_effect=AssertionError("Unexpected database access")):
            for host, path in (("rudamachorugby.com", "/"), ("www.rudamachorugby.com", "/"), ("localhost", "/club")):
                with self.subTest(host=host), app.app.test_client() as client:
                    response = client.get(path, base_url=f"https://{host}")
                    self.assertEqual(response.status_code, 200)
                    html = response.get_data(as_text=True)
                    self.assertIn("Nuestra lucha<br>es <em>jugando.</em>", html)
                    self.assertIn('href="/postulate"', html)
                    self.assertIn('https://sig.rudamachorugby.com/login', html)

    def test_sig_and_private_routes_still_require_login(self):
        with app.app.test_client() as client:
            for host, path in (("sig.rudamachorugby.com", "/"), ("localhost", "/"), ("rudamachorugby.com", "/jugadores")):
                with self.subTest(host=host, path=path):
                    response = client.get(path, base_url=f"https://{host}")
                    self.assertEqual(response.status_code, 302)
                    self.assertTrue(response.location.endswith("/login"))

    def test_home_head(self):
        with app.app.test_client() as client:
            response = client.head("/", base_url="https://rudamachorugby.com")
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.data, b"")


if __name__ == "__main__":
    unittest.main()
