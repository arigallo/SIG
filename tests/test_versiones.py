import os
import unittest

os.environ["INIT_DB"] = "false"
os.environ.setdefault("SECRET_KEY", "test-only-secret-key")

import app


class _Result:
    def __init__(self, rows):
        self.rows = rows

    def fetchall(self):
        return self.rows


class _Connection:
    def __init__(self, rows):
        self.rows = rows

    def execute(self, sql, params=None):
        if "FROM versiones_publicadas" in sql:
            return _Result(self.rows)
        raise AssertionError(sql)


class VersionesTests(unittest.TestCase):
    def test_next_release_is_2_1_0(self):
        self.assertEqual(app.RELEASE_VERSION, "2.1.0")

    def test_counter_starts_at_zero_before_publication(self):
        resumen = app.resumen_versiones_publicadas(_Connection([]))
        self.assertEqual(resumen["total"], 0)

    def test_first_publication_counts_as_one(self):
        resumen = app.resumen_versiones_publicadas(_Connection([{"numero": 1}]))
        self.assertEqual(resumen["total"], 1)


if __name__ == "__main__":
    unittest.main()
