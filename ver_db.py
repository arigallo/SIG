import sqlite3

conn = sqlite3.connect("club.db")
conn.row_factory = sqlite3.Row

movimientos = conn.execute("""
    SELECT id, tipo, concepto, monto, fecha, referencia
    FROM movimientos
    ORDER BY id DESC
""").fetchall()

for m in movimientos:
    print(dict(m))

conn.close()