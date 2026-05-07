import os
import sqlite3
import psycopg2
import psycopg2.extras

SQLITE_DB = os.environ.get("SQLITE_DB", "club.db")
DB_NAME = os.environ.get("DB_NAME", "sig")
DB_USER = os.environ.get("DB_USER", "sig_user")
DB_PASS = os.environ.get("DB_PASS", "")
DB_HOST = os.environ.get("DB_HOST", "localhost")
DB_PORT = os.environ.get("DB_PORT", "5432")

TABLAS = [
    "jugadores",
    "usuarios",
    "cuotas",
    "fichas_medicas",
    "lesiones",
    "movimientos",
    "cierres_mensuales",
    "eventos_asistencia",
    "asistencias",
]


def pg_conn():
    return psycopg2.connect(
        dbname=DB_NAME,
        user=DB_USER,
        password=DB_PASS,
        host=DB_HOST,
        port=DB_PORT,
        cursor_factory=psycopg2.extras.DictCursor,
    )


def columnas_sqlite(conn, tabla):
    return [r[1] for r in conn.execute(f"PRAGMA table_info({tabla})").fetchall()]


def columnas_postgres(cur, tabla):
    cur.execute(
        """
        SELECT column_name
        FROM information_schema.columns
        WHERE table_schema = 'public'
          AND table_name = %s
        ORDER BY ordinal_position
        """,
        (tabla,),
    )
    return [r["column_name"] for r in cur.fetchall()]


def reset_sequence(cur, tabla):
    cur.execute(
        """
        SELECT pg_get_serial_sequence(%s, 'id') AS seq
        """,
        (tabla,),
    )
    row = cur.fetchone()
    seq = row["seq"] if row else None
    if seq:
        cur.execute(f"SELECT COALESCE(MAX(id), 1) FROM {tabla}")
        max_id = cur.fetchone()[0]
        cur.execute("SELECT setval(%s, %s, true)", (seq, max_id))


def main():
    sqlite = sqlite3.connect(SQLITE_DB)
    sqlite.row_factory = sqlite3.Row
    pg = pg_conn()
    cur = pg.cursor()

    for tabla in TABLAS:
        try:
            cols_sqlite = columnas_sqlite(sqlite, tabla)
            cols_pg = columnas_postgres(cur, tabla)
            cols = [c for c in cols_sqlite if c in cols_pg]
            if not cols:
                print(f"{tabla}: sin columnas compatibles, omitida")
                continue

            rows = sqlite.execute(f"SELECT {', '.join(cols)} FROM {tabla}").fetchall()
            if not rows:
                print(f"{tabla}: 0 registros")
                continue

            placeholders = ", ".join(["%s"] * len(cols))
            columnas = ", ".join(cols)
            conflict = "ON CONFLICT (id) DO NOTHING" if "id" in cols else ""
            sql = f"INSERT INTO {tabla} ({columnas}) VALUES ({placeholders}) {conflict}"

            for row in rows:
                cur.execute(sql, tuple(row[c] for c in cols))

            reset_sequence(cur, tabla)
            pg.commit()
            print(f"{tabla}: migrados {len(rows)} registros")
        except Exception as e:
            pg.rollback()
            print(f"ERROR migrando {tabla}: {e}")

    cur.close()
    pg.close()
    sqlite.close()


if __name__ == "__main__":
    main()
