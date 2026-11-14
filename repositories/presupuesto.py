"""Versiones guardadas del plan financiero, separadas de la contabilidad."""
import json


def crear_tablas(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS presupuesto_simulaciones (
        id SERIAL PRIMARY KEY,
        nombre TEXT NOT NULL,
        anio INTEGER NOT NULL,
        contenido JSONB NOT NULL,
        creado_por TEXT NOT NULL,
        creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
        archivado INTEGER NOT NULL DEFAULT 0
    )""")


def listar(conn):
    return conn.execute("""SELECT id, nombre, anio, creado_por, creado_en
        FROM presupuesto_simulaciones WHERE archivado = 0
        ORDER BY creado_en DESC, id DESC LIMIT 100""").fetchall()


def obtener(conn, identificador):
    return conn.execute("""SELECT * FROM presupuesto_simulaciones
        WHERE id = %s AND archivado = 0""", (identificador,)).fetchone()


def guardar(conn, nombre, simulacion, autor):
    return conn.execute("""INSERT INTO presupuesto_simulaciones
        (nombre, anio, contenido, creado_por) VALUES (%s, %s, %s::jsonb, %s)
        RETURNING id""", (nombre, simulacion['anio'],
                           json.dumps(simulacion, ensure_ascii=False, allow_nan=False), autor)).fetchone()['id']
