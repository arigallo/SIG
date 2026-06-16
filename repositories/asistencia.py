def listar_eventos_asistencia(conn):
    return conn.execute("""
        SELECT
            e.*,
            ce.id AS calendario_evento_id
        FROM eventos_asistencia e
        LEFT JOIN calendario_eventos ce ON ce.asistencia_evento_id = e.id
        ORDER BY
            CASE
                WHEN e.fecha >= CURRENT_DATE::text THEN 0
                ELSE 1
            END,
            CASE
                WHEN e.fecha >= CURRENT_DATE::text THEN e.fecha
            END ASC,
            CASE
                WHEN e.fecha < CURRENT_DATE::text THEN e.fecha
            END DESC,
            e.id DESC
    """).fetchall()


def obtener_evento_asistencia(conn, evento_id):
    return conn.execute("""
        SELECT *
        FROM eventos_asistencia
        WHERE id = %s
    """, (evento_id,)).fetchone()


def obtener_calendario_evento_por_asistencia(conn, evento_id):
    return conn.execute("""
        SELECT id
        FROM calendario_eventos
        WHERE asistencia_evento_id = %s
    """, (evento_id,)).fetchone()
