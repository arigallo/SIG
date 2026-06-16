def listar_eventos_deportivos_portal(conn):
    return conn.execute("""
        SELECT *
        FROM calendario_eventos
        WHERE COALESCE(publicar_portal, 0) = 1
          AND fecha >= CURRENT_DATE::text
        ORDER BY fecha ASC, COALESCE(hora_inicio, '') ASC, id ASC
        LIMIT 80
    """).fetchall()
