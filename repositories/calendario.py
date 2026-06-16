def crear_evento_asistencia_desde_calendario(conn, data):
    descripcion = data.get("descripcion") or data.get("titulo") or ""
    if data.get("ubicacion"):
        descripcion = f"{descripcion}\nLugar: {data['ubicacion']}".strip()
    if data.get("categoria"):
        descripcion = f"{descripcion}\nCategoria: {data['categoria']}".strip()

    fila = conn.execute("""
        INSERT INTO eventos_asistencia (fecha, tipo, descripcion)
        VALUES (%s, %s, %s)
        RETURNING id
    """, (data["fecha"], data["tipo"], descripcion)).fetchone()
    return fila["id"]


def existe_evento_calendario(conn, data):
    return conn.execute("""
        SELECT id
        FROM calendario_eventos
        WHERE fecha = %s
          AND tipo = %s
          AND titulo = %s
          AND COALESCE(hora_inicio, '') = COALESCE(%s, '')
          AND COALESCE(categoria, '') = COALESCE(%s, '')
        LIMIT 1
    """, (
        data["fecha"],
        data["tipo"],
        data["titulo"],
        data["hora_inicio"] or None,
        data["categoria"] or None,
    )).fetchone()


def crear_evento_calendario_desde_data(conn, data):
    asistencia_evento_id = None
    if data["crear_asistencia"]:
        asistencia_evento_id = crear_evento_asistencia_desde_calendario(conn, data)

    evento_id = conn.execute("""
        INSERT INTO calendario_eventos (
            fecha, tipo, titulo, descripcion, ubicacion, categoria,
            hora_inicio, duracion_minutos, publicar_portal, asistencia_evento_id,
            convocatoria_texto, convocatoria_cierre, minuta_post_evento
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        RETURNING id
    """, (
        data["fecha"],
        data["tipo"],
        data["titulo"],
        data["descripcion"],
        data["ubicacion"],
        data["categoria"],
        data["hora_inicio"] or None,
        data["duracion_minutos"],
        data["publicar_portal"],
        asistencia_evento_id,
        data["convocatoria_texto"] or None,
        data["convocatoria_cierre"] or None,
        data["minuta_post_evento"] or None,
    )).fetchone()["id"]
    return evento_id, asistencia_evento_id
