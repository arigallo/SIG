import json


def truncar_valor(value, max_length=300):
    text = "" if value is None else str(value)
    return text if len(text) <= max_length else text[:max_length] + "..."


def guardar_suscripcion_push(conn, subscription, actor, user_agent=None):
    endpoint = (subscription or {}).get("endpoint")
    if not endpoint:
        raise ValueError("Suscripcion push sin endpoint.")
    conn.execute("""
        INSERT INTO pwa_push_subscriptions (
            endpoint, actor_tipo, usuario_id, jugador_id, portal_token_hash,
            subscription_json, user_agent, enabled, actualizado_en
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, 1, CURRENT_TIMESTAMP)
        ON CONFLICT(endpoint) DO UPDATE SET
            actor_tipo = EXCLUDED.actor_tipo,
            usuario_id = EXCLUDED.usuario_id,
            jugador_id = EXCLUDED.jugador_id,
            portal_token_hash = EXCLUDED.portal_token_hash,
            subscription_json = EXCLUDED.subscription_json,
            user_agent = EXCLUDED.user_agent,
            enabled = 1,
            actualizado_en = CURRENT_TIMESTAMP
    """, (
        endpoint,
        actor["tipo"],
        actor.get("usuario_id"),
        actor.get("jugador_id"),
        actor.get("portal_token_hash"),
        json.dumps(subscription),
        truncar_valor(user_agent or "", 500),
    ))


def desactivar_suscripcion_push(conn, endpoint):
    if not endpoint:
        return
    conn.execute("""
        UPDATE pwa_push_subscriptions
        SET enabled = 0, actualizado_en = CURRENT_TIMESTAMP
        WHERE endpoint = %s
    """, (endpoint,))


def jugador_tiene_suscripcion_push_activa(conn, jugador_id):
    if not jugador_id:
        return False
    row = conn.execute("""
        SELECT 1 AS activa
        FROM pwa_push_subscriptions
        WHERE actor_tipo = 'portal'
          AND jugador_id = %s
          AND enabled = 1
        LIMIT 1
    """, (jugador_id,)).fetchone()
    return bool(row)


def obtener_destinatarios_push_manual(conn, destino, categoria=None, jugador_id=None):
    destino = (destino or "").strip()
    params = []
    filtros = ["s.enabled = 1"]
    joins = []

    if destino == "admins":
        joins.append("JOIN usuarios u ON u.id = s.usuario_id")
        filtros.append("s.actor_tipo = 'usuario'")
        filtros.append("u.rol = 'admin'")
    elif destino == "usuarios":
        filtros.append("s.actor_tipo = 'usuario'")
    elif destino == "portal_todos":
        filtros.append("s.actor_tipo = 'portal'")
    elif destino == "categoria":
        joins.append("JOIN jugadores j ON j.id = s.jugador_id")
        filtros.append("s.actor_tipo = 'portal'")
        if categoria == "Sin categoria":
            filtros.append("COALESCE(NULLIF(j.categoria, ''), 'Sin categoria') = 'Sin categoria'")
        else:
            filtros.append("COALESCE(j.categoria, '') = %s")
            params.append(categoria or "")
    elif destino == "jugador":
        filtros.append("s.actor_tipo = 'portal'")
        filtros.append("s.jugador_id = %s")
        params.append(jugador_id)
    else:
        return []

    join_sql = "\n".join(joins)
    where_sql = " AND ".join(filtros)
    return conn.execute(f"""
        SELECT DISTINCT
            s.endpoint,
            s.subscription_json,
            s.actor_tipo,
            s.usuario_id,
            s.jugador_id,
            s.actualizado_en
        FROM pwa_push_subscriptions s
        {join_sql}
        WHERE {where_sql}
        ORDER BY s.actualizado_en DESC
    """, params).fetchall()


def resumen_suscripciones_push(conn):
    return conn.execute("""
        SELECT
            COUNT(*) FILTER (WHERE enabled = 1) AS activas,
            COUNT(*) FILTER (WHERE enabled = 1 AND actor_tipo = 'portal') AS portal,
            COUNT(*) FILTER (WHERE enabled = 1 AND actor_tipo = 'usuario') AS usuarios,
            COUNT(*) FILTER (WHERE enabled = 1 AND actor_tipo = 'usuario' AND usuario_id IN (
                SELECT id FROM usuarios WHERE rol = 'admin'
            )) AS admins
        FROM pwa_push_subscriptions
    """).fetchone()


def obtener_comunicaciones_portal_dia(conn, jugador, hoy, limite=5):
    categoria = jugador.get("categoria") or ""
    return conn.execute("""
        SELECT id, titulo, mensaje, destino, categoria, jugador_id, url, creado_en, visible_hasta
        FROM pwa_push_envios
        WHERE COALESCE(mostrar_portal, 0) = 1
          AND (
              visible_hasta IS NULL
              OR visible_hasta = ''
              OR visible_hasta >= %s
          )
          AND (
              destino = 'portal_todos'
              OR (destino = 'categoria' AND COALESCE(categoria, '') = %s)
              OR (destino = 'jugador' AND jugador_id = %s)
          )
        ORDER BY creado_en DESC, id DESC
        LIMIT %s
    """, (hoy, categoria, jugador["id"], limite)).fetchall()
