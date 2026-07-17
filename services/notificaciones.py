import hashlib

from repositories.notificaciones import (
    desactivar_suscripcion_push,
    guardar_suscripcion_push,
    jugador_tiene_suscripcion_push_activa,
    obtener_comunicaciones_portal_dia,
    obtener_destinatarios_push_manual,
    resumen_suscripciones_push,
)


def hash_portal_token(token):
    token = (token or "").strip()
    return hashlib.sha256(token.encode("utf-8")).hexdigest() if token else None


def actor_push_actual(conn, portal_token=None, usuario_id=None):
    portal_token = (portal_token or "").strip()
    if portal_token:
        jugador = conn.execute("""
            SELECT id
            FROM jugadores
            WHERE portal_token = %s
              AND COALESCE(portal_activo, 0) = 1
        """, (portal_token,)).fetchone()
        if jugador:
            return {
                "tipo": "portal",
                "usuario_id": None,
                "jugador_id": jugador["id"],
                "portal_token_hash": hash_portal_token(portal_token),
            }

    if usuario_id:
        return {
            "tipo": "usuario",
            "usuario_id": usuario_id,
            "jugador_id": None,
            "portal_token_hash": None,
        }

    return None


def normalizar_url_push(valor, fallback="/"):
    valor = (valor or "").strip()
    if not valor:
        return fallback
    if valor.startswith("/") and not valor.startswith("//"):
        return valor
    if valor.startswith("https://") or valor.startswith("http://"):
        return valor
    return fallback
