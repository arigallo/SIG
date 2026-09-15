import unicodedata

from repositories.portal import listar_eventos_deportivos_portal
from services.calendario import categoria_evento_aplica


PORTAL_ASISTENCIA_ESTADOS = {"confirmado", "dudoso", "no_asiste"}
PORTAL_ASISTENCIA_LABELS = {
    "default": {
        "confirmado": "Voy",
        "dudoso": "Dudoso",
        "no_asiste": "No voy",
    },
    "partido": {
        "confirmado": "Voy y Juego",
        "dudoso": "Voy y no juego",
        "no_asiste": "No voy",
    },
}
PORTAL_ASISTENCIA_BADGES = {
    "confirmado": "badge-success",
    "dudoso": "badge-warning",
    "no_asiste": "badge-danger",
}

BIENESTAR_HORAS_OPCIONES = ["<5 h", "5-6 h", "6-7 h", "7-8 h", ">8 h"]
BIENESTAR_HORAS_SCORE = {"<5 h": 1, "5-6 h": 2, "6-7 h": 4, "7-8 h": 5, ">8 h": 4}
BIENESTAR_PESOS = {
    "sueno_calidad": 1.0,
    "horas_sueno": 0.75,
    "doms": 1.25,
    "fatiga": 1.5,
    "estres": 1.0,
    "animo": 0.75,
    "motivacion": 0.75,
    "recuperacion": 1.5,
}
BIENESTAR_DOLOR_ZONAS = [
    "No",
    "Cuello",
    "Hombro",
    "Brazo",
    "Zona lumbar",
    "Cadera",
    "Muslo",
    "Rodilla",
    "Pantorrilla",
    "Tobillo",
    "Pie",
    "Otro",
]


def es_evento_partido(evento):
    tipo = (evento.get("tipo") if evento else "") or ""
    tipo = unicodedata.normalize("NFKD", str(tipo).strip().lower())
    tipo = "".join(ch for ch in tipo if not unicodedata.combining(ch))
    return tipo in {"partido", "partidos"}


def asistencia_portal_labels(evento):
    return PORTAL_ASISTENCIA_LABELS["partido" if es_evento_partido(evento) else "default"]


def asistencia_portal_opciones(evento):
    labels = asistencia_portal_labels(evento)
    return [
        {"valor": estado, "label": labels[estado]}
        for estado in ("confirmado", "dudoso", "no_asiste")
    ]


def asistencia_portal_label(evento, estado):
    return asistencia_portal_labels(evento).get(estado, PORTAL_ASISTENCIA_LABELS["default"]["confirmado"])


def asistencia_portal_badge_class(estado):
    return PORTAL_ASISTENCIA_BADGES.get(estado, "badge-muted")


def horas_sueno_score(valor):
    return BIENESTAR_HORAS_SCORE.get((valor or "").strip(), 0)


def resumen_bienestar_confirmacion(confirmacion):
    if not confirmacion or confirmacion.get("sueno_calidad") is None:
        return None
    valores = {}
    for clave in ("sueno_calidad", "doms", "fatiga", "estres", "animo", "motivacion", "recuperacion"):
        try:
            valores[clave] = int(confirmacion.get(clave) or 0)
        except (TypeError, ValueError):
            valores[clave] = 0
    horas_score = horas_sueno_score(confirmacion.get("horas_sueno"))
    if horas_score:
        valores["horas_sueno"] = horas_score
    valores_validos = {clave: valor for clave, valor in valores.items() if valor}
    peso_total = sum(BIENESTAR_PESOS[clave] for clave in valores_validos)
    promedio = round(
        sum(valor * BIENESTAR_PESOS[clave] for clave, valor in valores_validos.items()) / peso_total,
        1,
    ) if peso_total else 0
    zonas = confirmacion.get("dolor_zonas_lista") or []
    zonas_alerta = [zona for zona in zonas if zona and zona not in {"No", "Otro"}]
    if confirmacion.get("dolor_otro"):
        zonas_alerta.append("Otro")
    dolor_relevante = len(zonas_alerta) >= 2 and (
        promedio < 3 or valores.get("doms", 5) <= 2 or valores.get("recuperacion", 5) <= 2
    )
    if (promedio and promedio < 2.25) or dolor_relevante:
        nivel = "danger"
        label = "Alerta roja"
    elif (promedio and promedio < 3) or zonas_alerta:
        nivel = "warning"
        label = "Alerta amarilla"
    else:
        nivel = "success"
        label = "Bienestar ok"
    return {
        "promedio": promedio,
        "nivel": nivel,
        "label": label,
        "dolores": zonas_alerta,
    }


def obtener_eventos_deportivos_portal(conn, jugador, limit=8):
    eventos = listar_eventos_deportivos_portal(conn)

    filtrados = []
    for evento in eventos:
        if categoria_evento_aplica(evento.get("categoria"), jugador.get("categoria")):
            filtrados.append(evento)
        if len(filtrados) >= limit:
            break
    return filtrados
