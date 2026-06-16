from datetime import datetime, timedelta

from repositories.calendario import (
    crear_evento_asistencia_desde_calendario,
    crear_evento_calendario_desde_data,
    existe_evento_calendario,
)


CALENDARIO_DEPORTIVO_TIPOS = {"Entrenamiento", "Partido", "Evento", "Otro"}
CALENDARIO_ASISTENCIA_TIPOS = {"Entrenamiento", "Partido"}


def normalizar_hora_evento(valor):
    valor = (valor or "").strip()
    if not valor:
        return ""
    for formato in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(valor, formato).strftime("%H:%M")
        except ValueError:
            continue
    return ""


def normalizar_duracion_evento(valor, default=90):
    try:
        duracion = int(valor or default)
    except (TypeError, ValueError):
        duracion = default
    return min(max(duracion, 15), 24 * 60)


def calendario_evento_es_deportivo(tipo):
    return (tipo or "").strip() in CALENDARIO_DEPORTIVO_TIPOS


def calendario_evento_requiere_asistencia(tipo):
    return (tipo or "").strip() in CALENDARIO_ASISTENCIA_TIPOS


def normalizar_categoria_calendario(categoria):
    return (categoria or "").strip()


def categoria_evento_aplica(categoria_evento, categoria_jugador):
    categoria_evento = normalizar_categoria_calendario(categoria_evento).lower()
    if not categoria_evento or categoria_evento in {"todo", "todos", "todo el club", "club", "general"}:
        return True
    categoria_jugador = normalizar_categoria_calendario(categoria_jugador).lower()
    return bool(categoria_jugador and categoria_jugador in categoria_evento)


def fecha_hora_evento(evento):
    fecha = evento.get("fecha")
    hora = normalizar_hora_evento(evento.get("hora_inicio") or evento.get("hora"))
    if hora:
        return f"{fecha} {hora}"
    return fecha


def formato_fecha_hora_evento(evento):
    fecha = evento.get("fecha") or ""
    hora = normalizar_hora_evento(evento.get("hora_inicio") or evento.get("hora"))
    return f"{fecha} {hora}" if hora else fecha


def generar_fechas_recurrentes_mes(mes, dias_semana):
    try:
        inicio = datetime.strptime(mes, "%Y-%m").date()
    except (TypeError, ValueError):
        return []

    dias = set()
    for dia in dias_semana or []:
        try:
            dia_numero = int(dia)
        except (TypeError, ValueError):
            continue
        if 0 <= dia_numero <= 6:
            dias.add(dia_numero)

    if not dias:
        return []

    if inicio.month == 12:
        fin = inicio.replace(year=inicio.year + 1, month=1)
    else:
        fin = inicio.replace(month=inicio.month + 1)

    fechas = []
    actual = inicio
    while actual < fin:
        if actual.weekday() in dias:
            fechas.append(actual.strftime("%Y-%m-%d"))
        actual += timedelta(days=1)
    return fechas


def crear_eventos_calendario(conn, data, fechas_recurrentes=None, crear_recurrentes=False):
    eventos_creados = []
    eventos_omitidos = []
    fechas_a_crear = fechas_recurrentes if crear_recurrentes else [data["fecha"]]

    for fecha_evento in fechas_a_crear:
        data_evento = dict(data)
        data_evento["fecha"] = fecha_evento
        if crear_recurrentes and existe_evento_calendario(conn, data_evento):
            eventos_omitidos.append(fecha_evento)
            continue
        evento_id, asistencia_evento_id = crear_evento_calendario_desde_data(conn, data_evento)
        eventos_creados.append({
            "id": evento_id,
            "fecha": fecha_evento,
            "asistencia_evento_id": asistencia_evento_id,
        })

    return eventos_creados, eventos_omitidos
