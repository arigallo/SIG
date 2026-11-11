"""Estados independientes para el seguimiento rápido de ahijados."""

CONTACTOS = {"pendiente": "Sin contactar", "contactado": "Ya le hablé"}
CONFIRMACIONES = {
    "pendiente": "Sin consultar",
    "esperando": "Esperando respuesta",
    "si": "Sí, viene",
    "no": "No viene",
}


def estados_rapidos(aspirante):
    etapa = aspirante.get("etapa_seguimiento") or "pendiente_contacto"
    contacto = aspirante.get("contacto_estado")
    if contacto not in CONTACTOS:
        contacto = "contactado" if etapa in {
            "contactado", "esperando_respuesta", "confirmo_asistencia", "en_seguimiento", "listo_ingresar"
        } else "pendiente"
    viene = aspirante.get("viene_estado")
    if viene not in CONFIRMACIONES:
        viene = {"confirmo_asistencia": "si", "esperando_respuesta": "esperando"}.get(etapa, "pendiente")
    return {"contacto_estado": contacto, "viene_estado": viene,
            "madrina_jugador_id": aspirante.get("madrina_jugador_id")}


def validar_cambios(payload):
    if not isinstance(payload, dict):
        raise ValueError("Datos inválidos.")
    cambios = payload.get("cambios", {})
    anteriores = payload.get("anteriores", {})
    nota = payload.get("nota", "")
    if not isinstance(cambios, dict) or not isinstance(anteriores, dict) or not isinstance(nota, str):
        raise ValueError("Datos inválidos.")
    nota = nota.strip()
    if len(nota) > 3000:
        raise ValueError("La nota admite hasta 3000 caracteres.")
    for campo, valor in cambios.items():
        if campo not in {"contacto_estado", "viene_estado", "madrina_jugador_id"} or campo not in anteriores:
            raise ValueError("Revisá los campos enviados.")
        if campo == "contacto_estado" and (not isinstance(valor, str) or valor not in CONTACTOS):
            raise ValueError("Contacto inválido.")
        if campo == "viene_estado" and (not isinstance(valor, str) or valor not in CONFIRMACIONES):
            raise ValueError("Confirmación inválida.")
        if campo == "madrina_jugador_id" and valor is not None and (type(valor) is not int or valor <= 0):
            raise ValueError("Madrina inválida.")
    if not cambios and not nota:
        raise ValueError("No hay cambios para guardar.")
    return cambios, anteriores, nota
