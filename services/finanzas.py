from datetime import datetime
import re

from repositories.finanzas import (
    obtener_cuenta_corriente_jugador,
    obtener_cuotas_impagas_para_plan,
)


def porcentaje_beca(valor):
    try:
        porcentaje = float(valor or 0)
    except (TypeError, ValueError):
        return None
    if porcentaje < 0 or porcentaje > 100:
        return None
    return round(porcentaje, 2)


def beca_vigente(jugador, periodo):
    if not jugador or not jugador.get("beca_activa"):
        return False

    porcentaje = porcentaje_beca(jugador.get("beca_porcentaje"))
    if not porcentaje or porcentaje <= 0:
        return False

    periodo = (periodo or "").strip()
    try:
        datetime.strptime(periodo, "%Y-%m")
    except ValueError:
        return False

    desde = (jugador.get("beca_desde") or "").strip()
    hasta = (jugador.get("beca_hasta") or "").strip()
    if desde and periodo < desde:
        return False
    if hasta and periodo > hasta:
        return False
    return True


def calcular_importe_con_beca(jugador, periodo, importe_base):
    importe_original = round(float(importe_base or 0), 2)
    if not beca_vigente(jugador, periodo):
        return {
            "importe_original": importe_original,
            "importe": importe_original,
            "descuento_beca": 0,
            "beca_porcentaje": 0,
            "beca_motivo": "",
            "becada": 0,
            "beca_total": 0,
        }

    porcentaje = porcentaje_beca(jugador.get("beca_porcentaje")) or 0
    descuento = round(importe_original * porcentaje / 100, 2)
    importe = max(0, round(importe_original - descuento, 2))
    return {
        "importe_original": importe_original,
        "importe": importe,
        "descuento_beca": descuento,
        "beca_porcentaje": porcentaje,
        "beca_motivo": jugador.get("beca_motivo") or "",
        "becada": 1,
        "beca_total": 1 if importe <= 0 else 0,
    }


def indice_periodo(periodo):
    try:
        fecha = datetime.strptime(str(periodo or "")[:7], "%Y-%m")
    except ValueError:
        return None
    return fecha.year * 12 + fecha.month


def periodo_inicio_plan(plan):
    fecha = plan.get("fecha_inicio") or plan.get("creado_en") or ""
    if isinstance(fecha, datetime):
        return fecha.strftime("%Y-%m")
    texto = str(fecha or "").strip()
    return texto[:7] if re.match(r"^[0-9]{4}-[0-9]{2}", texto) else None


def periodo_minimo(*periodos):
    validos = [periodo for periodo in periodos if indice_periodo(periodo) is not None]
    if not validos:
        return None
    return min(validos, key=lambda periodo: indice_periodo(periodo))


def adicional_plan_pago_para_periodo(conn, jugador_id, periodo):
    periodo_idx = indice_periodo(periodo)
    if periodo_idx is None:
        return {"monto": 0, "detalle": ""}

    planes = conn.execute("""
        SELECT id, fecha_inicio, monto_total, cantidad_cuotas, monto_cuota,
               descripcion, creado_en
        FROM planes_pago
        WHERE jugador_id = %s
          AND estado = 'Activo'
          AND COALESCE(monto_total, 0) > 0
          AND COALESCE(cantidad_cuotas, 0) > 0
        ORDER BY fecha_inicio ASC, id ASC
    """, (jugador_id,)).fetchall()

    monto_total_periodo = 0
    detalles = []
    for plan in planes:
        inicio = periodo_inicio_plan(plan)
        inicio_idx = indice_periodo(inicio)
        cantidad = int(plan["cantidad_cuotas"] or 0)
        if inicio_idx is None or cantidad <= 0:
            continue

        numero_cuota = periodo_idx - inicio_idx + 1
        if numero_cuota < 1 or numero_cuota > cantidad:
            continue

        monto_plan = round(float(plan["monto_cuota"] or 0), 2)
        if numero_cuota == cantidad:
            monto_total = round(float(plan["monto_total"] or 0), 2)
            monto_plan = round(monto_total - (monto_plan * (cantidad - 1)), 2)
        if monto_plan <= 0:
            continue

        monto_total_periodo = round(monto_total_periodo + monto_plan, 2)
        etiqueta = f"Plan #{plan['id']} cuota {numero_cuota}/{cantidad}"
        if plan.get("descripcion"):
            etiqueta = f"{etiqueta} - {plan['descripcion']}"
        detalles.append(etiqueta)

    return {
        "monto": monto_total_periodo,
        "detalle": "; ".join(detalles),
    }


def calcular_importe_cuota_mensual(conn, jugador, periodo, importe_base):
    cuota = calcular_importe_con_beca(jugador, periodo, importe_base)
    adicional_plan = adicional_plan_pago_para_periodo(conn, jugador["id"], periodo)
    monto_plan = round(float(adicional_plan["monto"] or 0), 2)
    cuota["plan_pago_monto"] = monto_plan
    cuota["plan_pago_detalle"] = adicional_plan["detalle"]
    if monto_plan:
        cuota["importe_original"] = round(cuota["importe_original"] + monto_plan, 2)
        cuota["importe"] = round(cuota["importe"] + monto_plan, 2)
        cuota["beca_total"] = 1 if cuota["importe"] <= 0 else 0
    return cuota


def recalcular_cuotas_planes_pago(conn, jugador_id, periodo_desde=None, hoy=None):
    jugador = conn.execute("SELECT * FROM jugadores WHERE id = %s", (jugador_id,)).fetchone()
    if jugador is None:
        return {"revisadas": 0, "actualizadas": 0}

    condiciones = ["jugador_id = %s", "pagado = 0", "COALESCE(anulada, 0) = 0"]
    parametros = [jugador_id]
    if periodo_desde:
        condiciones.append("periodo >= %s")
        parametros.append(periodo_desde)

    cuotas = conn.execute(f"""
        SELECT *
        FROM cuotas
        WHERE {" AND ".join(condiciones)}
        ORDER BY periodo ASC, id ASC
    """, parametros).fetchall()

    resultado = {"revisadas": len(cuotas), "actualizadas": 0}

    for cuota in cuotas:
        plan_pago_monto = round(float(cuota.get("plan_pago_monto") or 0), 2)
        importe_base = cuota.get("importe_original")
        if importe_base is None:
            importe_base = cuota.get("importe") or 0
        importe_base = max(0, round(float(importe_base or 0) - plan_pago_monto, 2))

        cuota_calculada = calcular_importe_cuota_mensual(conn, jugador, cuota["periodo"], importe_base)
        pagado = 1 if cuota_calculada["beca_total"] else 0
        fecha_pago = hoy if pagado else cuota.get("fecha_pago")
        metodo_pago = "Beca" if pagado else cuota.get("metodo_pago")
        referencia_pago = (
            f"Beca total {cuota_calculada['beca_porcentaje']:g}%"
            if pagado else cuota.get("referencia_pago")
        )

        conn.execute("""
            UPDATE cuotas
            SET importe = %s,
                importe_original = %s,
                descuento_beca = %s,
                beca_porcentaje = %s,
                beca_motivo = %s,
                becada = %s,
                pagado = %s,
                fecha_pago = %s,
                metodo_pago = %s,
                referencia_pago = %s,
                plan_pago_monto = %s,
                plan_pago_detalle = %s
            WHERE id = %s
        """, (
            cuota_calculada["importe"],
            cuota_calculada["importe_original"],
            cuota_calculada["descuento_beca"],
            cuota_calculada["beca_porcentaje"],
            cuota_calculada["beca_motivo"],
            cuota_calculada["becada"],
            pagado,
            fecha_pago,
            metodo_pago,
            referencia_pago,
            cuota_calculada["plan_pago_monto"],
            cuota_calculada["plan_pago_detalle"] or None,
            cuota["id"],
        ))
        resultado["actualizadas"] += 1

    return resultado
