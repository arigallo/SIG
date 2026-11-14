"""Indicadores de cobranza y simulaciones; nunca modifican cuotas ni sueldos."""
from decimal import Decimal, InvalidOperation
from math import ceil, floor
from services.finanzas import beca_vigente, importe_social_segun_miembro


def numero(valor, defecto=0, minimo=0, maximo=1_000_000_000_000):
    try:
        n = Decimal(str(valor))
        if not n.is_finite() or n < minimo or n > maximo:
            return float(defecto)
        return float(n)
    except (InvalidOperation, TypeError, ValueError):
        return float(defecto)


def validar_parametros(args):
    limites = {k: (0, 1_000_000_000_000) for k in
               ['sueldo', 'objetivo', 'cuotas', 'cuota_unitaria', 'reserva', 'otros']}
    limites['saldo'] = (-1_000_000_000_000, 1_000_000_000_000)
    limites['anio'] = (2020, 2100)
    enteros = {'anio', 'mes_cuota', 'mes_sueldo', 'mes_gastos', 'mes_recuperacion', 'plazo_recuperacion'}
    for k in enteros - {'anio'}:
        limites[k] = (1, 12)
    for k in ['cobranza', 'demora', 'recuperacion']:
        limites[k] = (0, 100)
    for k in ['aumento_cuota', 'aumento_gastos', 'cargas']:
        limites[k] = (0, 1000)
    for n in range(1, 13):
        for k in ['socios', 'descuento', 'cuota', 'emision', 'cobranza', 'otros', 'extra']:
            limites[f'{k}_{n}'] = (0, 100 if k in {'descuento', 'cobranza'} else 1000000 if k == 'socios' else 1_000_000_000_000)
        enteros.add(f'socios_{n}')
    for clave, (minimo, maximo) in limites.items():
        valor = args.get(clave)
        if valor is None or str(valor).strip() == '':
            continue
        try:
            n = Decimal(str(valor))
            if not n.is_finite() or not minimo <= n <= maximo or (clave in enteros and n != n.to_integral_value()):
                raise InvalidOperation
        except InvalidOperation:
            raise ValueError(f'Revisá el campo {clave}: debe estar entre {minimo} y {maximo}'
                             + (' y ser un número entero.' if clave in enteros else '.')) from None
    for clave, valores in [('fuente', {'socios', 'cuotas', 'presupuesto'}),
                           ('sueldo_incluido', {'si', 'no'}), ('aguinaldo', {'si', 'no'})]:
        if args.get(clave) is not None and args.get(clave) not in valores:
            raise ValueError(f'Revisá la selección de {clave}.')


def obtener_indicadores(conn, hoy, periodos):
    historial = conn.execute("""
        SELECT periodo AS mes, COUNT(*) AS cantidad,
            COALESCE(SUM(importe), 0) AS emitido,
            COALESCE(SUM(CASE WHEN pagado = 1 THEN importe ELSE 0 END), 0) AS cobrado,
            COALESCE(SUM(CASE WHEN pagado = 1 AND fecha_pago <= fecha_vencimiento
                THEN importe ELSE 0 END), 0) AS en_termino,
            COALESCE(SUM(CASE WHEN pagado = 1 AND substring(fecha_pago from 1 for 7) > periodo
                THEN importe ELSE 0 END), 0) AS tardio
        FROM cuotas WHERE periodo = ANY(%s) AND COALESCE(anulada, 0) = 0
            AND COALESCE(incobrable, 0) = 0 AND importe > 0
        GROUP BY periodo ORDER BY periodo
    """, (periodos,)).fetchall()
    por_mes = {r['mes']: r for r in historial}
    filas = []
    for mes in periodos:
        r = dict(por_mes.get(mes, {'mes': mes, 'cantidad': 0, 'emitido': 0, 'cobrado': 0, 'en_termino': 0}))
        r['tasa'] = float(r['cobrado']) / float(r['emitido']) * 100 if r['emitido'] else 0
        filas.append(r)
    deudores = conn.execute("""
        SELECT j.id, j.nombre, j.apellido, COUNT(*) AS cuotas,
            SUM(c.importe) AS deuda,
            SUM(CASE WHEN c.fecha_vencimiento < %s THEN c.importe ELSE 0 END) AS vencido,
            MIN(c.fecha_vencimiento) AS primer_vencimiento
        FROM cuotas c JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.pagado = 0 AND COALESCE(c.anulada, 0) = 0
            AND COALESCE(c.incobrable, 0) = 0 AND c.importe > 0
        GROUP BY j.id, j.nombre, j.apellido ORDER BY vencido DESC, deuda DESC, j.id
    """, (hoy.isoformat(),)).fetchall()
    caja = conn.execute("""
        SELECT COALESCE(SUM(CASE WHEN tipo = 'ingreso' THEN monto
            WHEN tipo = 'egreso' THEN -monto ELSE 0 END), 0) AS saldo
        FROM movimientos WHERE COALESCE(anulado, 0) = 0 AND fecha <= %s
    """, (hoy.isoformat(),)).fetchone()
    emitido = sum(float(r['emitido']) for r in filas)
    cobrado = sum(float(r['cobrado']) for r in filas)
    tardio = sum(float(r.get('tardio') or 0) for r in filas)
    return {'historial': filas, 'deudores': deudores,
            'deuda': sum(float(r['deuda']) for r in deudores),
            'vencido': sum(float(r['vencido']) for r in deudores),
            'saldo': float(caja['saldo']), 'emitido': emitido, 'cobrado': cobrado,
            'tasa': cobrado / emitido * 100 if emitido else 0,
            'tardio': tardio,
            'demora': tardio / cobrado * 100 if cobrado else 0,
            'base_cuotas': emitido / len(periodos) if periodos else 0,
            'meses_con_datos': sum(1 for r in filas if r['emitido'])}


def obtener_padron(conn, anio):
    jugadores = conn.execute("""SELECT tipo_miembro, beca_activa, beca_porcentaje,
        beca_desde, beca_hasta FROM jugadores
        WHERE estado = 'Activo' AND COALESCE(cobra_cuota, 1) = 1""").fetchall()
    meses = []
    for n in range(1, 13):
        equivalentes = 0
        becados = 0
        for jugador in jugadores:
            porcentaje = float(jugador.get('beca_porcentaje') or 0) if beca_vigente(jugador, f'{anio}-{n:02d}') else 0
            becados += int(porcentaje > 0)
            equivalentes += importe_social_segun_miembro(jugador, 1) * (1 - porcentaje / 100)
        meses.append({'socios': len(jugadores), 'becados': becados,
                      'descuento': (1 - equivalentes / len(jugadores)) * 100 if jugadores else 0})
    return meses


def obtener_comportamiento(conn, hoy, periodos):
    return conn.execute("""SELECT j.id, j.nombre, j.apellido,
        COUNT(*) AS emitidas,
        COUNT(CASE WHEN c.pagado = 1 AND c.fecha_pago <= c.fecha_vencimiento THEN 1 END) AS al_dia,
        COUNT(CASE WHEN c.pagado = 1 AND c.fecha_pago > c.fecha_vencimiento THEN 1 END) AS tardias,
        COUNT(CASE WHEN c.pagado = 1 AND (NULLIF(c.fecha_pago, '') IS NULL
            OR NULLIF(c.fecha_vencimiento, '') IS NULL) THEN 1 END) AS sin_fechas,
        COUNT(CASE WHEN c.pagado = 0 AND c.fecha_vencimiento < %s THEN 1 END) AS vencidas,
        COALESCE(SUM(CASE WHEN c.pagado = 0 THEN c.importe ELSE 0 END), 0) AS pendiente
        FROM cuotas c JOIN jugadores j ON j.id = c.jugador_id
        WHERE c.periodo = ANY(%s) AND COALESCE(c.anulada, 0) = 0
            AND COALESCE(c.incobrable, 0) = 0 AND c.importe > 0
        GROUP BY j.id, j.nombre, j.apellido
        ORDER BY vencidas DESC, tardias DESC, j.apellido, j.nombre
    """, (hoy.isoformat(), periodos)).fetchall()


def simular(items, indicadores, args, anio):
    anio = int(numero(args.get('anio'), anio, minimo=2020, maximo=2100))
    padron = indicadores.get('padron') or [{'socios': 0, 'descuento': 0}] * 12
    equivalentes = padron[0]['socios'] * (1 - padron[0]['descuento'] / 100)
    p = {
        'sueldo': numero(args.get('sueldo'), 550000),
        'objetivo': numero(args.get('objetivo'), 550000),
        'cuotas': numero(args.get('cuotas'), indicadores['base_cuotas']),
        'cuota_unitaria': numero(args.get('cuota_unitaria'), indicadores['base_cuotas'] / equivalentes if equivalentes else 0),
        'cobranza': numero(args.get('cobranza'), indicadores['tasa'], maximo=100),
        'demora': numero(args.get('demora'), indicadores.get('demora', 0), maximo=100),
        'aumento_cuota': numero(args.get('aumento_cuota'), 0, maximo=1000),
        'aumento_gastos': numero(args.get('aumento_gastos'), 0, maximo=1000),
        'cargas': numero(args.get('cargas'), 0, maximo=1000),
        'reserva': numero(args.get('reserva'), 0),
        'otros': numero(args.get('otros'), 0),
        'recuperacion': numero(args.get('recuperacion'), 0, maximo=100),
        'saldo': numero(args.get('saldo'), 0, minimo=-1_000_000_000_000),
        'sueldo_incluido': args.get('sueldo_incluido', 'no') == 'si',
        'aguinaldo': args.get('aguinaldo') == 'si',
        'fuente': args.get('fuente') if args.get('fuente') in {'presupuesto', 'socios', 'cuotas'} else ('socios' if equivalentes else 'cuotas'),
    }
    for clave, defecto in [('mes_cuota', 1), ('mes_sueldo', 1), ('mes_gastos', 1),
                           ('mes_recuperacion', 1), ('plazo_recuperacion', 6)]:
        p[clave] = int(numero(args.get(clave), defecto, minimo=1, maximo=12))
    p['objetivo'] = max(p['sueldo'], p['objetivo'])
    configuracion = []
    for n in range(1, 13):
        configuracion.append({
            'numero': n, 'mes': f'{anio}-{n:02d}',
            'socios': int(numero(args.get(f'socios_{n}'), padron[n - 1]['socios'], maximo=1000000)),
            'descuento': numero(args.get(f'descuento_{n}'), padron[n - 1]['descuento'], maximo=100),
            'cuota': numero(args.get(f'cuota_{n}'), p['cuota_unitaria']),
            'emision': numero(args.get(f'emision_{n}'), p['cuotas']),
            'cobranza': numero(args.get(f'cobranza_{n}'), p['cobranza'], maximo=100),
            'otros': numero(args.get(f'otros_{n}'), p['otros']),
            'extra': numero(args.get(f'extra_{n}'), 0),
        })
    escenarios = []
    for clave, nombre, variacion, recuperacion_factor in [
            ('conservador', 'Conservador', -15, 0.5),
            ('esperado', 'Esperado', 0, 1),
            ('optimista', 'Optimista', 10, 1)]:
        saldo = p['saldo']
        acumulado = sensibilidad_acum = exposicion = 0
        limites = []
        necesidad_cuota = 0
        cuota_calculable = p['fuente'] != 'presupuesto'
        deficit_previo = False
        filas = []
        factor = 1 + p['cargas'] / 100
        pendiente = sensibilidad_pendiente = 0
        deuda_inicial = max(0, float(indicadores.get('vencido') or 0))
        deuda_a_recuperar = deuda_inicial * p['recuperacion'] / 100 * recuperacion_factor
        plazo = p['plazo_recuperacion']
        for cfg in configuracion:
            n, mes = cfg['numero'], cfg['mes']
            tasa = min(100, max(0, cfg['cobranza'] + variacion))
            vigentes = [i for i in items if i.get('activo')
                        and (not i.get('mes_inicio') or i['mes_inicio'] <= mes)
                        and (not i.get('mes_fin') or i['mes_fin'] >= mes)]
            ingresos_presupuesto = sum(float(i['monto']) for i in vigentes if i['tipo'] == 'ingreso')
            gastos_presupuesto = sum(float(i['monto']) for i in vigentes if i['tipo'] == 'egreso')
            ajuste_gasto = 1 + (p['aumento_gastos'] / 100 if n >= p['mes_gastos'] else 0)
            gastos = gastos_presupuesto * ajuste_gasto + cfg['extra']
            emision_base = (cfg['socios'] * cfg['cuota'] * (1 - cfg['descuento'] / 100)
                            if p['fuente'] == 'socios' else cfg['emision'])
            ajuste_cuota = 1 + (p['aumento_cuota'] / 100 if n >= p['mes_cuota'] else 0)
            emision = emision_base * ajuste_cuota
            cobrable = emision * tasa / 100
            sensibilidad = emision_base * tasa / 100 / 100 if n >= p['mes_cuota'] else 0
            demora = p['demora'] / 100
            cobro = cobrable * (1 - demora) + pendiente
            sensibilidad_cobro = sensibilidad * (1 - demora) + sensibilidad_pendiente
            pendiente = cobrable * demora
            sensibilidad_pendiente = sensibilidad * demora
            recuperado = deuda_a_recuperar / plazo if p['mes_recuperacion'] <= n < p['mes_recuperacion'] + plazo else 0
            if p['fuente'] == 'presupuesto':
                ingresos = ingresos_presupuesto
                cobro = recuperado = 0
            else:
                ingresos = cfg['otros'] + cobro + recuperado
            sac_base = p['sueldo'] * 0.5 * factor if p['aguinaldo'] and n in {6, 12} else 0
            if not p['sueldo_incluido']:
                gastos += p['sueldo'] * factor
            gastos += sac_base
            coeficiente = factor * (1 + (0.5 if p['aguinaldo'] and n in {6, 12} else 0)) if n >= p['mes_sueldo'] else 0
            adicional = (p['objetivo'] - p['sueldo']) * coeficiente
            resultado = ingresos - gastos
            acumulado += resultado
            exposicion += coeficiente
            margen_saldo = p['saldo'] + acumulado - p['reserva']
            if exposicion:
                limites.append(margen_saldo / exposicion)
            elif margen_saldo < -0.005:
                deficit_previo = True
            sensibilidad_acum += sensibilidad_cobro if p['fuente'] != 'presupuesto' else 0
            faltante = -margen_saldo + (p['objetivo'] - p['sueldo']) * exposicion
            if faltante > 0.005:
                if sensibilidad_acum:
                    necesidad_cuota = max(necesidad_cuota, faltante / sensibilidad_acum)
                else:
                    cuota_calculable = False
            saldo += resultado - adicional
            filas.append({'mes': mes, 'socios': cfg['socios'], 'emision': emision,
                          'tasa': tasa, 'cobro': cobro, 'recuperado': recuperado,
                          'ingresos': ingresos, 'gastos': gastos, 'extra': cfg['extra'],
                          'adicional': adicional, 'resultado': resultado - adicional,
                          'saldo': saldo})
        margen = min(limites) if limites else 0
        aumento_maximo = floor(max(0, margen) * 100 + 1e-6) / 100 if not deficit_previo else 0
        primer_deficit = next((f['mes'] for f in filas if f['saldo'] < p['reserva'] - 0.005), None)
        escenarios.append({'clave': clave, 'nombre': nombre,
                           'tasa': min(100, max(0, p['cobranza'] + variacion)), 'filas': filas,
                           'aumento_maximo': aumento_maximo,
                           'sueldo_maximo': p['sueldo'] + aumento_maximo,
                           'deficit_base': margen < -0.005 or deficit_previo,
                           'viable': primer_deficit is None,
                           'primer_deficit': primer_deficit, 'saldo_final': saldo,
                           'saldo_minimo': min(f['saldo'] for f in filas),
                           'ingresos_total': sum(f['ingresos'] for f in filas),
                           'gastos_total': sum(f['gastos'] + f['adicional'] for f in filas),
                           'cuotas_por_cobrar_diciembre': pendiente if p['fuente'] != 'presupuesto' else 0,
                           'deuda_a_recuperar_despues': max(0, deuda_a_recuperar - sum(f['recuperado'] for f in filas)) if p['fuente'] != 'presupuesto' else 0,
                           'cuota_calculable': cuota_calculable,
                           'aumento_cuota_adicional': ceil(max(0, necesidad_cuota) * 100 - 1e-6) / 100})
    formulario = {k: ('si' if v else 'no') if isinstance(v, bool) else str(v) for k, v in p.items()}
    formulario['anio'] = str(anio)
    ajustes = {}
    for cfg in configuracion:
        for campo in ['socios', 'descuento', 'cuota', 'emision', 'cobranza', 'otros', 'extra']:
            clave = f'{campo}_{cfg["numero"]}'
            if args.get(clave) is not None and str(args.get(clave)).strip() != '':
                ajustes[clave] = str(cfg[campo])
                formulario[clave] = str(cfg[campo])
    valores = [f['saldo'] for e in escenarios for f in e['filas']] + [p['reserva'], 0]
    bajo, alto = min(valores), max(valores)
    rango = max(alto - bajo, 1)
    def y(valor):
        return round(180 - (valor - bajo) / rango * 150, 2)
    grafico = {'reserva_y': y(p['reserva']), 'cero_y': y(0)}
    for e in escenarios:
        e['puntos'] = ' '.join(f'{45 + n * 60},{y(f["saldo"])}' for n, f in enumerate(e['filas']))
    return {'parametros': p, 'escenarios': escenarios, 'anio': anio,
            'configuracion': configuracion, 'formulario': formulario, 'ajustes': ajustes, 'grafico': grafico,
            'base_esperada': escenarios[1],
            'meses_sin_gastos': [f'{anio}-{n:02d}' for n in range(1, 13) if not any(
                i.get('activo') and i['tipo'] == 'egreso' and float(i['monto'] or 0) > 0
                and (not i.get('mes_inicio') or i['mes_inicio'] <= f'{anio}-{n:02d}')
                and (not i.get('mes_fin') or i['mes_fin'] >= f'{anio}-{n:02d}') for i in items)],
            'referencia': {'deuda_vencida': indicadores.get('vencido', 0),
                           'tasa_historica': indicadores['tasa'],
                           'demora_historica': indicadores.get('demora', 0),
                           'meses_con_datos': indicadores.get('meses_con_datos', 0)}}
