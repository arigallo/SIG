"""Read models for the daily workspace. All changes still use SIG's existing routes."""
from datetime import timedelta


def ordenar_notas_financieras(items):
    return sorted((item for item in items if item['tipo'] == 'finanzas'),
                  key=lambda item: (bool(item.get('resuelto_en')), str(item.get('seguimiento_fecha') or '9999-12-31'), -item['id']))


def obtener_notas_financieras(conn, jugador_id):
    return ordenar_notas_financieras(conn.execute("""
        SELECT * FROM jugador_bitacora WHERE jugador_id = %s AND tipo = 'finanzas'
        ORDER BY creado_en DESC, id DESC
    """, (jugador_id,)).fetchall())


def preparar_gestion(item, hoy):
    """Explicit next step, with an agreed future date taking precedence over chasing debt."""
    item = dict(item)
    if item.get('archivo') and item.get('comprobante_estado') in (None, '', 'sin_comprobante'):
        item['comprobante_estado'] = 'pendiente'
    fecha = str(item.get('seguimiento_fecha') or '')
    activa = bool(item.get('nota')) and not item.get('resuelto_en')
    revision = bool(item.get('archivo')) and item.get('comprobante_estado') == 'pendiente'
    item['carril'] = 'seguimientos' if item.get('origen') == 'seguimiento' else 'revision' if revision else 'cobros'
    if item['carril'] == 'seguimientos':
        futura = fecha and fecha > hoy
        item.update(situacion='Acuerdo vigente' if futura else 'Seguimiento abierto',
                    proxima_accion='Esperar hasta ' + fecha if futura else 'Revisar y resolver la gestión',
                    prioridad=3 if futura else 1)
    elif revision:
        item.update(situacion='Comprobante por revisar', proxima_accion='Revisar comprobante', prioridad=0)
    elif activa and fecha and fecha > hoy:
        item.update(situacion='Acuerdo vigente', proxima_accion='Esperar hasta ' + fecha, prioridad=3)
    elif activa and fecha and fecha <= hoy:
        item.update(situacion='Seguimiento pendiente', proxima_accion='Retomar la gestión', prioridad=1)
    elif item.get('comprobante_estado') == 'rechazado':
        item.update(situacion='Comprobante rechazado', proxima_accion='Solicitar corrección del comprobante', prioridad=1)
    elif str(item.get('fecha_vencimiento') or '9999') < hoy:
        item.update(situacion='Saldo vencido', proxima_accion='Contactar y registrar un acuerdo', prioridad=2)
    else:
        item.update(situacion='Saldo pendiente', proxima_accion='Consultar vencimiento', prioridad=4)
    return item


def obtener_cola_finanzas(conn, hoy, vista=None, busqueda='', pagina=1, por_pagina=None):
    # Follow-ups are their own work items, even when the player's debt is zero.
    sql = """
        WITH pendientes AS (
            SELECT 'cuota' AS origen, c.id, c.jugador_id, NULL::integer AS gasto_id,
                   j.nombre, j.apellido, 'Cuota ' || c.periodo AS concepto, c.importe,
                   c.fecha_vencimiento::text, c.comprobante_estado, c.comprobante_drive_file_id AS archivo
            FROM cuotas c JOIN jugadores j ON j.id = c.jugador_id
            WHERE c.pagado = 0 AND COALESCE(c.anulada, 0) = 0
              AND COALESCE(c.incobrable, 0) = 0 AND c.importe > 0
            UNION ALL
            SELECT 'gasto', i.id, i.jugador_id, g.id,
                   COALESCE(j.nombre, i.nombre_manual, a.nombre), COALESCE(j.apellido, i.apellido_manual, a.apellido),
                   g.titulo, i.importe, g.fecha_vencimiento::text, i.comprobante_estado, i.comprobante_drive_file_id
            FROM gasto_compartido_items i JOIN gastos_compartidos g ON g.id = i.gasto_id
            LEFT JOIN jugadores j ON j.id = i.jugador_id LEFT JOIN aspirantes a ON a.id = i.aspirante_id
            WHERE i.estado = 'pendiente' AND i.importe > 0
            UNION ALL
            SELECT 'seguimiento', n.id, n.jugador_id, NULL::integer,
                   j.nombre, j.apellido, 'Seguimiento interno', 0,
                   NULL::text, NULL::text, NULL::text
            FROM jugador_bitacora n JOIN jugadores j ON j.id = n.jugador_id
            WHERE n.tipo = 'finanzas' AND n.resuelto_en IS NULL
        )
        , gestiones AS (
        SELECT p.*, n.nota, n.creado_en AS ultima_gestion, n.creado_por AS gestor,
               n.seguimiento_fecha, n.resuelto_en
        FROM pendientes p
        LEFT JOIN LATERAL (
            SELECT nota, creado_en, creado_por, seguimiento_fecha, resuelto_en
            FROM jugador_bitacora WHERE jugador_id = p.jugador_id AND tipo = 'finanzas'
              AND (p.origen <> 'seguimiento' OR id = p.id)
            ORDER BY (resuelto_en IS NULL) DESC, creado_en DESC, id DESC LIMIT 1
        ) n ON TRUE
        ), clasificados AS (
            SELECT *, CASE WHEN origen = 'seguimiento' THEN 'seguimientos'
                WHEN NULLIF(archivo, '') IS NOT NULL AND COALESCE(NULLIF(comprobante_estado, ''), 'sin_comprobante') IN ('pendiente', 'sin_comprobante') THEN 'revision'
                ELSE 'cobros' END AS carril_sql,
                CASE WHEN origen <> 'seguimiento' AND NULLIF(archivo, '') IS NOT NULL
                        AND COALESCE(NULLIF(comprobante_estado, ''), 'sin_comprobante') IN ('pendiente', 'sin_comprobante') THEN 0
                     WHEN resuelto_en IS NULL AND seguimiento_fecha::text > %(hoy)s THEN 3
                     WHEN origen = 'seguimiento' OR (resuelto_en IS NULL AND seguimiento_fecha::text <= %(hoy)s)
                        OR comprobante_estado = 'rechazado' THEN 1
                     WHEN fecha_vencimiento < %(hoy)s THEN 2 ELSE 4 END AS prioridad_sql
            FROM gestiones
        )
    """
    params = {'hoy': hoy}
    conditions = []
    if vista:
        conditions.append('carril_sql = %(vista)s')
        params['vista'] = vista
    if busqueda:
        conditions.append("concat_ws(' ', apellido, nombre, concepto, nota) ILIKE %(q)s")
        params['q'] = '%' + busqueda.replace('\\', '\\\\').replace('%', '\\%').replace('_', '\\_') + '%'
    where = ' WHERE ' + ' AND '.join(conditions) if conditions else ''
    page = None
    if por_pagina:
        total = conn.execute(sql + 'SELECT COUNT(*) AS total FROM clasificados' + where, params).fetchone()['total']
        page = paginacion(total, pagina, por_pagina)
        params.update(limite=page['por_pagina'], offset=page['offset'])
    query = sql + 'SELECT * FROM clasificados' + where + """ ORDER BY prioridad_sql,
        COALESCE(seguimiento_fecha::text, fecha_vencimiento, '9999'), apellido, nombre, origen, id"""
    if por_pagina:
        query += ' LIMIT %(limite)s OFFSET %(offset)s'
    rows = conn.execute(query, params).fetchall()
    items = [preparar_gestion(row, hoy) for row in rows]
    if page is not None:
        counts = conn.execute(sql + 'SELECT carril_sql, COUNT(*) AS total FROM clasificados GROUP BY carril_sql', {'hoy': hoy}).fetchall()
        return items, page, {row['carril_sql']: row['total'] for row in counts}
    return items


def paginacion(total, pagina, por_pagina=25):
    try:
        pagina = max(1, int(pagina))
    except (ValueError, TypeError):
        pagina = 1
    paginas = max(1, (total + por_pagina - 1) // por_pagina)
    pagina = min(pagina, paginas)
    return dict(total=total, pagina=pagina, paginas=paginas, por_pagina=por_pagina,
                offset=(pagina - 1) * por_pagina)


def agrupar_pendientes_inicio(items):
    groups = {}
    for item in items:
        if item['prioridad'] >= 3:
            continue
        key = ('jugador', item['jugador_id']) if item.get('jugador_id') else (item['origen'], item['id'])
        if key not in groups:
            groups[key] = {**item, 'importe': 0, 'conceptos': 0}
        groups[key]['importe'] += float(item['importe'] or 0)
        groups[key]['conceptos'] += item.get('origen') != 'seguimiento'
    return list(groups.values())


def obtener_inicio(conn, puede, username, modulos, hoy):
    data = {'tareas': [], 'finanzas': [], 'salud': [], 'madrinas': [], 'secretaria': [], 'eventos': []}
    if modulos and puede('tareas_ver', 'tareas_gestionar'):
        data['tareas'] = conn.execute("""SELECT * FROM tareas_sig WHERE estado = 'pendiente'
            AND modulo = ANY(%s) AND lower(COALESCE(NULLIF(responsable, ''), creado_por, '')) = lower(%s)
            ORDER BY fecha_vencimiento NULLS LAST, id DESC LIMIT 8""", (sorted(modulos), username)).fetchall()
    if puede('cuotas_ver', 'cuotas_gestionar'):
        data['finanzas'] = agrupar_pendientes_inicio(obtener_cola_finanzas(conn, hoy.isoformat()))
    if puede('salud_ver'):
        data['salud'] = conn.execute("""SELECT j.id, j.apellido, j.nombre,
            CASE WHEN f.id IS NULL OR COALESCE(f.presentada, 0) = 0 THEN 'Ficha médica faltante'
                 WHEN NULLIF(f.fecha_vencimiento::text, '') < %s THEN 'Ficha médica vencida'
                 ELSE 'Apto físico pendiente' END AS situacion
            FROM jugadores j LEFT JOIN fichas_medicas f ON f.jugador_id = j.id
            WHERE j.estado = 'Activo' AND (f.id IS NULL OR COALESCE(f.presentada, 0) = 0
                OR NULLIF(f.fecha_vencimiento::text, '') < %s
                OR (COALESCE(f.apto_fisico, 0) = 0 AND NULLIF(f.documento_drive_file_id, '') IS NULL))
            ORDER BY j.apellido, j.nombre LIMIT 8""", (hoy.isoformat(), hoy.isoformat())).fetchall()
    if puede('aspirantes_ver'):
        data['madrinas'] = conn.execute("""SELECT id, nombre, apellido, proxima_accion_fecha FROM aspirantes
            WHERE estado = 'Aspirante' AND (proxima_accion_fecha IS NULL OR proxima_accion_fecha <= %s)
            ORDER BY proxima_accion_fecha NULLS FIRST, id DESC LIMIT 8""", (hoy.isoformat(),)).fetchall()
    if puede('secretaria_ver', 'secretaria_gestionar'):
        data['secretaria'] = conn.execute("""SELECT id, titulo, fecha_vencimiento FROM secretaria_documentos
            WHERE NULLIF(fecha_vencimiento::text, '') <= %s ORDER BY fecha_vencimiento LIMIT 8""",
            ((hoy + timedelta(days=30)).isoformat(),)).fetchall()
    if puede('calendario_ver', 'asistencia_ver'):
        data['eventos'] = conn.execute("""SELECT id, titulo, fecha, hora_inicio FROM calendario_eventos
            WHERE fecha >= %s ORDER BY fecha, hora_inicio NULLS LAST LIMIT 5""", (hoy.isoformat(),)).fetchall()
    return data
