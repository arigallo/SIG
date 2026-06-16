def obtener_cuenta_corriente_jugador(conn, jugador_id, limite=30):
    return conn.execute("""
        SELECT *
        FROM (
            SELECT
                c.fecha_vencimiento AS fecha,
                'cuota' AS origen,
                c.id AS origen_id,
                'Cuota ' || c.periodo AS concepto,
                c.importe AS importe,
                CASE
                    WHEN COALESCE(c.anulada, 0) = 1 THEN 'anulado'
                    WHEN c.pagado = 1 THEN 'pagado'
                    ELSE 'pendiente'
                END AS estado,
                c.fecha_pago
            FROM cuotas c
            WHERE c.jugador_id = %s

            UNION ALL

            SELECT
                COALESCE(g.fecha_vencimiento, g.fecha_evento, g.creado_en::text) AS fecha,
                'gasto_compartido' AS origen,
                i.id AS origen_id,
                'Gasto compartido: ' || g.titulo AS concepto,
                i.importe AS importe,
                i.estado,
                i.fecha_pago
            FROM gasto_compartido_items i
            JOIN gastos_compartidos g ON g.id = i.gasto_id
            WHERE i.jugador_id = %s
        ) cuenta
        ORDER BY COALESCE(fecha_pago, fecha) DESC NULLS LAST, origen_id DESC
        LIMIT %s
    """, (jugador_id, jugador_id, limite)).fetchall()


def obtener_cuotas_impagas_para_plan(conn, jugador_id):
    return conn.execute("""
        SELECT id, periodo, importe, importe_original, fecha_vencimiento
        FROM cuotas
        WHERE jugador_id = %s
          AND pagado = 0
          AND COALESCE(anulada, 0) = 0
          AND COALESCE(importe, 0) > 0
        ORDER BY periodo ASC, id ASC
    """, (jugador_id,)).fetchall()
