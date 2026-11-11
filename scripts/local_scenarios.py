"""One-time review fixtures. Invoked exclusively by scripts/local.py, never production."""
import io
import secrets
from datetime import timedelta


def seed_review_scenarios(sig):
    conn = sig.get_connection()
    try:
        if conn.execute('SELECT 1 FROM sig_local_seed WHERE version = 3').fetchone():
            return
        today = sig.ahora_sig().date()
        category = sig.planteles_disponibles(conn)[0]
        pdf_stream = io.BytesIO()
        from reportlab.pdfgen import canvas
        pdf = canvas.Canvas(pdf_stream)
        pdf.drawString(45, 760, 'SIG LOCAL - COMPROBANTE FICTICIO PARA REVISION')
        pdf.drawString(45, 735, 'No acredita ninguna transferencia real.')
        pdf.save()
        media = sig.MediaIoBaseUpload(io.BytesIO(pdf_stream.getvalue()), mimetype='application/pdf', resumable=False)
        attachment = sig.drive_service().files().create(body={'name': 'comprobante-ficticio-revision.pdf', 'parents': ['local-root']}, media_body=media).execute()

        def period(offset):
            index = today.year * 12 + today.month - 1 + offset
            return f'{index // 12:04d}-{index % 12 + 1:02d}'

        cases = ['Mora de varios meses', 'Comprobante pendiente', 'Comprobante rechazado', 'Acuerdo vigente',
                 'Seguimiento vencido', 'Beca total', 'Beca parcial', 'Plan parcialmente abonado',
                 'Ficha faltante', 'Ficha vencida', 'Lesión en seguimiento', 'Historial extenso']
        players = []
        for index, case in enumerate(cases):
            player = conn.execute("""INSERT INTO jugadores
                (nombre, apellido, dni, email, telefono, categoria, fecha_ingreso, estado, tipo_miembro,
                 cobra_cuota, portal_activo, portal_token, observaciones, beca_activa, beca_porcentaje, beca_motivo)
                VALUES (%s, 'Demo revisión', %s, %s, '1100000099', %s, %s, 'Activo', 'Jugador', 1, 1, %s, %s, %s, %s, %s)
                RETURNING id""", (case, f'770000{index:02d}', f'escenario{index}@example.invalid', category,
                  today.isoformat(), secrets.token_urlsafe(24), 'Datos ficticios para revisar SIG local.',
                  int(index in (5, 6)), 100 if index == 5 else 40 if index == 6 else 0,
                  'Escenario de prueba local' if index in (5, 6) else None)).fetchone()
            player_id = player['id']
            players.append(player_id)
            if index != 7:
                months = range(-14, 1) if index == 11 else range(-2, 1) if index == 0 else [-1]
                for offset in months:
                    paid = index == 5 or (index == 11 and offset < -2)
                    amount = 0 if index == 5 else 12000 if index == 6 else 20000
                    conn.execute("""INSERT INTO cuotas (jugador_id, periodo, importe, pagado, fecha_vencimiento,
                        fecha_pago, importe_original, becada, beca_porcentaje, descuento_beca,
                        comprobante_drive_file_id, comprobante_nombre, comprobante_mime_type,
                        comprobante_tamano, comprobante_estado, comprobante_fecha, comprobante_observaciones)
                        VALUES (%s, %s, %s, %s, %s, %s, 20000, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                        (player_id, period(offset), amount, int(paid), (today - timedelta(days=35)).isoformat(),
                         (today - timedelta(days=40)).isoformat() if paid else None,
                         int(index in (5, 6)), 100 if index == 5 else 40 if index == 6 else 0,
                         20000 - amount if index in (5, 6) else 0,
                         attachment['id'] if index in (1, 2) else None,
                         attachment['name'] if index in (1, 2) else None,
                         'application/pdf' if index in (1, 2) else None,
                         attachment['size'] if index in (1, 2) else None,
                         'pendiente' if index == 1 else 'rechazado' if index == 2 else None,
                         sig.ahora_sig().isoformat() if index in (1, 2) else None,
                         'Archivo ilegible; solicitar una copia clara. Ejemplo local.' if index == 2 else None))
            if index in (3, 4, 11):
                text = 'Acordó pagar después de cobrar; esperar hasta la fecha indicada.' if index == 3 else 'Venció la fecha acordada. Contactar para confirmar el nuevo compromiso.'
                if index == 11:
                    text = ('Seguimiento de ejemplo con contexto extenso. Se revisaron los períodos adeudados, '
                            'los comprobantes disponibles y la propuesta de regularización. ' * 8)
                conn.execute("""INSERT INTO jugador_bitacora (jugador_id, tipo, nota, creado_por, seguimiento_fecha)
                    VALUES (%s, 'finanzas', %s, 'tesorero', %s)""",
                    (player_id, text, today + timedelta(days=10 if index == 3 else -3)))
                conn.execute("""INSERT INTO jugador_bitacora (jugador_id, tipo, nota, creado_por, resuelto_en, resuelto_por)
                    VALUES (%s, 'finanzas', 'Consulta anterior resuelta. Historial local.', 'admin', CURRENT_TIMESTAMP, 'tesorero')""", (player_id,))
            if index not in (8,):
                conn.execute("""INSERT INTO fichas_medicas (jugador_id, presentada, apto_fisico, fecha_vencimiento)
                    VALUES (%s, 1, 1, %s)""", (player_id, (today + timedelta(days=-5 if index == 9 else 180)).isoformat()))
            if index == 10:
                conn.execute("""INSERT INTO lesiones (jugador_id, fecha_lesion, tipo_lesion, zona_cuerpo, estado, proximo_control)
                    VALUES (%s, %s, 'Esguince de prueba', 'Tobillo', 'Activa', %s)""",
                    (player_id, (today - timedelta(days=12)).isoformat(), (today + timedelta(days=2)).isoformat()))

        plan = conn.execute("""INSERT INTO planes_pago (jugador_id, fecha_inicio, monto_total, cantidad_cuotas, monto_cuota,
            estado, descripcion, creado_por) VALUES (%s, %s, 60000, 3, 20000, 'Activo', 'Plan de prueba: una cuota pagada y dos pendientes.', 'tesorero') RETURNING id""",
            (players[7], period(-2) + '-01')).fetchone()
        for offset in (-2, -1, 0):
            conn.execute("""INSERT INTO cuotas (jugador_id, periodo, importe, importe_original, pagado, fecha_pago,
                fecha_vencimiento, plan_pago_id, plan_pago_monto, plan_pago_detalle)
                VALUES (%s, %s, 20000, 20000, %s, %s, %s, %s, 20000, 'Cuota del plan local')""",
                (players[7], period(offset), int(offset == -2), (today - timedelta(days=40)).isoformat() if offset == -2 else None,
                 period(offset) + '-10', plan['id']))
        for owner, module, title in [('admin', 'general', 'Revisar el paquete de mejoras local'),
                                     ('tesorero', 'finanzas', 'Revisar comprobantes de los escenarios demo'),
                                     ('medico', 'salud', 'Completar la ficha faltante de demostración'),
                                     ('entrenador', 'deportivo', 'Confirmar asistencia al próximo entrenamiento'),
                                     ('secretaria', 'secretaria', 'Revisar el vencimiento del documento demo')]:
            conn.execute("""INSERT INTO tareas_sig (titulo, descripcion, modulo, prioridad, estado, responsable,
                fecha_vencimiento, creado_por) VALUES (%s, 'Tarea ficticia para probar el inicio por rol.', %s, 'alta', 'pendiente', %s, %s, 'admin')""",
                (title, module, owner, today.isoformat()))
        conn.execute("""INSERT INTO secretaria_documentos (categoria, titulo, descripcion, fecha_documento,
            fecha_vencimiento, drive_file_id, archivo_nombre, archivo_mime_type, archivo_tamano, creado_por)
            VALUES ('Pruebas', 'Documento local por renovar', 'Archivo ficticio para revisar alertas.', %s, %s, %s, %s, 'application/pdf', %s, 'admin')""",
            ((today - timedelta(days=60)).isoformat(), (today + timedelta(days=4)).isoformat(), attachment['id'], attachment['name'], attachment['size']))
        conn.execute('INSERT INTO sig_local_seed VALUES (3)')
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def seed_financial_review(sig):
    conn = sig.get_connection()
    try:
        if conn.execute('SELECT 1 FROM sig_local_seed WHERE version = 4').fetchone():
            return
        today = sig.ahora_sig().date().isoformat()
        # Mirror only the fictitious paid installments in the local cash ledger.
        paid = conn.execute("""SELECT c.*, j.nombre FROM cuotas c JOIN jugadores j ON j.id = c.jugador_id
            WHERE j.apellido = 'Demo revisión' AND j.email LIKE %s
              AND c.pagado = 1 AND c.importe > 0""", ('escenario%@example.invalid',)).fetchall()
        for item in paid:
            conn.execute("""INSERT INTO movimientos (tipo, concepto, monto, fecha, referencia)
                VALUES ('ingreso', %s, %s, %s, 'Escenario ficticio de revisión')""",
                (f"Cuota demo {item['periodo']} - {item['nombre']}", item['importe'], item['fecha_pago']))
        expense = conn.execute("""INSERT INTO movimientos (tipo, concepto, monto, fecha, referencia)
            VALUES ('egreso', 'Elementos deportivos - demo local', 15000, %s, 'Escenario ficticio de revisión') RETURNING id""", (today,)).fetchone()
        product = conn.execute("""INSERT INTO venta_productos (nombre, descripcion, precio, stock)
            VALUES ('Remera de prueba local', 'Producto ficticio para revisar ventas y stock.', 22000, 7) RETURNING id""").fetchone()
        income = conn.execute("""INSERT INTO movimientos (tipo, concepto, monto, fecha, referencia)
            VALUES ('ingreso', 'Venta de prueba local', 22000, %s, 'Escenario ficticio de revisión') RETURNING id""", (today,)).fetchone()
        sale = conn.execute("""INSERT INTO ventas_club (fecha, comprador, medio_pago, total, notas, movimiento_id, creado_por)
            VALUES (%s, 'Comprador ficticio', 'Transferencia', 22000, 'Venta de demostración local.', %s, 'tesorero') RETURNING id""", (today, income['id'])).fetchone()
        conn.execute("""INSERT INTO venta_items (venta_id, producto_id, producto_nombre, cantidad, precio_unitario, subtotal)
            VALUES (%s, %s, 'Remera de prueba local', 1, 22000, 22000)""", (sale['id'], product['id']))
        conn.execute("""INSERT INTO presupuesto_items (tipo, categoria, concepto, monto, mes_inicio, notas)
            VALUES ('egreso', 'variable', 'Elementos deportivos - demo local', 18000, %s, 'Concepto ficticio para comparar con caja.')""", (today[:7],))
        event = conn.execute("""INSERT INTO gastos_compartidos (titulo, concepto, fecha_evento, fecha_vencimiento,
            modo_importe, monto_total, monto_por_jugador, estado, creado_por)
            VALUES ('Tercer tiempo - demo local', 'Escenario ficticio con participantes y un invitado.', %s, %s,
            'por_jugador', 30000, 10000, 'Activo', 'tesorero') RETURNING id""", (today, today)).fetchone()
        players = conn.execute("""SELECT id FROM jugadores WHERE apellido = 'Demo revisión'
            AND nombre IN ('Mora de varios meses', 'Acuerdo vigente') ORDER BY id""").fetchall()
        for player in players:
            conn.execute("INSERT INTO gasto_compartido_items (gasto_id, jugador_id, importe) VALUES (%s, %s, 10000)", (event['id'], player['id']))
        conn.execute("""INSERT INTO gasto_compartido_items (gasto_id, participante_tipo, nombre_manual, apellido_manual, importe)
            VALUES (%s, 'manual', 'Invitado de prueba', 'Demo revisión', 10000)""", (event['id'],))
        conn.execute('INSERT INTO sig_local_seed VALUES (4)')
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
