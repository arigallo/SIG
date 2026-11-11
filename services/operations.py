"""Durable protection against replaying an administrative form submission.

A reservation is committed before calling the business route. An interrupted
operation stays reserved: retrying an ambiguous payment/message is never automatic.
"""
import hashlib
import json
import re

from flask import g, request, session, redirect, render_template, flash


PROTECTED_ENDPOINTS = {
    'nuevo_movimiento', 'editar_movimiento', 'eliminar_movimiento',
    'nueva_cuota', 'pagar_cuota', 'revisar_comprobante_cuota',
    'registrar_pago_posterior_gasto_compartido', 'registrar_factura_recibida_en_caja',
    'actualizar_item_gasto_compartido',
    'registrar_venta_club', 'nuevo_plan_pago', 'pagos_masivos_debito_automatico',
    'nueva_bitacora_jugador', 'resolver_nota_finanzas',
    'responder_whatsapp_inbox', 'solicitar_comprobante_jugador', 'solicitar_comprobante_bbva',
    'enviar_email_comunicacion_moroso', 'enviar_email_comunicacion_morosos_lote',
    'enviar_whatsapp_comunicacion_moroso', 'enviar_whatsapp_comunicacion_morosos_lote',
    'enviar_recordatorio_cuota', 'enviar_recordatorio_cuota_whatsapp',
    'enviar_recordatorio_ficha', 'enviar_recordatorio_ficha_whatsapp',
    'enviar_recordatorios_cuotas_lote', 'enviar_recordatorios_cuotas_whatsapp_lote',
    'enviar_recordatorios_fichas_lote', 'enviar_recordatorios_fichas_whatsapp_lote',
    'enviar_recordatorio_gasto_compartido', 'enviar_recordatorios_gasto_compartido_lote',
    'enviar_recordatorio_evento_calendario', 'enviar_notificacion_app_manual',
}


def crear_tabla_operaciones(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS operaciones_formulario (
        usuario TEXT NOT NULL, token TEXT NOT NULL, huella TEXT NOT NULL,
        estado TEXT NOT NULL DEFAULT 'procesando', destino TEXT,
        creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
        terminado_en TIMESTAMPTZ, PRIMARY KEY (usuario, token)
    )""")


def huella_solicitud():
    fields = [(key, values) for key, values in sorted(request.form.lists())
              if key not in {'_csrf_token', '_operation_token'}]
    digest = hashlib.sha256(json.dumps([request.path, fields], ensure_ascii=False).encode())
    for key, files in sorted(request.files.lists()):
        for file in files:
            digest.update(json.dumps([key, file.filename, file.mimetype]).encode())
            position = file.stream.tell()
            try:
                while chunk := file.stream.read(65536):
                    digest.update(chunk)
            finally:
                file.stream.seek(position)
    return digest.hexdigest()


def instalar_proteccion(app, connection_factory):
    @app.before_request
    def reservar_operacion():
        # Registered after authentication/CSRF/maintenance checks.
        if request.method != 'POST' or request.endpoint not in PROTECTED_ENDPOINTS or not session.get('user_id'):
            return
        token = request.form.get('_operation_token', '')
        if not re.fullmatch(r'[A-Za-z0-9_-]{20,100}', token):
            return render_template('operation_conflict.html', motivo='formulario'), 409
        user = str(session['user_id'])
        fingerprint = huella_solicitud()
        conn = connection_factory()
        try:
            reserved = conn.execute("""INSERT INTO operaciones_formulario (usuario, token, huella)
                VALUES (%s, %s, %s) ON CONFLICT DO NOTHING RETURNING token""",
                (user, token, fingerprint)).fetchone()
            if reserved:
                conn.commit()
                g.operacion_formulario = (user, token)
                return
            previous = conn.execute('SELECT * FROM operaciones_formulario WHERE usuario = %s AND token = %s',
                                    (user, token)).fetchone()
            if previous['huella'] != fingerprint:
                return render_template('operation_conflict.html', motivo='cambio'), 409
            if previous['estado'] == 'terminada' and previous['destino']:
                flash('Esta operación ya fue procesada. No se volvió a registrar ni enviar.', 'ok')
                return redirect(previous['destino'], code=303)
            return render_template('operation_conflict.html', motivo='procesando'), 409
        finally:
            conn.close()

    @app.after_request
    def terminar_operacion(response):
        operation = getattr(g, 'operacion_formulario', None)
        if not operation:
            return response
        conn = connection_factory()
        try:
            destination = response.headers.get('Location', '')
            if 300 <= response.status_code < 400 and destination.startswith('/') and not destination.startswith('//'):
                conn.execute("""UPDATE operaciones_formulario SET estado = 'terminada', destino = %s,
                    terminado_en = CURRENT_TIMESTAMP WHERE usuario = %s AND token = %s""", (destination, *operation))
            elif response.status_code < 500:
                # Validation errors render the form without performing the operation.
                conn.execute('DELETE FROM operaciones_formulario WHERE usuario = %s AND token = %s', operation)
            # A 5xx or process interruption is ambiguous. Keep the reservation.
            conn.commit()
        finally:
            conn.close()
        response.headers['Cache-Control'] = 'no-store'
        return response
