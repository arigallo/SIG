"""Verificación funcional contra PostgreSQL local y las rutas reales de SIG."""
import io
import secrets
import re
import socket
from html.parser import HTMLParser
from urllib.parse import urlsplit
from werkzeug.exceptions import HTTPException


def check(sig, config):
    app = sig.app
    app.testing = True
    client = app.test_client()
    count = 0

    def expect(condition, description):
        nonlocal count
        if not condition:
            raise AssertionError(description)
        count += 1
        print("OK " + description, flush=True)

    def login(username):
        token = client.get("/login/csrf").json["token"]
        response = client.post("/login", data={"username": username, "password": config["admin_password"], "_csrf_token": token})
        expect(response.status_code == 302 and response.location == "/", "Login real: " + username)

    def post(url, data=None):
        with client.session_transaction() as session:
            token = session["_csrf_token"]
        return client.post(url, data={"_operation_token": secrets.token_urlsafe(24), **(data or {}), "_csrf_token": token})

    def row(sql, parameters=()):
        conn = sig.get_connection()
        try:
            return conn.execute(sql, parameters).fetchone()
        finally:
            conn.close()

    from services.operations import PROTECTED_ENDPOINTS
    broken_forms = []
    class FormProtectionParser(HTMLParser):
        def __init__(self, path):
            super().__init__()
            self.path, self.current = path, None
        def handle_starttag(self, tag, attrs):
            attributes = dict(attrs)
            if tag == 'form':
                self.current = [attributes.get('method', 'get').lower(), attributes.get('action') or self.path, False]
            elif tag == 'input' and self.current and attributes.get('name') == '_operation_token' and attributes.get('value'):
                self.current[2] = True
        def handle_endtag(self, tag):
            if tag == 'form' and self.current:
                method, action, protected = self.current
                self.current = None
                if method != 'post':
                    return
                try:
                    endpoint, _ = app.url_map.bind('localhost').match(urlsplit(action).path, method='POST')
                except HTTPException:
                    return
                if endpoint in PROTECTED_ENDPOINTS and not protected:
                    broken_forms.append((self.path, endpoint))

    def verify_forms(path, response):
        FormProtectionParser(path).feed(response.get_data(as_text=True))

    expect(client.post("/jugadores/1/bitacora", data={"tipo": "finanzas", "nota": "Sin CSRF"}).status_code == 400,
           "CSRF sigue activo")
    login("admin")
    pages = ["/", "/jugadores", "/jugadores/nuevo", "/ahijadxs", "/ahijadxs/nuevo", "/operacion",
             "/caja", "/presupuesto", "/gastos-compartidos", "/finanzas/cobranzas", "/finanzas/excepciones",
             "/finanzas/ventas", "/finanzas/facturas-recibidas", "/finanzas/transferencias-bbva",
             "/cuotas/comprobantes", "/cuotas/conciliacion", "/calendario", "/asistencia", "/tests",
             "/reportes", "/comunicaciones", "/comunicacion/whatsapp", "/comunicaciones/app", "/movimientos/nuevo", "/secretaria/documentos",
             "/admin/sistema", "/admin/versiones", "/admin/encuestas", "/admin/sugerencias-recomendaciones",
             "/__local/outbox"]
    for url in pages:
        response = client.get(url, follow_redirects=True)
        expect(response.status_code == 200, "Pantalla " + url)
        title = re.search(r"<title>(.*?)</title>", response.get_data(as_text=True), re.S)
        expect(title is not None and "<" not in title.group(1), "Título sin contenido de formularios: " + url)
        verify_forms(url, response)

    unique = secrets.token_hex(5)
    conn = sig.get_connection()
    categoria = sig.planteles_disponibles(conn)[0]
    conn.close()
    response = post("/jugadores/nuevo", {"nombre": "Verificación", "apellido": "Local " + unique,
        "dni": str(80000000 + int(unique, 16) % 9000000),
        "categoria": categoria, "email": "verificacion@example.invalid", "telefono": "1100000000"})
    expect(response.status_code == 302, "Alta de jugador por formulario real")
    jugador = row("SELECT * FROM jugadores WHERE apellido = %s", ("Local " + unique,))
    expect(jugador is not None and jugador["portal_activo"], "Jugador y portal persistidos en PostgreSQL")
    jugador_id = jugador["id"]
    cuota_id = None

    try:
        expect(client.get(f"/jugadores/{jugador_id}").status_code == 200, "Perfil real")
        response = post(f"/jugadores/{jugador_id}/bitacora", {"tipo": "finanzas", "nota": "Paga más tarde " + unique,
            "seguimiento_fecha": "2026-12-20"})
        nota = row("SELECT * FROM jugador_bitacora WHERE jugador_id = %s", (jugador_id,))
        expect(response.status_code == 302 and nota["nota"].endswith(unique), "Nota financiera guardada")
        expect(str(nota["seguimiento_fecha"]) == "2026-12-20", "Fecha de seguimiento persistida")
        for page in (f'/jugadores/{jugador_id}', f'/jugadores/{jugador_id}/cuotas'):
            html = client.get(page).get_data(as_text=True)
            expect(html.count('Paga más tarde ' + unique) == 1, 'Historial financiero único en ' + page)
        resolution = post(f'/jugadores/{jugador_id}/notas-finanzas/{nota["id"]}/resolver', {'destino_notas': 'finanzas'})
        expect(resolution.location.endswith(f'/jugadores/{jugador_id}/cuotas#notas-finanzas'), 'Resolución vuelve a la pantalla financiera de origen')
        nota = row("SELECT * FROM jugador_bitacora WHERE id = %s", (nota["id"],))
        expect(nota["resuelto_en"] is not None and nota["resuelto_por"] == "admin", "Resolución con autor e historial")

        # La URL pública puede variar; se obtiene del mapa de rutas de la aplicación.
        with app.test_request_context():
            portal = client.get(sig.url_for("portal_jugador", token=jugador["portal_token"]))
        expect(portal.status_code == 200 and ("Paga más tarde " + unique).encode() not in portal.data,
               "Portal del jugador sin notas internas")

        from reportlab.pdfgen import canvas
        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer)
        pdf.drawString(40, 750, "Documento de verificacion local " + unique)
        pdf.save()
        content = buffer.getvalue()
        response = post(f"/jugadores/{jugador_id}/cuotas/nueva", {
            "periodo": sig.ahora_sig().strftime("%Y-%m"), "importe": "12345", "fecha_vencimiento": "2026-12-20"})
        cuota = row("SELECT * FROM cuotas WHERE jugador_id = %s", (jugador_id,))
        expect(response.status_code == 302 and cuota is not None, "Cuota creada desde el formulario real")
        cuota_id = cuota["id"]
        for path in (f'/jugadores/{jugador_id}', f'/jugadores/{jugador_id}/cuotas',
                     f'/jugadores/{jugador_id}/cuotas/nueva', f'/cuotas/{cuota_id}/pagar'):
            verify_forms(path, client.get(path))
        expect(not broken_forms, 'Formularios protegidos incluyen token de operación: ' + str(broken_forms))
        search = client.get('/api/buscar-jugadores', query_string={'q': 'Verificación Local ' + unique})
        expect(search.status_code == 200 and any(item['id'] == jugador_id and 'notas' in item for item in search.json['resultados']),
               'Buscador global por nombre con accesos financieros autorizados')
        for field in ('dni', 'numero_socio'):
            found = client.get('/api/buscar-jugadores', query_string={'q': jugador[field]})
            expect(any(item['id'] == jugador_id for item in found.json['resultados']), 'Buscador global encuentra por ' + field)
        note_token = secrets.token_urlsafe(24)
        followup_data = {'tipo': 'finanzas', 'nota': 'Acuerdo de verificación ' + unique,
            'seguimiento_fecha': sig.ahora_sig().date().isoformat(), '_operation_token': note_token}
        post(f'/jugadores/{jugador_id}/bitacora', followup_data)
        replay = post(f'/jugadores/{jugador_id}/bitacora', followup_data)
        expect(replay.status_code == 303 and row("SELECT COUNT(*) AS total FROM jugador_bitacora WHERE jugador_id = %s AND nota = %s",
            (jugador_id, followup_data['nota']))['total'] == 1, 'Reenvío de formulario no duplica acuerdos')
        followup = row("SELECT * FROM jugador_bitacora WHERE jugador_id = %s AND resuelto_en IS NULL", (jugador_id,))
        pay_token = secrets.token_urlsafe(24)
        response = post(f"/cuotas/{cuota_id}/pagar", {"metodo_pago": "Transferencia",
            "_operation_token": pay_token, "referencia_pago": "Verificación local " + unique, "comprobante_pago": (io.BytesIO(content), "comprobante.pdf")})
        cuota = row("SELECT * FROM cuotas WHERE id = %s", (cuota_id,))
        expect(response.status_code == 302 and cuota["pagado"] == 1, "Pago persistido")
        expect(cuota["comprobante_drive_file_id"].startswith("local-"), "Comprobante financiero guardado localmente")
        movimiento = row("SELECT id FROM movimientos WHERE concepto = %s",
            (f"Cuota {cuota['periodo']} - {jugador['apellido']}, {jugador['nombre']}",))
        expect(movimiento is not None, "Pago registrado en caja")
        replay = post(f"/cuotas/{cuota_id}/pagar", {"metodo_pago": "Transferencia", "_operation_token": pay_token,
            "referencia_pago": "Verificación local " + unique, "comprobante_pago": (io.BytesIO(content), "comprobante.pdf")})
        expect(replay.status_code == 303 and row('SELECT COUNT(*) AS total FROM movimientos WHERE concepto = %s',
            (f"Cuota {cuota['periodo']} - {jugador['apellido']}, {jugador['nombre']}",))['total'] == 1,
            'Reintentar el mismo pago con archivo conserva un solo movimiento')
        expect(row('SELECT resuelto_en FROM jugador_bitacora WHERE id = %s', (followup['id'],))['resuelto_en'] is None,
               'El pago no cierra un acuerdo automáticamente')
        queue = client.get('/finanzas/cobranzas', query_string={'vista': 'seguimientos', 'q': unique})
        expect(followup_data['nota'].encode() in queue.data and b'1 resultados' in queue.data,
               'Seguimiento sin deuda permanece en su propia cola')
        expect('Hay 1 gestión(es) abierta(s)' in client.get(response.location).get_data(as_text=True),
               'Al registrar un pago se ofrece revisar los acuerdos abiertos')
        post(f'/jugadores/{jugador_id}/notas-finanzas/{followup["id"]}/resolver', {'destino_notas': 'finanzas'})
        expect(followup_data['nota'].encode() not in client.get('/finanzas/cobranzas', query_string={'vista': 'seguimientos', 'q': unique}).data,
               'Resolver retira la gestión de la cola y conserva su historial')

        # Simultaneous submissions from the same form must make one cash entry.
        from concurrent.futures import ThreadPoolExecutor
        from threading import Barrier
        barrier = Barrier(2)
        operation = secrets.token_urlsafe(24)
        cookie_name = app.config['SESSION_COOKIE_NAME']
        cookie = client.get_cookie(cookie_name)
        with client.session_transaction() as current_session:
            csrf = current_session['_csrf_token']
        movement_data = dict(tipo='ingreso', concepto='Doble envío local ' + unique, monto='15',
            fecha=sig.ahora_sig().date().isoformat(), _csrf_token=csrf, _operation_token=operation)
        def concurrent_submit(_):
            isolated_client = app.test_client()
            isolated_client.set_cookie(cookie_name, cookie.value)
            barrier.wait(timeout=10)
            return isolated_client.post('/movimientos/nuevo', data=movement_data).status_code
        with ThreadPoolExecutor(max_workers=2) as executor:
            statuses = list(executor.map(concurrent_submit, range(2)))
        expect(302 in statuses and all(status in (302, 303, 409) for status in statuses) and
            row('SELECT COUNT(*) AS total FROM movimientos WHERE concepto = %s', (movement_data['concepto'],))['total'] == 1,
            'Dos solicitudes simultáneas registran un solo movimiento real')
        changed = post('/movimientos/nuevo', {**movement_data, 'monto': '99'})
        expect(changed.status_code == 409, 'Reutilizar una operación con otros datos se rechaza')
        correction_data = {**movement_data, '_operation_token': secrets.token_urlsafe(24),
            'concepto': 'Corrección local ' + unique, 'fecha': 'fecha inválida'}
        invalid = post('/movimientos/nuevo', correction_data)
        corrected = post('/movimientos/nuevo', {**correction_data, 'fecha': sig.ahora_sig().date().isoformat()})
        expect(invalid.status_code == 200 and corrected.status_code == 302,
            'Un formulario inválido permite corregir y guardar con la misma operación')

        from services.operations import huella_solicitud
        uncertain_data = {**movement_data, '_operation_token': secrets.token_urlsafe(24), 'concepto': 'Interrumpida local ' + unique}
        with app.test_request_context('/movimientos/nuevo', method='POST', data=uncertain_data):
            fingerprint = huella_solicitud()
        with client.session_transaction() as current_session:
            user_id = str(current_session['user_id'])
        conn = sig.get_connection()
        try:
            conn.execute('INSERT INTO operaciones_formulario (usuario, token, huella) VALUES (%s, %s, %s)',
                (user_id, uncertain_data['_operation_token'], fingerprint))
            conn.commit()
        finally:
            conn.close()
        expect(post('/movimientos/nuevo', uncertain_data).status_code == 409 and
            row('SELECT COUNT(*) AS total FROM movimientos WHERE concepto = %s', (uncertain_data['concepto'],))['total'] == 0,
            'Una operación interrumpida no se vuelve a ejecutar a ciegas')

        message_data = dict(telefono='1100000000', mensaje='Mensaje sin duplicar ' + unique,
            _operation_token=secrets.token_urlsafe(24))
        first_message = post('/comunicacion/whatsapp/responder', message_data)
        second_message = post('/comunicacion/whatsapp/responder', message_data)
        expect(first_message.status_code == 302 and second_message.status_code == 303 and
            row('SELECT COUNT(*) AS total FROM whatsapp_envios WHERE mensaje = %s', (message_data['mensaje'],))['total'] == 1,
            'Reenviar el mismo formulario produce un solo WhatsApp simulado')

        conn = sig.get_connection()
        try:
            conn.execute("""INSERT INTO cuotas (jugador_id, periodo, importe, pagado, fecha_vencimiento)
                SELECT %s, 'PRUEBA-PAGINA-' || serie::text, 1, 0, '2026-01-01'
                FROM generate_series(1, 310) serie""", (jugador_id,))
            conn.commit()
        finally:
            conn.close()
        paged = client.get('/finanzas/cobranzas', query_string={'vista': 'cobros', 'q': unique, 'pagina': 13})
        expect(b'301' in paged.data and b'310 resultados' in paged.data and b'PAGINA-' in paged.data,
               'Paginación alcanza registros posteriores al antiguo límite de 300')
        profile = client.get(f'/jugadores/{jugador_id}/cuotas').get_data(as_text=True)
        expect('pagina=13' in profile and unique in profile, 'Volver desde finanzas conserva búsqueda, vista y página')
        expect(b'310 resultados' in client.get('/finanzas/cobranzas').data, 'La cola recuerda los filtros al regresar por el menú')
        client.get('/finanzas/cobranzas?limpiar=1')
        receipt = client.get(f"/cuotas/{cuota_id}/recibo")
        expect(receipt.status_code == 200 and receipt.data.startswith(b"%PDF"), "Recibo PDF generado y descargado")
        response = post("/secretaria/documentos/nuevo", {"categoria": sig.SECRETARIA_CATEGORIAS[0],
            "titulo": "Verificación local " + unique, "archivo": (io.BytesIO(content), "prueba-local.pdf")})
        document = row("SELECT * FROM secretaria_documentos WHERE titulo = %s", ("Verificación local " + unique,))
        expect(response.status_code == 302 and document is not None, "Archivo cargado por formulario de secretaría")
        downloaded = client.get(f'/secretaria/documentos/{document["id"]}/archivo')
        expect(downloaded.status_code == 200 and downloaded.data == content, "Descarga del mismo archivo desde disco local")
        post(f'/secretaria/documentos/{document["id"]}/eliminar')

        with app.test_request_context():
            sent, reason = sig.enviar_email("prueba@example.invalid", "Verificación local " + unique, "Mensaje de prueba")
            whatsapp = sig.enviar_whatsapp_jugador(jugador, "Verificación local " + unique)
        expect(sent and reason is None, "Correo capturado sin SMTP externo")
        expect(whatsapp[0], "WhatsApp simulado con historial real")
        outbox = client.get("/__local/outbox")
        expect(unique.encode() in outbox.data, "Envíos visibles en bandeja local")
        try:
            socket.create_connection(("203.0.113.1", 443), timeout=1)
        except ConnectionRefusedError:
            expect(True, "Conexiones externas bloqueadas")
        else:
            raise AssertionError("La salida de red debería estar bloqueada")

        for role in ("tesorero", "medico", "madrinas", "secretaria", "entrenador"):
            login(role)
            home = client.get('/')
            expect(home.status_code == 200, "Panel con permisos de " + role)
            if role in {'medico', 'madrinas', 'secretaria', 'entrenador'}:
                expect('Pagos y cobros' not in home.get_data(as_text=True), 'Inicio sin información financiera para ' + role)
            if role in {'medico', 'entrenador'}:
                search = client.get('/api/buscar-jugadores', query_string={'q': unique})
                expect(search.status_code == 200 and all('finanzas' not in item and 'notas' not in item for item in search.json['resultados']),
                    'Búsqueda global sin accesos financieros para ' + role)
        response = post(f'/jugadores/{jugador_id}/notas-finanzas/{nota["id"]}/resolver')
        expect(response.status_code == 403, "Entrenador no puede resolver notas de finanzas")
    finally:
        login("admin")
        # Solo el movimiento que corresponde al jugador temporal de esta ejecución.
        conn = sig.get_connection()
        conn.execute('DELETE FROM movimientos WHERE concepto = %s', ('Doble envío local ' + unique,))
        conn.execute('DELETE FROM movimientos WHERE concepto = %s', ('Corrección local ' + unique,))
        conn.execute("DELETE FROM movimientos WHERE concepto = %s", (
            f"Cuota {sig.ahora_sig().strftime('%Y-%m')} - {jugador['apellido']}, {jugador['nombre']}",))
        conn.commit()
        conn.close()
        response = post(f"/jugadores/{jugador_id}/eliminar", {"confirmacion": f"ELIMINAR {jugador_id}"})
        expect(response.status_code == 302 and row("SELECT id FROM jugadores WHERE id = %s", (jugador_id,)) is None,
               "Limpieza del jugador creado por esta verificación")
    print(f"\n{count} verificaciones funcionales correctas contra PostgreSQL local.")
