"""Lectura de avisos BBVA; nunca contabiliza pagos ni envía mensajes."""
import hashlib
import html
import re
from datetime import date, timedelta
from decimal import Decimal
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr


def crear_tablas(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS bbva_remitentes (
        cuit TEXT NOT NULL, cbu TEXT NOT NULL, jugador_id INTEGER NOT NULL REFERENCES jugadores(id),
        PRIMARY KEY (cuit, cbu))""")
    conn.execute("""CREATE TABLE IF NOT EXISTS bbva_transferencias (
        id SERIAL PRIMARY KEY, mensaje_clave TEXT NOT NULL UNIQUE, fecha DATE NOT NULL,
        cuenta TEXT NOT NULL, cuit TEXT NOT NULL, cbu TEXT NOT NULL, importe NUMERIC(14,2) NOT NULL,
        jugador_id INTEGER REFERENCES jugadores(id), estado TEXT NOT NULL DEFAULT 'pendiente',
        creado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
        solicitado_en TIMESTAMPTZ, solicitado_por TEXT)""")


def parsear_aviso(raw):
    mensaje = BytesParser(policy=policy.default).parsebytes(raw)
    if parseaddr(str(mensaje.get('From', '')))[1].lower() != 'avisos@bbva.com.ar':
        return None
    if 'AVISO TRANSFERENCIA INMEDIATA ACREDITADA' not in str(mensaje.get('Subject', '')).upper():
        return None
    textos = []
    for parte in mensaje.walk():
        if parte.get_content_type() in ('text/plain', 'text/html') and parte.get_content_disposition() != 'attachment':
            contenido = parte.get_content()
            contenido = re.sub(r'<(script|style)\b[^>]*>.*?</\1>', '', contenido, flags=re.S|re.I)
            textos.append(html.unescape(re.sub(r'<[^>]+>', ' ', contenido)))
    texto = re.sub(r'\s+', ' ', ' '.join(textos))
    patrones = {
        'fecha': r'Fecha\s+(\d{2}/\d{2}/\d{4})',
        'cuenta': r'N[uú]mero cuenta\s+(.+?)\s+CBU origen',
        'cbu': r'CBU origen\s+(\d{22})',
        'cuit': r'CUIT origen\s+(\d{11})',
        'importe': r'Importe\s+ARS\s+([\d.]+,\d{2})(?!\d)',
    }
    datos = {}
    for campo, patron in patrones.items():
        encontrado = re.search(patron, texto, re.I)
        if not encontrado:
            raise ValueError('Aviso BBVA incompleto: ' + campo)
        datos[campo] = encontrado.group(1).strip()
    dia, mes, anio = map(int, datos['fecha'].split('/'))
    datos['fecha'] = date(anio, mes, dia)
    datos['importe'] = Decimal(datos['importe'].replace('.', '').replace(',', '.'))
    if datos['importe'] <= 0:
        raise ValueError('Importe BBVA inválido')
    identidad = str(mensaje.get('Message-ID') or '').strip() or hashlib.sha256(raw).hexdigest()
    datos['mensaje_clave'] = hashlib.sha256(identidad.encode()).hexdigest()
    return datos


def sugerir_jugador(datos, jugadores, vinculos):
    for vinculo in vinculos:
        if (vinculo['cuit'], vinculo['cbu']) == (datos['cuit'], datos['cbu']):
            return vinculo['jugador_id'], 'Remitente vinculado previamente'
    dni = datos['cuit'][2:10].lstrip('0')
    candidatos = [j for j in jugadores if re.sub(r'\D', '', j.get('dni') or '').lstrip('0') == dni]
    if len(candidatos) == 1:
        return candidatos[0]['id'], 'Coincidencia de DNI con CUIT de origen; confirmar'
    return None, 'Sin coincidencia única; seleccionar jugador'


def sincronizar(conn, configs, imaplib):
    nuevas = 0
    errores = 0
    cuentas = [c for c in configs if c['user'].lower() == 'tesoreria@rudamachorugby.com']
    if not cuentas:
        raise RuntimeError('La casilla de tesorería no está configurada.')
    for config in cuentas:
        cliente = None
        try:
            clase = imaplib.IMAP4_SSL if config['use_ssl'] else imaplib.IMAP4
            cliente = clase(config['host'], config['port'], timeout=30)
            cliente.login(config['user'], config['password'])
            estado, _ = cliente.select(config['folder'], readonly=True)
            if estado != 'OK':
                raise RuntimeError('No se pudo abrir la casilla de tesorería.')
            desde = (date.today()-timedelta(days=config['search_days'])).strftime('%d-%b-%Y')
            estado, datos = cliente.search(None, 'SINCE', desde, 'FROM', '"avisos@bbva.com.ar"', 'SUBJECT', '"AVISO TRANSFERENCIA INMEDIATA ACREDITADA"')
            if estado != 'OK':
                raise RuntimeError('No se pudieron buscar los avisos BBVA.')
            # No limitar a los últimos mensajes: evitar perder ingresos en días de mucho movimiento.
            for uid in (datos[0] or b'').split():
                estado, partes = cliente.fetch(uid, '(BODY.PEEK[])')
                if estado != 'OK':
                    errores += 1
                    continue
                raw = next((p[1] for p in partes if isinstance(p, tuple)), None)
                if not raw:
                    errores += 1
                    continue
                try:
                    aviso = parsear_aviso(raw)
                except (ValueError, TypeError, UnicodeError):
                    errores += 1
                    continue
                if aviso:
                    fila = conn.execute("""INSERT INTO bbva_transferencias
                        (mensaje_clave, fecha, cuenta, cuit, cbu, importe)
                        VALUES (%s,%s,%s,%s,%s,%s) ON CONFLICT (mensaje_clave) DO NOTHING RETURNING id""",
                        tuple(aviso[k] for k in ('mensaje_clave','fecha','cuenta','cuit','cbu','importe'))).fetchone()
                    nuevas += bool(fila)
        finally:
            if cliente:
                try:
                    cliente.logout()
                except Exception:
                    pass
    return {'nuevas': nuevas, 'errores': errores}
