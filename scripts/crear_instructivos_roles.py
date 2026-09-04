"""Genera las guías de los roles base. Ejecutar desde la raíz con reportlab."""
from pathlib import Path
import shutil
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

ROOT = Path(__file__).resolve().parents[1]
FONT = Path('C:/Windows/Fonts')
if (FONT / 'segoeui.ttf').exists():
    pdfmetrics.registerFont(TTFont('Guia', str(FONT / 'segoeui.ttf')))
    pdfmetrics.registerFont(TTFont('GuiaBold', str(FONT / 'segoeuib.ttf')))
    pdfmetrics.registerFontFamily('Guia', normal='Guia', bold='GuiaBold')
    NORMAL, BOLD = 'Guia', 'GuiaBold'
else:
    NORMAL, BOLD = 'Helvetica', 'Helvetica-Bold'

# Rutas y acciones contrastadas con ROLE_PRESETS y las plantillas de SIG.
GUIAS = {
    'madrinas': ('Madrinas', 'Acompañá a tus ahijadxs durante su incorporación al club.', [
        ('Asignación y seguimiento', [
            ('1. Identificar tu acceso', 'Si tenés una cuenta de gestión habilitada, ingresá con tu usuario y clave y abrí <b>Club > Ahijadxs</b>. Si solo accedés al portal del jugador o no aparece esa sección, pedí a administración el acceso o la información necesaria para el acompañamiento.'),
            ('2. Encontrar a tus ahijadxs', 'En el listado de <b>Ahijadxs</b>, buscá por nombre, DNI, categoría o madrina. Abrí el detalle de la persona y verificá el apartado <b>Madrina/Padrino</b>. La búsqueda ayuda a encontrar las asignaciones; no supone que tu cuenta vea únicamente tus ahijadxs.'),
            ('3. Verificar la asignación', 'Revisá que el nombre y la categoría de la madrina/padrino sean correctos. Si tenés permiso de gestión y el registro está en seguimiento, entrá en <b>Editar</b>, seleccioná <b>Madrina/Padrino</b> y tocá <b>Guardar</b>. Si no aparece la edición, solicitá la corrección a quien gestione ahijadxs.'),
            ('4. Revisar los datos de contacto', 'En el detalle, consultá teléfono, email, categoría y fecha de postulación para organizar el acompañamiento. Si faltan datos, coordiná su actualización. En la edición también están disponibles las observaciones y los <b>Entrenamientos requeridos</b>; este objetivo debe reflejar lo acordado por el club.'),
            ('Sobre los permisos', 'La asignación como madrina/padrino en un registro no otorga por sí sola acceso al sistema. Consultar ahijadxs requiere permiso de consulta y modificarlos requiere permiso de gestión. Esta guía describe el circuito disponible y distingue las acciones que dependen de esos accesos.'),
        ]),
        ('Entrenamientos e ingreso', [
            ('5. Consultar el progreso', 'En el detalle del ahijadx, revisá <b>Asistencias</b>, <b>Progreso</b> e <b>Historial de asistencia</b>. Para el objetivo de ingreso solo cuentan eventos de tipo <b>Entrenamiento</b>. Revisá las fechas y los estados del historial antes de consultar por una diferencia en el total.'),
            ('6. Coordinar la asistencia', 'Coordiná con quien toma asistencia que el ahijadx figure en el evento correcto. Si tu cuenta también tiene permisos de asistencia, abrí <b>Club > Asistencia</b>, registrá el estado observado y usá <b>Guardar asistencia</b>. Si no tenés ese acceso, solicitá la carga o corrección al responsable.'),
            ('7. Preparar el ingreso al club', 'Cuando se alcance el objetivo, el detalle puede mostrar <b>Listx para ingresar</b>. Revisá identidad, DNI y datos de contacto y coordiná el ingreso con el responsable. Si faltan entrenamientos, continuá el seguimiento; no reduzcas el objetivo solo para habilitar la conversión.'),
            ('8. Confirmar la incorporación', 'Con permiso de gestión y un ahijadx en seguimiento que cumpla el objetivo, aparece <b>Ingresar como jugador</b>. La persona responsable debe revisar la <b>Fecha de ingreso</b> y confirmar. Después, comprobá el estado <b>Ingresado</b> y la fecha. Si aparece un aviso de DNI existente, pedí revisar el padrón antes de repetir.'),
            ('Rutina de acompañamiento', 'Revisá tus asignaciones y los próximos pasos con cada ahijadx. Después de cada entrenamiento, consultá el avance y derivá las diferencias al responsable de asistencia. Ante una desvinculación, coordiná la actualización con quien gestiona el registro; <b>Dar de baja</b> cambia su estado a Baja.'),
        ]),
    ]),
    'admin': ('Administrador', 'Administrá los accesos y coordiná la operación del club.', [
        ('Accesos y configuración', [
            ('1. Ingresar y revisar el panel', 'Ingresá con tu usuario y clave. En el menú de tu cuenta, usá <b>Mi clave</b> para cambiarla. Revisá <b>Panel</b> y <b>Operación</b> para identificar pendientes. El rol admin tiene todos los permisos base; usá cada módulo según la tarea que debas resolver.'),
            ('2. Crear o actualizar usuarios', 'Abrí <b>Admin > Usuarios</b> y elegí el alta o la edición. Completá usuario, email de recuperación y rol; para un alta, definí la clave inicial. Usá <b>Crear usuario</b> o <b>Guardar usuario</b> y verificá el registro en el listado. Asigná el rol que corresponda a las tareas de esa persona.'),
            ('3. Administrar los roles', 'Entrá en <b>Admin > Roles</b>. Usá <b>Nuevo rol</b> para crear uno o <b>Editar</b> para revisar sus permisos. Guardá y comprobá los permisos seleccionados. Los roles base también pueden modificarse: un cambio afecta el alcance de las cuentas que utilizan ese rol.'),
            ('4. Configurar el funcionamiento', 'En <b>Admin > Sistema</b>, revisá las opciones de automatización y su estado antes de habilitarlas. En <b>Email facturas</b>, configurá la casilla para importar facturas. Usá <b>Avisos login</b> para los avisos de acceso y <b>Mantenimiento</b> cuando necesites restringir temporalmente el uso del sistema.'),
        ]),
        ('Supervisión y tareas del club', [
            ('5. Mantener el padrón y el portal', 'En <b>Club > Jugadores</b>, buscá primero por los datos de la persona para evitar duplicados. Creá o editá el registro y verificá los datos guardados. Desde las acciones del jugador, administrá sus enlaces de portal cuando corresponda. Compartí cada enlace únicamente con su destinatario.'),
            ('6. Coordinar las áreas', 'Revisá <b>Finanzas</b> para cuotas, comprobantes, caja, ventas y presupuesto; <b>Salud</b> para fichas, documentos y lesiones; y <b>Club</b> para calendario, asistencia y tests. Para la carga cotidiana, seguí la guía específica de tesorería, medicina o entrenamiento.'),
            ('7. Gestionar documentación y devoluciones', 'En <b>Secretaría > Documentos</b>, cargá y categorizá los archivos administrativos. En <b>Admin > Sugerencias</b>, revisá los envíos y actualizá su seguimiento. En <b>Admin > Encuestas</b>, administrá campañas y consultá resultados. Verificá destinatarios y contenido antes de enviar comunicaciones.'),
            ('8. Revisar actividad y respaldo', 'Consultá <b>Admin > Auditoría</b> y <b>Seguridad</b> ante cambios o accesos que requieran revisión. En <b>Backup</b>, consultá la información disponible; los backups manuales se solicitan desde <b>Sistema</b>. Verificá el resultado informado, sin dar por completado un respaldo solo por haberlo solicitado.'),
            ('Cierre de la jornada', 'Confirmá que las altas y cambios quedaron guardados, derivá los pendientes al área responsable y cerrá sesión desde <b>Salir</b>. Antes de eliminar datos, revisá que la baja definitiva sea la acción necesaria.'),
        ]),
    ]),
    'tesorero': ('Tesorero', 'Registrá cobros y gastos, revisá comprobantes y controlá la caja.', [
        ('Cuotas y cobranzas', [
            ('1. Revisar los pendientes', 'Ingresá con tu usuario y clave. Abrí <b>Finanzas > Cobranzas</b> y revisá el estado de deuda. Consultá al jugador en <b>Club > Jugadores</b> para verificar su identidad y su cuenta corriente antes de imputar un cobro.'),
            ('2. Generar cuotas', 'Entrá en <b>Finanzas > Generar cuotas</b>. Completá las opciones que solicita el formulario y revisá período, importe y destinatarios antes de confirmar. Después, consultá las cuotas generadas. Si el resultado es incierto, verificá el listado antes de repetir la operación.'),
            ('3. Registrar un pago', 'Abrí las cuotas del jugador y seleccioná la que corresponde. Completá el formulario de pago y comprobante, según corresponda. Guardá y verificá el estado de la cuota y el ingreso asociado en caja. Para cobros por lote, utilizá <b>Débitos automáticos</b> y revisá la selección antes de confirmar.'),
            ('4. Revisar comprobantes del portal', 'Abrí <b>Finanzas > Comprobantes</b> y filtrá <b>Pendientes</b>. Usá <b>Ver</b> para examinar el archivo y contrastá persona, importe y pago recibido. Elegí <b>Aceptar</b> o <b>Rechazar</b> según el resultado. Un archivo subido requiere revisión; verificá el estado final luego de la acción.'),
            ('5. Conciliar y acordar regularizaciones', 'En <b>Conciliación</b>, revisá las opciones y coincidencias antes de confirmar imputaciones. Usá <b>Planes</b> para registrar y seguir acuerdos de pago. Consultá nuevamente la cuenta corriente para comprobar qué deuda continúa pendiente.'),
        ]),
        ('Caja y control financiero', [
            ('6. Registrar ingresos y egresos', 'Entrá en <b>Finanzas > Caja</b>. Cargá el movimiento con fecha, tipo, importe y detalle que correspondan; adjuntá el respaldo cuando aplique. Revisá el listado y el saldo después de guardar. Antes de cargar un ingreso, comprobá si ya fue generado por un cobro o una venta.'),
            ('7. Procesar facturas recibidas', 'En <b>Facturas recibidas</b>, sincronizá y revisá los archivos detectados. Contrastá proveedor, importe y fecha antes de registrarlos como egreso. Verificá el movimiento en caja. Si falla la conexión de email, pedí al administrador que revise <b>Admin > Email facturas</b>.'),
            ('8. Administrar gastos, ventas y presupuesto', 'En <b>Gastos compartidos / terceros tiempos</b>, revisá participantes, distribución y cobros. Cerrar un gasto no elimina las deudas pendientes. En <b>Ventas del club</b>, controlá producto y stock antes de registrar ventas. En <b>Presupuesto</b>, mantené las previsiones de ingresos y gastos.'),
            ('9. Controlar y comunicar', 'Consultá <b>Reportes</b> y <b>Salud > Alertas</b> para los avisos financieros disponibles. Usá <b>Comunicación</b> para revisar destinatarios y preparar avisos de deuda. Verificá importes y destinatarios antes del envío. El rol también permite gestionar calendario y asistencia desde <b>Club</b>.'),
            ('Antes de cerrar el mes', 'Contrastá cobros, egresos y respaldos; resolvé diferencias antes del cierre en caja. El rol base permite consultar jugadores, pero no editar su padrón ni administrar usuarios. Para esos cambios, contactá a administración.'),
        ]),
    ]),
    'medico': ('Médico', 'Mantené las fichas médicas, los documentos y el seguimiento de lesiones.', [
        ('Fichas y documentación', [
            ('1. Identificar prioridades', 'Ingresá con tu usuario y clave y abrí <b>Salud > Panel salud</b>. Revisá fichas vencidas, faltantes y lesiones activas. Consultá <b>Salud > Alertas</b> para organizar el seguimiento. Corroborá siempre la identidad del jugador antes de modificar su información.'),
            ('2. Abrir la ficha del jugador', 'En <b>Club > Jugadores</b>, buscá a la persona y abrí su ficha médica desde el detalle. Revisá la información existente y el documento asociado antes de editar. El rol base permite consultar el padrón; los cambios generales del jugador deben solicitarse a administración.'),
            ('3. Actualizar la ficha médica', 'En la edición, revisá los indicadores del formulario, la <b>Fecha de vencimiento</b>, el contacto y teléfono de emergencia y las observaciones. Adjuntá el <b>Documento BDUAR / ficha escaneada</b> si corresponde. Usá <b>Guardar</b> y volvé a abrir la ficha para comprobar datos y archivo.'),
            ('4. Controlar documentos', 'Entrá en <b>Salud > Documentos</b> para revisar vencimientos y faltantes. Para una carga o edición manual, seleccioná el jugador correcto y completá los datos del documento. Guardá y verificá que el archivo se abra y la fecha registrada corresponda al respaldo.'),
        ]),
        ('Carga por lote y seguimiento', [
            ('5. Cargar varias fichas', 'Abrí <b>Salud > Fichas batch</b> y cargá los archivos admitidos por la pantalla. Revisá la propuesta de asociación y los datos en la etapa de revisión antes de confirmar. Si no podés identificar a una persona con certeza, resolvé esa asociación antes de incorporarla a su ficha.'),
            ('6. Registrar y actualizar lesiones', 'Abrí las lesiones desde el jugador. Creá el registro con los campos solicitados, o editá uno existente para actualizar su seguimiento. Guardá y verificá que la lesión corresponda a la persona correcta y que su estado refleje la información disponible.'),
            ('7. Comprobar el resultado', 'Volvé al <b>Panel salud</b> y revisá la ficha o lesión modificada. Si sigue apareciendo un vencimiento, comprobá la fecha guardada y el documento respaldatorio. No reemplaces fechas solo para quitar una alerta: registrá lo que indique la documentación.'),
            ('Rutina sugerida', 'Al inicio, revisá alertas y priorizá pendientes. Durante la jornada, registrá documentación y seguimiento. Al finalizar, confirmá que los cambios quedaron guardados y comunicá al área responsable qué documentación falta, evitando difundir información clínica innecesaria.'),
            ('Alcance y consultas', 'Este rol administra salud y documentos; no incluye caja, cuotas, tests, asistencia ni usuarios. Si falta una opción necesaria, solicitá que administración revise tus permisos. Para salir, abrí el menú de tu cuenta y elegí <b>Salir</b>.'),
        ]),
    ]),
    'entrenador': ('Entrenador', 'Organizá el plantel, las actividades, la asistencia y las mediciones.', [
        ('Plantel y agenda', [
            ('1. Revisar el plantel', 'Ingresá con tu usuario y clave y abrí <b>Club > Jugadores</b>. Buscá al jugador antes de darlo de alta. Para incorporar una persona, usá <b>Nuevo jugador</b>; para corregir datos, abrí su edición. Guardá y comprobá categoría y datos de contacto en el detalle.'),
            ('2. Importar jugadores', 'Entrá en <b>Club > Importar</b> y seguí las indicaciones de formato de la pantalla. Revisá el resultado de la importación y buscá los registros incorporados. Si hubo errores o el resultado es incierto, verificá los jugadores existentes antes de volver a cargar el archivo.'),
            ('3. Dar seguimiento a ahijadxs', 'En <b>Club > Ahijadxs</b>, creá o editá el registro y completá su seguimiento. Registrá la asistencia correspondiente y revisá la información antes de convertirlo a jugador activo. Después de la conversión, comprobá el alta en el listado de jugadores.'),
            ('4. Organizar el calendario', 'Abrí <b>Club > Calendario</b> y creá el evento. Completá fecha, horario, lugar y demás opciones disponibles. Revisá categoría y convocatoria cuando correspondan. Guardá y comprobá que la actividad aparezca correctamente en el calendario.'),
        ]),
        ('Asistencia y tests', [
            ('5. Preparar la asistencia', 'En <b>Club > Asistencia</b>, abrí el evento correspondiente o creá uno si falta. Confirmá fecha, categoría y participantes. Revisá las <b>Respuestas portal</b> y el bienestar disponible para preparar la actividad; después registrá la asistencia efectivamente observada.'),
            ('6. Guardar y cerrar el evento', 'Marcá el estado de cada participante y agregá observaciones cuando corresponda. Tocá <b>Guardar asistencia</b> y verificá los registros. Al terminar la revisión, usá <b>Cerrar evento</b>. Si necesitás corregir uno cerrado, utilizá <b>Reabrir evento</b>, realizá los cambios y guardá nuevamente.'),
            ('7. Cargar mediciones deportivas', 'Abrí <b>Club > Tests</b>, seleccioná el test y entrá a la carga. Revisá la fecha y la unidad de medida antes de ingresar resultados por jugador. Usá <b>Guardar mediciones</b> y comprobá los valores. Para importar, revisá la propuesta antes de confirmar las mediciones.'),
            ('8. Consultar la evolución', 'Desde <b>Tests</b>, consultá los gráficos, rankings y mediciones disponibles. Compará resultados del mismo test y revisá fechas y unidades. Si un dato parece incorrecto, contrastalo con el registro original antes de usarlo para planificar la actividad.'),
            ('Cierre y límites del rol', 'Verificá asistencia guardada, agenda y mediciones pendientes. El entrenador puede crear y editar jugadores, pero no eliminarlos definitivamente. El rol base no habilita la gestión médica, las finanzas ni los usuarios; derivá esos cambios al área correspondiente. Cerrá sesión desde <b>Salir</b>.'),
        ]),
    ]),
}

def generate():
    for key, (name, subtitle, pages) in GUIAS.items():
        out = ROOT / 'output/pdf' / f'instructivo_{key}.pdf'
        out.parent.mkdir(parents=True, exist_ok=True)
        pdf = canvas.Canvas(str(out), pagesize=(595.28, 841.89))
        pdf.setTitle(f'SIG | Instructivo para {name.lower()}')
        pdf.setAuthor('SIG')
        for number, (heading, sections) in enumerate(pages, 1):
            pdf.setFillColor(HexColor('#10231e'))
            pdf.rect(0, 735, 596, 107, fill=1, stroke=0)
            pdf.setFillColor(HexColor('#d6a443'))
            pdf.setFont(BOLD, 10)
            pdf.drawString(44, 810, f'SIG / GUÍA DE USO / {name.upper()}')
            pdf.setFillColor(HexColor('#ffffff'))
            pdf.setFont(BOLD, 21)
            pdf.drawString(44, 774, heading)
            y = 714
            def paragraph(text, size=10.5, bold=False, color='#17212f', gap=9):
                nonlocal y
                style = ParagraphStyle('body', fontName=BOLD if bold else NORMAL,
                    fontSize=size, leading=size * 1.42, textColor=HexColor(color))
                obj = Paragraph(text, style)
                _, height = obj.wrap(507, 700)
                y -= height
                if y < 64:
                    raise ValueError(f'Desborde: {key}, página {number}: {text}')
                obj.drawOn(pdf, 44, y)
                y -= gap
            paragraph(subtitle, 11, color='#0f766e', gap=12)
            for title, body in sections:
                paragraph(title, 12, bold=True, gap=5)
                paragraph(body, gap=12)
            if number == 1 and key != 'madrinas':
                paragraph('Esta guía corresponde a los permisos base del rol. Si el club los modificó o creó roles personalizados, las opciones visibles pueden variar. Pedí a administración que confirme tu acceso.', 9, color='#0f766e')
            pdf.setStrokeColor(HexColor('#0f766e'))
            pdf.line(44, 49, 551, 49)
            pdf.setFillColor(HexColor('#17212f'))
            pdf.setFont(NORMAL, 8)
            pdf.drawString(44, 33, 'SIG · Guías por función · Septiembre 2026')
            pdf.drawRightString(551, 33, f'{number} / {len(pages)}')
            pdf.showPage()
        pdf.save()
        target = ROOT / 'static/docs' / out.name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(out, target)
        print(out)

if __name__ == '__main__':
    generate()
