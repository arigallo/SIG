from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path
pdfmetrics.registerFont(TTFont('Segoe','C:/Windows/Fonts/segoeui.ttf'))
pdfmetrics.registerFont(TTFont('SegoeB','C:/Windows/Fonts/segoeuib.ttf'))
pdfmetrics.registerFontFamily('Segoe',normal='Segoe',bold='SegoeB')
out=Path('output/pdf/instructivo_portal_del_jugador.pdf')
c=canvas.Canvas(str(out),pagesize=(595.28,841.89))
c.setTitle('Portal del jugador | Guía para nuevos jugadores')
c.setAuthor('SIG')
navy=HexColor('#10231e'); teal=HexColor('#0f766e'); gray=HexColor('#17212f')
style=ParagraphStyle('body',fontName='Segoe',fontSize=11,leading=16,textColor=gray)
y=0

def p(text,size=11,color=gray,space=12):
 global y
 st=ParagraphStyle('x',parent=style,fontSize=size,leading=size*1.45,textColor=color)
 obj=Paragraph(text,st); w,h=obj.wrap(491,700)
 obj.drawOn(c,52,y-h); y-=h+space
 assert y>53, (text,y)

def title(n,t,sub):
 global y
 c.setFillColor(navy);c.rect(0,744,596,98,fill=1,stroke=0)
 c.setFillColor(HexColor('#d6a443'));c.rect(0,741,596,3,fill=1,stroke=0)
 c.drawImage('static/img/logo.png',464,757,width=79,height=72,preserveAspectRatio=True,mask='auto')
 c.setFillColor(HexColor('#d6a443'));c.setFont('SegoeB',10);c.drawString(52,810,'SIG  /  PORTAL DEL JUGADOR')
 c.setFillColor(HexColor('#FFFFFF'));c.setFont('SegoeB',22);c.drawString(52,774,t)
 y=717;p(sub,12)
 c.setFillColor(teal);c.rect(52,49,491,1,fill=1,stroke=0)
 c.setFont('Segoe',9);c.setFillColor(gray);c.drawString(52,32,'Guía para nuevos jugadores · Septiembre 2026');c.drawRightString(543,32,f'{n} / 5')

def h(t): p('<b>'+t+'</b>',15,navy,8)
def step(n,t,body): h(f'{n}. {t}');p(body)
def note(t): p(t,11,teal,15)

title(1,'Tu primer ingreso','Tu primera tarea al ingresar al SIG: actualizar todos tus datos personales.')
h('Primero: actualizá todos tus datos personales')
p('En tu primer ingreso al SIG, bajá hasta <b>Datos personales</b> y revisá, completá o corregí <b>todos los campos</b>: nombre, apellido, DNI, fecha de nacimiento, teléfono, email, dirección, obra social y número de afiliado; también contacto familiar, parentesco, teléfono y email familiar. Tocá <b>Actualizar datos</b> y verificá el mensaje de guardado. Si todavía no entraste, seguí estos pasos para acceder.')
step(1,'Abrí el portal','Usá el enlace que te comparta el club. Si estás en la pantalla de acceso de SIG, elegí <b>Portal del Jugador</b>. También podés ingresar desde tu enlace personal, si administración te lo envió.')
step(2,'Ingresá con tu DNI','En la pantalla <b>Ingresá con tu DNI</b>, escribí solo los números y tocá <b>Entrar al portal</b>. Tu portal debe estar habilitado por administración.')
step(3,'Activá los avisos','Si aparece la pantalla de notificaciones, tocá <b>Activar notificaciones</b> y elegí <b>Permitir</b>. Al completar la activación se abrirá tu portal. En la página siguiente tenés los pasos según tu teléfono.')

note('<b>Tu enlace es personal.</b> Guardalo para volver a entrar y evitá compartirlo: permite acceder a tu información.')
h('Antes de tu primera actividad')
p('Revisá el próximo evento, el estado de tu ficha médica y si tenés cuotas pendientes. Las secciones se recorren desplazándote hacia abajo en la misma página.')
c.showPage()

title(2,'Notificaciones y agenda','Recibí comunicaciones y consultá las actividades de tu categoría.')
h('Activar notificaciones en tu teléfono')
p('<b>Android:</b> abrí el portal en Chrome o Edge, tocá <b>Activar notificaciones</b> y permití el envío cuando el navegador lo solicite. La pantalla del portal indica que no hace falta instalar SIG para activarlas.')
p('<b>iPhone o iPad:</b> seguí la ayuda que muestra el portal: abrilo en Safari, tocá <b>Compartir</b> y elegí <b>Agregar a inicio</b>. Abrí SIG desde el nuevo ícono y activá las notificaciones desde allí.')
p('Una vez dentro, buscá <b>Notificaciones del portal</b>. Usá <b>Enviar prueba</b> para comprobar que recibís el aviso en ese dispositivo.')
note('Si no podés activarlas, la <b>X</b> de la pantalla inicial permite continuar temporalmente durante 2 horas. Después, el aviso puede volver a aparecer.')
h('Consultar la agenda')
p('En <b>Calendario deportivo</b>, revisá la fecha, la hora, la ubicación y el detalle de cada evento. Prestá atención a la convocatoria y a la fecha límite para confirmar, si aparecen.')
p('En la parte superior tenés los accesos <b>iPhone</b>, <b>Android</b>, <b>Google Calendar</b> y <b>Descargar ICS</b>. Elegí el que corresponda y seguí las indicaciones de tu calendario. Antes de asistir, revisá también el evento en el portal.')
h('Leer las comunicaciones')
p('Consultá <b>Comunicaciones del día</b> y los avisos de <b>Notificaciones</b> cuando aparezcan. Si una comunicación incluye un enlace, tocá <b>Abrir</b> para ver el contenido.')
c.showPage()

title(3,'Asistencia y bienestar','Confirmá cada actividad desde su tarjeta en el calendario deportivo.')
step(1,'Buscá el evento correcto','Verificá el título y la fecha. Si el evento admite confirmación, vas a encontrar un selector con las opciones disponibles.')
step(2,'Elegí tu respuesta y guardala','Para un <b>partido</b>, elegí tu opción y tocá <b>Guardar confirmacion</b>. Los partidos no requieren cuestionario de bienestar.<br/><br/>Para un <b>entrenamiento u otra actividad</b>, elegí tu opción y tocá <b>Continuar</b>. Si indicás que no asistís, no se requiere bienestar. Las demás respuestas te llevan al cuestionario.')
step(3,'Completá el bienestar cuando se solicite','Respondé sobre calidad y horas de sueño, dolor muscular, fatiga, estrés, ánimo, motivación y recuperación. Leé las etiquetas de cada escala: describen qué significa cada valor.<br/><br/>Marcá las zonas con dolor o molestias, si corresponde, y agregá comentarios que el cuerpo técnico deba conocer.')
step(4,'Finalizá el envío','Tocá <b>Guardar asistencia y bienestar</b>. Al volver al portal, verificá la confirmación y el resumen de bienestar del evento. Si figura <b>Falta cuestionario de bienestar</b>, ingresá nuevamente y completalo.')
note('<b>No alcanza con abrir el cuestionario.</b> Para registrar la respuesta tenés que terminar de completarlo y guardar.')
p('El <b>Historial de asistencia</b> muestra registros como A tiempo, Tarde o Ausente. Consultalo para revisar tus actividades anteriores.')
c.showPage()

title(4,'Cuotas y comprobantes','Consultá tu deuda y enviá el comprobante del pago correspondiente.')
step(1,'Identificá qué estás abonando','Revisá <b>Deuda actual</b> y <b>Cuenta corriente</b>. Después ubicá el período en <b>Cuotas y comprobantes</b> o el concepto en <b>Gastos compartidos / terceros tiempos</b>. La deuda actual ya incluye cuotas y gastos compartidos.')
step(2,'Seleccioná el archivo','Una vez realizado el pago por el medio indicado por el club, buscá la fila correspondiente. Tocá el área para adjuntar el comprobante y elegí un archivo <b>PDF, JPG, JPEG o PNG</b>. En computadora también podés arrastrarlo o pegarlo. Para las cuotas, podés completar el campo <b>Referencia</b>.')
step(3,'Enviá y comprobá el estado','Tocá <b>Enviar</b>. Revisá que aparezca el nombre del archivo y el estado <b>En revisión</b>. Podés usar <b>Ver comprobante</b> para comprobar lo que enviaste.')
h('Qué significa cada estado')
p('<b>Pendiente:</b> el pago todavía no figura registrado.<br/><b>En revisión:</b> el comprobante fue recibido y falta su validación.<br/><b>Rechazado:</b> leé las observaciones, elegí el archivo correcto y tocá <b>Reemplazar</b>.<br/><b>Pagada / Pagado:</b> el pago ya está registrado. En cuotas, el botón <b>Recibo</b> permite descargarlo.')
note('<b>Enviar un comprobante no acredita el pago automáticamente.</b> La deuda puede seguir visible mientras administración lo revisa.')
p('Si un gasto compartido figura <b>Cerrado con saldo pendiente</b>, podés enviar el comprobante para regularizarlo. Los planes de pago se consultan en <b>Planes de pago</b>; para acordar o corregir uno, contactá a administración.')
c.showPage()

title(5,'Tu información y ayuda','Mantené tus datos al día y consultá tus registros cuando lo necesites.')
h('Salud, documentos y evolución')
p('<b>Ficha médica:</b> revisá si está presentada, su vencimiento y el apto físico. En <b>Salud y disponibilidad</b> podés consultar lesiones y recuperación registradas. Para entregar documentación o corregir estos registros, contactá al club.')
p('<b>Documentos:</b> tocá <b>Abrir</b> cuando haya un archivo disponible. Usá <b>Constancia PDF</b>, arriba o en <b>Datos del club</b>, para descargar tu constancia.')
p('<b>Evolución de tests:</b> consultá puntajes, fechas y observaciones. En <b>Análisis de tests</b>, elegí el test y, si querés, un rango Desde/Hasta; tocá <b>Actualizar grafico</b>. <b>Limpiar fechas</b> quita el filtro de fechas. Leé si el test indica que un valor mayor o menor es mejor.')
h('Si algo no funciona')
p('<b>No encuentra tu DNI:</b> revisá los números. Si el mensaje indica que no hay un portal activo o que hay más de uno, pedí a administración que revise tu registro y habilitación.')
p('<b>El enlace no abre:</b> solicitá a administración que verifique que tu portal siga activo y te comparta el acceso vigente.')
p('<b>No recibís avisos:</b> revisá los permisos del navegador o teléfono y volvé a abrir el portal. Seguí la ayuda de activación que aparece en pantalla y usá <b>Enviar prueba</b>.')
p('<b>El pago sigue pendiente o un dato está mal:</b> revisá el estado del comprobante y las observaciones. Contactá a administración indicando el período o concepto; para datos deportivos, consultá al cuerpo técnico.')
note('<b>Tu rutina:</b> leé los avisos, revisá el próximo evento, confirmá asistencia y bienestar cuando corresponda, y controlá pagos y vencimientos.')
c.save()
print(out.resolve())
