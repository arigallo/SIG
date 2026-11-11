# Recorrido de revisión del paquete local

Abrir http://127.0.0.1:5080. El entorno mantiene guardado real en PostgreSQL local.
No hay despliegue ni conexión con producción. Los mensajes se capturan localmente.

## Inicio por rol

Probar `admin`, `tesorero`, `medico`, `entrenador`, `secretaria` y `madrinas`
con las credenciales locales. El inicio muestra solo los módulos permitidos y las
tareas asignadas al usuario. Las tareas de ejemplo indican su carácter ficticio.
Madrinas no recibe el bloque de tareas cuando no tiene ese permiso.

## Pagos y cobros

En Finanzas → Cobranzas:

1. Revisar comprobantes: abrir el comprobante de **Comprobante pendiente**.
   La acción lleva a su fila de revisión, con un PDF ficticio descargable.
2. Seguir cobros: comparar **Acuerdo vigente** (esperar la fecha pactada),
   **Seguimiento vencido** (retomar gestión) y **Comprobante rechazado**
   (pedir corrección). También hay cuotas, planes y gastos compartidos.
3. Acuerdos y seguimientos: muestra cada nota abierta aunque no haya deuda.
   Después de registrar un pago se ofrece revisar las gestiones; cerrarlas es
   una acción explícita y se conserva su historial.
4. Caja: consultar movimientos históricos de cuotas, una venta y un gasto
   ficticios; revisar la conexión con Ventas y Presupuesto.

Jugadores, colas de Finanzas, excepciones y comprobantes muestran 25 registros
por página, con cantidad total y enlaces de navegación. Se retiraron los topes
silenciosos de 300 conceptos y 150 comprobantes. Los filtros se conservan en la
sesión. Desde una ficha o cuenta financiera, **Volver al listado con mis filtros**
regresa a la vista, búsqueda y página de origen. El enlace de revisión de un
comprobante abre la página que contiene ese registro.

## Buscador global

El encabezado permite buscar por nombre, apellido, DNI o número de socio sin
salir del módulo actual. En móvil se abre con **Buscar**. Elegir la ficha o los
accesos a Finanzas y Notas internas; estos últimos solo aparecen con permisos
financieros. Enter abre el listado completo. Escape cierra los resultados y
flecha abajo permite pasar al primer resultado con el teclado.

## Protección de operaciones

Los formularios bloquean envíos repetidos mientras guardan. Pagos, movimientos,
ventas y envíos de comunicación también reservan una operación en PostgreSQL:
reenviar el mismo formulario no repite la acción, incluso con solicitudes
simultáneas. Una validación fallida permite corregir y volver a guardar.

Si una respuesta se interrumpe y su resultado es incierto, ese formulario no se
ejecuta automáticamente otra vez. Se indica revisar el registro o historial antes
de abrir una operación nueva. Las reservas no almacenan archivos ni cuerpos de
mensajes: conservan la huella de los datos y el destino de la respuesta.

## Historial único

Buscar **Demo revisión** en Jugadores. En la persona **Seguimiento vencido**:

- Abrir Finanzas en la ficha y consultar sus notas.
- Abrir la cuenta financiera: aparecen los mismos registros, con los mismos ID.
- Registrar una gestión con fecha de seguimiento; comprobar ambas pantallas.
- Resolverla; comprobar que se conserva en el historial y vuelve a la pantalla
  desde donde se realizó la acción.
- Comprobar que las notas no aparecen en el portal público del jugador.

**Historial extenso** tiene quince cuotas y una nota larga; **Beca total** y
**Beca parcial** prueban descuentos. **Plan parcialmente abonado** tiene una
cuota del plan pagada y dos pendientes. Esto representa avance del plan: no
agrega un mecanismo nuevo de pagos parciales sobre una misma cuota.

## Alertas y acciones

- **Ficha faltante** y **Ficha vencida** muestran alertas para los roles de salud.
- Las filas con varias operaciones mantienen la acción principal visible y
  agrupan las secundarias en Más acciones. Abrir el desplegable antes de probar
  operaciones de anulación o eliminación de datos ficticios.
- Una tabla sin registros muestra un mensaje breve; no se ocultan alertas por
  documentación obligatoria.

Los escenarios se agregan una sola vez y no restablecen lo que se edite después.
Se puede verificar el paquete con `./scripts/local.ps1 check` y las pruebas
unitarias descritas en `LOCAL.md`.
