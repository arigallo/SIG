# Preparación de publicación — revisión aprobada en local

El usuario completó la revisión funcional local. Las comprobaciones iniciales
fueron de solo lectura. El 11/11/2026 autorizó commit y push para desplegar una
revisión nueva sin tráfico, manteniendo la revisión actual en servicio.

## Estrategia autorizada para este despliegue

- El usuario fijó el tráfico al 100 % en `club-jugadores-00211-sg6`; se verificó
  que ya no apunta automáticamente a `LATEST`.
- `cloudbuild.json` conserva los pasos del trigger y agrega `--no-traffic` y
  `--tag=revision-candidata` al paso Deploy. Antes del push se debe cargar y
  verificar esa misma configuración en el trigger.
- El push a `main` construye y despliega la candidata; el pase de tráfico
  requiere una decisión posterior y no forma parte de este despliegue.
- La candidata comparte Cloud SQL y las integraciones del servicio actual.
  No es una preproducción aislada. Al arrancar aplica el `init_db()` habitual,
  incluidas las nuevas columnas de bitácora y la tabla de operaciones. Los
  envíos o pagos realizados por su URL actúan sobre los datos reales.
- Se deben comprobar el estado del build, la revisión lista y los porcentajes
  de tráfico después del push. No ejecutar allí el verificador local ni cargar
  los escenarios ficticios.

## Entorno identificado por consultas de solo lectura

- Proyecto: `rudamacho-app-publica-2026`.
- Cloud Run: `club-jugadores`, región `us-central1`.
- Revisión que sirve el 100 %: `club-jugadores-00211-sg6`.
- Commit desplegado y base del checkout: `c5cc757dab2ad52f7da3de9d4a5917e8a698b261`.
- Cloud SQL: `sig-db-us`, PostgreSQL 15, estado `RUNNABLE`.
- Respaldos automáticos habilitados, retención de 14 respaldos y recuperación
  a un momento específico habilitada con 7 días de registros de transacciones.
- Se encontró un respaldo bajo demanda `SUCCESSFUL`, ID `1794364952872`,
  terminado el 11/11/2026 a las 02:44:04 UTC (10/11, 23:44 en Buenos Aires).
  Confirmar vigencia antes de publicar; consultar un respaldo exitoso no prueba
  que su restauración ya haya sido ensayada.
- El trigger de Cloud Build `dd6897c1-9ce1-4a0f-84db-9d070725f93f` responde a
  push sobre `^main$` en `arigallo/SIG`. No enviar cambios a esa rama durante
  la preparación: hacerlo inicia la publicación automática.

## Entrega cerrada

`scripts/prepare_release.py` empaqueta los archivos de aplicación, recursos y
pruebas del checkout actual; agrega los archivos nuevos necesarios y genera un
manifiesto SHA-256. Extrae esa misma entrega y ejecuta sus pruebas unitarias,
sin importar la aplicación con una base productiva. No hace commit, push,
build remoto ni despliegue.

El resultado queda en `.local/releases/`, excluido de Git y del envío a Cloud.
Se excluyen los adaptadores locales, escenarios ficticios, credenciales,
respaldos, archivos temporales y exportaciones. El ZIP es un candidato de
fuentes; todavía no es una imagen validada en Cloud Run.

## Comprobaciones pendientes antes del pase de tráfico

1. Ensayar restauración y actualización con una copia aislada de la base real
   usando PostgreSQL 15. El entorno funcional local usa PostgreSQL 17. La nueva
   versión agrega `seguimiento_fecha`, `resuelto_en` y `resuelto_por` a
   `jugador_bitacora`, y la tabla `operaciones_formulario`. Son adiciones;
   comprobar también que el `init_db()` completo de arranque termina sin errores.
2. Construir y probar el contenedor de la entrega, con base y archivos de prueba
   separados. Probar Drive, correo y WhatsApp con destinatarios de prueba
   explícitos; las verificaciones locales simulan esas integraciones.
3. Confirmar respaldo reciente y revisión de retorno. Una revisión nueva sin
   tráfico que apunte a la base productiva también puede ejecutar `init_db()`;
   no sirve como entorno de ensayo aislado.
4. Publicar únicamente la entrega validada después de la autorización final;
   comprobar acceso, permisos, búsqueda, pagos, notas, archivos y comunicaciones.

## Retorno preparado, sin ejecutar

La revisión anterior puede recibir el tráfico nuevamente si falla la nueva.
Verificar que sigue disponible y comprobar compatibilidad del esquema antes de
usar el comando. Este cambio de tráfico no deshace movimientos de negocio ni
datos nuevos y no restaura una base de datos:

```powershell
gcloud run services update-traffic club-jugadores --project=rudamacho-app-publica-2026 --region=us-central1 --to-revisions=club-jugadores-00211-sg6=100
```

No borrar columnas, tablas o notas como parte de un retorno del código. Si el
problema afecta los datos, evaluar la restauración por separado.
