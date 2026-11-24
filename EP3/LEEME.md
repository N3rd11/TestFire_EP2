# TestFire EP3 en Rocketbot

## Cómo saber que terminó

Para ejecutar el robot principal, abra **EJECUTAR_EP3.bat**. Se abre Edge con Rocketbot Studio y se marcan los comandos de la ejecución. Studio muestra **FINALIZADO** al pie de la ventana. La consola registra las etapas, el tiempo del robot, los conteos, el resultado del correo y la carpeta del Excel. Cierre esa ventana de Edge cuando termine de revisar, y se abrirá la carpeta. Con **Run** directamente en Studio, consulte **Variables → resultado_ep3** y el archivo **execution.json**.

El lanzador usa la consola para mostrar el progreso y evita abrir la ventana auxiliar del depurador. Si pulsa **Run** directamente en una instancia de Studio abierta desde una ruta con guiones, esa ventana puede volver a mostrar el error de ruta.

Para enviar un Excel por cada corrida del robot principal, la configuración local de Gmail SMTP en modo **ENVIO** admite el campo **destino_duoc**. El robot envía el archivo recién generado a ese único destinatario y añade **HISTORICO** al asunto si usó la captura de Altoro Mutual. **SMTP_ACEPTADO** registra que el servidor de salida aceptó el mensaje. La recepción se puede ver en la bandeja de Duoc. El UUID identifica la corrida y su reserva evita repetir el envío de ese mismo archivo ante un resultado incierto.

El correo resume el origen, el período, los conteos y la referencia del Banco Mundial. **Reporte_TestFire_EP3.xlsx** incluye los movimientos y la referencia externa. **Analisis_Valores_EP3.xlsx** contiene los totales por tipo y cuenta. **Detalle_Movimientos_Altoro_Mutual.xlsx** contiene una sola hoja con la tabla completa de depósitos y retiros lista para filtrar. **outlook.json** registra la huella de los tres adjuntos.


Al confirmar el envío, **EJECUTAR_EP3.bat** abre Gmail en el navegador predeterminado y busca el correo de esa corrida por su identificador. Si Gmail pide acceso, inicia sesión en la cuenta Duoc. La ventana de Edge del robot se abre con la traducción y los avisos de notificaciones desactivados.

La validación y la conciliación por lotes conservan las mismas reglas, los rechazos y el registro por etapa. En la prueba histórica con 96 filas válidas y 4 rechazadas, el tiempo interno de Rocketbot bajó de **66,59 s a 1,30 s**. El tiempo de apertura de Studio y las esperas del login o del correo dependen del entorno.

## Demostración con correo local

Ejecute **ATAJOS/EP3/EJECUTAR_EVIDENCIA_EP3.bat** o abra **ATAJOS/EP3/INICIAR_EP3.bat -CorreoLocal** y pulse **Run** en el robot **DEMO_CORREO_LOCAL_EP3**. Esta ruta incluye datos sintéticos, consulta HTTPS del Banco Mundial, generación y apertura del XLSX con openpyxl, y envío SMTP con adjunto a un buzón de prueba en el mismo PC. Requiere Internet para consultar el Banco Mundial. El servicio de correo inicia y termina con el robot y funciona sin cuenta ni aplicaciones de Office.

**ATAJOS/EP3/EJECUTAR_EVIDENCIA_EP3.bat** también muestra la demo en Edge. **Correo_EP3.eml** contiene el mensaje original, **outlook.json** guarda la respuesta SMTP y la huella del adjunto y **audit.jsonl** registra cada paso. También está disponible la vista **Correo_EP3.html** del buzón local. La recepción queda dentro del PC. El envío institucional corresponde a la ruta de correo externo.

**Correo sin Google Cloud:** configura [SMTP e IMAP](CORREO_SMTP.md) para el envío y la recepción.

**Respaldo para este PC:** si falta el login o TestFire no responde, el flujo principal puede leer la captura de octubre de 2026 incluida en EP2. El reporte indica `ARCHIVO_LOCAL_HISTORICO` y se verifica con openpyxl. Con `destino_duoc` configurado, el asunto identifica el reporte como histórico y el robot lo envía al terminar la corrida.

**Sin Office de escritorio:** use la variante [Gmail institucional](CORREO_INSTITUCIONAL.md).

Para ejecutar la demo con un doble clic en este equipo, abra **ATAJOS/EP3/EJECUTAR_DEMO_EP3.bat**. El lanzador inicia el robot desde Studio y abre la carpeta de resultados. La ejecución automática usa Microsoft Edge y Playwright. En equipos sin ese entorno, use **ATAJOS/EP3/INICIAR_EP3.bat -Demo** y pulse **Run**.

Abra `ATAJOS/EP3/INICIAR_EP3.bat` para el flujo web o `ATAJOS/EP3/INICIAR_EP3.bat -Demo` para trabajar con transacciones sintéticas. En Studio pulse **Run**. Los dos robots consultan el Banco Mundial por HTTPS en cada ejecución. EP3 no usa la API de EP2. Para una entrega en modo ENVIO configure su coordinador y la cola según [Operación autónoma](OPERACION.md).

## Preparación en Windows

1. Use Rocketbot Studio para Windows. La base `rocketbot/MAIN_TESTFIRE_EP3/robot.db` usa el formato editable de la exportación EP2. Mantenga la ubicación relativa dentro de **TestFire_EP2** y abra ese proyecto si Studio ya está iniciado.
2. Use Excel de escritorio y Outlook clásico con una cuenta de prueba autenticada. Outlook nuevo carece de la automatización COM usada por este proyecto. Antes de la clase pruebe el modo **BORRADOR**. Copie `EP3/outlook.example.json` a `EP3/outlook.local.json`, reemplace `mailbox` por la cuenta de prueba del perfil y deje `mode` como `BORRADOR`. Para la última ejecución de demostración configure la cola compartida, cambie a `ENVIO` y vuelva a pulsar Run. El destinatario es la misma cuenta configurada. Se usa un solo destinatario por ejecución.
3. Para el flujo web, copie `.env.example` a `.env` si todavía no existe. Configure `TESTFIRE_USER` y `TESTFIRE_PASSWORD` con la cuenta de demostración autorizada. En el equipo actual se observó el certificado HTTPS vencido de TestFire. Para la presentación use el robot sintético o espere a que TestFire restablezca HTTPS. La configuración puede apuntar al entorno de demostración autorizado del laboratorio. Mantenga credenciales reales fuera de la demo.
4. En `.env` seleccione un maestro válido con `MASTER_EXCEL_PATH`. El snapshot de EP2 de `data/master_snapshot_testfire.xlsx` trae su periodo en la hoja **Metadata**. Para el snapshot incluido el periodo es del 1 al 31 de octubre de 2026. Configure esos valores en `PERIOD_START` y `PERIOD_END` antes de usarlo. El modo sintético usa su maestro y su periodo de prueba automáticamente. Si Studio no encuentra alguno de los módulos Python usados por EP2, aplique `docs/CORRECCION_DEPENDENCIAS_ROCKETBOT.md`. Para la prueba Python externa se reutilizan las dependencias de EP2.

## Resultados y comprobación

La hoja **Movimientos** del reporte incluye la cuenta, el ID de transacción, la fecha y hora, la descripción y el importe de cada fila válida. La moneda queda vacía cuando la fuente no la informa. **Resumen** registra el origen y el periodo. **Referencia externa** contiene las observaciones del Banco Mundial, la URL y la fecha de consulta. En el flujo principal, `WEB` indica una extracción desde la sesión de Altoro Mutual. `ARCHIVO_LOCAL_HISTORICO` indica que los movimientos provienen de una captura anterior y que el envío quedó omitido.

Cada corrida queda en `EP3/ejecuciones/<UUID>/`. El libro `Reporte_TestFire_EP3.xlsx` contiene **Resumen** y **Referencia externa**. `api_banco_mundial.json` contiene el JSON recibido, `artefacto.json` registra el SHA-256, `excel.json` confirma la apertura de escritorio y `outlook.json` registra el estado sin incluir la dirección. `execution.json` consolida etapas y resultado. `OK` exige recepción clasificada y un flujo sin errores. Un borrador o una confirmación pendiente se informa como `PARCIAL`.

Si Outlook tarda más de 48 segundos, el siguiente ciclo del mismo trabajador comprueba el estado de envío. Para una comprobación adicional puede ejecutar desde PowerShell en la raíz:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\EP3\reanudar_outlook.ps1 -RunId UUID
```

La reanudación solo está disponible para un intento de envío registrado y vuelve a buscar el mismo asunto en Enviados y Entrada. La carpeta del UUID original mantiene la evidencia, incluso si el robot reportó un error tras solicitar el envío.

Para verificar la API real y la creación del XLSX fuera de Studio en un PC con Python y las dependencias EP2:

```powershell
.\.venv-ep2\Scripts\python.exe -m auxiliaries.ep3
```

## Entrega

- `entrega/Informe_Tecnico_EP3.docx` y `entrega/Presentacion_EP3.pptx`.
- `entrega/Guion_EP3.md` para la exposición de 11 minutos.
- `evidencia/` con el libro, el JSON, el registro de comandos y una captura de la ejecución verificada en Rocketbot Studio 2025.07.01. La fuente bancaria fue sintética y la consulta al Banco Mundial fue real.
- Para adjuntar evidencia de escritorio, capture el árbol y el resultado de Studio, las hojas de Excel y el mensaje en Enviados y Recibidos con la categoría **RPA EP3 Procesado**. Oculte la dirección de la cuenta y mantenga visible el mismo UUID en todas las capturas. Grabe el proceso si lo solicita el docente.
