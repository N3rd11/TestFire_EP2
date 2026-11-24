# Ejecución completa con Gmail institucional

`MAIN_TESTFIRE_EP3` es el robot principal. Desde la licencia de desarrollo, abre `ATAJOS/EP3/INICIAR_EP3.bat` y pulsa **Run** una vez en Studio. El flujo reclama el periodo, inicia sesión en TestFire, extrae y concilia los movimientos, guarda el registro local, consulta la API pública, genera el Excel de entrega y lo envía a Gmail cuando la cuenta y la cola ya están configuradas.

- **Registro_Movimientos_EP3.xlsx** guarda el detalle de movimientos, conciliación y rechazos para revisar localmente.
- **Reporte_TestFire_EP3.xlsx** es el adjunto. Contiene conteos, origen y referencia externa. El robot lo vuelve a abrir con openpyxl, valida la huella, y comprueba en Gmail el mensaje enviado, el recibido y la etiqueta **RPA EP3 Procesado**.
- **DEMO_TESTFIRE_EP3** ejecuta las mismas etapas desde el registro con datos sintéticos. La prueba sintética y el flujo web tienen claves de cola diferentes.
- Si TestFire falla, se puede configurar una API HTTPS autorizada que entregue los movimientos. El archivo histórico de laboratorio genera resultados con su procedencia explícita y omite el despacho.

## Preparación única

1. El flujo web necesita un TestFire de demostración con HTTPS vigente, una cuenta de prueba en `.env` y un maestro para el periodo. La configuración y el formato del maestro están en `LEEME.md`.
2. Obtén un **cliente OAuth de escritorio autorizado** para Gmail API, abre `ATAJOS/EP3/CONECTAR_CORREO_EP3.bat` y selecciona el JSON de Google que descargaste. Escribe tu correo institucional e inicia sesión en Google. Workspace puede requerir autorización del administrador. El token queda local y se excluye de Git.
3. Abre `ATAJOS/EP3/INICIAR_EP3.bat` y pulsa **Run**. El conector prepara una cola SQLite local por defecto. Para compartir la cola entre varios equipos usa el coordinador HTTPS de `OPERACION.md`.

El mismo botón **Run** recupera un acuse pendiente sin reenviar el correo. La carpeta `EP3/ejecuciones/<UUID>` contiene el registro, el reporte, la bitácora y el estado del correo de esa ejecución.

[Configuración OAuth de Gmail](https://developers.google.com/workspace/gmail/api/quickstart/python) · [Permisos de Gmail](https://developers.google.com/workspace/gmail/api/auth/scopes).
