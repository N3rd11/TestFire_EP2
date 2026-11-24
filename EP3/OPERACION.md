# Operación autónoma de TestFire EP3

## Problema y solución

Dos equipos pueden tomar el mismo reporte y Outlook puede aceptar el envío antes de que el robot reciba el acuse. El robot registra un trabajo por proceso y periodo en una cola SQL central. Reserva el envío antes de solicitarlo a Outlook. Si el correo queda pendiente, en el siguiente ciclo vuelve a buscar el asunto en Enviados y Entrada. Esta automatización trabaja con reportes de prueba. El cobro mencionado en clase ilustra la duplicidad de una acción irreversible.

## Diseño técnico

| Término | Aplicación en Rocketbot |
|---|---|
| Template | La base editable y el patrón reutilizable padre e hijos. `operacion.example.json` define los parámetros por ambiente y equipo. |
| Encolar | `EP3_ENCOLAR` hace POST al coordinador. La restricción UNIQUE sobre la clave conserva un solo trabajo aunque dos equipos lo registren juntos. |
| Leer cola | `EP3_LEER_COLA` pide el trabajo de su proceso y periodo. Una transacción `BEGIN IMMEDIATE` otorga la toma a un equipo y entrega un identificador temporal. Una lectura pasiva de SQL por sí sola no otorgaría exclusividad. |
| Caso web lento o en blanco | El comando de espera de Studio busca elementos por hasta 30 s. Si falla la sesión o la extracción carece de tabla válida, se llama `EP3_RESPALDO_API`. |
| Regex | Se valida cuenta con `\d{1,32}`, la fecha y hora con `\d{4}-\d{2}-\d{2} \d{2}:\d{2}` y el importe mediante el patrón delimitado `MONEY_US`. Después se comprueban fechas reales, periodo, precisión y rango. Se rechaza un identificador de movimiento repetido en la API de respaldo. |
| Autogestión | Hasta tres tomas antes del despacho, espera progresiva y vencimiento de la toma tras 15 minutos. El envío reservado queda bloqueado y el equipo original verifica el acuse con el UUID original. |

La clave distingue prueba de entrega y vista previa. Cada grupo usa el mismo modo y periodo en sus equipos. Una reejecución con el mismo periodo conserva el trabajo y sus estados. Para correcciones use un proceso controlado en el que el responsable compruebe el envío previo antes de autorizar otra entrega. Si falta el archivo local de un intento reservado, la cola conserva el bloqueo. Outlook no ofrece una transacción atómica con la cola: se prioriza evitar el reenvío ante un resultado incierto.

### Viabilidad técnica

La extracción API y la generación XLSX son automatizables por tener contrato, reglas deterministas y resultados reproducibles. La extracción DOM es condicional a un inicio de sesión autorizado, TLS válido y una tabla estable. La integración de escritorio requiere Excel instalado y Outlook clásico con una cuenta autenticada en el perfil de ejecución. Una pantalla en blanco, un bloqueo interactivo de Outlook o credenciales caducadas afectan la disponibilidad. El servicio central y los trabajadores deben compartir conectividad; mantenga los tokens fuera del proyecto exportado y use un único origen de cola.

### Respaldo de movimientos

`source_api_url` acepta un endpoint HTTPS autorizado. El robot añade `start` y `end` en formato ISO y valida un JSON con esas dos fechas y `rows`. Cada fila contiene las cadenas `account`, `date_raw`, `description`, `amount_raw` y `web_id`. Use la misma cobertura que la conciliación esperada. El Banco Mundial solo aporta inflación y permanece como referencia externa.

En laboratorio, active `servir_archivo_demo` en el servidor y `usar_archivo_demo` en el trabajador para servir el snapshot TestFire de octubre de 2026 que acompaña la EP2. Use `MASTER_EXCEL_PATH=data/master_snapshot_testfire.xlsx` y el periodo `2026-10-01` a `2026-10-31`. Este respaldo histórico demuestra el cambio de ruta. El Excel declara `API_ARCHIVO_HISTORICO` y el robot suprime el correo de esa corrida.

## Pasos de laboratorio

1. En el coordinador copie `operacion.example.json` a `operacion.local.json`. Cambie el token por uno aleatorio de 32 caracteres o más. En la raíz arranque ` .\.venv-ep2\Scripts\python.exe -m api.ep3_queue` después de preparar las dependencias de EP2. La demostración escucha solo en loopback, puerto 18103.
2. Para dos PCs, publique el coordinador como un único servicio HTTPS autorizado y protegido y use esa misma URL y token en ambos. Mantenga SQLite local al servidor. Asigne un `worker` distinto por PC. El servidor de desarrollo de Flask es para la prueba local. Para despliegue real use un servidor WSGI y un proxy HTTPS autorizado en la red del laboratorio.
3. Cargue el `robot.db` actualizado con `ATAJOS/EP3/INICIAR_EP3.bat`. Seleccione `MAIN_TESTFIRE_EP3` o use `-Demo` para la versión sintética. `cola_entrada.json` y `cola_lectura.json` muestran quién obtuvo el trabajo. Cualquier envío real requiere también el modo `ENVIO` en Outlook y una reserva central. Complete primero la configuración del buzón y confirme la apertura de Excel.
4. Para autogestión en producción, ejecute una vez `EP3/programar_ciclo.ps1 -RocketbotExe RUTA` para programar un ciclo cada cinco minutos, con sesión de Windows iniciada, una tarea por PC e ignorar una instancia nueva si la anterior aún se ejecuta. Deje activo el coordinador. Este ciclo reintenta casos previos al envío y consulta la recepción de casos ya reservados. La licencia de Studio que permite abrir el editor puede ser distinta de la necesaria para ejecutar por CLI sin supervisión.

**Próximo paso de aceptación.** En el laboratorio, abra el mismo periodo desde dos equipos, active la cuenta de prueba e identifique el único correo recibido. Confirme que una segunda toma conserva el estado de la cola, simule lentitud en la web y verifique la procedencia del XLSX y el resultado en la bitácora.

Fuentes oficiales. [Robot Template](https://docs.rocketbot.com/2024/11/27/documentacion-robot-template/) · [Procesos y colas](https://docs.rocketbot.com/2024/08/20/orquestador-rocketbot-procesos/) · [Programador de tareas](https://docs.rocketbot.com/2019/12/19/crear-tarea-programada/).
