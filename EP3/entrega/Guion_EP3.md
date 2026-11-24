# Guion de exposición EP3

Duración objetivo de 11 minutos. Complete los nombres del equipo en la portada. En la presentación individual explique su aporte real y acuerde los turnos con el equipo.

| Diapositiva | Tiempo | Explicación y demostración |
|---|---:|---|
| 1 · Portada | 0:30 | Presentar al equipo y el objetivo. |
| 2 · Propósito | 1:00 | Explicar cómo se repite la conciliación y qué contexto añade el Banco Mundial. Mostrar la distinción entre datos sintéticos y fuente pública real. |
| 3 · Arquitectura | 1:15 | Abrir el árbol de Studio y señalar login, extracción, API, integración, apertura de Excel y Outlook. Mostrar EP3_ENCOLAR y EP3_LEER_COLA. Explicar el respaldo ante tabla vacía, el template y las dependencias técnicas de Excel y Outlook. |
| 4 · API pública | 1:00 | Mostrar el indicador, la consulta HTTPS y el año de la observación más reciente. Explicar el formato de porcentaje. |
| 5 · Evidencia API | 1:00 | Abrir el JSON de la ejecución y el log. Mostrar HTTP 200, CHL, el identificador del indicador y rechazo de un dato inválido. |
| 6 · Excel | 1:15 | Comparar la conciliación base con Resumen y Referencia externa de EP3. En la ejecución de laboratorio mostrar excel.json. |
| 7 · Outlook | 1:45 | Mostrar en Outlook clásico Enviados y Bandeja de entrada con el mismo UUID, el adjunto y la categoría. Si la prueba no se pudo cerrar, explicar el estado real y mostrar la evidencia disponible. |
| 8 · Evitar duplicados | 1:10 | Explicar la clave de negocio, la toma SQL y la reserva. Demostrar que el segundo cliente no obtiene el trabajo y el ciclo recupera el acuse sin reenviar. |
| 9 · Riesgos | 1:10 | Cubrir minimización, destinatario único, certificado vencido y límite de reintentos. |
| 10 · Cierre | 0:55 | Mostrar resultados locales, límites y próximos pasos del ciclo operativo. |

## Preguntas de defensa

**¿Qué sucede si el Banco Mundial responde HTTP 200 pero cambia el formato?** El código valida la estructura, el indicador, el país, el año y los valores. Registra el error y corta la integración. Cambiar el contrato requiere revisar el parser y volver a probar.

**¿Por qué se confirma la recepción si Outlook ya ejecutó Send?** Send solicita el despacho. La confirmación exige el mismo asunto en Elementos enviados y Bandeja de entrada, más el adjunto previsto y la aplicación de la categoría. Si se pierde la confirmación, el identificador original permite retomar la búsqueda.
