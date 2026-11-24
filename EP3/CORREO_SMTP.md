# Correo SMTP e IMAP para Rocketbot EP3

Esta opción no requiere crear un proyecto en Google Cloud, activar facturación ni instalar Outlook o Excel de escritorio. Rocketbot prepara el XLSX, verifica la huella y usa SMTP para enviarlo a la misma cuenta autorizada. Después consulta Enviados e Inbox por IMAP, compara el adjunto y aplica la etiqueta **RPA EP3 Procesado**.

## Conectar una cuenta una sola vez

1. Abre [Contraseñas de aplicación de Google](https://myaccount.google.com/apppasswords) con tu cuenta Duoc. Google exige verificación en dos pasos. Si la cuenta institucional no muestra la opción o bloquea IMAP, usa una cuenta Gmail personal propia para una prueba autocontenida, o solicita al laboratorio una cuenta que permita SMTP e IMAP. [Requisitos de Google](https://support.google.com/accounts/answer/185833)
2. Genera la contraseña de aplicación en Google. Abre `ATAJOS/EP3/CONECTAR_CORREO_SMTP_EP3.bat`, escribe tu correo y pega allí los 16 caracteres sin espacios. El campo oculta la entrada. Windows protege la credencial en `EP3/smtp-credencial.local.clixml` para el usuario de ese PC y el archivo se excluye de Git. La configuración primero comprueba el acceso a SMTP e IMAP y no envía un correo.
3. Abre `ATAJOS/EP3/EJECUTAR_DEMO_EP3.bat`. El robot usa datos sintéticos, envía al buzón configurado y vuelve a buscar el mismo mensaje si la confirmación queda pendiente. Guarda el resultado en `EP3/ejecuciones/<UUID>/outlook.json` bajo el canal `GMAIL_SMTP`. Usa Enviados y Recibidos de Gmail para capturar el mismo asunto y el adjunto.

`MAIN_TESTFIRE_EP3` puede usar el mismo transporte para un reporte obtenido en vivo. Las corridas que recurren al archivo histórico omiten el envío. Los modos Gmail API y Outlook de escritorio siguen disponibles en el proyecto.

Gmail SMTP usa `smtp.gmail.com:465` con TLS (si la conexión se cierra antes del envío prueba 587 con STARTTLS) e IMAP usa `imap.gmail.com:993` con TLS. [Manual de Email de Rocketbot](https://docs.rocketbot.com/category/manual/email/) · [Extensiones de IMAP de Gmail](https://developers.google.com/workspace/gmail/imap/imap-extensions).
