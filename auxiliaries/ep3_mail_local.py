"""Buzón SMTP de evidencia, accesible únicamente desde el mismo equipo."""
from datetime import datetime, timezone
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from email.utils import format_datetime
from hashlib import sha256
from html import escape
import smtplib
from socketserver import StreamRequestHandler, TCPServer
from threading import Thread

SENDER = "robot@ep3.test"
RECIPIENT = "buzon@ep3.test"
MAX_BYTES = 15 * 1024 * 1024


class Inbox(StreamRequestHandler):
    def reply(self, line):
        self.wfile.write(line.encode("ascii") + b"\r\n")

    def handle(self):
        self.connection.settimeout(8)
        self.reply("220 EP3 SMTP local")
        for command, answer in (
            (b"ehlo ep3.test", "250 EP3 ready"),
            (f"mail FROM:<{SENDER}>".encode(), "250 Sender OK"),
            (f"rcpt TO:<{RECIPIENT}>".encode(), "250 Recipient OK"),
            (b"data", "354 End with dot"),
        ):
            if self.rfile.readline(1024).strip().lower() != command.lower():
                self.reply("503 Unexpected SMTP command")
                return
            self.reply(answer)

        received = bytearray()
        while len(received) <= MAX_BYTES:
            line = self.rfile.readline(1024)
            if line == b".\r\n":
                break
            if not line or not line.endswith(b"\r\n"):
                return
            received.extend(line[1:] if line.startswith(b"..") else line)
        else:
            self.reply("552 Message too large")
            return

        try:
            self.server.store_receipt(bytes(received))
        except (ValueError, OSError):
            self.reply("554 Receipt validation failed")
            return
        self.reply("250 Stored and attachment verified")
        if self.rfile.readline(1024).strip().lower() == b"quit":
            self.reply("221 Closing connection")


def deliver(folder, report, digest):
    payload = report.read_bytes()
    if sha256(payload).hexdigest() != digest:
        raise ValueError("El adjunto cambió después de validar el Excel")
    message = EmailMessage(policy=policy.SMTP)
    message["From"], message["To"] = SENDER, RECIPIENT
    message["Subject"] = f"EP3 | Reporte TestFire | {folder.name}"
    message["Message-ID"] = f"<{folder.name}@ep3.test>"
    message["Date"] = format_datetime(datetime.now(timezone.utc))
    message.set_content("Resultado de la conciliación sintética EP3. Se adjunta el reporte Excel.")
    message.add_attachment(payload, maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename=report.name)
    raw = message.as_bytes()
    if len(raw) > MAX_BYTES:
        raise ValueError("El correo supera el tamaño permitido por el buzón de prueba")

    # La reserva evita enviar de nuevo el mismo reporte, incluso tras un cierre incierto.
    with (folder / "correo_local.reserva").open("x", encoding="ascii") as reservation:
        reservation.write(digest)
    receipt_path = folder / "Correo_EP3.eml"

    def store_receipt(received):
        parsed = BytesParser(policy=policy.default).parsebytes(received)
        attachments = list(parsed.iter_attachments())
        if (parsed["Message-ID"] != message["Message-ID"] or parsed["To"] != RECIPIENT
                or parsed["From"] != SENDER or len(attachments) != 1
                or attachments[0].get_filename() != report.name
                or sha256(attachments[0].get_payload(decode=True) or b"").hexdigest() != digest):
            raise ValueError("El correo recibido no coincide con el comprobante esperado")
        with receipt_path.open("xb") as receipt:
            receipt.write(received)

    with TCPServer(("127.0.0.1", 0), Inbox) as sink:
        sink.store_receipt = store_receipt
        worker = Thread(target=sink.serve_forever, kwargs={"poll_interval": 0.1}, daemon=True)
        worker.start()
        try:
            with smtplib.SMTP("127.0.0.1", sink.server_address[1], local_hostname="ep3.test", timeout=8) as client:
                for command, args in ((client.ehlo, ()), (client.mail, (SENDER,)), (client.rcpt, (RECIPIENT,))):
                    code, reply = command(*args)
                    if code != 250:
                        raise smtplib.SMTPResponseException(code, reply)
                code, reply = client.data(raw)
                if code != 250:
                    raise smtplib.SMTPDataError(code, reply)
        finally:
            sink.shutdown()
            worker.join()

    received = BytesParser(policy=policy.default).parsebytes(receipt_path.read_bytes())
    reply_text = reply.decode("ascii")
    body = escape(received.get_body(preferencelist=("plain",)).get_content().strip())
    (folder / "Correo_EP3.html").write_text(f"""<!doctype html>
<html lang="es"><meta charset="utf-8"><title>Correo recibido | EP3</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>body{{margin:0;background:#eef2f6;font:16px system-ui;color:#182536}}
main{{max-width:740px;margin:54px auto;padding:32px;background:white;border-radius:16px}}
h1{{margin:4px 0 14px}}p{{line-height:1.65}}code{{overflow-wrap:anywhere}}
a{{color:#064e86}}small{{font-size:14px;color:#43617c}}</style>
<main><small>EP3 · Buzón SMTP local</small><h1>Correo recibido</h1>
<p><strong>De</strong> {escape(str(received['From']))}<br>
<strong>Para</strong> {escape(str(received['To']))}<br>
<strong>Asunto</strong> {escape(str(received['Subject']))}<br>
<strong>Fecha del mensaje (UTC)</strong> {escape(str(received['Date']))}</p>
<p>{body}</p>
<p>Adjunto <a href="{escape(report.name, quote=True)}">{escape(report.name)}</a></p>
<p><strong>SMTP</strong> {code} {escape(reply_text)}<br>
<strong>SHA-256 del adjunto recibido</strong> <code>{digest}</code></p>
<p><a href="Correo_EP3.eml" download>Descargar mensaje original (.eml)</a></p></main></html>""", encoding="utf-8")
    return dict(canal="SMTP_LOCAL", estado="CLASIFICADO", host="127.0.0.1",
                destinatario=RECIPIENT, asunto=str(received["Subject"]),
                message_id=str(received["Message-ID"]), adjunto=report.name, sha256=digest,
                codigo_smtp=code, respuesta_smtp=reply_text, comprobante=receipt_path.name)
