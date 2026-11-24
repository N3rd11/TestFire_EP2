"""Correo EP3 con SMTP/IMAP de Gmail y una credencial protegida por Windows."""
from contextlib import contextmanager
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr
from hashlib import sha256
import imaplib
import json
from pathlib import Path
import re
import smtplib
import ssl
import subprocess
import time

from auxiliaries.ep3_gmail import LABEL, REPORT, SUBJECT_PREFIX, build_message, read_state, write_state


ADDRESS = r"[\w.+-]+@(gmail[.]com|duocuc[.]cl)"


def client(root, expected, credential=None):
    folder = Path(root) / "EP3"
    file = credential or folder / "smtp-credencial.local.clixml"
    result = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
         "-File", str(folder / "leer_credencial_smtp.ps1"), "-Ruta", str(file)],
        capture_output=True, timeout=15, check=False, text=True, encoding="utf-8",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if result.returncode:
        raise RuntimeError("No se pudo abrir la credencial SMTP del usuario de Windows")
    data = json.loads(result.stdout)
    if data["username"].lower() != expected.lower():
        raise ValueError("La credencial SMTP pertenece a otra cuenta")
    return GmailSmtp(expected, data["password"])


def imap_ok(result, step):
    status, values = result
    if status != "OK":
        raise RuntimeError(f"Gmail IMAP no confirmó {step}")
    return values


class GmailSmtp:
    def __init__(self, address, password):
        if not re.fullmatch(ADDRESS, address, re.ASCII | re.I):
            raise ValueError("Se requiere una sola cuenta propia Gmail o Duoc")
        if not re.fullmatch(r"[A-Za-z0-9]{16}", password):
            raise ValueError("Usa una contraseña de aplicación de Google de 16 caracteres sin espacios")
        self.address = address.lower()
        self.password = password
        self.ssl = ssl.create_default_context()

    def authenticate_smtp(self, port):
        server = None
        try:
            if port == 465:
                server = smtplib.SMTP_SSL("smtp.gmail.com", port, context=self.ssl, timeout=12)
            else:
                server = smtplib.SMTP("smtp.gmail.com", port, timeout=12)
                server.ehlo()
                server.starttls(context=self.ssl)
                server.ehlo()
            server.login(self.address, self.password)
            return server
        except Exception:
            if server is not None:
                server.close()
            raise

    @contextmanager
    def smtp(self):
        for port in (465, 587):
            try:
                server = self.authenticate_smtp(port)
                break
            except ssl.SSLCertVerificationError:
                raise RuntimeError("No se pudo verificar el certificado TLS de Gmail SMTP.") from None
            except smtplib.SMTPAuthenticationError as exc:
                raise RuntimeError(f"Gmail rechazó la autenticación SMTP ({exc.smtp_code}). "
                                   "Usa el correo que generó la contraseña de aplicación.") from None
            except (OSError, smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError):
                if port == 587:
                    raise RuntimeError("Gmail cerró la conexión SMTP antes del envío en 465 y 587. "
                                       "La cuenta no quedó conectada.") from None
            except smtplib.SMTPException:
                raise RuntimeError(f"Gmail no pudo iniciar SMTP seguro en el puerto {port}.") from None
        try:
            yield server
        except (smtplib.SMTPException, OSError):
            raise RuntimeError("SMTP se interrumpió durante el envío. Comprueba el acuse antes de reintentar.") from None
        finally:
            server.close()

    @contextmanager
    def imap(self):
        try:
            with imaplib.IMAP4_SSL("imap.gmail.com", 993, ssl_context=self.ssl, timeout=12) as server:
                server.login(self.address, self.password)
                yield server
        except ssl.SSLCertVerificationError:
            raise RuntimeError("No se pudo verificar el certificado TLS de Gmail IMAP.") from None
        except (OSError, imaplib.IMAP4.abort):
            raise RuntimeError("Gmail cerró la conexión IMAP en el puerto 993.") from None
        except imaplib.IMAP4.error:
            raise RuntimeError("Gmail rechazó el acceso IMAP o no admite esta operación para la cuenta.") from None

    def check_account(self):
        with self.smtp():
            pass
        with self.imap() as server:
            self.sent_folder(server)

    @staticmethod
    def sent_folder(server):
        for item in imap_ok(server.list(), "las carpetas"):
            if not isinstance(item, bytes):
                continue
            match = re.fullmatch(rb'\(([^)]*)\) (?:NIL|"[^"]*") (.+)', item)
            if match and b"\\sent" in match[1].lower().split():
                return match[2].decode("ascii")
        raise RuntimeError("No se encontró la carpeta especial Enviados")

    def send(self, folder, digest, origin="SINTETICO"):
        if read_state(folder):
            raise RuntimeError("Esta ejecución ya tiene un estado de correo")
        message = build_message(folder, digest, self.address, origin)
        with self.smtp() as server:
            write_state(folder, "INTENTO_REGISTRADO", channel="GMAIL_SMTP")
            result = server.send_message(message, from_addr=self.address, to_addrs=[self.address])
            if result:
                raise RuntimeError("El servidor SMTP rechazó la dirección")
            write_state(folder, "ENVIADO_PENDIENTE", channel="GMAIL_SMTP")

    def send_to_duoc(self, folder, digest, recipient, origin, *, summary, extra_attachments):
        if not isinstance(recipient, str) or not re.fullmatch(r"[\w.+-]+@duocuc[.]cl", recipient, re.ASCII | re.I):
            raise ValueError("El destino debe ser una sola cuenta institucional Duoc")
        if read_state(folder):
            raise RuntimeError("Esta ejecución ya tiene un intento de correo")
        message = build_message(folder, digest, self.address, origin, recipient=recipient,
                                summary=summary, extra_attachments=extra_attachments)
        details = dict(destinatario=recipient, asunto=str(message["Subject"]),
                       message_id=str(message["Message-ID"]), adjunto=REPORT, sha256=digest,
                       adjuntos=dict(((REPORT, digest), *extra_attachments)))
        with self.smtp() as server:
            # Reservar antes de transmitir, incluso si se pierde la respuesta del servidor.
            with (folder / "correo_duoc.reserva").open("x", encoding="ascii") as reservation:
                reservation.write(digest)
            write_state(folder, "INTENTO_REGISTRADO", channel="GMAIL_SMTP_DUOC", **details)
            if server.send_message(message, from_addr=self.address, to_addrs=[recipient]):
                raise RuntimeError("El servidor SMTP rechazó el destinatario Duoc")
            write_state(folder, "SMTP_ACEPTADO", channel="GMAIL_SMTP_DUOC", **details)

    @staticmethod
    def find(server, box, run_id, readonly=True):
        imap_ok(server.select(box, readonly=readonly), "la carpeta")
        uids = imap_ok(server.uid("SEARCH", None, "HEADER", "Message-ID",
                                 f'"<{run_id}@testfire-ep3.invalid>"'), "la búsqueda")
        matches = (uids[0] or b"").split() if uids else []
        if len(matches) > 1:
            raise RuntimeError("Se encontró más de un mensaje con el identificador de ejecución")
        return matches[0].decode("ascii") if matches else None

    def verify_attachment(self, server, uid, folder, digest):
        parts = imap_ok(server.uid("FETCH", uid, "(BODY.PEEK[])"), "la lectura del mensaje")
        raw = next((item[1] for item in parts if isinstance(item, tuple)), None)
        if raw is None:
            raise RuntimeError("Gmail no devolvió el mensaje")
        message = BytesParser(policy=policy.default).parsebytes(raw)
        if (message["Subject"] != SUBJECT_PREFIX + folder.name
                or parseaddr(str(message["To"]))[1].lower() != self.address):
            raise ValueError("El asunto o el destinatario no coinciden con la ejecución")
        attached = [part for part in message.iter_attachments() if part.get_filename() == REPORT]
        if len(attached) != 1 or sha256(attached[0].get_payload(decode=True)).hexdigest() != digest:
            raise ValueError("El adjunto recibido no coincide con el XLSX")

    @staticmethod
    def classify(server, uid):
        quoted = '"' + LABEL + '"'
        existing = imap_ok(server.list('""', quoted), "la lista de etiquetas")
        if not any(isinstance(item, bytes) and item.endswith(quoted.encode()) for item in existing):
            imap_ok(server.create(quoted), "la creación de la etiqueta")
        imap_ok(server.uid("STORE", uid, "+X-GM-LABELS", f"({quoted})"), "la etiqueta")
        labels = imap_ok(server.uid("FETCH", uid, "(X-GM-LABELS)"), "la verificación de la etiqueta")
        if not any(isinstance(item, bytes) and quoted.encode() in item for item in labels):
            raise RuntimeError("Gmail no confirmó la etiqueta de la ejecución")

    def confirm(self, folder, digest, attempts=6):
        if read_state(folder).get("estado") not in ("INTENTO_REGISTRADO", "ENVIADO_PENDIENTE"):
            return
        with self.imap() as server:
            sent_box = self.sent_folder(server)
            for index in range(attempts):
                sent = self.find(server, sent_box, folder.name)
                received = self.find(server, "INBOX", folder.name, readonly=False)
                if sent and received:
                    self.verify_attachment(server, received, folder, digest)
                    self.classify(server, received)
                    write_state(folder, "CLASIFICADO", channel="GMAIL_SMTP", enviados=1, recibidos=1,
                                categoria=LABEL, asunto=SUBJECT_PREFIX + folder.name, adjunto=REPORT, sha256=digest)
                    return
                if index + 1 < attempts:
                    time.sleep(3)
