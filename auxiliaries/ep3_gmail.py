"""Transporte Gmail autorizado para el reporte agregado de EP3."""
import base64
from email.message import EmailMessage
from email.policy import SMTP
from email.utils import formatdate
from hashlib import sha256
import json
from pathlib import Path
import re
import time

SCOPE = "https://www.googleapis.com/auth/gmail.modify"
BASE = "https://gmail.googleapis.com/gmail/v1/users/me/"
LABEL = "RPA EP3 Procesado"
REPORT = "Reporte_TestFire_EP3.xlsx"
SUBJECT_PREFIX = "[RPA EP3] TestFire "


def client(root, expected):
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2.credentials import Credentials

    token = Path(root) / "EP3" / "gmail-token.local.json"
    if not token.is_file():
        raise RuntimeError("Conectar Gmail institucional con ATAJOS/EP3/CONECTAR_CORREO_EP3.bat")
    credentials = Credentials.from_authorized_user_file(str(token), [SCOPE])
    gmail = Gmail(AuthorizedSession(credentials), expected)
    gmail.check_account()
    return gmail


def read_state(folder):
    path = folder / "outlook.json"
    return json.loads(path.read_text(encoding="utf-8-sig")) if path.is_file() else {}


def write_state(folder, state, channel="GMAIL_API", **fields):
    path = folder / "outlook.json"
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(dict(estado=state, canal=channel, **fields), indent=2), encoding="utf-8")
    temp.replace(path)


def build_message(folder, digest, address, origin, recipient=None, *, summary=None, extra_attachments=()):
    message = EmailMessage(policy=SMTP)
    message["Date"] = formatdate(usegmt=True)
    message["From"] = address
    message["To"] = recipient or address
    historical = origin in ("ARCHIVO_LOCAL_HISTORICO", "API_ARCHIVO_HISTORICO")
    message["Subject"] = SUBJECT_PREFIX + ("HISTORICO " if historical else "") + folder.name
    message["Message-ID"] = f"<{folder.name}@testfire-ep3.invalid>"
    source = "una captura histórica de Altoro Mutual" if historical else {
        "SINTETICO": "datos sintéticos", "WEB": "TestFire de demostración",
        "API_OPERATIVA": "la API de movimientos autorizada",
    }[origin]
    body = f"Reporte EP3 generado por Rocketbot con {source}. Ejecución {folder.name}."
    message.set_content(summary if summary is not None else f"{body}\nAdjunto {REPORT}.")
    for name, expected_digest in ((REPORT, digest), *extra_attachments):
        payload = (folder / name).read_bytes()
        if sha256(payload).hexdigest() != expected_digest:
            raise ValueError(f"La huella de {name} no coincide")
        message.add_attachment(payload, maintype="application",
                               subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=name)
    return message


class Gmail:
    def __init__(self, session, address):
        self.session = session
        if not isinstance(address, str) or not re.fullmatch(r"[\w.+-]+@duocuc[.]cl", address, re.ASCII):
            raise ValueError("Indicar un único destinatario institucional de Duoc")
        self.address = address.lower()

    def request(self, endpoint, *, data=None, params=None):
        response = self.session.request("POST" if data is not None else "GET", BASE + endpoint,
                                        json=data, params=params, timeout=12, allow_redirects=False)
        if not 200 <= response.status_code < 300:
            raise RuntimeError(f"Gmail API respondió HTTP {response.status_code}")
        return response.json()

    def check_account(self):
        if self.request("profile").get("emailAddress", "").lower() != self.address:
            raise RuntimeError("La cuenta de Gmail autorizada no coincide con el destinatario configurado")

    def send(self, folder, digest, origin="SINTETICO"):
        if read_state(folder):
            raise RuntimeError("Esta ejecución ya tiene un estado de correo")
        message = build_message(folder, digest, self.address, origin)
        write_state(folder, "INTENTO_REGISTRADO")
        result = self.request("messages/send", data={"raw": base64.urlsafe_b64encode(message.as_bytes()).decode()})
        write_state(folder, "ENVIADO_PENDIENTE", gmail_id=result["id"])

    def find(self, folder, label):
        result = self.request("messages", params={
            "q": f"rfc822msgid:{folder.name}@testfire-ep3.invalid",
            "labelIds": label, "maxResults": 5,
        })
        matches = result.get("messages", [])
        if len(matches) > 1 or result.get("nextPageToken"):
            raise RuntimeError(f"Gmail devolvió más de un mensaje en {label}")
        return matches[0]["id"] if matches else None

    def verify_attachment(self, message_id, folder, digest):
        message = self.request(f"messages/{message_id}", params={"format": "full"})
        payload = message["payload"]
        headers = {h["name"].lower(): h["value"] for h in payload.get("headers", [])}
        if headers.get("subject") != SUBJECT_PREFIX + folder.name:
            raise ValueError("El asunto no coincide con la ejecución")
        attachments = [p for p in payload.get("parts", []) if p.get("filename") == REPORT]
        if len(attachments) != 1:
            raise ValueError("El adjunto esperado no está en el mensaje")
        body = attachments[0]["body"]
        data = body.get("data")
        if data is None:
            data = self.request(f"messages/{message_id}/attachments/{body['attachmentId']}")["data"]
        if sha256(base64.urlsafe_b64decode(data + "=" * (-len(data) % 4))).hexdigest() != digest:
            raise ValueError("El adjunto recibido no coincide con el XLSX")

    def confirm(self, folder, digest, attempts=6):
        state = read_state(folder)
        if state.get("estado") not in ("INTENTO_REGISTRADO", "ENVIADO_PENDIENTE"):
            return
        for index in range(attempts):
            sent = self.find(folder, "SENT")
            received = self.find(folder, "INBOX")
            if sent and received:
                self.verify_attachment(received, folder, digest)
                labels = self.request("labels").get("labels", [])
                label = next((item for item in labels if item["name"] == LABEL), None)
                if label is None:
                    label = self.request("labels", data={"name": LABEL})
                result = self.request(f"messages/{received}/modify", data={"addLabelIds": [label["id"]]})
                if label["id"] not in result.get("labelIds", []):
                    raise RuntimeError("Gmail no confirmó la etiqueta")
                write_state(folder, "CLASIFICADO", enviados=1, recibidos=1,
                            categoria=LABEL, asunto=SUBJECT_PREFIX + folder.name, adjunto=REPORT,
                            sha256=digest)
                return
            if index + 1 < attempts:
                time.sleep(3)
