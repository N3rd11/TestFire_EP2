"""Conexión explícita de Gmail para la demo. El usuario autoriza en Google."""
import argparse
import json
from pathlib import Path
import re
from secrets import token_urlsafe
import socket
import sys

ROOT = Path(__file__).resolve().parents[1]
FOLDER = ROOT / "EP3"
RUNTIME = FOLDER / ".runtime-correo"


def save(path, content):
    temporary = path.with_suffix(".tmp")
    temporary.write_text(content, encoding="utf-8")
    temporary.replace(path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("mailbox")
    parser.add_argument("client_json")
    args = parser.parse_args()
    if not re.fullmatch(r"[\w.+-]+@duocuc[.]cl", args.mailbox, re.ASCII):
        raise SystemExit("Indique una sola cuenta institucional de Duoc.")
    sys.path.insert(0, str(RUNTIME))
    from google.auth.transport.requests import AuthorizedSession
    from google_auth_oauthlib.flow import InstalledAppFlow
    from auxiliaries.ep3_gmail import Gmail, SCOPE

    mail_config = FOLDER / "outlook.local.json"
    queue_config = FOLDER / "operacion.local.json"
    if mail_config.exists():
        current = json.loads(mail_config.read_text(encoding="utf-8-sig"))
        if current.get("transport") != "GMAIL_API" or current.get("mailbox", "").lower() != args.mailbox.lower():
            raise SystemExit("Ya existe una configuración de correo distinta. Revísela antes de conectar Gmail.")
    flow = InstalledAppFlow.from_client_secrets_file(args.client_json, [SCOPE])
    credentials = flow.run_local_server(port=0, open_browser=False, timeout_seconds=210,
        authorization_prompt_message="Abre esta URL en tu navegador e inicia sesión con Duoc:\n{url}")
    Gmail(AuthorizedSession(credentials), args.mailbox).check_account()
    save(FOLDER / "gmail-token.local.json", credentials.to_json())
    existing = json.loads(queue_config.read_text(encoding="utf-8-sig")) if queue_config.exists() else {}
    if not existing or existing.get("solo_respaldo"):
        worker = re.sub(r"[^A-Za-z0-9_-]", "-", socket.gethostname())[:64] or "pc-ep3"
        save(queue_config, json.dumps(dict(embedded=True, token=token_urlsafe(32),
                                           worker=worker,
                                           usar_archivo_local=existing.get("usar_archivo_local", False)), indent=2))
    save(mail_config, json.dumps(dict(mode="ENVIO", transport="GMAIL_API", mailbox=args.mailbox.lower()), indent=2))
    print("Gmail conectado. ATAJOS/EP3/EJECUTAR_DEMO_EP3.bat iniciará la demo con un envío para el periodo.")


if __name__ == "__main__":
    main()
