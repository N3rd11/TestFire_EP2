"""Valida SMTP/IMAP y configura la entrega por Gmail en este Windows."""
import argparse
import json
from pathlib import Path
import re
from secrets import token_urlsafe
import socket
import sys

from auxiliaries.ep3_conectar_gmail import save
from auxiliaries.ep3_smtp import ADDRESS, client


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("mailbox")
    parser.add_argument("credential", type=Path)
    args = parser.parse_args()
    mailbox = args.mailbox.lower()
    if not re.fullmatch(ADDRESS, mailbox, re.ASCII):
        raise ValueError("Indica tu cuenta Gmail o Duoc")
    root = Path(__file__).resolve().parents[1]
    folder = root / "EP3"
    mail_path = folder / "outlook.local.json"
    queue_path = folder / "operacion.local.json"
    mail = json.loads(mail_path.read_text(encoding="utf-8-sig")) if mail_path.exists() else {}
    if mail and mail.get("mailbox", "").lower() != mailbox:
        raise ValueError("Ya hay otro destinatario configurado en EP3/outlook.local.json")
    queue = json.loads(queue_path.read_text(encoding="utf-8-sig")) if queue_path.exists() else {}
    client(root, mailbox, args.credential).check_account()
    if not queue or queue.pop("solo_respaldo", False):
        queue.update(embedded=True, token=token_urlsafe(32),
                     worker=re.sub(r"[^A-Za-z0-9_-]", "-", socket.gethostname())[:64] or "pc-ep3")
        save(queue_path, json.dumps(queue, indent=2))
    args.credential.replace(folder / "smtp-credencial.local.clixml")
    save(mail_path, json.dumps(dict(mode="ENVIO", transport="GMAIL_SMTP", mailbox=mailbox), indent=2))
    print("Cuenta verificada por SMTP e IMAP. La configuración no envió un correo.")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, OSError, RuntimeError) as exc:
        print(exc, file=sys.stderr)
        sys.exit(1)
