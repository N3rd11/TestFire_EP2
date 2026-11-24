from contextlib import nullcontext
from hashlib import sha256
from openpyxl import load_workbook
import pytest

from auxiliaries import ep3_analysis, ep3_gmail, ep3_smtp


def test_duoc_delivery_uses_run_attachment_recipient_and_unique_reservation(tmp_path, monkeypatch):
    report = b"reporte de prueba"
    (tmp_path / ep3_gmail.REPORT).write_bytes(report)
    digest = sha256(report).hexdigest()
    movements = [
        dict(account="080003", currency="CLP", amount="100.25", description="Deposit", date="2026-10-06", web_id="1"),
        dict(account="080003", currency="CLP", amount="-27.50", description="Withdrawal", date="2026-10-07", web_id="2"),
        dict(account="080003", currency="USD", amount="8.00", description="Deposit", date="2026-10-08", web_id="3"),
        dict(account="080003", currency=None, amount="-4.00", description="Withdrawal", date="2026-10-09", web_id="4"),
    ]
    groups = ep3_analysis.summarize(movements)
    classifications = {"CONCILIADO": 2, "SOLO_WEB": 2}
    start, end = "2026-10-01", "2026-10-31"
    origin = "ARCHIVO_LOCAL_HISTORICO"
    analysis = ep3_analysis.write_analysis(tmp_path, origin, start, end, movements, groups, classifications)
    analysis_bytes = analysis.read_bytes()
    analysis_digest = sha256(analysis_bytes).hexdigest()
    book = load_workbook(analysis, read_only=True)
    try:
        assert book.active.title == "Por tipo"
        assert list(book["Por tipo"].values)[1:] == [
            ("Deposit", "CLP", 1, 100.25),
            ("Deposit", "USD", 1, 8),
            ("Withdrawal", "CLP", 1, -27.50),
            ("Withdrawal", "No informada", 1, -4),
        ]
        detail = list(book["Detalle de operaciones"].values)[1:]
        assert len(detail) == 4
        assert detail[1][0].strftime("%Y-%m-%d") == "2026-10-07"
        assert detail[1][1:] == ("080003", "2", "Withdrawal", -27.50, "CLP")
        assert list(book["Por cuenta"].values)[1:] == [
            ("080003", "CLP", 2, 100.25, 27.50, 72.75),
            ("080003", "No informada", 1, 0, 4, -4),
            ("080003", "USD", 1, 8, 0, 8),
        ]
    finally:
        book.close()
    standalone = ep3_analysis.write_movement_attachment(analysis, len(movements))
    standalone_bytes = standalone.read_bytes()
    standalone_digest = sha256(standalone_bytes).hexdigest()
    book = load_workbook(standalone)
    try:
        assert book.sheetnames == ["Movimientos"]
        assert book.active.tables["TablaMovimientos"].autoFilter.ref == "A1:F5"
        assert list(book.active.values)[0] == ("Fecha y hora", "Cuenta", "ID", "Operación", "Importe", "Moneda")
        assert list(book.active.values)[1:] == detail
    finally:
        book.close()
    extra_attachments = ((analysis.name, analysis_digest), (standalone.name, standalone_digest))
    overview = ep3_analysis.email_summary(origin, start, end, movements, 1,
                                          ([("2025", 4.21)], "2026-11-23T20:00:00Z"))
    client = ep3_smtp.GmailSmtp("demo@gmail.com", "a" * 16)
    recipient = "demo@duocuc.cl"
    sent = []

    class Smtp:
        def send_message(self, message, *, from_addr, to_addrs):
            assert from_addr == "demo@gmail.com" and to_addrs == [recipient]
            assert ep3_gmail.read_state(tmp_path)["estado"] == "INTENTO_REGISTRADO"
            assert str(message["To"]) == recipient and "HISTORICO" in str(message["Subject"])
            body = message.get_body().get_content()
            assert body.strip() == overview
            assert len(body.splitlines()) == 10
            assert "2026-10-01 al 2026-10-31" in body
            assert "4 movimientos válidos · 1 rechazados" in body
            assert "2 depósitos · 2 retiros" in body
            assert "080003" not in body
            assert {part.get_filename(): part.get_payload(decode=True) for part in message.iter_attachments()} == {
                ep3_gmail.REPORT: report, analysis.name: analysis_bytes, standalone.name: standalone_bytes,
            }
            sent.append(message["Message-ID"])
            return {}

    monkeypatch.setattr(client, "smtp", lambda: nullcontext(Smtp()))
    with pytest.raises(ValueError, match="huella"):
        client.send_to_duoc(tmp_path, digest, recipient, origin, summary=overview,
                            extra_attachments=((analysis.name, analysis_digest), (standalone.name, "alterado")))
    client.send_to_duoc(tmp_path, digest, recipient, origin, summary=overview, extra_attachments=extra_attachments)
    state = ep3_gmail.read_state(tmp_path)
    assert state["estado"] == "SMTP_ACEPTADO"
    assert state["adjuntos"] == {ep3_gmail.REPORT: digest, **dict(extra_attachments)}
    with pytest.raises(RuntimeError, match="intento"):
        client.send_to_duoc(tmp_path, digest, recipient, origin, summary=overview, extra_attachments=extra_attachments)
    assert len(sent) == 1


def test_smtp_delivers_and_imap_classifies_the_same_report(tmp_path, monkeypatch):
    folder = tmp_path / "55d0944a-b3c7-464d-90de-9a8cf4d10af0"
    folder.mkdir()
    report = b"synthetic XLSX"
    (folder / ep3_gmail.REPORT).write_bytes(report)
    digest = sha256(report).hexdigest()
    client = ep3_smtp.GmailSmtp("demo@gmail.com", "a" * 16)
    mail = []
    calls = []

    class Smtp:
        def send_message(self, message, *, from_addr, to_addrs):
            assert from_addr == "demo@gmail.com" and to_addrs == [from_addr]
            assert ep3_gmail.read_state(folder)["estado"] == "INTENTO_REGISTRADO"
            mail.append(message)
            return {}

    class Imap:
        current = None
        def list(self, _reference=None, _name=None):
            return "OK", [b'(\\HasNoChildren \\Sent) "/" "[Gmail]/Sent Mail"',
                          b'(\\HasNoChildren) "/" "RPA EP3 Procesado"']

        def select(self, name, readonly=True):
            self.current = name
            calls.append((name, readonly))
            return "OK", [b"1"]

        def uid(self, action, *_args):
            match action:
                case "SEARCH":
                    assert folder.name in _args[-1]
                    return "OK", [b"1"]
                case "FETCH" if _args[-1] == "(BODY.PEEK[])":
                    return "OK", [(b"1", mail[0].as_bytes())]
                case "STORE":
                    assert self.current == "INBOX" and _args[-1] == '("RPA EP3 Procesado")'
                    return "OK", [b"1"]
                case "FETCH":
                    return "OK", [b'1 (X-GM-LABELS ("RPA EP3 Procesado"))']
                case _:
                    raise AssertionError(action)

    monkeypatch.setattr(client, "smtp", lambda: nullcontext(Smtp()))
    monkeypatch.setattr(client, "imap", lambda: nullcontext(Imap()))
    client.check_account()
    client.send(folder, digest)
    client.confirm(folder, digest)
    assert calls == [('"[Gmail]/Sent Mail"', True), ("INBOX", False)]
    assert ep3_gmail.read_state(folder)["canal"] == "GMAIL_SMTP"
    assert ep3_gmail.read_state(folder)["estado"] == "CLASIFICADO"
    client.confirm(folder, digest)
    assert len(mail) == 1


def test_smtp_retries_connection_on_starttls_without_repeating_send(monkeypatch):
    from smtplib import SMTPServerDisconnected

    client = ep3_smtp.GmailSmtp("demo@gmail.com", "a" * 16)
    calls = []
    def failure(*_args, **_kwargs):
        calls.append(465)
        raise SMTPServerDisconnected("connection closed")

    class StartTls:
        def __init__(self, host, port, *, timeout):
            assert host == "smtp.gmail.com" and timeout == 12
            calls.append(port)

        def ehlo(self):
            calls.append("ehlo")

        def starttls(self, *, context):
            assert context is client.ssl
            calls.append("tls")

        def login(self, _address, _secret):
            calls.append("login")

        def close(self):
            calls.append("close")

    monkeypatch.setattr(ep3_smtp.smtplib, "SMTP_SSL", failure)
    monkeypatch.setattr(ep3_smtp.smtplib, "SMTP", StartTls)
    with client.smtp():
        calls.append("send")
    assert calls == [465, 587, "ehlo", "tls", "ehlo", "login", "send", "close"]
