from base64 import urlsafe_b64decode, urlsafe_b64encode
from email import policy
from email.parser import BytesParser
from hashlib import sha256
from types import SimpleNamespace

import pytest

from auxiliaries import ep3_gmail as mail, ep3_ops


def test_gmail_mime_and_confirm_without_resending(tmp_path):
    run = tmp_path / "6a438eb1-ab92-4d18-b026-16ba793bea99"
    run.mkdir()
    xlsx = b"test-report"
    (run / mail.REPORT).write_bytes(xlsx)
    digest = sha256(xlsx).hexdigest()
    addr = "demo@duocuc.cl"
    transmissions = []

    class Session:
        def request(self, method, url, *, json, params, **kwargs):
            rest = url.removeprefix(mail.BASE)
            if rest == "profile":
                data = {"emailAddress": addr}
            elif rest == "messages/send":
                assert method == "POST"
                assert mail.read_state(run)["estado"] == "INTENTO_REGISTRADO"
                message = BytesParser(policy=policy.default).parsebytes(urlsafe_b64decode(json["raw"]))
                assert message["To"] == message["From"] == addr
                assert message["Subject"] == mail.SUBJECT_PREFIX + run.name
                assert list(message.iter_attachments())[0].get_payload(decode=True) == xlsx
                transmissions.append(message)
                data = {"id": "sent"}
            elif rest == "messages":
                assert params["q"] == f"rfc822msgid:{run.name}@testfire-ep3.invalid"
                data = {"messages": [{"id": "sent" if params["labelIds"] == "SENT" else "received"}]}
            elif rest == "messages/received":
                data = {"payload": {"headers": [{"name": "Subject", "value": mail.SUBJECT_PREFIX + run.name}],
                    "parts": [{"filename": mail.REPORT, "body": {"data": urlsafe_b64encode(xlsx).decode()}}]}}
            elif rest == "labels":
                data = {"labels": [{"name": mail.LABEL, "id": "Label_1"}]}
            elif rest == "messages/received/modify":
                assert json == {"addLabelIds": ["Label_1"]}
                data = {"labelIds": ["INBOX", "Label_1"]}
            else:
                raise AssertionError(rest)
            return SimpleNamespace(status_code=200, json=lambda: data)

    client = mail.Gmail(Session(), addr)
    client.check_account()
    client.send(run, digest)
    client.confirm(run, digest, attempts=1)
    assert mail.read_state(run)["estado"] == "CLASIFICADO"
    assert mail.read_state(run)["sha256"] == digest
    with pytest.raises(RuntimeError, match="estado"):
        client.send(run, digest)
    client.confirm(run, digest)
    assert len(transmissions) == 1


def test_embedded_queue_excludes_second_worker(tmp_path):
    def worker(name, run_id):
        return {"cfg": {"PROJECT_ROOT": str(tmp_path)},
                "ops": {"embedded": True, "token": "x" * 40, "worker": name},
                "job_key": "demo:2026-10-01:2026-10-31",
                "audit": SimpleNamespace(id=run_id)}
    one = worker("uno", "285fa16f-cbf8-484d-916f-adc349685487")
    two = worker("dos", "0e2cc953-b4f5-4f51-a705-016baf19e280")
    assert ep3_ops.job_request(one, "enqueue")["added"]
    assert not ep3_ops.job_request(two, "enqueue")["added"]
    assert ep3_ops.job_request(one, "claim")["claimed"]
    assert not ep3_ops.job_request(two, "claim")["claimed"]


@pytest.mark.parametrize("transport", ["GMAIL_API", "GMAIL_SMTP"])
def test_main_web_report_uses_gmail_without_desktop(tmp_path, monkeypatch, transport):
    from auxiliaries import bridge, ep3
    from openpyxl import load_workbook

    original_load = ep3.load_config
    config = original_load(ep3.ROOT, require_secrets=False)
    config.update(PROJECT_ROOT=str(tmp_path), TESTFIRE_URL="https://demo.testfire.net/")
    monkeypatch.setattr(ep3, "load_config", lambda *args, **kwargs: config)
    monkeypatch.setattr(ep3_ops, "mailbox_mode", lambda _root: "ENVIO")
    eid = ep3.initialize(tmp_path)
    s = bridge.session(eid)
    s.update(
        mail_config={"transport": transport, "mailbox": "demo@duocuc.cl"},
        master={},
        valid=[dict(account="0800003", web_id="00123", date_raw="2026-10-18 06:36",
                    description="=1+1", amount="-25.75", currency=None)],
        reference=([("2025", 4.21)], "2026-11-23T15:00:00Z"),
        web_ok=True,
        claim="valid-test-claim",
        job_key="testfire:2026-10-01:2026-10-31",
    )
    actions = []
    def queue(_session, action, **_kwargs):
        actions.append(action)
        return {"state": "ENVIO_RESERVADO" if action == "reserve" else "COMPLETADO"}

    def send(folder, digest, origin):
        assert origin == "WEB" and digest == ep3.report_digest(folder)
        assert actions == ["reserve"]
        actions.append("send")

    def confirm(folder, digest):
        actions.append("confirm")
        mail.write_state(folder, "CLASIFICADO", channel=transport, sha256=digest)

    monkeypatch.setattr(ep3_ops, "job_request", queue)
    stub = SimpleNamespace(send=send, confirm=confirm, check_account=lambda: None)
    monkeypatch.setattr(ep3, "gmail_client", lambda _session: stub)
    monkeypatch.setattr(ep3.ep3_smtp, "client", lambda _root, _mailbox: stub)
    try:
        ep3.register_movements(eid)
        assert s["register"].is_file()
        ep3.integrate_excel(eid)
        ep3.open_excel(eid)
        report = load_workbook(s["report"], read_only=True)
        try:
            assert report.sheetnames == ["Movimientos", "Resumen", "Referencia externa"]
            assert report.active.title == "Movimientos"
            movements = report["Movimientos"]
            assert movements["A2"].value == "0800003"
            assert movements["B2"].value == "00123"
            assert movements["C2"].value.isoformat(timespec="minutes") == "2026-10-18T06:36"
            assert movements["D2"].value == "=1+1" and movements["D2"].data_type == "s"
            assert movements["E2"].value == -25.75
            assert movements["F2"].value == "No informada"
        finally:
            report.close()
        assert s["xlsx_verified"] and not s["excel_open"]
        s["origin"] = "API_ARCHIVO_HISTORICO"
        ep3.mail_and_classify(eid)
        assert s["mail"] == "ARCHIVO_SIN_DESPACHO" and not actions
        s["origin"] = "WEB"
        ep3.mail_and_classify(eid)
        assert ep3.finalize(eid)["status"] == "OK"
        assert actions == ["reserve", "send", "confirm", "finish"]
    finally:
        bridge._sessions.pop(eid, None)

