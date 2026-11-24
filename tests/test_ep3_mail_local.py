from concurrent.futures import ThreadPoolExecutor
from email import policy
from email.parser import BytesParser
from hashlib import sha256

from openpyxl import Workbook
import pytest

from auxiliaries import ep3_mail_local


def test_delivery_preserves_attachment_and_blocks_concurrent_duplicates(tmp_path):
    report = tmp_path / "Reporte_TestFire_EP3.xlsx"
    workbook = Workbook()
    workbook.save(report)
    workbook.close()
    payload = report.read_bytes()
    digest = sha256(payload).hexdigest()

    def send():
        try:
            return ep3_mail_local.deliver(tmp_path, report, digest)
        except FileExistsError:
            return None

    with ThreadPoolExecutor(max_workers=2) as pool:
        outcomes = list(pool.map(lambda _: send(), range(2)))

    (receipt,) = [outcome for outcome in outcomes if outcome is not None]
    assert receipt["codigo_smtp"] == 250
    message = BytesParser(policy=policy.default).parsebytes((tmp_path / receipt["comprobante"]).read_bytes())
    assert message["Message-ID"] == receipt["message_id"]
    (attachment,) = message.iter_attachments()
    assert attachment.get_filename() == report.name
    assert attachment.get_payload(decode=True) == payload


def test_delivery_refuses_modified_attachment_before_reservation(tmp_path):
    report = tmp_path / "Reporte_TestFire_EP3.xlsx"
    report.write_bytes(b"changed")
    with pytest.raises(ValueError, match="adjunto"):
        ep3_mail_local.deliver(tmp_path, report, sha256(b"original").hexdigest())
    assert not (tmp_path / "correo_local.reserva").exists()
    assert not (tmp_path / "Correo_EP3.eml").exists()
