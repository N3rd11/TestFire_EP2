from copy import deepcopy
import json
from pathlib import Path
from urllib.error import HTTPError

import pytest

from auxiliaries import bridge, ep3


def payload():
    return [{"pages": 1}, [
        {"countryiso3code": "CHL", "indicator": {"id": ep3.INDICATOR}, "date": "2023", "value": 7.58},
        {"countryiso3code": "CHL", "indicator": {"id": ep3.INDICATOR}, "date": "2025", "value": 4.21},
        {"countryiso3code": "CHL", "indicator": {"id": ep3.INDICATOR}, "date": "2024", "value": None},
    ]]


def test_reference_requires_expected_country_and_finite_unique_observations():
    value = payload()
    assert ep3.parse_reference(value) == [("2025", 4.21), ("2023", 7.58)]
    for key, invalid in (("countryiso3code", "ARG"), ("date", "2023"), ("value", float("inf"))):
        corrupted = deepcopy(value)
        corrupted[1][1][key] = invalid
        with pytest.raises(ValueError):
            ep3.parse_reference(corrupted)


@pytest.fixture
def run_id(tmp_path, monkeypatch):
    real = ep3.load_config
    def config(*args, **kwargs):
        return dict(real(ep3.ROOT, require_secrets=False), PROJECT_ROOT=str(tmp_path))
    monkeypatch.setattr(ep3, "load_config", config)
    run_id = ep3.initialize(tmp_path)
    yield run_id
    bridge._sessions.pop(run_id, None)


def test_api_stops_after_non_retryable_error_and_records_it(run_id, monkeypatch):
    attempts = []
    def open_failed(request, timeout):
        attempts.append(request.full_url)
        raise HTTPError(request.full_url, 404, "Not found", {}, None)
    monkeypatch.setattr(ep3, "urlopen", open_failed)
    with pytest.raises(RuntimeError, match="404"):
        ep3.extract_reference(run_id)
    assert attempts == [ep3.ENDPOINT]
    assert ep3.finalize(run_id)["status"] == "ERROR"


def test_historical_preview_does_not_consume_live_queue_attempts(run_id, monkeypatch):
    s = bridge.session(run_id)
    s["ops"] = {"usar_archivo_local": True}
    s["cfg"]["TESTFIRE_PASSWORD"] = "REEMPLAZAR"
    monkeypatch.setattr(ep3.ep3_ops, "job_request", lambda *_: pytest.fail("Unexpected queue request"))
    ep3.enqueue(run_id)
    assert ep3.claim(run_id)
    assert s["queue_state"] == "VISTA_HISTORICA"


def test_mail_requires_excel_and_borrador_stays_partial(run_id, monkeypatch):
    s = bridge.session(run_id)
    s["master"] = {}
    s["reference"] = (ep3.parse_reference(payload()), "2026-11-23T15:00:00Z")
    ep3.integrate_excel(run_id)
    called = []
    def office(action, folder):
        called.append(action)
        if action == "Outlook":
            (folder / "outlook.json").write_text(json.dumps({"estado": "BORRADOR"}))
    monkeypatch.setattr(ep3, "office_action", office)
    monkeypatch.setattr(ep3.ep3_ops, "mailbox_mode", lambda _: "BORRADOR")
    with pytest.raises(ValueError, match="Microsoft Excel"):
        ep3.mail_and_classify(run_id)
    assert not called
    ep3.open_excel(run_id)
    ep3.mail_and_classify(run_id)
    outcome = ep3.finalize(run_id)
    assert called == ["Excel", "Outlook"]
    assert outcome["status"] == "PARCIAL"
    assert outcome["outlook"] == "BORRADOR"
