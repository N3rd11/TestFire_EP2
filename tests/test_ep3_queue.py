from concurrent.futures import ThreadPoolExecutor
from uuid import uuid4

import pytest

from api.ep3_queue import LEASE, create_app
from auxiliaries import ep3_ops

TOKEN = "y" * 40
KEY = "testfire:2026-10-01:2026-10-31"
HEADER = {"Authorization": "Bearer " + TOKEN}


@pytest.fixture
def service(tmp_path):
    now = [1000.0]
    return create_app(tmp_path / "jobs.db", TOKEN, clock=lambda: now[0]), now


def job(worker):
    return dict(key=KEY, worker=worker, run_id=str(uuid4()))


def request(app, action, data):
    with app.test_client() as client:
        response = client.post("/queue/" + action, headers=HEADER, json=data)
        return response.status_code, response.json


def test_two_pcs_enqueue_and_claim_one_job(service):
    app, _ = service
    pc_a, pc_b = job("pc-a"), job("pc-b")
    with ThreadPoolExecutor(2) as pool:
        catalog = list(pool.map(lambda p: request(app, "enqueue", p), [pc_a, pc_b]))
        claims = list(pool.map(lambda p: request(app, "claim", p), [pc_a, pc_b]))
    assert sorted(code for code, _ in catalog) == [200, 201]
    assert sum(bool(r.get("claimed")) for _, r in claims) == 1
    active, won = next((p, r) for p, (_, r) in zip([pc_a, pc_b], claims) if r.get("claimed"))
    permit = dict(active, claim=won["claim"])
    assert request(app, "reserve", permit)[1]["state"] == "ENVIO_RESERVADO"
    assert request(app, "reserve", permit)[0] == 409
    pc_other = pc_a if active == pc_b else pc_b
    assert request(app, "claim", pc_other)[1]["claimed"] is False
    recovery = request(app, "claim", active)[1]
    assert recovery["state"] == "RECUPERAR_CORREO" and recovery["run_id"] == active["run_id"]
    assert request(app, "finish", dict(permit, outcome="CLASIFICADO"))[1]["state"] == "COMPLETADO"
    assert request(app, "enqueue", pc_other)[1] == dict(state="COMPLETADO", added=False)


def test_lease_expiry_fences_stale_worker_and_caps_retries(service):
    app, clock = service
    pc_a, pc_b = job("pc-a"), job("pc-b")
    request(app, "enqueue", pc_a)
    old = request(app, "claim", pc_a)[1]["claim"]
    clock[0] += LEASE + 1
    for attempt in (2, 3):
        current = request(app, "claim", pc_b)[1]
        assert current["attempt"] == attempt
        assert request(app, "reserve", dict(pc_a, claim=old))[0] == 409
        response = request(app, "finish", dict(pc_b, claim=current["claim"], outcome="REINTENTAR"))[1]
        assert response["state"] == ("PENDIENTE" if attempt == 2 else "AGOTADO")
        clock[0] += 120
    assert not request(app, "claim", pc_a)[1]["claimed"]


def test_reserved_send_never_expires_to_another_pc(service):
    app, clock = service
    pc_a, pc_b = job("pc-a"), job("pc-b")
    request(app, "enqueue", pc_a)
    claim = request(app, "claim", pc_a)[1]["claim"]
    request(app, "reserve", dict(pc_a, claim=claim))
    clock[0] += LEASE * 10
    assert request(app, "claim", pc_b)[1] == dict(state="ENVIO_RESERVADO", claimed=False)


def test_backup_rejects_duplicate_and_labels_archive(monkeypatch):
    s = {"ops": {"queue_url": "http://127.0.0.1:18103", "token": TOKEN, "usar_archivo_demo": True},
         "cfg": {"PERIOD_START": "2026-10-01", "PERIOD_END": "2026-10-31"}}
    item = dict(account="123", date_raw="2026-10-18 12:00", description="Demo",
                amount_raw="20.00", web_id="TX1")
    payload = dict(start="2026-10-01", end="2026-10-31", rows=[item])
    monkeypatch.setattr(ep3_ops, "fetch", lambda *a, **kw: payload)
    source, rows = ep3_ops.load_backup(s)
    assert source == "API_ARCHIVO_HISTORICO" and rows[0]["row_index"] == 1
    payload["rows"].append(item)
    with pytest.raises(ValueError, match="duplicado"):
        ep3_ops.load_backup(s)


def test_queue_rejects_unauthorized_requests(service):
    app, _ = service
    with app.test_client() as client:
        assert client.post("/queue/enqueue", json=job("pc-a")).status_code == 401



def test_local_archive_has_distinct_origin_and_all_rows():
    from pathlib import Path

    settings = {
        "ops": {"usar_archivo_local": True, "solo_respaldo": True},
        "cfg": {
            "PROJECT_ROOT": str(Path(__file__).resolve().parents[1]),
            "PERIOD_START": "2026-10-01",
            "PERIOD_END": "2026-10-31",
        },
    }
    source, rows = ep3_ops.load_backup(settings)
    assert source == "ARCHIVO_LOCAL_HISTORICO"
    assert len(rows) == 100
    assert len({r["web_id"] for r in rows}) == 100

