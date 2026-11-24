"""Cola EP3 central. SQLite se aloja únicamente en el servidor, detrás de la API."""
from contextlib import closing
from datetime import date
import json
import hmac
from pathlib import Path
import re
import sqlite3
import time
from uuid import uuid4

from flask import Flask, abort, jsonify, request

ROOT = Path(__file__).resolve().parents[1]
MAX_ATTEMPTS = 3
LEASE = 900
TOKEN_PATTERN = r"[A-Za-z0-9_-]{32,128}"
KEY_PATTERN = r"(testfire|demo|vista-testfire|vista-demo):(20[0-9]{2}-[0-9]{2}-[0-9]{2}):(20[0-9]{2}-[0-9]{2}-[0-9]{2})"
WORKER_PATTERN = r"[a-zA-Z0-9_-]{1,64}"
UUID_PATTERN = r"[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}"


def create_app(path, secret, clock=time.time, archive=None):
    if not re.fullmatch(TOKEN_PATTERN, secret):
        raise ValueError("Configurar un token compartido aleatorio de al menos 32 caracteres")
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with closing(sqlite3.connect(path)) as db, db:
        db.execute("PRAGMA journal_mode=WAL")
        db.execute("""CREATE TABLE IF NOT EXISTS jobs(
            id INTEGER PRIMARY KEY, work_key TEXT UNIQUE NOT NULL,
            state TEXT NOT NULL DEFAULT 'PENDIENTE',
            worker TEXT, run_id TEXT, claim TEXT, attempts INTEGER NOT NULL DEFAULT 0,
            available REAL NOT NULL, lease REAL, detail TEXT NOT NULL DEFAULT ''
        )""")
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 4096

    @app.before_request
    def authenticate():
        header = request.headers.get("Authorization", "")
        if not hmac.compare_digest(header.encode(), ("Bearer " + secret).encode()):
            abort(401)

    @app.get("/archivo-demo")
    def get_archive():
        start, end = request.args.get("start"), request.args.get("end")
        if archive is None or (start, end) != ("2026-10-01", "2026-10-31"):
            abort(404)
        rows = json.loads(Path(archive).read_text(encoding="utf-8"))
        return jsonify(start=start, end=end, rows=rows)

    @app.post("/queue/<action>")
    def operate(action):
        data = request.get_json()
        if not isinstance(data, dict):
            abort(400)
        key, worker, run_id, claim = (data.get(k, "") for k in ("key", "worker", "run_id", "claim"))
        matched = re.fullmatch(KEY_PATTERN, key) if isinstance(key, str) else None
        if not matched:
            abort(400)
        try:
            if date.fromisoformat(matched[2]) > date.fromisoformat(matched[3]):
                abort(400)
        except ValueError:
            abort(400)
        for value, pattern in ((worker, WORKER_PATTERN), (run_id, UUID_PATTERN)):
            if not isinstance(value, str) or not re.fullmatch(pattern, value):
                abort(400)
        now = clock()
        with closing(sqlite3.connect(path, timeout=8)) as db, db:
            db.row_factory = sqlite3.Row
            db.execute("BEGIN IMMEDIATE")
            db.execute("""
                UPDATE jobs SET state=CASE WHEN attempts>=? THEN 'AGOTADO' ELSE 'PENDIENTE' END,
                                available=?, detail='LEASE_VENCIDO', claim=NULL
                WHERE work_key=? AND state='EN_PROCESO' AND lease<=?
            """, (MAX_ATTEMPTS, now, key, now))
            job = db.execute("SELECT * FROM jobs WHERE work_key=?", (key,)).fetchone()
            if action == "enqueue":
                if job:
                    return jsonify(state=job["state"], added=False)
                db.execute("INSERT INTO jobs(work_key, available) VALUES (?,?)", (key, now))
                return jsonify(state="PENDIENTE", added=True), 201
            if job is None:
                abort(404)
            if action == "claim":
                if job["state"] == "ENVIO_RESERVADO" and job["worker"] == worker:
                    return jsonify(state="RECUPERAR_CORREO", run_id=job["run_id"], claim=job["claim"])
                if job["state"] != "PENDIENTE" or job["available"] > now:
                    return jsonify(state=job["state"], claimed=False)
                claim = str(uuid4())
                db.execute("""UPDATE jobs SET state='EN_PROCESO', worker=?, run_id=?, claim=?,
                               attempts=attempts+1, lease=? WHERE id=?""",
                           (worker, run_id, claim, now + LEASE, job["id"]))
                return jsonify(state="EN_PROCESO", claimed=True, claim=claim, attempt=job["attempts"] + 1)
            if (not isinstance(claim, str) or not job["claim"] or
                    (worker, run_id, claim) != (job["worker"], job["run_id"], job["claim"])):
                abort(409)
            if action == "reserve":
                if job["state"] != "EN_PROCESO":
                    abort(409)
                state, detail, available = "ENVIO_RESERVADO", "ESPERA_ACUSE", now
            elif action == "finish":
                outcome = data.get("outcome")
                if outcome == "CLASIFICADO" and job["state"] in ("ENVIO_RESERVADO", "COMPLETADO"):
                    state, detail, available = "COMPLETADO", "ACUSE_CLASIFICADO", now
                elif outcome == "SIN_ENVIO" and job["state"] == "EN_PROCESO":
                    state, detail, available = "COMPLETADO", "CORRIDA_SIN_DESPACHO", now
                elif outcome == "REINTENTAR" and job["state"] == "EN_PROCESO":
                    state = "AGOTADO" if job["attempts"] >= MAX_ATTEMPTS else "PENDIENTE"
                    detail, available = "FALLO_ANTES_DEL_DESPACHO", now + 30 * job["attempts"]
                elif outcome == "REINTENTAR" and job["state"] == "ENVIO_RESERVADO":
                    return jsonify(state=job["state"])
                else:
                    abort(409)
            else:
                abort(404)
            db.execute("UPDATE jobs SET state=?, detail=?, available=? WHERE id=?",
                       (state, detail, available, job["id"]))
            return jsonify(state=state)

    return app


if __name__ == "__main__":
    from auxiliaries.ep3_ops import load_local
    settings = load_local(ROOT)
    archive = ROOT / "data/movimientos_historicos_altoro.json"
    create_app(ROOT / "EP3" / "cola" / "ep3.sqlite3", settings["token"],
               archive=archive if settings.get("servir_archivo_demo") else None).run(
        host="127.0.0.1", port=18103, debug=False, use_reloader=False)

