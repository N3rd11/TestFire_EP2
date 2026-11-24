"""Cliente de cola, tipado del contrato de respaldo y política de recuperación."""
import json
import sys
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener

from auxiliaries.validation import parse_date


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def load_local(root):
    path = Path(root) / "EP3" / "operacion.local.json"
    if not path.exists():
        return {}
    settings = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(settings, dict):
        raise ValueError("Configuración de operación inválida")
    return settings


def fetch(url, token, payload=None, limit=1_000_000, https_only=False):
    parsed = urlsplit(url)
    local = parsed.hostname in ("127.0.0.1", "localhost")
    if (not parsed.hostname or parsed.username or parsed.password or parsed.fragment
            or parsed.scheme != "https" and (https_only or not local or parsed.scheme != "http")):
        raise ValueError("Usar HTTPS o HTTP de loopback para la cola local")
    headers = {"Accept": "application/json"}
    if token:
        headers["Authorization"] = "Bearer " + token
    raw = None
    if payload is not None:
        raw = json.dumps(payload).encode()
        headers["Content-Type"] = "application/json"
    try:
        with build_opener(NoRedirect()).open(Request(url, data=raw, headers=headers), timeout=15) as response:
            if response.headers.get_content_type() != "application/json":
                raise ValueError("Respuesta no JSON")
            body = response.read(limit + 1)
    except HTTPError as exc:
        raise RuntimeError(f"Servicio de operación respondió HTTP {exc.code}") from None
    if len(body) > limit:
        raise ValueError("Respuesta supera el límite")
    return json.loads(body)


def mailbox_mode(root):
    path = Path(root) / "EP3" / "outlook.local.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8-sig"))["mode"]
    return "OMITIDO"


def job_request(s, action, **fields):
    settings = s["ops"]
    payload = dict(key=s["job_key"], worker=settings["worker"], run_id=s["audit"].id,
                   claim=s.get("claim", ""))
    payload.update(fields)
    if settings.get("embedded") is True:
        root = Path(s["cfg"]["PROJECT_ROOT"])
        if "queue_app" not in s:
            runtime = root / "EP3" / ".runtime-correo"
            if runtime.is_dir() and str(runtime) not in sys.path:
                sys.path.append(str(runtime))
            from api.ep3_queue import create_app
            s["queue_app"] = create_app(root / "EP3" / "cola" / "gmail-demo.sqlite3", settings["token"])
        with s["queue_app"].test_client() as local:
            result = local.post("/queue/" + action, json=payload,
                                headers={"Authorization": f"Bearer {settings['token']}"})
            if not 200 <= result.status_code < 300:
                raise RuntimeError(f"La cola local devolvió HTTP {result.status_code}")
            return result.json
    return fetch(settings["queue_url"].rstrip("/") + "/queue/" + action, settings["token"], payload)


def load_backup(s):
    settings = s["ops"]
    cfg = s["cfg"]
    start, end = cfg["PERIOD_START"], cfg["PERIOD_END"]
    if settings.get("source_api_url"):
        url = settings["source_api_url"]
        # El endpoint autorizado recibe inicio y fin y devuelve el contrato documentado.
        join = "&" if "?" in url else "?"
        data = fetch(f"{url}{join}start={start}&end={end}", settings.get("source_api_token"), https_only=True)
        origin = "API_OPERATIVA"
    elif settings.get("usar_archivo_demo") is True:
        url = settings["queue_url"].rstrip("/") + f"/archivo-demo?start={start}&end={end}"
        data = fetch(url, settings["token"])
        origin = "API_ARCHIVO_HISTORICO"
    elif settings.get("usar_archivo_local") is True and (start, end) == ("2026-10-01", "2026-10-31"):
        archive = Path(cfg["PROJECT_ROOT"]) / "data" / "movimientos_historicos_altoro.json"
        data = dict(start=start, end=end, rows=json.loads(archive.read_text(encoding="utf-8")))
        origin = "ARCHIVO_LOCAL_HISTORICO"
    else:
        raise ValueError("No hay API de movimientos autorizada para el respaldo")
    if (not isinstance(data, dict) or data.get("start") != start or data.get("end") != end
            or not isinstance(data.get("rows"), list) or not 1 <= len(data["rows"]) <= 5000):
        raise ValueError("La API de respaldo no cumple el contrato del periodo")
    rows = []
    seen = set()
    fields = ("account", "date_raw", "description", "amount_raw", "web_id")
    for index, item in enumerate(data["rows"], 1):
        if (not isinstance(item, dict) or any(not isinstance(item.get(k), str) or len(item[k]) > 255
                                               for k in fields)):
            raise ValueError("Fila de respaldo incompleta")
        day = parse_date(item["date_raw"], "testfire_minute")
        if not start <= day <= end or not item["web_id"].strip() or item["web_id"] in seen:
            raise ValueError("Fecha fuera del periodo o identificador duplicado en el respaldo")
        seen.add(item["web_id"])
        rows.append(dict({k: item[k] for k in fields}, row_index=index))
    return origin, rows

