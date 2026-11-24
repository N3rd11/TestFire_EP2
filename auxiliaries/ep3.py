"""Etapas del robot EP3, compartidas por Studio y la prueba fuera de Studio."""

import json
import math

from datetime import datetime, timezone
from decimal import Decimal
from hashlib import sha256
import sys
from pathlib import Path
import subprocess
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
from urllib.request import Request, urlopen

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from auxiliaries import bridge, ep3_analysis, ep3_gmail, ep3_mail_local, ep3_ops, ep3_smtp, studio_web
from auxiliaries.audit import Audit, utc_now
from auxiliaries.config import load_config
from auxiliaries.excel import write_report
from auxiliaries.reconciliation import classify_pair, counts

ROOT = Path(__file__).resolve().parents[1]
INDICATOR = "FP.CPI.TOTL.ZG"
ENDPOINT = f"https://api.worldbank.org/v2/country/CHL/indicator/{INDICATOR}?format=json&mrnev=3"
REFERENCE = "Referencia externa"
MOVEMENTS = "Movimientos"
ATTACHMENT = "Reporte_TestFire_EP3.xlsx"
ARCHIVE_ORIGINS = {"API_ARCHIVO_HISTORICO", "ARCHIVO_LOCAL_HISTORICO"}
GMAIL_TRANSPORTS = {"GMAIL_API", "GMAIL_SMTP"}


def initialize(root, demo=False, local_mail=False):
    if local_mail and not demo:
        raise ValueError("El correo local requiere datos de prueba")
    cfg = load_config(root, require_secrets=False)
    cfg.update(output_path=str(root / "EP3" / "ejecuciones"), ROBOT_VERSION="EP3-1.12")
    if demo:
        cfg.update(PERIOD_START="2026-10-01", PERIOD_END="2026-10-31",
                   ALLOW_SYNTHETIC_MASTER="true",
                   master_path=str(root / "data" / "conciliacion_master.xlsx"))
    audit = Audit(cfg, mode="SINTETICO" if demo else "TESTFIRE_WEB")
    mail = {"transport": "SMTP_LOCAL"} if local_mail else mail_settings(root)
    bridge._sessions[audit.id] = dict(
        cfg=cfg, audit=audit, valid=[], rejected=[], results=[], rules={}, master=None,
        report=None, register=None, reference=None, excel_open=False, xlsx_verified=False, mail="OMITIDO",
        mail_config=mail, duoc_recipient=mail.get("destino_duoc") if not demo else None,
        origin="SINTETICO" if demo else "WEB",
        ops={} if local_mail else ep3_ops.load_local(root), claim="", job_key="", queue_state="LOCAL",
    )
    return audit.id


def mail_settings(root):
    path = Path(root) / "EP3" / "outlook.local.json"
    if path.is_file():
        return json.loads(path.read_text(encoding="utf-8-sig"))
    return {}


def gmail_client(s):
    runtime = str(ROOT / "EP3" / ".runtime-correo")
    if runtime not in sys.path and Path(runtime).is_dir():
        sys.path.append(runtime)
    return ep3_gmail.client(ROOT, s["mail_config"]["mailbox"])


def webmail_client(s):
    if s["mail_config"].get("transport") == "GMAIL_SMTP":
        client = ep3_smtp.client(ROOT, s["mail_config"]["mailbox"])
        client.check_account()
        return client
    return gmail_client(s)


def report_digest(folder):
    artifact = json.loads((folder / "artefacto.json").read_text(encoding="utf-8"))
    if artifact["execution_id"] != folder.name:
        raise ValueError("El manifiesto corresponde a otra ejecución")
    report = folder / ATTACHMENT
    if sha256(report.read_bytes()).hexdigest() != artifact["sha256"]:
        raise ValueError("El XLSX ya no coincide con su huella de auditoría")
    return artifact["sha256"]


def enqueue(eid):
    s = bridge.session(eid)
    if s.get("duoc_recipient"):
        s["queue_state"] = "RESERVA_POR_EJECUCION"
        return
    if not s["ops"] or s["ops"].get("solo_respaldo"):
        return
    if (s["origin"] == "WEB" and missing_testfire_credentials(s["cfg"])
            and not s["ops"].get("source_api_url")
            and (s["ops"].get("usar_archivo_demo") or s["ops"].get("usar_archivo_local"))):
        s["queue_state"] = "VISTA_HISTORICA"
        return
    with s["audit"].step("ENCOLAR"):
        cfg = s["cfg"]
        kind = "demo" if s["origin"] == "SINTETICO" else "testfire"
        if ep3_ops.mailbox_mode(ROOT) != "ENVIO":
            kind = "vista-" + kind
        s["job_key"] = f"{kind}:{cfg['PERIOD_START']}:{cfg['PERIOD_END']}"
        result = ep3_ops.job_request(s, "enqueue")
        s["queue_state"] = result["state"]
        s["audit"].write_json("cola_entrada.json", dict(clave=s["job_key"], **result))


def claim(eid):
    s = bridge.session(eid)
    if not s["job_key"]:
        return True
    with s["audit"].step("LEER_COLA"):
        result = ep3_ops.job_request(s, "claim")
        s["queue_state"] = result["state"]
        if result.get("claimed"):
            s["claim"] = result["claim"]
        if result["state"] == "RECUPERAR_CORREO":
            recover_receipt(s, result)
        s["audit"].write_json("cola_lectura.json", dict(clave=s["job_key"], estado=s["queue_state"]))
        return bool(s["claim"])


def recover_receipt(s, result):
    from auxiliaries.validation import valid_uuid

    audit = s["audit"]
    prior_id = valid_uuid(result["run_id"])
    folder = audit.folder.parent / prior_id
    state = folder / "outlook.json"
    if not state.is_file():
        return
    if state.stat().st_size == 0:
        return
    stored = ep3_gmail.read_state(folder)
    current = stored["estado"]
    if current in ("INTENTO_REGISTRADO", "ENVIADO_PENDIENTE"):
        if stored.get("canal") in GMAIL_TRANSPORTS:
            if stored["canal"] != s["mail_config"].get("transport"):
                raise RuntimeError("Restablece el mismo transporte usado para el envío pendiente")
            webmail_client(s).confirm(folder, report_digest(folder))
        else:
            office_action("Outlook", folder, receipt_only=True)
        current = json.loads(state.read_text(encoding="utf-8-sig"))["estado"]
    if current == "CLASIFICADO":
        response = ep3_ops.job_request(s, "finish", run_id=prior_id,
                                      claim=result["claim"], outcome="CLASIFICADO")
        s["queue_state"] = response["state"]
        if stored.get("canal") in GMAIL_TRANSPORTS:
            s.update(mail=current, report=folder / ATTACHMENT, xlsx_verified=True, mail_recovered_from=prior_id)


def recover_source(eid):
    s = bridge.session(eid)
    audit = s["audit"]
    for event in list(audit.active):
        audit.end(event, "ADVERTENCIA", "Login o extracción web no disponible")
    with audit.step("RESPALDO_MOVIMIENTOS"):
        s["origin"], rows = ep3_ops.load_backup(s)
        audit.write_json("origen_movimientos.json", {"origen": s["origin"], "filas": len(rows)})
    return rows


def missing_testfire_credentials(cfg):
    return [key for key in ("TESTFIRE_USER", "TESTFIRE_PASSWORD")
            if not cfg.get(key) or cfg[key] == "REEMPLAZAR"]


def use_web_source(eid):
    s = bridge.session(eid)
    has_backup = any(s["ops"].get(key) for key in
                     ("source_api_url", "usar_archivo_demo", "usar_archivo_local"))
    if has_backup and missing_testfire_credentials(s["cfg"]):
        audit = s["audit"]
        event = audit.start("CONFIGURACION_TESTFIRE")
        audit.end(event, "OMITIDO", "TestFire sin cuenta de demostración. Se usa el respaldo configurado.")
        return False
    return True


def login_config(eid):
    s = bridge.session(eid)
    cfg = s["cfg"]
    missing = missing_testfire_credentials(cfg)
    if missing:
        event = s["audit"].start("CONFIGURACION_TESTFIRE")
        detail = "Falta configurar " + " y ".join(missing) + " en .env"
        s["audit"].end(event, "ADVERTENCIA", detail)
        raise ValueError(detail)
    return cfg


def reconcile_demo(eid):
    with bridge.session(eid)["audit"].step("EXTRACCION_SINTETICA"):
        raw = json.loads((ROOT / "data" / "filas_demo.json").read_text(encoding="utf-8"))
    reconcile_movements(eid, raw, date_format="iso")


def reconcile_movements(eid, rows, date_format="testfire_minute"):
    audit = bridge.session(eid)["audit"]
    with audit.step("VALIDATE_DATA"):
        if not isinstance(rows, list) or not rows:
            raise ValueError("La fuente de movimientos no devolvió filas")
        for row in rows:
            verdict = bridge.validate(eid, row, date_format)
            if verdict["valid"]:
                bridge.append_valid(eid, verdict["row"])
            else:
                bridge.append_rejected(eid, verdict)
    with audit.step("READ_REFERENCES"):
        studio_web.read_reference(eid)
    with audit.step("RECONCILE"):
        for pair in bridge.pairs(eid):
            bridge.append_result(eid, pair, classify_pair(pair))


def parse_reference(payload):
    if (not isinstance(payload, list) or len(payload) != 2
            or not isinstance(payload[0], dict) or not isinstance(payload[1], list)
            or payload[0].get("pages") != 1):
        raise ValueError("Estructura o paginación inesperada del Banco Mundial")
    values = {}
    for record in payload[1]:
        if not isinstance(record, dict):
            raise ValueError("Observación no válida")
        indicator = record.get("indicator")
        if (not isinstance(indicator, dict) or indicator.get("id") != INDICATOR
                or record.get("countryiso3code") != "CHL"):
            raise ValueError("País o indicador distinto del solicitado")
        value = record.get("value")
        if value is None:
            continue
        year = record.get("date", "")
        if (not isinstance(year, str) or len(year) != 4 or not year.isdigit()
                or not 1960 <= int(year) <= datetime.now(timezone.utc).year
                or year in values or type(value) not in (int, float) or not math.isfinite(value)):
            raise ValueError("Año duplicado o dato anual inválido")
        values[year] = value
    if not 1 <= len(values) <= 3:
        raise ValueError("Cantidad de observaciones utilizables fuera del contrato")
    return sorted(values.items(), reverse=True)


def extract_reference(eid):
    s = bridge.session(eid)
    audit = s["audit"]
    with audit.step("API_PUBLICA"):
        request = Request(ENDPOINT, headers={"Accept": "application/json"})
        for attempt in range(1, 4):
            try:
                with urlopen(request, timeout=12) as response:
                    final_url = urlsplit(response.url)
                    if (final_url.scheme, final_url.hostname) != ("https", "api.worldbank.org"):
                        raise ValueError("La respuesta no proviene del endpoint HTTPS esperado")
                    code = response.status
                    if code != 200 or response.headers.get_content_type() != "application/json":
                        raise ValueError(f"Respuesta inesperada de API, HTTP {code}")
                    raw = response.read(100_001)
                    if len(raw) > 100_000:
                        raise ValueError("Respuesta excede el límite de tamaño")
                    payload = json.loads(raw)
            except (HTTPError, URLError, TimeoutError) as exc:
                code = exc.code if isinstance(exc, HTTPError) else None
                retryable = code is None or code in {408, 429, 500, 502, 503, 504}
                event = audit.start("HTTP_BANCO_MUNDIAL", attempt)
                audit.end(event, "ADVERTENCIA", f"HTTP {code or 'sin respuesta'}")
                if not retryable or attempt == 3:
                    raise RuntimeError(f"Banco Mundial no disponible, HTTP {code or 'sin respuesta'}") from exc
                time.sleep(attempt)
            else:
                event = audit.start("HTTP_BANCO_MUNDIAL", attempt)
                audit.end(event, detail="HTTP 200")
                break
        observations = parse_reference(payload)
        audit.write_json("api_banco_mundial.json", payload)
        s["reference"] = (observations, utc_now())
        print(f"EP3 API HTTP 200 | {len(observations)} observaciones")


def register_movements(eid):
    s = bridge.session(eid)
    audit = s["audit"]
    with audit.step("REGISTRO_LOCAL_MOVIMIENTOS"):
        if s["master"] is None:
            raise ValueError("Falta la conciliación")
        register = audit.folder / "Registro_Movimientos_EP3.xlsx"
        write_report(register, s["valid"], s["results"], s["rejected"], audit.events,
                     {"ejecucion": eid, "origen": s["origin"], "validos": len(s["valid"]),
                      "rechazados": len(s["rejected"])})
        s["register"] = register
        audit.write_json("registro_artefacto.json", audit.artifact(register, "REGISTRO_LOCAL"))


def integrate_excel(eid):
    s = bridge.session(eid)
    audit = s["audit"]
    with audit.step("INTEGRACION_EXCEL"):
        if s["master"] is None:
            raise ValueError("Falta la extracción y conciliación")
        observations, fetched_at = s["reference"]
        wb = Workbook()
        movements = wb.active
        movements.title = MOVEMENTS
        summary = wb.create_sheet("Resumen")
        summary.append(("Indicador", "Valor"))
        source = "Altoro Mutual (captura anterior)" if s["origin"] in ARCHIVE_ORIGINS else {
            "WEB": "Altoro Mutual",
            "SINTETICO": "Datos de prueba",
            "API_OPERATIVA": "API operativa de movimientos",
        }[s["origin"]]
        for row in (
            ("Ejecución", eid), ("Origen", s["origin"]),
            ("Filas válidas", len(s["valid"])), ("Filas rechazadas", len(s["rejected"])),
            *sorted(counts(s["results"]).items()),
            ("Periodo desde", s["cfg"]["PERIOD_START"]), ("Periodo hasta", s["cfg"]["PERIOD_END"]),
            ("Fuente de movimientos", source),
        ):
            summary.append(row)
        movements.append(("Cuenta", "ID de transacción", "Fecha y hora", "Operación original", "Importe", "Moneda"))
        for movement in s["valid"]:
            stamp = movement.get("date_raw") or movement["date"]
            movements.append((movement["account"], movement.get("web_id"), datetime.fromisoformat(stamp),
                              movement["description"], Decimal(movement["amount"]), ep3_analysis.currency_label(movement)))
            account, transaction, date_cell, description, amount, currency = movements[movements.max_row]
            for cell in (account, transaction, description, currency):
                if isinstance(cell.value, str):
                    cell.data_type = "s"
            date_cell.number_format = "yyyy-mm-dd hh:mm:ss" if len(stamp) > 10 else "yyyy-mm-dd"
            amount.number_format = "#,##0.00;[Red]-#,##0.00"
        external = wb.create_sheet(REFERENCE)
        external.append(("País", "Indicador", "Año", "Inflación anual", "Fuente", "Consulta UTC"))
        for year, percent in observations:
            external.append(("Chile", INDICATOR, int(year), percent / 100, ENDPOINT, fetched_at))
            external.cell(external.max_row, 4).number_format = "0.00%"
        for ws, widths in ((summary, [30, 42]), (movements, [16, 22, 24, 35, 21, 19]),
                           (external, [14, 23, 12, 20, 90, 35])):
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell, width in zip(ws[1], widths):
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="16324F")
                ws.column_dimensions[cell.column_letter].width = width
        path = audit.folder / ATTACHMENT
        temporary = path.with_suffix(".tmp.xlsx")
        try:
            wb.save(temporary)
            saved = load_workbook(temporary, read_only=True)
            try:
                if (saved.sheetnames != [MOVEMENTS, "Resumen", REFERENCE] or saved.active.title != MOVEMENTS
                        or saved["Resumen"]["B2"].value != eid
                        or saved[MOVEMENTS].max_row != 1 + len(s["valid"])
                        or list(saved[REFERENCE].values)[1][2] != int(observations[0][0])):
                    raise ValueError("Verificación de persistencia XLSX falló")
            finally:
                saved.close()
            temporary.replace(path)
        finally:
            wb.close()
            temporary.unlink(missing_ok=True)
        s["report"] = path
        audit.write_json("artefacto.json", audit.artifact(path, "EXCEL"))
        print(f"EP3 Excel | {path}")


def office_action(action, folder, allow_send=False, receipt_only=False):
    command = [
        "powershell.exe", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass",
        "-File", str(ROOT / "EP3" / "office_ep3.ps1"), "-Action", action,
        "-RunFolder", str(folder),
    ]
    if allow_send:
        command.append("-AllowSend")
    if receipt_only:
        command.append("-ReceiptOnly")
    try:
        subprocess.run(command, timeout=100, check=True)
    except subprocess.CalledProcessError as exc:
        if exc.returncode == 12:
            raise RuntimeError(f'{action} de escritorio no esta registrado en Windows') from None
        raise RuntimeError(f'Error en {action} de escritorio, codigo {exc.returncode}') from None


def open_excel(eid):
    s = bridge.session(eid)
    library_transports = GMAIL_TRANSPORTS | {"SMTP_LOCAL"}
    if s["origin"] in ARCHIVE_ORIGINS or s["mail_config"].get("transport") in library_transports:
        with s["audit"].step("APERTURA_XLSX_LIBRERIA"):
            folder = s["audit"].folder
            digest = report_digest(folder)
            wb = load_workbook(folder / ATTACHMENT, read_only=True, data_only=True)
            try:
                if wb["Resumen"]["B2"].value != eid or wb["Resumen"]["B3"].value != s["origin"]:
                    raise ValueError("El XLSX no corresponde a la ejecución o a su origen")
            finally:
                wb.close()
            s["xlsx_verified"] = True
            s["audit"].write_json("excel.json", {"estado": "LEIDO_Y_VALIDADO", "motor": "openpyxl", "sha256": digest})
        return
    with s["audit"].step("APERTURA_EXCEL"):
        office_action("Excel", s["audit"].folder)
        s["excel_open"] = True


def mail_and_classify(eid):
    s = bridge.session(eid)
    if s.get("duoc_recipient"):
        send_duoc(s)
        return
    if s["origin"] in ARCHIVE_ORIGINS:
        with s["audit"].step("CORREO_OMITIDO_ARCHIVO_HISTORICO"):
            s["mail"] = "ARCHIVO_SIN_DESPACHO"
        return
    if s["mail_config"].get("transport") == "SMTP_LOCAL":
        audit = s["audit"]
        with audit.step("CORREO_SMTP_LOCAL"):
            if not s["xlsx_verified"] or s["origin"] != "SINTETICO":
                raise ValueError("El buzon local requiere un XLSX de prueba validado")
            digest = report_digest(audit.folder)
            receipt = ep3_mail_local.deliver(audit.folder, s["report"], digest)
            audit.write_json("outlook.json", receipt)
            s["mail"] = receipt["estado"]
            print(f"EP3 SMTP local | {receipt['codigo_smtp']} | {s['mail']}")
        return
    if s["mail_config"].get("transport") in GMAIL_TRANSPORTS:
        gmail_and_classify(s)
        return
    if not (s["excel_open"] and s["report"]):
        raise ValueError("Microsoft Excel debe abrir el reporte antes del envío")
    audit = s["audit"]
    with audit.step("OUTLOOK"):
        send = ep3_ops.mailbox_mode(ROOT) == "ENVIO"
        if send:
            if not s["claim"] or s["job_key"].startswith("vista-"):
                raise RuntimeError("ENVIO requiere una reserva vigente en la cola compartida")
            office_action("OutlookCheck", audit.folder)
            response = ep3_ops.job_request(s, "reserve")
            s["queue_state"] = response["state"]
            if s["queue_state"] != "ENVIO_RESERVADO":
                raise RuntimeError("La cola no autorizó el despacho")
        try:
            if send:
                office_action("Outlook", audit.folder, allow_send=True)
            else:
                office_action("Outlook", audit.folder)
        finally:
            state = audit.folder / "outlook.json"
            if state.is_file():
                s["mail"] = json.loads(state.read_text(encoding="utf-8-sig"))["estado"]
        print(f"EP3 Outlook | {s['mail']}")


def send_duoc(s):
    audit = s["audit"]
    with audit.step("CORREO_SMTP_DUOC"):
        settings = s["mail_config"]
        if settings.get("transport") != "GMAIL_SMTP" or settings.get("mode") != "ENVIO":
            raise ValueError("El envío a Duoc requiere una cuenta Gmail SMTP en modo ENVIO")
        if not s["xlsx_verified"] or not s["valid"]:
            raise ValueError("Se requiere un Excel validado con movimientos")
        if s["origin"] == "WEB":
            site = urlsplit(s["cfg"]["TESTFIRE_URL"])
            if site.scheme != "https" or site.hostname != "demo.testfire.net" or not s.get("web_ok"):
                raise ValueError("El envío web requiere una extracción HTTPS válida del TestFire de demostración")
        elif s["origin"] not in ARCHIVE_ORIGINS | {"API_OPERATIVA"}:
            raise ValueError("El origen de los movimientos no está habilitado para Duoc")
        start, end = s["cfg"]["PERIOD_START"], s["cfg"]["PERIOD_END"]
        groups = ep3_analysis.summarize(s["valid"])
        classifications = counts(s["results"])
        analysis = ep3_analysis.write_analysis(audit.folder, s["origin"], start, end, s["valid"], groups,
                                             classifications)
        artifact = audit.artifact(analysis, "EXCEL_ANALISIS")
        audit.write_json("analisis_artefacto.json", artifact)
        s["analysis"] = analysis
        detail = ep3_analysis.write_movement_attachment(analysis, len(s["valid"]))
        detail_artifact = audit.artifact(detail, "EXCEL_DETALLE")
        audit.write_json("detalle_artefacto.json", detail_artifact)
        s["detail_report"] = detail
        body = ep3_analysis.email_summary(s["origin"], start, end, s["valid"], len(s["rejected"]), s["reference"])
        digest = report_digest(audit.folder)
        smtp = ep3_smtp.client(Path(s["cfg"]["PROJECT_ROOT"]), settings["mailbox"])
        try:
            smtp.send_to_duoc(audit.folder, digest, s["duoc_recipient"], s["origin"],
                              summary=body, extra_attachments=((analysis.name, artifact["sha256"]),
                                                               (detail.name, detail_artifact["sha256"])))
        finally:
            s["mail"] = ep3_gmail.read_state(audit.folder).get("estado", "OMITIDO")


def gmail_and_classify(s):
    audit = s["audit"]
    with audit.step("CORREO_" + s["mail_config"]["transport"]):
        if not s["xlsx_verified"]:
            raise ValueError("Se requiere un XLSX validado")
        if s["origin"] == "WEB":
            site = urlsplit(s["cfg"]["TESTFIRE_URL"])
            if site.scheme != "https" or site.hostname != "demo.testfire.net" or not s.get("web_ok"):
                raise ValueError("El envío web requiere una extracción HTTPS válida del TestFire de demostración")
        elif s["origin"] not in ("SINTETICO", "API_OPERATIVA"):
            raise ValueError("El origen del reporte no está habilitado para Gmail")
        if ep3_ops.mailbox_mode(ROOT) != "ENVIO":
            s["mail"] = "OMITIDO"
            return
        if not s["claim"] or s["job_key"].startswith("vista-"):
            raise RuntimeError("Gmail requiere reserva en la cola central")
        digest = report_digest(audit.folder)
        gmail = webmail_client(s)
        response = ep3_ops.job_request(s, "reserve")
        s["queue_state"] = response["state"]
        if s["queue_state"] != "ENVIO_RESERVADO":
            raise RuntimeError("No se obtuvo reserva de Gmail")
        try:
            gmail.send(audit.folder, digest, origin=s["origin"])
            gmail.confirm(audit.folder, digest)
        finally:
            state = ep3_gmail.read_state(audit.folder)
            s["mail"] = state.get("estado", "OMITIDO")


def audit_failure(eid):
    audit = bridge.session(eid)["audit"]
    for event in list(audit.active):
        audit.end(event, "ERROR", "Revisar el comando y la consola de Studio")
    audit.write_json("fallo_robot.json", {"estado": "ERROR", "etapa": "Flujo principal"})


def finalize(eid):
    s = bridge.session(eid)
    audit = s["audit"]
    if audit.active:
        audit_failure(eid)
    failure = any(e["status"] == "ERROR" for e in audit.events) or (audit.folder / "fallo_robot.json").exists()
    if s["claim"]:
        outcome = "CLASIFICADO" if s["mail"] == "CLASIFICADO" else "REINTENTAR"
        if (not failure and s["report"] and (s["excel_open"] or s["xlsx_verified"]) and s["job_key"].startswith("vista-")):
            outcome = "SIN_ENVIO"
        try:
            s["queue_state"] = ep3_ops.job_request(s, "finish", outcome=outcome)["state"]
        except (OSError, ValueError, RuntimeError):
            s["queue_state"] = "SIN_CONFIRMACION_COLA"
    mail_complete = s["mail"] == "CLASIFICADO" or (s.get("duoc_recipient") and s["mail"] == "SMTP_ACEPTADO")
    complete = all((s["report"], s["excel_open"] or s["xlsx_verified"], mail_complete,
                    s["queue_state"] != "SIN_CONFIRMACION_COLA"))
    status = "ERROR" if failure and not s["report"] else "OK" if complete and not failure else "PARCIAL"
    audit.summary.update(
        status=status, ended_at=utc_now(),
        duration_ms=round((time.perf_counter_ns() - audit.started_ns) / 1e6, 3),
        valid_rows=len(s["valid"]), rejected_rows=len(s["rejected"]),
        classifications=counts(s["results"]), api_ok=s["reference"] is not None,
        report=s["report"].name if s["report"] else None,
        register=s["register"].name if s["register"] else None,
        analysis=s["analysis"].name if s.get("analysis") else None,
        detail_report=s["detail_report"].name if s.get("detail_report") else None,
        excel_open=s["excel_open"], xlsx_verified=s["xlsx_verified"], outlook=s["mail"],
        mail_transport="GMAIL_SMTP_DUOC" if s.get("duoc_recipient") else s["mail_config"].get("transport", "OUTLOOK_DESKTOP"),
        origin=s["origin"], queue=s["queue_state"], job_key=s["job_key"],
        mail_recovered_from=s.get("mail_recovered_from"),
    )
    audit.summary["resultado"] = (
        f"FINALIZADO | {status} | {audit.summary['duration_ms'] / 1000:.2f} s | "
        f"{len(s['valid'])} movimientos | Correo: {s['mail']}"
    )
    audit.save_summary()
    print(f"EP3 {eid} | {audit.summary['resultado']}")
    return audit.summary


def run_without_studio():
    eid = initialize(ROOT, demo=True)
    try:
        reconcile_demo(eid)
        register_movements(eid)
        extract_reference(eid)
        integrate_excel(eid)
    finally:
        finalize(eid)


if __name__ == "__main__":
    run_without_studio()
