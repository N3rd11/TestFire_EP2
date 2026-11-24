"""Genera el proyecto editable para Studio a partir del export EP2."""
from base64 import b64decode, b64encode
from copy import deepcopy
import json
from pathlib import Path
import sqlite3
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[1]
ORIGINAL = ROOT / "rocketbot" / "MAIN_TESTFIRE_EP2" / "EXPORT_TESTFIRE_EP2.db"
TARGET = ROOT / "rocketbot" / "MAIN_TESTFIRE_EP3" / "robot.db"
MAIN = "MAIN_TESTFIRE_EP3"
DEMO = "DEMO_TESTFIRE_EP3"
LOCAL_DEMO = "DEMO_CORREO_LOCAL_EP3"


def build():
    with sqlite3.connect(ORIGINAL.as_uri() + "?mode=ro", uri=True) as source:
        schema, = source.execute("SELECT sql FROM sqlite_master WHERE name='bots'").fetchone()
        projects = {
            name: json.loads(b64decode(data))["project"]
            for name, data in source.execute("SELECT name,data FROM bots")
        }
    base = projects.pop("MAIN_TESTFIRE_EP2")
    template = base["commands"][1]["children"]

    def command(model, value):
        cmd = deepcopy(model)
        cmd.update(id=str(uuid4()), command=value, children=[], **{"else": []})
        return cmd

    def py(code):
        return command(base["commands"][0], code)

    def call(name):
        return command(template[1], name)

    def operation(name):
        return py(f"from auxiliaries import ep3\nep3.{name}(GetVar('execution_id'))")

    def project(name, description, commands, variables=None):
        robot = deepcopy(base)
        robot.update(commands=commands, vars=variables or [], ifs=[])
        robot["profile"].update(name=name, description=description)
        return robot

    projects["EP3_REGISTRAR_MOVIMIENTOS"] = project(
        "EP3_REGISTRAR_MOVIMIENTOS", "Guarda movimientos, rechazos y conciliación en un registro local",
        [operation("register_movements")])
    stages = {
        "EP3_API_PUBLICA": ("Consulta y validación JSON del Banco Mundial", "extract_reference"),
        "EP3_INTEGRAR_EXCEL": ("Reporte XLSX agregado con referencia externa", "integrate_excel"),
        "EP3_ABRIR_EXCEL": ("Abre y valida el XLSX con Excel de escritorio o una libreria", "open_excel"),
        "EP3_OUTLOOK": ("Envía o verifica la recepción por SMTP local, Gmail u Outlook", "mail_and_classify"),
        "EP3_SINTETICO": ("Valida y concilia filas de prueba", "reconcile_demo"),
    }
    for name, (description, operation_name) in stages.items():
        projects[name] = project(name, description, [operation(operation_name)])
    projects["EP3_ENCOLAR"] = project("EP3_ENCOLAR", "Registra la clave del trabajo en la cola compartida",
                                     [operation("enqueue")])
    projects["EP3_LEER_COLA"] = project("EP3_LEER_COLA", "Reclama de forma atómica o recupera el acuse",
        [py("from auxiliaries import ep3\nSetVar('ep3_job_ok',ep3.claim(GetVar('execution_id')))")])
    projects["EP3_RESPALDO_API"] = project("EP3_RESPALDO_API", "Obtiene movimientos cuando falla la web",
        [py("from auxiliaries import ep3\nSetVar('filas_web',ep3.recover_source(GetVar('execution_id')))")])
    projects["EP3_CONCILIAR"] = project("EP3_CONCILIAR", "Valida y concilia los movimientos en un solo comando",
        [py("from auxiliaries import ep3\nfrom auxiliaries.studio_demo import variable\n"
            "ep3.reconcile_movements(GetVar('execution_id'), variable(GetVar('filas_web')))")])

    def attempt(commands, on_error=None):
        flow = command(base["commands"][1], "")
        flow["children"] = commands
        flow["else"] = on_error or [operation("audit_failure")]
        return flow

    flow_close = deepcopy(base["commands"][2])
    flow_close["else"] = [py("print('EP3: navegador ya cerrado o no disponible')")]
    def fallback():
        return attempt([call("EP3_RESPALDO_API"), py("SetVar('ep3_source_ok',True)")])

    web = attempt(deepcopy(template[:4]) + [py("SetVar('ep3_source_ok',True)")], [fallback()])
    source = command(template[9], "{ep3_live_web}")
    source["children"] = [web, flow_close, py("SetVar('usuario','')\nSetVar('clave','')")]
    source["else"] = [fallback()]
    choose_source = py("from auxiliaries import ep3\nSetVar('ep3_live_web',ep3.use_web_source(GetVar('execution_id')))")
    gate = command(template[9], "{ep3_source_ok}")
    gate["children"] = [call("EP3_CONCILIAR"), call("EP3_REGISTRAR_MOVIMIENTOS"),
                        *[call(name) for name in list(stages)[:4]]]
    pipeline = attempt([gate])
    root_setup = (
        "import importlib\nimport sys\nfrom pathlib import Path\n"
        "root=Path(r'%project_path%').resolve().parents[1]\n"
        "if str(root) not in sys.path: sys.path.insert(0,str(root))\n"
        "from auxiliaries import ep3_analysis, ep3_gmail, ep3_ops\nimportlib.reload(ep3_analysis)\nimportlib.reload(ep3_gmail)\nimportlib.reload(ep3_ops)\nfrom auxiliaries import ep3_smtp\nimportlib.reload(ep3_smtp)\nfrom auxiliaries import ep3\nimportlib.reload(ep3)\nSetVar('ep3_source_ok',False)\nSetVar('execution_id','')\nSetVar('tabla_web','')\n"
    )
    live_login = (
        "from auxiliaries import ep3\nc=ep3.login_config(GetVar('execution_id'))\n"
        "SetVar('url_login',c['TESTFIRE_URL'].rstrip('/')+'/login.jsp')\n"
        "SetVar('usuario',c['TESTFIRE_USER'])\nSetVar('clave',c['TESTFIRE_PASSWORD'])"
    )
    web["children"].insert(1, py(live_login))

    def initial(demo, local_mail=False):
        return py(root_setup + "SetVar('resultado_ep3','EN CURSO')\n"
                  f"SetVar('execution_id',ep3.initialize(root,demo={demo},local_mail={local_mail}))")

    def finish():
        cmd = py("from auxiliaries import ep3\n"
                 "SetVar('resultado_ep3',ep3.finalize(GetVar('execution_id'))['resultado'])")
        cmd["description"] = "Fin de ejecución. Ver resultado_ep3 en Variables"
        return cmd

    clean_vars = [
        item for item in base["vars"]
        if item["name"] not in {"respuesta", "http_valid", "payload", "intento", "retryable",
                                "acuse_ok", "reporte_rel", "PROJECT_ROOT"}
    ] + [dict(name="ep3_source_ok", data="", type="string", category="", disabled=False)]
    clean_vars.append(dict(name="ep3_job_ok", data="", type="string", category="", disabled=False))
    clean_vars.append(dict(name="ep3_live_web", data="", type="string", category="", disabled=False))
    clean_vars.append(dict(name="resultado_ep3", data="", type="string", category="", disabled=False))
    queue_commands = [call("EP3_ENCOLAR"), call("EP3_LEER_COLA")]
    live_job = command(template[9], "{ep3_job_ok}")
    live_job["children"] = [choose_source, source, pipeline]
    projects[MAIN] = project(MAIN, "TestFire, registro, Excel y correo institucional",
        [initial(False), attempt(queue_commands + [live_job]), finish()], clean_vars)
    demo_job = command(template[9], "{ep3_job_ok}")
    demo_job["children"] = [call("EP3_SINTETICO"), call("EP3_REGISTRAR_MOVIMIENTOS"), *[call(name) for name in list(stages)[:4]]]
    projects[DEMO] = project(DEMO, "Cola opcional, datos sintéticos y consulta API real",
        [initial(True), attempt(deepcopy(queue_commands) + [demo_job]), finish()], deepcopy(clean_vars))

    projects[LOCAL_DEMO] = project(LOCAL_DEMO, "Conciliacion sintetica, API publica, XLSX y buzon SMTP local",
        [initial(True, local_mail=True), attempt(deepcopy(demo_job["children"])), finish()],
        deepcopy(clean_vars))

    TARGET.parent.mkdir(parents=True, exist_ok=True)
    temporary = TARGET.with_suffix(".build.db")
    if temporary.exists():
        raise FileExistsError(temporary)
    target = sqlite3.connect(temporary)
    try:
        with target:
            target.execute(schema)
            for name in (MAIN, DEMO, LOCAL_DEMO, *[n for n in projects if n not in (MAIN, DEMO, LOCAL_DEMO)]):
                data = b64encode(json.dumps({"project": projects[name]}, ensure_ascii=True).encode()).decode()
                target.execute("INSERT INTO bots(name,data,description,version,father,data_type) VALUES (?,?,?,?,?,?)",
                               (name, data, projects[name]["profile"]["description"], "1.0.0", "", "normal"))
    finally:
        target.close()
    temporary.replace(TARGET)
    print(TARGET)


if __name__ == "__main__":
    build()
