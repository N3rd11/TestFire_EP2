const fs = require('node:fs/promises');
const path = require('node:path');
const { setTimeout: wait } = require('node:timers/promises');
const { chromium } = require(process.env.EP3_PLAYWRIGHT || 'playwright');

const root = path.resolve(__dirname, '..');
const project = process.argv[2] || 'DEMO_TESTFIRE_EP3';
const visible = process.argv.includes('--visible');
const closeAfterRun = process.argv.includes('--close-after-run');
if (!['MAIN_TESTFIRE_EP3', 'DEMO_TESTFIRE_EP3', 'DEMO_CORREO_LOCAL_EP3'].includes(project)) {
  throw new Error('Robot EP3 desconocido.');
}
const database = path.join(root, 'rocketbot', 'MAIN_TESTFIRE_EP3', 'robot.db');
const editor = 'http://127.0.0.1:5000/editor/#/edit/';
const guid = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i;
const stages = {
  WEB_LOGIN_EXTRACT: 'Login y extraccion de TestFire',
  CONFIGURACION_TESTFIRE: 'Comprobacion de acceso a TestFire',
  RESPALDO_MOVIMIENTOS: 'Lectura de movimientos historicos',
  EXTRACCION_SINTETICA: 'Lectura de movimientos de prueba',
  VALIDATE_DATA: 'Validacion de movimientos',
  RECONCILE: 'Conciliacion de movimientos',
  API_PUBLICA: 'Consulta al Banco Mundial',
  INTEGRACION_EXCEL: 'Generacion del Excel',
  APERTURA_XLSX_LIBRERIA: 'Verificacion del Excel',
  APERTURA_EXCEL: 'Apertura del Excel',
  CORREO_OMITIDO_ARCHIVO_HISTORICO: 'Sin envio para el reporte historico',
  CORREO_SMTP_LOCAL: 'Envio y recepcion SMTP local',
  CORREO_SMTP_DUOC: 'Envio al correo Duoc por SMTP',
  CORREO_GMAIL_SMTP: 'Envio y recepcion Gmail',
  CORREO_GMAIL_API: 'Envio y recepcion Gmail',
  OUTLOOK: 'Correo en Outlook',
};

async function showProgress(folder, reported) {
  let log;
  try {
    log = await fs.readFile(path.join(folder, 'audit.jsonl'), 'utf8');
  } catch (error) {
    if (error.code === 'ENOENT') return;
    throw error;
  }
  for (const line of log.split('\n').slice(0, -1)) {
    const event = JSON.parse(line);
    const stage = stages[event.step_id];
    if (stage && !reported.has(event.event_id)) {
      process.stdout.write(`[${event.status}] ${stage}\n`);
      reported.add(event.event_id);
    }
  }
}

async function readSummary(folder) {
  try {
    return JSON.parse(await fs.readFile(path.join(folder, 'execution.json'), 'utf8'));
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
    return null;
  }
}

async function showResult(page, summary) {
  await page.evaluate(({ outcome, origin }) => {
    const result = document.createElement('aside');
    result.setAttribute('role', 'status');
    result.textContent = outcome + (origin === 'ARCHIVO_LOCAL_HISTORICO'
      ? '\nMovimientos de respaldo histórico' : '');
    Object.assign(result.style, {
      position: 'fixed', bottom: '42px', right: '22px', zIndex: '10000',
      padding: '14px 18px', color: 'white', background: '#16324f',
      font: '600 15px/1.5 system-ui, sans-serif', whiteSpace: 'pre-line',
      borderRadius: '8px', maxWidth: 'calc(100vw - 44px)',
    });
    document.body.append(result);
  }, { outcome: summary.resultado, origin: summary.origin });
}

async function run() {
  const browser = await chromium.launch({
    channel: 'msedge',
    headless: !visible,
    args: ['--disable-translate', '--disable-notifications'],
  });
  try {
    const page = await browser.newPage({ viewport: visible ? null : undefined });
    // El lanzador muestra audit.jsonl en consola. El depurador nativo es opcional.
    await page.route(/^http:\/\/(?:localhost|127\.0\.0\.1):5000\/debug\/start(?:\?.*)?$/,
      route => route.abort());
    const reported = new Set();
    let executionId;
    let latestError;
    page.on('response', async response => {
      if (new URL(response.url()).pathname !== '/execute') return;
      try {
        const body = await response.json();
        const id = body.vars?.find(v => v.name === 'execution_id')?.data;
        if (typeof id === 'string' && guid.test(id) && !executionId) executionId = id;
        if (String(body.status).toLowerCase() === 'false' && typeof body.message === 'string') {
          latestError = body.message;
        }
      } catch (error) {
        latestError ||= 'Studio devolvio una respuesta incompleta.';
      }
    });
    const encoded = Buffer.from(database, 'utf8').toString('base64');
    await page.goto(editor + project + '/db/' + encoded, { waitUntil: 'domcontentloaded' });
    const runButton = page.getByRole('button', { name: /^(Run|Ejecutar)$/i });
    const heading = page.getByText(new RegExp('Sequence of commands - ' + project + '|Secuencia de comandos - ' + project, 'i'));
    await heading.waitFor({ timeout: 25000 });
    if (visible) await page.bringToFront();
    await runButton.click({ timeout: 12000 });
    process.stdout.write('EN CURSO | Rocketbot inicio la ejecucion.\n');
    const deadline = Date.now() + 210000;
    let nextUpdate = Date.now() + 10000;
    while (Date.now() < deadline) {
      if (executionId) {
        const folder = path.join(root, 'EP3', 'ejecuciones', executionId);
        const summary = await readSummary(folder);
        await showProgress(folder, reported);
        if (summary?.ended_at && summary.execution_id === executionId) {
          if (visible) await showResult(page, summary).catch(() => {});
          if (visible || project === 'DEMO_CORREO_LOCAL_EP3') {
            await page.screenshot({ path: path.join(folder, 'Rocketbot_Studio.png') })
              .catch(() => {});
          }
          process.stdout.write(JSON.stringify({
            folder: guid.test(summary.mail_recovered_from || '')
              ? path.join(root, 'EP3', 'ejecuciones', summary.mail_recovered_from)
              : folder,
            status: summary.status, report: summary.report, duration_ms: summary.duration_ms,
            outlook: summary.outlook, valid_rows: summary.valid_rows, rejected_rows: summary.rejected_rows,
            origin: summary.origin, queue: summary.queue, message: summary.status === 'ERROR' ? latestError : null,
          }) + '\n');
          if (visible && !closeAfterRun && !page.isClosed()) {
            process.stdout.write('Studio queda abierto para revisar la ejecucion. Cierra esa ventana al terminar.\n');
            await page.waitForEvent('close', { timeout: 0 }).catch(() => {});
          }
          return;
        }
      }
      if (latestError && !executionId) throw new Error(latestError);
      if (page.isClosed()) throw new Error('Se cerro la ventana de Studio antes de confirmar el resultado.');
      if (Date.now() >= nextUpdate) {
        process.stdout.write('EN CURSO | Esperando que Studio termine.\n');
        nextUpdate += 10000;
      }
      await wait(500);
    }
    throw new Error('Studio no confirmo el cierre. Revisa el editor antes de volver a iniciar.');
  } finally {
    await browser.close();
  }
}

run().catch(error => {
  console.error(error.message);
  process.exitCode = 1;
});
