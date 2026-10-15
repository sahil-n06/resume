const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const ROOT = __dirname;
const PORT = 4174;
const PDF = path.join(os.homedir(), 'Downloads', 'Sahil_Nawaz_Resume.pdf');
const CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';

function waitForReady() {
  return new Promise((resolve, reject) => {
    const started = Date.now();
    const tick = () => {
      const req = http.get('http://127.0.0.1:' + PORT + '/status', (res) => {
        let body = '';
        res.on('data', (chunk) => { body += chunk; });
        res.on('end', () => {
          try {
            const status = JSON.parse(body);
            if (status.ready) return resolve();
            if (status.error) return reject(new Error(status.error));
          } catch (err) {
            return reject(err);
          }
          if (Date.now() - started > 20000) return reject(new Error('Resume build timed out'));
          setTimeout(tick, 200);
        });
      });
      req.on('error', () => {
        if (Date.now() - started > 20000) return reject(new Error('Preview server did not start'));
        setTimeout(tick, 200);
      });
    };
    tick();
  });
}

function waitForDebugger() {
  return new Promise((resolve, reject) => {
    const started = Date.now();
    const tick = () => {
      const req = http.get('http://127.0.0.1:9222/json/version', (res) => {
        res.resume();
        if (res.statusCode === 200) resolve();
        else retry();
      });
      req.on('error', retry);
      function retry() {
        if (Date.now() - started > 15000) reject(new Error('Chrome did not open its debug port'));
        else setTimeout(tick, 200);
      }
    };
    tick();
  });
}

function cdp(wsUrl) {
  const ws = new WebSocket(wsUrl);
  let nextId = 0;
  const pending = new Map();
  const opened = new Promise((resolve, reject) => {
    ws.addEventListener('open', resolve);
    ws.addEventListener('error', () => reject(new Error('Could not connect to Chrome')));
  });
  ws.addEventListener('message', (event) => {
    const message = JSON.parse(event.data);
    if (!message.id || !pending.has(message.id)) return;
    const { resolve, reject } = pending.get(message.id);
    pending.delete(message.id);
    if (message.error) reject(new Error(message.error.message));
    else resolve(message.result);
  });
  return {
    async send(method, params = {}) {
      await opened;
      const id = ++nextId;
      ws.send(JSON.stringify({ id, method, params }));
      return new Promise((resolve, reject) => pending.set(id, { resolve, reject }));
    },
    close() { ws.close(); },
  };
}

async function printPdf(chrome) {
  const version = await fetch('http://127.0.0.1:9222/json/version').then((res) => res.json());
  const browser = cdp(version.webSocketDebuggerUrl);
  const { targetId } = await browser.send('Target.createTarget', {
    url: 'http://127.0.0.1:' + PORT + '/?export=1',
  });
  const list = await fetch('http://127.0.0.1:9222/json/list').then((res) => res.json());
  const page = list.find((target) => target.id === targetId);
  const session = cdp(page.webSocketDebuggerUrl);
  await session.send('Page.enable');
  await session.send('Runtime.enable');

  const started = Date.now();
  while (Date.now() - started < 20000) {
    const state = await session.send('Runtime.evaluate', {
      expression: 'JSON.stringify({ ready: !!window.__resumeReady, error: window.__resumeError || "" })',
      returnByValue: true,
    });
    const value = JSON.parse(state.result.value);
    if (value.error) throw new Error(value.error);
    if (value.ready) break;
    await new Promise((resolve) => setTimeout(resolve, 200));
  }
  if (Date.now() - started >= 20000) throw new Error('Resume did not finish rendering');

  await session.send('Emulation.setEmulatedMedia', { media: 'print' });

  const printed = await session.send('Page.printToPDF', {
    printBackground: true,
    paperWidth: 8.5,
    paperHeight: 11,
    marginTop: 0,
    marginBottom: 0,
    marginLeft: 0,
    marginRight: 0,
    preferCSSPageSize: true,
    displayHeaderFooter: false,
  });
  fs.mkdirSync(path.dirname(PDF), { recursive: true });
  fs.writeFileSync(PDF, Buffer.from(printed.data, 'base64'));
  session.close();
  browser.close();
  chrome.kill();
}

async function main() {
  if (!fs.existsSync(CHROME)) {
    throw new Error('Google Chrome is required to export the PDF');
  }

  const server = spawn(process.execPath, ['preview-server.js'], {
    cwd: ROOT,
    env: { ...process.env, PORT: String(PORT) },
    stdio: 'inherit',
  });
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'resume-pdf-'));
  const chrome = spawn(CHROME, [
    '--headless=new',
    '--disable-gpu',
    '--remote-debugging-port=9222',
    '--remote-allow-origins=*',
    '--user-data-dir=' + profile,
    '--no-first-run',
    '--no-default-browser-check',
  ], { stdio: ['ignore', 'ignore', 'pipe'] });
  let chromeLog = '';
  chrome.stderr.on('data', (chunk) => { chromeLog += chunk; });

  const stop = () => {
    server.kill();
    chrome.kill();
    try { fs.rmSync(profile, { recursive: true, force: true }); } catch (_) {}
  };

  try {
    await waitForReady();
    await waitForDebugger();
    await printPdf(chrome);
    stop();
    console.log('Wrote ' + PDF);
  } catch (err) {
    stop();
    if (chromeLog.trim()) console.error(chromeLog.trim().split('\n').slice(-8).join('\n'));
    throw err;
  }
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
