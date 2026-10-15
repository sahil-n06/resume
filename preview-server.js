const http = require('http');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const ROOT = __dirname;
const PORT = Number(process.env.PORT) || 4173;
const DOCX = path.join(ROOT, 'output', 'Sahil_Nawaz_Resume.docx');
const HTML = path.join(ROOT, 'preview', 'index.html');

const VENDOR = {
  '/vendor/jszip.min.js': path.join(ROOT, 'node_modules', 'jszip', 'dist', 'jszip.min.js'),
  '/vendor/docx-preview.min.js': path.join(ROOT, 'node_modules', 'docx-preview', 'dist', 'docx-preview.min.js'),
};

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};

let stamp = '';
let error = '';
let ready = false;
let building = false;
let queued = false;

function generate() {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [path.join(ROOT, 'sahil.js')], {
      cwd: ROOT,
      stdio: 'inherit',
    });
    child.on('error', reject);
    child.on('exit', (code) => {
      if (code === 0) resolve();
      else reject(new Error('sahil.js exited with code ' + code));
    });
  });
}

async function rebuild() {
  if (building) {
    queued = true;
    return;
  }
  building = true;
  try {
    await generate();
    error = '';
    ready = fs.existsSync(DOCX);
    stamp = String(Date.now());
  } catch (err) {
    error = err.message;
    stamp = String(Date.now());
    console.error(err);
  } finally {
    building = false;
    if (queued) {
      queued = false;
      rebuild();
    }
  }
}

function send(res, status, body, type) {
  res.writeHead(status, {
    'Content-Type': type,
    'Cache-Control': 'no-store',
  });
  res.end(body);
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  const pathname = decodeURIComponent(url.pathname);

  if (pathname === '/status') {
    send(res, 200, JSON.stringify({ stamp, error, ready }), 'application/json; charset=utf-8');
    return;
  }

  if (pathname === '/' || pathname === '/index.html') {
    send(res, 200, fs.readFileSync(HTML), TYPES['.html']);
    return;
  }

  if (VENDOR[pathname]) {
    if (!fs.existsSync(VENDOR[pathname])) {
      send(res, 404, 'missing vendor file', 'text/plain; charset=utf-8');
      return;
    }
    send(res, 200, fs.readFileSync(VENDOR[pathname]), TYPES['.js']);
    return;
  }

  if (pathname === '/Sahil_Nawaz_Resume.docx') {
    if (!fs.existsSync(DOCX)) {
      send(res, 404, 'not built yet', 'text/plain; charset=utf-8');
      return;
    }
    send(res, 200, fs.readFileSync(DOCX), TYPES['.docx']);
    return;
  }

  send(res, 404, 'not found', 'text/plain; charset=utf-8');
});

let watchTimer;
fs.watch(path.join(ROOT, 'sahil.js'), () => {
  clearTimeout(watchTimer);
  watchTimer = setTimeout(rebuild, 250);
});

server.listen(PORT, () => {
  const url = 'http://localhost:' + PORT;
  console.log('Resume preview at ' + url);
  console.log('Editing sahil.js refreshes the page.');
  rebuild();
});
