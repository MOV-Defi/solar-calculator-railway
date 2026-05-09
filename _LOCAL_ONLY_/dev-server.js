const http = require('http');
const fs = require('fs');
const path = require('path');
const { spawn, exec } = require('child_process');

const DEFAULT_PORT = Number(process.env.PORT || 3001);
const ROOT_DIR = path.join(__dirname, '..'); // Сервер у папці _LOCAL_ONLY_, тому корінь на рівень вище
const DATA_DIR = path.join(ROOT_DIR, 'data');
const TEMPLATES_FILE = path.join(DATA_DIR, 'templates_catalog.json');

const MIME_TYPES = {
  '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css',
  '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpg',
  '.svg': 'image/svg+xml', '.pdf': 'application/pdf',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

const ensureTemplatesFile = () => {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(TEMPLATES_FILE)) {
    fs.writeFileSync(TEMPLATES_FILE, JSON.stringify({
      schemaVersion: 1,
      updatedAt: new Date().toISOString(),
      templates: []
    }, null, 2));
  }
};

const readTemplatesCatalog = () => {
  try {
    ensureTemplatesFile();
    const raw = fs.readFileSync(TEMPLATES_FILE, 'utf-8');
    const parsed = JSON.parse(raw);
    return {
      schemaVersion: 1,
      updatedAt: parsed?.updatedAt || new Date().toISOString(),
      templates: Array.isArray(parsed?.templates) ? parsed.templates : []
    };
  } catch (_) {
    return {
      schemaVersion: 1,
      updatedAt: new Date().toISOString(),
      templates: []
    };
  }
};

const writeTemplatesCatalog = (templates) => {
  ensureTemplatesFile();
  const payload = {
    schemaVersion: 1,
    updatedAt: new Date().toISOString(),
    templates: Array.isArray(templates) ? templates : []
  };
  fs.writeFileSync(TEMPLATES_FILE, JSON.stringify(payload, null, 2), 'utf-8');
  return payload;
};

const server = http.createServer((req, res) => {
  const rawPath = (req.url || '/').split('?')[0].split('#')[0];

  // Підтримка відкриття папок локально
  if (req.method === 'POST' && rawPath === '/api/open-folder') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      try {
        const parsed = JSON.parse(body || '{}');
        const folderName = String(parsed.projectFolderName || '').trim().replace(/[<>:"\/\\|?*\x00-\x1F]/g, "_");
        if (!folderName) return res.end(JSON.stringify({ ok: false }));

        // Спрощений пошук папки в документах/завантаженнях
        const home = require('os').homedir();
        const searchRoots = [ROOT_DIR, path.join(home, 'Documents'), path.join(home, 'Downloads')];
        let foundPath = null;

        for (const root of searchRoots) {
          const candidate = path.join(root, folderName);
          if (fs.existsSync(candidate) && fs.statSync(candidate).isDirectory()) {
            foundPath = candidate;
            break;
          }
        }

        if (foundPath) {
          const cmd = process.platform === 'darwin' ? 'open' : 'explorer';
          exec(`${cmd} "${foundPath}"`);
          res.end(JSON.stringify({ ok: true }));
        } else {
          res.end(JSON.stringify({ ok: false, error: 'not_found' }));
        }
      } catch (e) { res.end(JSON.stringify({ ok: false })); }
    });
    return;
  }

  if (rawPath === '/api/templates') {
    if (req.method === 'GET') {
      const payload = readTemplatesCatalog();
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ ok: true, data: payload }));
      return;
    }
    if (req.method === 'POST') {
      let body = '';
      req.on('data', chunk => { body += chunk; });
      req.on('end', () => {
        try {
          const parsed = JSON.parse(body || '{}');
          const templates = Array.isArray(parsed?.templates) ? parsed.templates : null;
          if (!templates) {
            res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ ok: false, error: 'invalid_templates_payload' }));
            return;
          }
          const saved = writeTemplatesCatalog(templates);
          res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ ok: true, data: saved }));
        } catch (_) {
          res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ ok: false, error: 'invalid_json' }));
        }
      });
      return;
    }
    res.writeHead(405, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ ok: false, error: 'method_not_allowed' }));
    return;
  }

  // Роздача статичних файлів з основної папки
  let relativePath = rawPath === '/' ? 'index.html' : rawPath.replace(/^\//, '');
  const filePath = path.resolve(ROOT_DIR, relativePath);

  fs.readFile(filePath, (error, content) => {
    if (error) {
      res.writeHead(404);
      res.end('File not found');
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, { 'Content-Type': MIME_TYPES[ext] || 'application/octet-stream' });
    res.end(content, 'utf-8');
  });
});

const startServer = (port, maxAttempts = 20) => {
  server.listen(port, () => {
    const url = `http://localhost:${port}`;
    console.log(`Локальний сервер для розробки запущено: ${url}`);
    console.log(`ЦЮ ПАПКУ (_LOCAL_ONLY_) НЕ ПОТРІБНО ЗАВАНТАЖУВАТИ НА ХОСТИНГ!`);
    exec(process.platform === 'darwin' ? `open ${url}` : `start ${url}`);
  });

  server.once('error', (err) => {
    const retryable = err && (err.code === 'EADDRINUSE' || err.code === 'EPERM' || err.code === 'EACCES');
    if (retryable && maxAttempts > 0) {
      const nextPort = port + 1;
      console.warn(`Порт ${port} недоступний (${err.code}). Пробую ${nextPort}...`);
      setTimeout(() => startServer(nextPort, maxAttempts - 1), 120);
      return;
    }
    console.error('Не вдалося запустити dev-сервер:', err);
    process.exit(1);
  });
};

startServer(DEFAULT_PORT);
