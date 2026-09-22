/* eslint-disable no-console */
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(process.cwd(), 'dist-standalone');
const HOST = process.env.HOST || '127.0.0.1';
const START_PORT = Number(process.env.PORT || 5173);
const MAX_PORT_TRIES = Number(process.env.PORT_TRIES || 20);

const MIME_TYPES = {
  '.css': 'text/css; charset=utf-8',
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.webp': 'image/webp'
};

function send(res, statusCode, body, headers = {}) {
  res.writeHead(statusCode, {
    'Cache-Control': 'no-store',
    ...headers
  });
  res.end(body);
}

function resolveRequestPath(urlPath) {
  const decodedPath = decodeURIComponent(urlPath.split('?')[0] || '/');
  const requestedPath = decodedPath === '/' ? '/index.html' : decodedPath;
  const normalizedPath = path.normalize(requestedPath).replace(/^(\.\.[/\\])+/, '');
  const fullPath = path.resolve(ROOT, `.${normalizedPath}`);

  if (!fullPath.startsWith(ROOT)) {
    return null;
  }

  return fullPath;
}

function createServer() {
  return http.createServer((req, res) => {
    const fullPath = resolveRequestPath(req.url || '/');

    if (!fullPath) {
      send(res, 403, 'Forbidden');
      return;
    }

    fs.stat(fullPath, (statError, stats) => {
      if (statError || !stats.isFile()) {
        send(res, 404, 'Not found');
        return;
      }

      const extension = path.extname(fullPath).toLowerCase();
      res.writeHead(200, {
        'Cache-Control': 'no-store',
        'Content-Type': MIME_TYPES[extension] || 'application/octet-stream'
      });
      fs.createReadStream(fullPath).pipe(res);
    });
  });
}

function listenOnAvailablePort(port, attemptsLeft) {
  const server = createServer();

  server.on('error', (error) => {
    if (error.code === 'EADDRINUSE' && attemptsLeft > 1) {
      console.warn(`[serve:standalone] port ${port} is busy; trying ${port + 1}`);
      listenOnAvailablePort(port + 1, attemptsLeft - 1);
      return;
    }

    console.error(`[serve:standalone] failed to start: ${error.message}`);
    process.exit(1);
  });

  server.listen(port, HOST, () => {
    const address = server.address();
    console.log(`[serve:standalone] serving ${ROOT}`);
    console.log(`[serve:standalone] http://${HOST}:${address.port}`);
  });
}

listenOnAvailablePort(START_PORT, MAX_PORT_TRIES);
