'use strict';

const https = require('https');
const fs    = require('fs');
const path  = require('path');

const PORT = 3000;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'application/javascript; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.png':  'image/png',
  '.ico':  'image/x-icon',
  '.json': 'application/json'
};

const ROOT = path.resolve(__dirname);

function requestHandler(req, res) {
  const urlPath  = req.url.split('?')[0];
  const filePath = path.resolve(__dirname, urlPath === '/' ? 'taskpane.html' : urlPath.slice(1));

  // Reject any path that escapes the project root (path traversal guard)
  if (!filePath.startsWith(ROOT + path.sep) && filePath !== ROOT) {
    res.writeHead(403, { 'Content-Type': 'text/plain' });
    res.end('403: Forbidden');
    return;
  }

  const ext      = path.extname(filePath).toLowerCase();
  const mimeType = MIME[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('404: Not Found');
      return;
    }
    res.writeHead(200, {
      'Content-Type':                mimeType,
      'Cache-Control':               'no-cache',
      // Restrict add-in to known origins only
      'Access-Control-Allow-Origin': 'https://localhost:3000',
      // Prevent the page from being embedded in unexpected frames
      'X-Frame-Options':             'SAMEORIGIN',
      'X-Content-Type-Options':      'nosniff'
    });
    res.end(data);
  });
}

function loadCredentials() {
  const certDir  = path.join(__dirname, 'certs');
  const keyPath  = path.join(certDir, 'localhost.key');
  const certPath = path.join(certDir, 'localhost.crt');

  if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    console.error('ERROR: No es troben els certificats a certs/');
    console.error('Executa: powershell -File setup-certs.ps1');
    process.exit(1);
  }

  return {
    key:  fs.readFileSync(keyPath),
    cert: fs.readFileSync(certPath)
  };
}

const server = https.createServer(loadCredentials(), requestHandler);

server.listen(PORT, () => {
  console.log('');
  console.log('╔══════════════════════════════════════════════════╗');
  console.log('║   RichCell – Servidor HTTPS  ✓                   ║');
  console.log('╠══════════════════════════════════════════════════╣');
  console.log(`║  https://localhost:${PORT}/taskpane.html              ║`);
  console.log('╚══════════════════════════════════════════════════╝');
  console.log('');
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`Port ${PORT} en ús. Executa: powershell -Command "Stop-Process -Id (Get-NetTCPConnection -LocalPort ${PORT}).OwningProcess -Force"`);
  } else {
    console.error('Error:', err.message);
  }
  process.exit(1);
});
