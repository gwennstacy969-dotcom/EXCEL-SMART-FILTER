const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { exec } = require('child_process');

const PORT = 3000;
const APP_DIR = path.join(__dirname, 'app');

const MIME_TYPES = {
  '.html': 'text/html',
  '.css': 'text/css',
  '.js': 'text/javascript',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xls': 'application/vnd.ms-excel'
};

const server = http.createServer((req, res) => {
  let reqPath = decodeURI(req.url.split('?')[0]);
  if (reqPath === '/') reqPath = '/index.html';

  let filePath = path.join(APP_DIR, reqPath);

  // Fallback to project root if requesting Sample_Data_Bulanan.xlsx
  if (!fs.existsSync(filePath)) {
    const rootPath = path.join(__dirname, reqPath);
    if (fs.existsSync(rootPath)) {
      filePath = rootPath;
    }
  }

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('404 Not Found');
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    res.writeHead(200, {
      'Content-Type': contentType,
      'Access-Control-Allow-Origin': '*'
    });
    fs.createReadStream(filePath).pipe(res);
  });
});

function getLocalIp() {
  const interfaces = os.networkInterfaces();
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      if (iface.family === 'IPv4' && !iface.internal) {
        return iface.address;
      }
    }
  }
  return 'localhost';
}

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.log(`[INFO] Port ${PORT} sudah berjalan. Membuka http://localhost:${PORT} di browser...`);
    const startCmd = process.platform === 'win32' ? `start http://localhost:${PORT}` : `open http://localhost:${PORT}`;
    exec(startCmd);
  } else {
    console.error('Server error:', err);
  }
});

server.listen(PORT, () => {
  const localIp = getLocalIp();
  console.log('===============================================================');
  console.log('       EXCEL SMART FILTER - RESPONSIVE WEB & MOBILE APP        ');
  console.log('===============================================================');
  console.log(`[+] Komputer / Laptop : http://localhost:${PORT}`);
  console.log(`[+] HP / Smartphone   : http://${localIp}:${PORT}  (Pastikan 1 Wi-Fi)`);
  console.log('===============================================================');
  console.log('Aplikasi otomatis terbuka di browser default Anda...\n');

  // Buka otomatis di browser
  const startCmd = process.platform === 'win32' ? `start http://localhost:${PORT}` : `open http://localhost:${PORT}`;
  exec(startCmd);
});

