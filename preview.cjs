// Local preview also serves read-only configs/*.json for automatic cloud loading.
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const assets = new Map([
  ['/', ['index.html', 'text/html']],
  ['/index.html', ['index.html', 'text/html']],
  ['/vendor/jszip.min.js', ['vendor/jszip.min.js', 'text/javascript']],
  ['/vendor/xlsx.full.min.js', ['vendor/xlsx.full.min.js', 'text/javascript']],
  ['/configs/axes.json', ['configs/axes.json', 'application/json']],
  ['/configs/regex.json', ['configs/regex.json', 'application/json']],
  ['/configs/selection.json', ['configs/selection.json', 'application/json']]
]);
const server = http.createServer((req, res) => {
  if (!['GET', 'HEAD'].includes(req.method)) { res.writeHead(405, { Allow: 'GET, HEAD' }); res.end('Read only'); return; }
  const pathname = new URL(req.url, 'http://localhost').pathname;
  const asset = assets.get(pathname);
  if (!asset) { res.writeHead(404); res.end('Not found'); return; }
  const stream = fs.createReadStream(path.join(__dirname, asset[0]));
  stream.on('open', () => {
    res.writeHead(200, { 'Content-Type': asset[1] + '; charset=utf-8', 'Cache-Control': 'no-store' });
    if (req.method === 'HEAD') { stream.destroy(); res.end(); } else stream.pipe(res);
  });
  stream.on('error', () => { res.writeHead(404); res.end('Not found'); });
});
server.listen(8765, '127.0.0.1', () => console.log('Local: http://127.0.0.1:8765'));
