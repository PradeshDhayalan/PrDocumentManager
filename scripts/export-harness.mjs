import { build } from 'esbuild';
import archiver from 'archiver';
import fs from 'node:fs';
import path from 'node:path';
const destination = path.resolve('artifacts/dms-harness');
fs.rmSync(destination, { recursive: true, force: true });
fs.mkdirSync(path.join(destination, 'api'), { recursive: true });
fs.cpSync('harness/dist', path.join(destination, 'web'), { recursive: true });
fs.cpSync('mock-api/seed-files', path.join(destination, 'seed-files'), { recursive: true });
await build({ entryPoints: ['mock-api/src/server.ts'], outfile: path.join(destination, 'api/server.cjs'), bundle: true, platform: 'node', target: 'node20', format: 'cjs', legalComments: 'eof' });
fs.writeFileSync(path.join(destination, 'start.cjs'), `const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { app, store } = require('./api/server.cjs');
if (!fs.existsSync(store.statePath)) store.reset();
const web = path.resolve(__dirname, 'web');
const port = Number(process.env.DMS_HARNESS_PORT || 5173);
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg' };
const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  if (/^\\/(api|__mock|mock-sharepoint)(\\/|$)/.test(url.pathname)) return app(req, res);
  let name;
  try { name = decodeURIComponent(url.pathname); } catch { res.writeHead(400); return res.end(); }
  if (name === '/') name = '/index.html';
  const file = path.resolve(web, '.' + name);
  if (!file.startsWith(web + path.sep)) { res.writeHead(403); return res.end(); }
  if (!fs.existsSync(file) || !fs.statSync(file).isFile()) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
  const stream = fs.createReadStream(file);
  stream.on('error', () => res.destroy());
  stream.pipe(res);
});
server.on('error', error => { process.stderr.write(error.code === 'EADDRINUSE' ? 'Port is busy. Set DMS_HARNESS_PORT to another port.\\n' : 'Harness server could not start.\\n'); process.exitCode = 1; });
server.listen(port, '127.0.0.1', () => process.stdout.write('DMS harness: http://localhost:' + port + '\\n'));
`);
fs.writeFileSync(path.join(destination, 'README.txt'), `DMS Grid portable harness — M0/M1

1. Install Node.js 20 or later if needed.
2. Extract this ZIP. Open a terminal in the dms-harness folder.
3. Run: node start.cjs
4. Open http://localhost:5173

No npm install, ffmpeg, tenant connection or internet access is needed to run this prebuilt package.
The React/Fluent PCF source was compiled into the harness. The mock API and generated sample files are included and share the same local port.
The current UI is the four-theme M0 shell; the API is M1. The production sample-data grid and UI action wiring are still pending.
First launch restores the seed into data/. Changes persist locally. POST /__mock/reset restores fixtures.
Use DMS_HARNESS_PORT to choose another port. Stop with Ctrl+C.
`);
const output = fs.createWriteStream(path.resolve('artifacts/dms-harness.zip'));
const archive = archiver('zip', { zlib: { level: 9 } });
const completed = new Promise((resolve, reject) => { output.on('close', resolve); output.on('error', reject); archive.on('error', reject); });
archive.pipe(output);
archive.directory(destination, 'dms-harness');
await archive.finalize();
await completed;
process.stdout.write('Created artifacts/dms-harness.zip\n');
