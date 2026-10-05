import { build } from 'esbuild';
import archiver from 'archiver';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
const revision = execFileSync('git', ['rev-parse', '--short', 'HEAD'], { encoding: 'utf8' }).trim();
const destination = path.resolve('artifacts/dms-harness');
fs.rmSync(destination, { recursive: true, force: true });
fs.mkdirSync(path.join(destination, 'api'), { recursive: true });
fs.cpSync('harness/dist', path.join(destination, 'web'), { recursive: true });
fs.cpSync('mock-api/seed-files', path.join(destination, 'seed-files'), { recursive: true });
fs.writeFileSync(path.join(destination, 'web/build-info.json'), JSON.stringify({ revision, app: 'DMS Grid interactive review' }));
await build({ entryPoints: ['mock-api/src/server.ts'], outfile: path.join(destination, 'api/server.cjs'), bundle: true, platform: 'node', target: 'node20', format: 'cjs', legalComments: 'eof' });
const licenses = path.join(destination, 'licenses');
fs.mkdirSync(licenses, { recursive: true });
fs.copyFileSync('mockup/public/icons/MICROSOFT-LICENSE.txt', path.join(licenses, 'microsoft-file-icons.txt'));
const packages = ['react','react-dom','express','cors','@noble/hashes'];
for (const scope of ['@fluentui','@griffel']) {
  const directory=path.join('node_modules',scope);
  if (fs.existsSync(directory)) packages.push(...fs.readdirSync(directory).map(name=>scope+'/'+name));
}
for(const name of packages){
 for(const root of ['node_modules','control/node_modules']){
  const directory=path.join(root,name);
  if(!fs.existsSync(directory))continue;
  const license=fs.readdirSync(directory).find(file=>/^licen[sc]e(?:\.|$)/i.test(file));
  if(license&&fs.statSync(path.join(directory,license)).isFile()) fs.copyFileSync(path.join(directory,license),path.join(licenses,name.replaceAll('/','-')+'-'+path.basename(root)+'.txt'));
 }
}
fs.writeFileSync(path.join(destination, 'start.cjs'), `const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { app, store } = require('./api/server.cjs');
if (!fs.existsSync(store.statePath)) store.reset();
const web = path.resolve(__dirname, 'web');
const port = Number(process.env.DMS_HARNESS_PORT || 5180);
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
server.listen(port, '127.0.0.1', () => {
  const url = 'http://localhost:' + port + '/?review=${revision}';
  process.stdout.write('DMS Grid review ${revision}: ' + url + '\\n');
  if (process.platform === 'darwin' && process.argv.includes('--open')) require('node:child_process').spawn('open', [url], {stdio:'ignore'}).on('error', () => {});
});
`);
fs.writeFileSync(path.join(destination, 'launch-mac.command'), `#!/bin/bash
cd -- "$(dirname -- "$0")"
node start.cjs --open
`, {mode:0o755});
fs.writeFileSync(path.join(destination, 'README.txt'), `DMS Grid interactive review ${revision}

1. Extract this ZIP.
2. Open a terminal in the extracted dms-harness folder.
3. Run: node start.cjs --open
4. Open http://localhost:5180 if the browser does not open automatically.

Mac shortcut: launch-mac.command starts the server and opens your browser.
Node.js 20 or later is required. No npm install, ffmpeg, tenant connection or internet access is needed.
Includes 22 Contoso documents, List/Tile views, Fluent Personas/thumbnails and mock upload/download/edit/delete.
Graph photos remain deferred. Office/PDF cards use illustrative covers; stored images, PDF and text can be previewed.
First launch copies the sample data into data/. Your changes persist there.
Stop with Ctrl+C. Set DMS_HARNESS_PORT to choose another port.
This package uses port 5180 so it does not reuse an older development server on 5173.
`);
const output = fs.createWriteStream(path.resolve('artifacts/dms-harness.zip'));
const archive = archiver('zip', { zlib: { level: 9 } });
const completed = new Promise((resolve, reject) => { output.on('close', resolve); output.on('error', reject); archive.on('error', reject); });
archive.pipe(output);
archive.directory(destination, 'dms-harness');
await archive.finalize();
await completed;
process.stdout.write('Created artifacts/dms-harness.zip\n');
