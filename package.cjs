// Build the static website and full source package. No npm dependencies are needed.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
const JSZip = require('./vendor/jszip.min.js');

async function main() {
  const root = __dirname;
  execFileSync(process.execPath, [path.join(root, 'build.cjs')], { stdio: 'inherit' });
  const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  assert.ok(!/\/\*__[A-Z]+__\*\//.test(html), 'Unresolved build marker');
  assert.ok(!/<(?:script|link)\b[^>]*(?:src|href)\s*=\s*["'](?:https?:)?\/\//i.test(html), 'External runtime dependency');
  const scripts = [...html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)];
  assert.equal(scripts.length, 7, 'Expected two vendor scripts and five application scripts');
  const runtimeFiles = [];
  scripts.forEach((match, index) => {
    const source = match[1].match(/\bsrc\s*=\s*["']([^"']+)["']/i);
    let code = match[2], filename = `index.html:script-${index + 1}`;
    if (source) {
      filename = source[1];
      assert.ok(/^vendor\/[a-z0-9.-]+\.js$/i.test(filename), `Expected local vendor path: ${filename}`);
      assert.equal(code.trim(), '', `External script must not contain inline code: ${filename}`);
      assert.ok(!/\b(?:async|defer)\b/i.test(match[1]), 'Vendor scripts must load before inline application code');
      assert.equal(index, runtimeFiles.length, 'Vendor scripts must precede application scripts');
      code = fs.readFileSync(path.join(root, filename), 'utf8');
      runtimeFiles.push(filename);
    }
    new vm.Script(code, { filename });
  });
  assert.deepEqual(runtimeFiles, ['vendor/jszip.min.js', 'vendor/xlsx.full.min.js'], 'Expected local dependencies in load order');
  for (const filename of runtimeFiles) assert.ok(!html.includes(fs.readFileSync(path.join(root, filename), 'utf8')), `Vendor bundle is still embedded: ${filename}`);

  const zip = new JSZip();
  const files = new Map();
  function add(relative) {
    const absolute = path.join(root, relative);
    if (fs.statSync(absolute).isDirectory()) {
      for (const name of fs.readdirSync(absolute).sort()) add(`${relative}/${name}`);
    } else {
      const bytes = fs.readFileSync(absolute);
      files.set(relative, bytes);
      zip.file(relative, bytes);
    }
  }
  for (const entry of ['index.html', 'README.md', 'RELEASE_NOTES.md', 'UPDATED_FILES.md', 'THIRD_PARTY_NOTICES.md', 'package.json', 'build.cjs', 'package.cjs', 'preview.cjs', 'src', 'tests', 'vendor', 'configs']) add(entry);
  const { parseCloudDocument, FILES } = require('./src/config-store.js');
  for (const [type, name] of Object.entries(FILES)) parseCloudDocument(fs.readFileSync(path.join(root, 'configs', name), 'utf8'), type);
  if (fs.existsSync(path.join(root, 'package-lock.json'))) add('package-lock.json');
  for (const filename of runtimeFiles) assert.ok(files.has(filename), `Missing runtime dependency in ZIP: ${filename}`);
  const bytes = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 9 } });
  const output = path.join(root, 'curve-viewer.zip');
  fs.writeFileSync(output, bytes);
  const readback = await JSZip.loadAsync(fs.readFileSync(output), { checkCRC32: true });
  for (const [name, expected] of files) assert.deepEqual(await readback.file(name).async('nodebuffer'), expected, `Archive mismatch: ${name}`);
  console.log(`Verified ${files.size} files; ${bytes.length.toLocaleString()} bytes: ${output}`);
  const website = new JSZip(), websiteFiles = new Map([...files].filter(([name]) => name === 'index.html' || name === 'THIRD_PARTY_NOTICES.md' || name.startsWith('vendor/') || name.startsWith('configs/')));
  for (const [name, content] of websiteFiles) website.file(name, content);
  const websiteOutput = path.join(root, 'curve-viewer-static.zip');
  fs.writeFileSync(websiteOutput, await website.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 9 } }));
  const websiteReadback = await JSZip.loadAsync(fs.readFileSync(websiteOutput), { checkCRC32: true });
  for (const [name, content] of websiteFiles) assert.deepEqual(await websiteReadback.file(name).async('nodebuffer'), content, `Static archive mismatch: ${name}`);
  console.log(`Verified ${websiteFiles.size} static files: ${websiteOutput}`);
}
main().catch(error => { console.error(error); process.exitCode = 1; });
