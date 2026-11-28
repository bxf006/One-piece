const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { Writable } = require('node:stream');

const root = path.resolve(__dirname, '..');
let handler;
vm.runInNewContext(fs.readFileSync(path.join(root, 'preview.cjs'), 'utf8'), {
  __dirname: root, URL, console: { log() {} },
  require(name) {
    if (name === 'node:http') return { createServer(callback) { handler = callback; return { listen() {} }; } };
    return require(name);
  }
}, { filename: 'preview.cjs' });
function request(url, method = 'GET') {
  return new Promise((resolve, reject) => {
    const chunks = [];
    const res = new Writable({ write(chunk, encoding, done) { chunks.push(Buffer.from(chunk)); done(); } });
    res.writeHead = (status, headers = {}) => { res.status = status; res.headers = headers; };
    res.on('error', reject);
    res.on('finish', () => resolve({ status: res.status, headers: res.headers, body: Buffer.concat(chunks) }));
    try { handler({ url, method }, res); } catch (error) { reject(error); }
  });
}

test('preview serves the three cloud JSON bundles unchanged without caching', async () => {
  for (const type of ['axes', 'regex', 'selection']) {
    const response = await request(`/configs/${type}.json?refresh=1`);
    assert.equal(response.status, 200);
    assert.equal(response.headers['Content-Type'], 'application/json; charset=utf-8');
    assert.equal(response.headers['Cache-Control'], 'no-store');
    assert.deepEqual(response.body, fs.readFileSync(path.join(root, 'configs', type + '.json')));
  }
});

test('preview HEAD returns JSON metadata without a response body', async () => {
  const response = await request('/configs/axes.json', 'HEAD');
  assert.equal(response.status, 200);
  assert.match(response.headers['Content-Type'], /application\/json/);
  assert.equal(response.body.length, 0);
});

test('preview rejects write/delete methods and unlisted files without changing cloud files', async () => {
  const before = fs.readFileSync(path.join(root, 'configs', 'axes.json'));
  for (const method of ['POST', 'PUT', 'PATCH', 'DELETE']) {
    const response = await request('/configs/axes.json', method);
    assert.equal(response.status, 405);
    assert.equal(response.headers.Allow, 'GET, HEAD');
  }
  for (const url of ['/src/config-store.js', '/configs/../package.json', '/configs/README.md']) {
    assert.equal((await request(url)).status, 404);
  }
  assert.deepEqual(fs.readFileSync(path.join(root, 'configs', 'axes.json')), before);
});
