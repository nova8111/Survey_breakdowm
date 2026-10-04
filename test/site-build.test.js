import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

await import('../build-site.mjs');
const { default: worker } = await import('../dist/server/index.js');

test('build packages worker dependencies and serves unchanged binary assets at both routes', async () => {
  for (const name of ['survey-core.js','data-io.js','report-worker.js','data-worker.js','vendor/xlsx.full.min.js']) {
    assert.deepEqual(await readFile(new URL(`../dist/${name}`, import.meta.url)), await readFile(new URL(`../${name}`, import.meta.url)));
  }
  const logo = await readFile(new URL('../assets/survey-response-logo.png', import.meta.url));
  for (const prefix of ['', '/Survey_breakdowm']) {
    const response = await worker.fetch(new Request(`https://example.test${prefix}/assets/survey-response-logo.png?v=1`));
    assert.equal(response.headers.get('content-type'), 'image/png');
    assert.deepEqual(Buffer.from(await response.arrayBuffer()), logo);
    const html = await worker.fetch(new Request(`https://example.test${prefix}/`));
    assert.equal(html.status, 200);
    assert.match(html.headers.get('content-security-policy'), /script-src 'self'/);
    assert.match(await html.text(), /vendor\/xlsx.full.min.js/);
  }
  assert.equal((await worker.fetch(new Request('https://example.test/missing.js'))).status, 404);
});

test('vendored library bytes match the pinned checksums', async () => {
  const { createHash } = await import('node:crypto');
  const checksums = (await readFile(new URL('../vendor/SHA256SUMS', import.meta.url), 'utf8')).replace(/^\uFEFF/, '').trim().split(/\r?\n/);
  for (const line of checksums) {
    const [expected, name] = line.split(/\s+/);
    const bytes = await readFile(new URL(`../vendor/${name}`, import.meta.url));
    assert.equal(createHash('sha256').update(bytes).digest('hex'), expected, name);
  }
});
