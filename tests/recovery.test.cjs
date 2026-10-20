const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const { createRecoveryStore } = require('../electron/recovery.cjs');
const limits = require('../electron/document-limits.json');

test('native recovery serializes writes, preserves previous version and rejects path traversal', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-recovery-test-'));
  const store = createRecoveryStore(directory);
  const doc = (id, text) => JSON.stringify({ format: 'word-studio', id, title: 'same title', text });
  await Promise.all([store.save(doc('a', 'first')), store.save(doc('a', 'second')), store.save(doc('b', 'other'))]);
  const records = await store.list();
  assert.equal(records.length, 2);
  const a = records.find(r => JSON.parse(r.current).id === 'a');
  assert.equal(JSON.parse(a.current).text, 'second');
  assert.equal(JSON.parse(a.previous).text, 'first');
  await assert.rejects(store.save(doc('../outside', 'invalid')));
  await assert.rejects(store.save('{bad json'));
  await store.save(doc('a', 'third'));
  assert.equal(JSON.parse((await store.list()).find(r => JSON.parse(r.current).id === 'a').previous).text, 'second');
  assert.equal((await fs.readdir(directory)).filter(f => f.endsWith('.tmp')).length, 0);
  // Delete only this test's known files, leaving recursive cleanup unnecessary.
  for (const name of ['a.json', 'b.json']) await fs.unlink(path.join(directory, name));
  await fs.rmdir(directory);
});

test('expanded picture documents above 40 MiB retain current and previous native recovery data', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-large-recovery-'));
  const store = createRecoveryStore(directory);
  try {
    const first = JSON.stringify({ format: 'word-studio', id: 'large-images', title: 'before' });
    const current = JSON.stringify({ format: 'word-studio', id: 'large-images', title: 'after', content: { type: 'doc', content: [{ type: 'image', attrs: { src: 'data:image/png;base64,' + 'A'.repeat(41 * 1024 * 1024) } }] } });
    await store.save(first); await store.save(current);
    let records = await store.list();
    assert.equal(records.length, 1); assert.equal(records[0].current, current); assert.equal(records[0].previous, first);
    await assert.rejects(store.save('x'.repeat(limits.studioDocumentBytes + 1)), /128 MiB/);
    records = await store.list(); assert.equal(records[0].current, current); assert.equal(records[0].previous, first);
  } finally { await fs.rm(directory, { recursive: true, force: true }); }
});
