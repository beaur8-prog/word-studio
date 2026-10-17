const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const { createRecoveryStore } = require('../electron/recovery.cjs');

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
