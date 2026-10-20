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
    const current = JSON.stringify({ format: 'word-studio', id: 'large-images', title: 'after', content: { type: 'doc', content: [{ type: 'image', attrs: { src: 'data:image/png;base64,' + 'A'.repeat(65 * 1024 * 1024) } }] } });
    await store.save(first); await store.save(current);
    let records = await store.list();
    assert.equal(records.length, 1); assert.equal(records[0].current, current); assert.equal(records[0].previous, first);
    const manifest = JSON.parse(await fs.readFile(path.join(directory, 'large-images.json'), 'utf8'));
    assert.equal(manifest.version, 2); assert.ok(manifest.currentFile.endsWith('.snapshot')); assert.ok(manifest.previousFile.endsWith('.snapshot'));
    await assert.rejects(store.save('x'.repeat(limits.studioDocumentBytes + 1)), /384 MiB/);
    records = await store.list(); assert.equal(records[0].current, current); assert.equal(records[0].previous, first);
    const rename = fs.rename;
    try {
      fs.rename = async () => { throw new Error('simulated atomic commit failure'); };
      await assert.rejects(store.save(JSON.stringify({ format: 'word-studio', id: 'large-images', title: 'not committed' })), /atomic commit failure/);
    } finally { fs.rename = rename; }
    records = await store.list(); assert.equal(records[0].current, current); assert.equal(records[0].previous, first);
    assert.equal((await fs.readdir(directory)).filter(name => name.endsWith('.snapshot')).length, 2);
    await store.save(JSON.stringify({ format: 'word-studio', id: 'large-images', title: 'next' }));
    records = await store.list(); assert.equal(JSON.parse(records[0].current).title, 'next'); assert.equal(records[0].previous, current);
    assert.equal((await fs.readdir(directory)).filter(name => name.endsWith('.snapshot')).length, 2);
  } finally { await fs.rm(directory, { recursive: true, force: true }); }
});

test('recovery manifests reject paths outside the document and keep a valid previous snapshot available', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-recovery-paths-'));
  try {
    const previousFile = 'safe.00000000-0000-4000-8000-000000000000.snapshot';
    const previous = JSON.stringify({ format: 'word-studio', id: 'safe', title: 'previous valid' });
    await fs.writeFile(path.join(directory, previousFile), previous);
    await fs.writeFile(path.join(directory, 'outside'), 'LEAVE ALONE');
    await fs.writeFile(path.join(directory, 'safe.json'), JSON.stringify({ version: 2, currentFile: '../outside', previousFile, savedAt: '2026-10-20' }));
    const store = createRecoveryStore(directory), records = await store.list();
    assert.equal(records.length, 1); assert.equal(records[0].current, ''); assert.equal(records[0].previous, previous);
    await store.save(JSON.stringify({ format: 'word-studio', id: 'safe', title: 'recovered' }));
    assert.equal((await store.list())[0].previous, previous); assert.equal(await fs.readFile(path.join(directory, 'outside'), 'utf8'), 'LEAVE ALONE');
  } finally { await fs.rm(directory, { recursive: true, force: true }); }
});
