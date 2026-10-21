'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { createFileDialogDirectory } = require('../electron/file-dialog-directory.cjs');

async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'word-dialog-directory-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const documents = path.join(root, 'Documents'), first = path.join(root, '한글 작업 폴더'), second = path.join(root, 'Other folder');
  await Promise.all([documents, first, second].map(dir => fs.mkdir(dir)));
  return { root, documents, first, second, settings: path.join(root, 'profile', 'file-dialog-directory.json') };
}

test('new profiles start in Documents; successful file locations survive restart', async t => {
  const f = await fixture(t), store = createFileDialogDirectory(f.settings, f.documents);
  assert.equal(await store.current(), f.documents);
  await store.rememberFile(path.join(f.first, '문서.docx'));
  assert.equal(await store.current(), f.first);
  assert.equal(await createFileDialogDirectory(f.settings, f.documents).current(), f.first);
  await store.rememberFile(path.join(f.second, '다른 이름.pdf'));
  assert.equal(await store.current(), f.second);
  assert.equal(await createFileDialogDirectory(f.settings, f.documents).current(), f.second);
});

test('missing folders and malformed preferences fall back without creating document folders', async t => {
  const f = await fixture(t), store = createFileDialogDirectory(f.settings, f.documents);
  await store.rememberFile(path.join(f.first, '문서.hwp'));
  await fs.rm(f.first, { recursive: true });
  assert.equal(await store.current(), f.documents);
  assert.equal(await createFileDialogDirectory(f.settings, f.documents).current(), f.documents);
  await assert.rejects(fs.stat(f.first), { code: 'ENOENT' });
  for (const contents of ['{bad json', '{"version":1,"directory":"relative/path"}', '{"version":1,"directory":null}', '{"version":2,"directory":"/"}']) {
    await fs.writeFile(f.settings, contents);
    assert.equal(await createFileDialogDirectory(f.settings, f.documents).current(), f.documents);
  }
  await fs.rm(f.documents, { recursive: true });
  assert.equal(await store.current(), undefined);
});

test('unavailable preferences still remember the folder in memory', async t => {
  const f = await fixture(t);
  await fs.writeFile(path.join(f.root, 'blocked'), 'not a directory');
  const store = createFileDialogDirectory(path.join(f.root, 'blocked', 'settings.json'), f.documents);
  await store.rememberFile(path.join(f.first, '문서.txt'));
  assert.equal(await store.current(), f.first);
});

test('invalid paths do not replace the folder; ordered updates leave no temporary settings files', async t => {
  const f = await fixture(t), store = createFileDialogDirectory(f.settings, f.documents);
  await Promise.all([
    store.rememberFile(path.join(f.first, '문서.docx')),
    store.rememberFile(path.join(f.second, '문서.wordstudio')),
  ]);
  for (const filename of [null, 'relative.docx', path.join(f.root, 'missing', 'document.docx')]) await store.rememberFile(filename);
  assert.equal(await store.current(), f.second);
  assert.deepEqual(await fs.readdir(path.dirname(f.settings)), ['file-dialog-directory.json']);
});
