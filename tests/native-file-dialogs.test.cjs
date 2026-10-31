'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs/promises');
const vm = require('node:vm');
const { createNativeFileDialogs } = require('../electron/native-file-dialogs.cjs');

for (const kind of ['open', 'save']) test(`macOS ${kind} replaces a bookmark with an app-specific plain path before panel creation`, async () => {
  const calls = [], parent = {}, directory = path.resolve('한글 작업 폴더');
  const options = { defaultPath: kind === 'save' ? path.join(directory, '문서.docx') : directory };
  let setting = Buffer.from('old bookmark');
  const dialogs = createNativeFileDialogs({ platform: 'darwin', app: { getPath: () => path.resolve('Documents') },
    systemPreferences: { setUserDefault(...args) { calls.push(args); setting = args[2]; } },
    dialog: { showOpenDialog(window, passed) { assert.equal(window, parent); assert.equal(passed, options); assert.equal(setting, directory); return { canceled: true }; },
      showSaveDialog(window, passed) { assert.equal(window, parent); assert.equal(passed, options); assert.equal(setting, directory); return { canceled: true }; } } });
  assert.deepEqual(await dialogs[kind](parent, options), { canceled: true });
  assert.deepEqual(calls, [['NSOSPLastRootDirectory', 'string', directory]]);
});
test('missing panel paths use Documents without reading/removing shared defaults', async () => {
  const directory = path.resolve('Documents'), calls = [];
  const dialogs = createNativeFileDialogs({ platform: 'darwin', app: { getPath: name => { assert.equal(name, 'documents'); return directory; } },
    systemPreferences: { setUserDefault: (...args) => calls.push(args) },
    dialog: { showOpenDialog: async () => null, showSaveDialog: async () => false } });
  await dialogs.open(null, {}); await dialogs.save(null, { defaultPath: '문서.docx' });
  assert.deepEqual(calls, Array(2).fill(['NSOSPLastRootDirectory', 'string', directory]));
});
for (const platform of ['win32', 'linux']) test(`${platform} keeps its existing dialog options and never changes macOS preferences`, async () => {
  const options = { defaultPath: path.resolve('Documents') }, calls = [];
  const dialogs = createNativeFileDialogs({ platform, app: {},
    systemPreferences: { setUserDefault() { throw Error('Must not change preferences'); } },
    dialog: { showOpenDialog: async (...args) => { calls.push(args); return null; }, showSaveDialog: async (...args) => { calls.push(args); return false; } } });
  await dialogs.open(null, options); await dialogs.save(null, options);
  assert.deepEqual(calls, [[null, options], [null, options]]);
});
test('a preference preparation failure never creates a native panel', () => {
  let panels = 0;
  const dialogs = createNativeFileDialogs({ platform: 'darwin', app: { getPath: () => path.resolve('Documents') },
    systemPreferences: { setUserDefault() { throw Error('preference unavailable'); } },
    dialog: { showOpenDialog() { panels++; } } });
  assert.throws(() => dialogs.open(null, {}), /preference unavailable/); assert.equal(panels, 0);
});

async function preload(invoke) {
  let bridge;
  vm.runInNewContext(await fs.readFile(path.resolve(__dirname, '../electron/preload.cjs'), 'utf8'), { process,
    require: () => ({ contextBridge: { exposeInMainWorld: (_name, value) => { bridge = value; } },
      ipcRenderer: { invoke, on() {}, removeListener() {} } }) });
  return bridge;
}
test('all native panel types share a duplicate guard and clear waiting state after cancel or error', async () => {
  let complete;
  const calls = [], states = [];
  const bridge = await preload((...args) => { calls.push(args); return new Promise((resolve, reject) => { complete = { resolve, reject }; }); });
  const off = bridge.onFileOperation(value => states.push(value));
  const opening = bridge.openFile();
  for (const other of [() => bridge.openFile(), () => bridge.saveFile({}), () => bridge.exportPdf({}), () => bridge.print()])
    await assert.rejects(other(), /다른 파일 작업 중/);
  assert.equal(calls.length, 1); complete.resolve(null); assert.equal(await opening, null);
  const saving = bridge.saveFile({ name: '문서' }); complete.reject(Error('panel failed'));
  await assert.rejects(saving, /panel failed/);
  const retry = bridge.openFile(); complete.resolve({ name: 'selected.txt' }); assert.equal((await retry).name, 'selected.txt');
  assert.deepEqual(states, [null, 'open', null, 'save', null, 'open', null]);
  off();
});
test('a synchronous IPC failure clears the panel guard for the next attempt', async () => {
  let fail = true;
  const states = [];
  const bridge = await preload(() => { if (fail) throw Error('IPC unavailable'); return Promise.resolve(false); });
  bridge.onFileOperation(value => states.push(value));
  await assert.rejects(bridge.exportPdf({}), /IPC unavailable/);
  fail = false; assert.equal(await bridge.saveFile({}), false);
  assert.deepEqual(states, [null, 'pdf', null, 'save', null]);
});
