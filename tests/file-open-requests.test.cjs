'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');
const { EventEmitter } = require('node:events');
const { createRequire } = require('node:module');
const { pathToFileURL } = require('node:url');
const { FileOpenRequests } = require('../electron/file-open-requests.cjs');

async function desktop(t, platform, argv = []) {
  const profile = await fs.mkdtemp(path.join(os.tmpdir(), 'word-os-open-'));
  t.after(() => fs.rm(profile, { recursive: true, force: true }));
  const filename = path.resolve(__dirname, '../electron/main.cjs'), require = createRequire(filename);
  let resolveReady, resolveDialog, windows = 0;
  const whenReady = new Promise(resolve => { resolveReady = resolve; });
  const events = new EventEmitter(), handlers = new Map(), notices = [];
  const frame = { url: pathToFileURL(path.resolve(__dirname, '../dist/index.html')).href };
  const contents = Object.assign(new EventEmitter(), { mainFrame: frame, setWindowOpenHandler() {},
    send: (...args) => notices.push(args) });
  const app = Object.assign(events, {
    isPackaged: true, setName() {}, setPath() {}, setAppUserModelId() {}, getPath: () => profile,
    commandLine: { getSwitchValue: () => '', hasSwitch: () => false, appendSwitch() {} },
    getVersion: () => 'test', requestSingleInstanceLock: () => true, quit() {}, whenReady: () => whenReady,
    setAboutPanelOptions() {}, dock: { setIcon() {} },
  });
  vm.runInNewContext(await fs.readFile(filename, 'utf8'), {
    __dirname: path.dirname(filename), process: { ...process, platform, argv: ['Word Studio', ...argv] },
    Buffer, Uint8Array, URL, setTimeout, clearTimeout,
    require: id => id === 'electron' ? {
      app, systemPreferences: { setUserDefault() {} },
      BrowserWindow: class extends EventEmitter {
        webContents = contents;
        constructor() { super(); windows++; }
        isDestroyed() { return false; } isMinimized() { return false; }
        show() {} focus() {} setMenuBarVisibility() {} async loadFile() {}
      },
      ipcMain: { handle: (name, fn) => handlers.set(name, fn), on() {} },
      dialog: { showOpenDialog: () => new Promise(resolve => { resolveDialog = resolve; }) },
      Menu: { buildFromTemplate: value => value, setApplicationMenu() {} },
      session: { defaultSession: { setPermissionRequestHandler() {}, setPermissionCheckHandler() {}, on() {},
        webRequest: { onBeforeRequest() {}, onHeadersReceived() {} } } },
    } : id === './studio-updater.cjs' ? { createStudioUpdater: () => ({ onChange() {}, start() {}, menuTemplate: () => [] }) } : require(id),
  });
  const event = { sender: contents, senderFrame: frame };
  return { profile, events, notices, handlers, event, count: () => windows,
    async ready() { resolveReady(); await new Promise(resolve => setTimeout(resolve, 20)); },
    take: (...args) => handlers.get('word-studio:take-requested-file')(event, ...args),
    finishDialog: result => resolveDialog(result),
    request: filename => events.emit('open-file', { preventDefault() {} }, filename),
  };
}

test('queue accepts supported local files and file URLs, ignores switches, remote URLs and duplicates, and stays bounded', () => {
  const queue = new FileOpenRequests(), cwd = path.resolve('한글 폴더');
  for (const ext of ['DOCX', 'wordstudio', 'hwp', 'HWPX', 'html', 'htm', 'txt']) assert.equal(queue.add(`공백 문서.${ext}`, cwd), true);
  for (const name of ['--log-file=a.docx', 'https://host/a.docx', 'file:bad%ZZ.docx', 'a.exe', '', null]) assert.equal(queue.add(name, cwd), false);
  assert.equal(queue.add(pathToFileURL(path.join(cwd, '공백 문서.DOCX')).href), false);
  assert.equal(queue.take(), path.join(cwd, '공백 문서.DOCX'));
  for (let i = 0; i < 100; i++) queue.add(`document-${i}.txt`, cwd);
  assert.equal(queue.size, 64);
});

test('Finder paths received before startup survive until the editor is ready and share the normal working folder', async t => {
  const app = await desktop(t, 'darwin');
  const filename = path.join(app.profile, '한글과 공백.DOCX'); await fs.writeFile(filename, 'document');
  app.request(filename); app.request(filename);
  assert.equal(app.count(), 0);
  await app.ready(); assert.equal(app.count(), 1);
  assert.equal((await app.take()).name, '한글과 공백.DOCX');
  assert.equal(await app.take(), null);
  const preferences = JSON.parse(await fs.readFile(path.join(app.profile, 'file-dialog-directory.json')));
  assert.equal(preferences.directory, app.profile);
});

for (const platform of ['win32', 'linux']) test(`${platform} startup and second-instance arguments open in order in one editor`, async t => {
  const app = await desktop(t, platform);
  const first = path.join(app.profile, '첫 문서.docx'), second = path.join(app.profile, '다음 문서.txt');
  await fs.writeFile(first, 'first'); await fs.writeFile(second, 'second');
  app.events.emit('second-instance', {}, ['Word Studio', first], app.profile);
  await app.ready(); assert.equal((await app.take()).name, '첫 문서.docx');
  app.events.emit('second-instance', {}, ['Word Studio', '--test=x.txt', '다음 문서.txt', pathToFileURL(first).href], app.profile);
  assert.ok(app.notices.some(notice => notice[0] === 'word-studio:open-requested'));
  assert.equal((await app.take()).name, '다음 문서.txt');
  assert.equal((await app.take()).name, '첫 문서.docx');
  assert.equal(await app.take(), null); assert.equal(app.count(), 1);
});

test('initial launch arguments are available without a renderer startup notification', async t => {
  const filename = path.join(os.tmpdir(), `word-startup-${process.pid}.txt`);
  await fs.writeFile(filename, 'startup'); t.after(() => fs.rm(filename, { force: true }));
  const app = await desktop(t, 'win32', [filename]); await app.ready();
  assert.equal((await app.take()).name, path.basename(filename)); assert.equal(await app.take(), null);
});

test('untrusted IPC cannot read or consume paths; missing, non-file and oversized documents do not block the next file', async t => {
  const app = await desktop(t, 'darwin'); await app.ready();
  const valid = path.join(app.profile, 'valid.txt'); await fs.writeFile(valid, 'valid');
  app.request(path.join(app.profile, 'missing.docx')); app.request(valid);
  const take = app.handlers.get('word-studio:take-requested-file');
  assert.throws(() => take({ ...app.event, sender: {} }), /editor window/);
  assert.throws(() => take({ ...app.event, senderFrame: { url: 'https://untrusted/' } }), /editor window/);
  assert.throws(() => app.take('/arbitrary/file.docx'), /경로나 인수/);
  await assert.rejects(app.take(), { code: 'ENOENT' }); assert.equal((await app.take()).name, 'valid.txt');
  const directory = path.join(app.profile, 'directory.docx'); await fs.mkdir(directory); app.request(directory);
  await assert.rejects(app.take(), /no larger/);
  const huge = path.join(app.profile, 'too-large.docx'); await fs.writeFile(huge, '');
  await fs.truncate(huge, require('../electron/document-limits.json').nativeFileBytes + 1);
  app.request(huge); app.request(valid);
  await assert.rejects(app.take(), /no larger/); assert.equal((await app.take()).name, 'valid.txt');
});

test('OS file requests wait for an existing native chooser and are notified again when it closes', async t => {
  const app = await desktop(t, 'darwin'); await app.ready();
  const pending = app.handlers.get('word-studio:open-file')(app.event);
  await new Promise(resolve => setTimeout(resolve, 20));
  const filename = path.join(app.profile, 'after-dialog.txt'); await fs.writeFile(filename, 'after');
  app.request(filename); assert.equal(app.take(), null);
  const before = app.notices.length;
  app.finishDialog({ canceled: true, filePaths: [] }); await pending;
  assert.ok(app.notices.length > before); assert.equal((await app.take()).name, 'after-dialog.txt');
});

test('preload hides event arguments and never accepts renderer-supplied file paths', async () => {
  let bridge; const listeners = new Map(), calls = [];
  vm.runInNewContext(await fs.readFile(path.resolve(__dirname, '../electron/preload.cjs'), 'utf8'), { process,
    require: () => ({ contextBridge: { exposeInMainWorld: (_name, value) => { bridge = value; } },
      ipcRenderer: { invoke: (...args) => { calls.push(args); return Promise.resolve(null); },
        on: (channel, fn) => listeners.set(channel, fn), removeListener: (channel, fn) => {
          assert.equal(listeners.get(channel), fn); listeners.delete(channel);
        } } }),
  });
  await bridge.takeRequestedFile('/arbitrary/file.docx');
  assert.deepEqual(calls, [['word-studio:take-requested-file']]);
  let args; const off = bridge.onOpenRequested((...received) => { args = received; });
  listeners.get('word-studio:open-requested')({ sender: 'private' }, '/private/path');
  assert.deepEqual(args, []); off(); assert.equal(listeners.size, 0);
});
