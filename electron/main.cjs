'use strict';

const { app, BrowserWindow, clipboard, dialog, ipcMain, Menu, session, shell, Notification } = require('electron');
const fs = require('node:fs/promises');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { fileURLToPath } = require('node:url');
const { createRecoveryStore } = require('./recovery.cjs');
const { createFileDialogDirectory } = require('./file-dialog-directory.cjs');
const { installedOutlineFonts } = require('./outline-fonts.cjs');
const { createStudioUpdater } = require('./studio-updater.cjs');

const PRODUCT_NAME = 'Word Studio';
const DEV_URL = 'http://127.0.0.1:5191';
const IS_DEV = !app.isPackaged && process.env.WORD_STUDIO_DEV === '1';
const DIST_PATH = path.resolve(__dirname, '..', 'dist');
const INDEX_PATH = path.join(DIST_PATH, 'index.html');
const MAX_FILE_BYTES = require('./document-limits.json').nativeFileBytes;
const SAVE_EXTENSIONS = new Set(['wordstudio', 'docx', 'html', 'txt', 'hwp', 'hwpx']);
const OPEN_EXTENSIONS = new Set([...SAVE_EXTENSIONS, 'htm', 'hwp', 'hwpx']);
const FORMAT_NAMES = {
  wordstudio: '편집본 (.wordstudio)', docx: 'Word 문서 (.docx)',
  html: 'HTML 문서 (.html)', txt: '텍스트 문서 (.txt)',
  hwp: '한글 문서 (.hwp)', hwpx: '한글 표준 문서 (.hwpx)',
};
const CHROME_THEMES = {
  dark: { color: '#151515', symbolColor: '#f2f2f2', height: 40 },
  light: { color: '#f3f3f3', symbolColor: '#242424', height: 40 },
};

let mainWindow = null;
let nativeOperationPending = false;
let chromeTheme = 'dark';
let printBackgroundActive = false;

app.setName(PRODUCT_NAME);
const customUserData = app.commandLine.getSwitchValue('user-data-dir');
app.setPath('userData', customUserData && path.isAbsolute(customUserData)
  ? path.resolve(customUserData) : path.join(app.getPath('appData'), PRODUCT_NAME));
if (process.platform === 'win32') app.setAppUserModelId('local.wordstudio.app');
const updates = createStudioUpdater({ app, dialog, shell, Notification, id: 'word', getWindow: () => mainWindow });
updates.onChange(installMenu);
// Older portal file choosers ignore defaultPath; let Electron use GTK/KDE
// when the portal cannot honor the document's working folder.
if (process.platform === 'linux' && !app.commandLine.hasSwitch('xdg-portal-required-version')) {
  app.commandLine.appendSwitch('xdg-portal-required-version', '4');
}

function isTrustedURL(value) {
  try {
    const url = new URL(value);
    if (IS_DEV) return url.origin === DEV_URL && url.pathname === '/';
    return url.protocol === 'file:' && path.resolve(fileURLToPath(url)) === INDEX_PATH;
  } catch { return false; }
}

function requireTrustedSender(event) {
  if (!mainWindow || mainWindow.isDestroyed() || event.sender !== mainWindow.webContents ||
      event.senderFrame !== mainWindow.webContents.mainFrame || !isTrustedURL(event.senderFrame.url)) {
    throw new Error('This operation is available only to the editor window.');
  }
}

function safeFilename(input, extension) {
  let name = path.basename(String(input || 'Untitled document'))
    .replace(/[\x00-\x1f<>:"/\\|?*]/g, '-')
    .replace(/^\.+/, '').slice(0, 160).trim().replace(/[. ]+$/, '') || 'Untitled document';
  if (/^(?:CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\.|$)/i.test(name)) name = `_${name}`;
  return name.toLowerCase().endsWith(`.${extension}`) ? name : `${name}.${extension}`;
}

function decodeFilePayload(payload) {
  if (!payload || typeof payload !== 'object' || typeof payload.name !== 'string' ||
      !SAVE_EXTENSIONS.has(payload.extension)) throw new TypeError('Choose a supported document format.');
  if (!(payload.data instanceof Uint8Array) || payload.data.byteLength > MAX_FILE_BYTES) {
    throw new TypeError(`Choose a document no larger than ${MAX_FILE_BYTES / 1024 / 1024} MiB.`);
  }
  return { name: safeFilename(payload.name, payload.extension), extension: payload.extension,
    data: Buffer.from(payload.data.buffer, payload.data.byteOffset, payload.data.byteLength) };
}

async function withNativeDialog(event, task) {
  requireTrustedSender(event);
  if (nativeOperationPending) throw new Error('Finish the current file or print dialog first.');
  nativeOperationPending = true;
  try { return await task(); }
  finally { nativeOperationPending = false; }
}

async function atomicWrite(filename, buffer) {
  const temporary = path.join(path.dirname(filename), `.${path.basename(filename)}.${randomUUID()}.tmp`);
  try {
    await fs.writeFile(temporary, buffer, { flag: 'wx', mode: 0o600 });
    await fs.rename(temporary, filename);
  } finally { await fs.rm(temporary, { force: true }).catch(() => {}); }
}

function sendMenuAction(action) {
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.webContents.send('word-studio:menu', action);
}

async function preparePrint(contents) {
  await contents.executeJavaScript(`(async () => {
    await document.fonts.ready;
    await Promise.all(Array.from(document.images, image => image.decode().catch(() => {})));
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  })()`);
}

async function withWhitePrintBackground(task) {
  const window = mainWindow;
  let printStyle;
  printBackgroundActive = true;
  // Chromium paints the native window color into the @page margin area.
  // A dark editor theme must never turn the exported paper margin black.
  window.setBackgroundColor('#ffffff');
  try {
    printStyle = await window.webContents.insertCSS('@media print { :root { color-scheme: only light !important; background: #ffffff !important; } }', { cssOrigin: 'user' });
    return await task();
  }
  finally {
    printBackgroundActive = false;
    if (!window.isDestroyed()) {
      if (printStyle) await window.webContents.removeInsertedCSS(printStyle).catch(() => {});
      window.setBackgroundColor(CHROME_THEMES[chromeTheme].color);
    }
  }
}

function pdfOptions(payload) {
  if (payload.paginated === true) return { printBackground: true, preferCSSPageSize: true,
    margins: { top: 0, right: 0, bottom: 0, left: 0 }, displayHeaderFooter: false };
  const header = payload.header ?? '';
  const footer = payload.footer ?? '';
  const pageNumbers = payload.pageNumbers ?? false;
  const margin = payload.margin ?? 25.4;
  if (typeof header !== 'string' || typeof footer !== 'string' || header.length > 4096 || footer.length > 4096 ||
      typeof pageNumbers !== 'boolean' || typeof margin !== 'number' || !Number.isFinite(margin) || margin < 0 || margin > 80) {
    throw new TypeError('Invalid PDF header, footer, or page margin.');
  }
  const escapeHTML = value => value.replace(/[&<>"']/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  })[character]).replace(/\r?\n/g, '<br>');
  const horizontal = margin.toFixed(3);
  const vertical = Math.max(0, margin / 2 - 4).toFixed(3);
  const textHeight = Math.max(3, margin - 8).toFixed(3);
  const style = `width:100%;box-sizing:border-box;font-family:Arial,'Malgun Gothic',sans-serif;font-size:8pt;line-height:1.35;color:#737c86;`;
  return {
    printBackground: true, preferCSSPageSize: true,
    // @page owns the content margins. API margins must not add a second inset.
    margins: { top: 0, right: 0, bottom: 0, left: 0 },
    displayHeaderFooter: Boolean(header || footer || pageNumbers),
    headerTemplate: header ? `<div style="${style}padding:${vertical}mm ${horizontal}mm 0;"><div style="max-height:${textHeight}mm;overflow:hidden;overflow-wrap:anywhere;">${escapeHTML(header)}</div></div>` : '<span></span>',
    footerTemplate: footer || pageNumbers ? `<div style="${style}padding:0 ${horizontal}mm ${vertical}mm;display:flex;align-items:flex-end;gap:12px;"><div style="flex:1;max-height:${textHeight}mm;overflow:hidden;overflow-wrap:anywhere;">${escapeHTML(footer)}</div>${pageNumbers ? '<span class="pageNumber" style="flex-shrink:0;"></span>' : ''}</div>` : '<span></span>',
  };
}

function installMenu() {
  const template = [
    ...(process.platform === 'darwin' ? [{ label: PRODUCT_NAME, submenu: [
      { role: 'about' }, { type: 'separator' }, { role: 'services' }, { type: 'separator' },
      { role: 'hide' }, { role: 'hideOthers' }, { role: 'unhide' }, { type: 'separator' }, { role: 'quit' },
    ] }] : []),
    { label: 'File', submenu: [
      { label: 'New Document', accelerator: 'CmdOrCtrl+N', click: () => sendMenuAction('new') },
      { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => sendMenuAction('open') },
      { type: 'separator' },
      { label: 'Save Editing Copy (.wordstudio)…', accelerator: 'CmdOrCtrl+S', click: () => sendMenuAction('save') },
      { label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S', click: () => sendMenuAction('save-as') },
      { label: 'Export Word Document (.docx)…', accelerator: 'CmdOrCtrl+Shift+E', click: () => sendMenuAction('export-docx') },
      { label: 'Export Hangul Document (HWP)…', click: () => sendMenuAction('export-hwp') },
      { label: 'Export Hangul Document (HWPX)…', click: () => sendMenuAction('export-hwpx') },
      { label: 'Export PDF…', click: () => sendMenuAction('export-pdf') },
      { label: 'Print…', accelerator: 'CmdOrCtrl+P', click: () => sendMenuAction('print') },
      { type: 'separator' }, process.platform === 'darwin' ? { role: 'close' } : { role: 'quit' },
    ] },
    { label: 'Edit', submenu: [
      { role: 'undo' }, { role: 'redo' }, { type: 'separator' },
      { role: 'cut' }, { role: 'copy' }, { role: 'paste' }, { role: 'selectAll' },
    ] },
    { label: 'View', submenu: [
      { role: 'togglefullscreen' },
      ...(IS_DEV ? [{ type: 'separator' }, { role: 'reload' }, { role: 'toggleDevTools' }] : []),
    ] },
    { label: 'Window', submenu: [{ role: 'minimize' }, { role: 'zoom' },
      ...(process.platform === 'darwin' ? [{ role: 'front' }] : [])] },
    { label: '도움말', submenu: updates.menuTemplate() },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function installIPC() {
  const recovery = createRecoveryStore(path.join(app.getPath('userData'), 'recovery'));
  const fileDirectory = createFileDialogDirectory(path.join(app.getPath('userData'), 'file-dialog-directory.json'), app.getPath('documents'));
  ipcMain.handle('word-studio:copy-text', (event, text) => {
    requireTrustedSender(event);
    if (typeof text !== 'string' || Buffer.byteLength(text, 'utf8') > 4 * 1024 * 1024) throw new TypeError('Copy text no larger than 4 MiB.');
    clipboard.writeText(text);
  });
  ipcMain.handle('word-studio:paste', event => {
    requireTrustedSender(event);
    // Chromium delivers a normal paste event, including table HTML, through
    // the editor's existing sanitization path. No general clipboard API is exposed.
    event.sender.paste();
  });
  ipcMain.handle('word-studio:outline-fonts', event => { requireTrustedSender(event); return installedOutlineFonts(); });
  ipcMain.handle('word-studio:read-recovery', event => { requireTrustedSender(event); return recovery.list(); });
  ipcMain.handle('word-studio:save-recovery', (event, serialized) => { requireTrustedSender(event); return recovery.save(serialized); });
  ipcMain.on('word-studio:set-theme', (event, theme) => {
    try {
      requireTrustedSender(event);
      if (theme !== 'dark' && theme !== 'light') return;
      chromeTheme = theme;
      if (process.platform === 'win32') mainWindow.setTitleBarOverlay(CHROME_THEMES[theme]);
      if (!printBackgroundActive) mainWindow.setBackgroundColor(CHROME_THEMES[theme].color);
    } catch { /* Invalid frames cannot change window chrome. */ }
  });

  ipcMain.handle('word-studio:save-file', (event, payload) => withNativeDialog(event, async () => {
    const file = decodeFilePayload(payload);
    const directory = await fileDirectory.current();
    const result = await dialog.showSaveDialog(mainWindow, {
      title: `${FORMAT_NAMES[file.extension]} 저장`,
      defaultPath: directory ? path.join(directory, file.name) : file.name,
      filters: [{ name: FORMAT_NAMES[file.extension], extensions: [file.extension] }],
      properties: ['createDirectory', 'showOverwriteConfirmation'],
    });
    if (result.canceled || !result.filePath) return false;
    await atomicWrite(result.filePath, file.data);
    await fileDirectory.rememberFile(result.filePath);
    return true;
  }));

  ipcMain.handle('word-studio:open-file', event => withNativeDialog(event, async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: '문서 열기',
      defaultPath: await fileDirectory.current(),
      filters: [{ name: '지원하는 문서', extensions: [...OPEN_EXTENSIONS] },
        { name: FORMAT_NAMES.wordstudio, extensions: ['wordstudio'] },
        { name: FORMAT_NAMES.docx, extensions: ['docx'] },
        { name: '한글 문서 (.hwp, .hwpx)', extensions: ['hwp', 'hwpx'] },
        { name: 'HTML·텍스트 (.html, .htm, .txt)', extensions: ['html', 'htm', 'txt'] }],
      properties: ['openFile'],
    });
    if (result.canceled || result.filePaths.length !== 1) return null;
    const filename = result.filePaths[0];
    if (!OPEN_EXTENSIONS.has(path.extname(filename).slice(1).toLowerCase())) throw new Error('Choose a supported document file.');
    const handle = await fs.open(filename, 'r');
    try {
      const stat = await handle.stat();
      if (!stat.isFile() || stat.size > MAX_FILE_BYTES) throw new Error(`Choose a document no larger than ${MAX_FILE_BYTES / 1024 / 1024} MiB.`);
      const buffer = await handle.readFile();
      if (buffer.length > MAX_FILE_BYTES) throw new Error(`Choose a document no larger than ${MAX_FILE_BYTES / 1024 / 1024} MiB.`);
      await fileDirectory.rememberFile(filename);
      return { name: path.basename(filename), data: new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength) };
    } finally { await handle.close(); }
  }));

  ipcMain.handle('word-studio:export-pdf', (event, payload) => withNativeDialog(event, async () => {
    if (!payload || typeof payload.name !== 'string') throw new TypeError('The PDF is missing a name.');
    const options = pdfOptions(payload);
    const directory = await fileDirectory.current();
    const name = safeFilename(payload.name, 'pdf');
    const result = await dialog.showSaveDialog(mainWindow, {
      title: 'PDF 문서 (.pdf) 저장', defaultPath: directory ? path.join(directory, name) : name,
      filters: [{ name: 'PDF 문서 (.pdf)', extensions: ['pdf'] }],
      properties: ['createDirectory', 'showOverwriteConfirmation'],
    });
    if (result.canceled || !result.filePath) return false;
    const data = await withWhitePrintBackground(async () => {
      await preparePrint(mainWindow.webContents);
      return mainWindow.webContents.printToPDF(options);
    });
    await atomicWrite(result.filePath, data);
    await fileDirectory.rememberFile(result.filePath);
    return true;
  }));

  ipcMain.handle('word-studio:print', event => withNativeDialog(event, () => withWhitePrintBackground(async () => {
    await preparePrint(mainWindow.webContents);
    return new Promise((resolve, reject) => {
      mainWindow.webContents.print({ silent: false, printBackground: true }, (success, reason) => {
        if (success) resolve(true);
        else if (/cancel/i.test(reason || '')) resolve(false);
        else reject(new Error(reason || 'The document could not be printed.'));
      });
    });
  })));
}

function hardenSession() {
  const editorSession = session.defaultSession;
  editorSession.setPermissionRequestHandler((_contents, _permission, callback) => callback(false));
  editorSession.setPermissionCheckHandler(() => false);
  editorSession.on('will-download', event => event.preventDefault());
  editorSession.webRequest.onBeforeRequest((details, callback) => {
    try {
      const url = new URL(details.url);
      const localFile = url.protocol === 'file:' && path.resolve(fileURLToPath(url)).startsWith(DIST_PATH + path.sep);
      const localData = ['data:', 'blob:', 'devtools:'].includes(url.protocol);
      const development = IS_DEV && ['http:', 'ws:'].includes(url.protocol) && url.hostname === '127.0.0.1' && url.port === '5191';
      callback({ cancel: !(localFile || localData || development) });
    } catch { callback({ cancel: true }); }
  });
  editorSession.webRequest.onHeadersReceived((details, callback) => {
    if (!isTrustedURL(details.url)) return callback({ responseHeaders: details.responseHeaders });
    const policy = [
      "default-src 'self'", `script-src 'self' 'wasm-unsafe-eval'${IS_DEV ? " 'unsafe-inline'" : ''}`,
      "style-src 'self' 'unsafe-inline'", "img-src 'self' data: blob:", "font-src 'self' data:",
      `connect-src 'self' ${IS_DEV ? 'http://127.0.0.1:5191 ws://127.0.0.1:5191' : ''}`,
      "worker-src 'self' blob:",
      "object-src 'none'", "frame-src 'none'", "base-uri 'none'", "form-action 'none'",
    ].join('; ');
    callback({ responseHeaders: { ...details.responseHeaders, 'Content-Security-Policy': [policy] } });
  });
}

async function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1540, height: 1000, minWidth: 1080, minHeight: 720, title: PRODUCT_NAME,
    icon: path.join(__dirname, process.platform === 'linux' ? 'WordStudio.png' : 'WordStudio.ico'),
    backgroundColor: CHROME_THEMES[chromeTheme].color, show: false,
    autoHideMenuBar: process.platform === 'win32',
    ...(process.platform === 'win32' ? { titleBarStyle: 'hidden', titleBarOverlay: CHROME_THEMES[chromeTheme] } : {}),
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true, nodeIntegration: false,
      sandbox: true, webSecurity: true, allowRunningInsecureContent: false,
      navigateOnDragDrop: false, webviewTag: false, spellcheck: true,
    },
  });
  if (process.platform === 'win32') mainWindow.setMenuBarVisibility(false);
  mainWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  mainWindow.webContents.on('before-input-event', (event, input) => {
    const primaryModifier = process.platform === 'darwin' ? input.meta : input.control;
    if (input.type !== 'keyDown' || !primaryModifier || input.alt) return;
    const key = input.key.toLowerCase();
    const action = input.shift ? { s: 'save-as', e: 'export-docx' }[key]
      : { n: 'new', o: 'open', s: 'save', p: 'print' }[key];
    if (!action) return;
    // Consume the key before both the DOM shortcut and the native accelerator
    // can handle it, then dispatch one shared renderer action.
    event.preventDefault();
    if (!input.isAutoRepeat) sendMenuAction(action);
  });
  mainWindow.webContents.on('will-navigate', (event, url) => { if (!isTrustedURL(url)) event.preventDefault(); });
  mainWindow.webContents.on('will-attach-webview', event => event.preventDefault());
  mainWindow.webContents.on('will-prevent-unload', event => {
    const response = dialog.showMessageBoxSync(mainWindow, {
      type: 'warning', title: 'Close Word Studio?', message: 'Some changes may not have been saved.',
      detail: 'Keep editing to save your document, or close without saving the pending changes.',
      buttons: ['Keep Editing', 'Close Without Saving'], defaultId: 0, cancelId: 0, noLink: true,
    });
    if (response === 1) event.preventDefault();
    else void updates.cancelInstall();
  });
  mainWindow.on('closed', () => { mainWindow = null; });
  mainWindow.once('ready-to-show', () => mainWindow?.show());
  try {
    if (IS_DEV) await mainWindow.loadURL(DEV_URL);
    else await mainWindow.loadFile(INDEX_PATH);
  } catch (error) {
    dialog.showErrorBox('Word Studio could not start', IS_DEV
      ? 'Start the development server with npm run dev, then run npm start -- --dev.'
      : `The editor could not be loaded. Rebuild with npm run build.\n\n${error.message}`);
    app.quit();
  }
}

if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.show(); mainWindow.focus();
    }
  });
  app.whenReady().then(async () => {
    if (process.platform === 'darwin') app.dock.setIcon(path.join(__dirname, 'WordStudio.png'));
    app.setAboutPanelOptions({ applicationName: PRODUCT_NAME, applicationVersion: app.getVersion(),
      copyright: 'An independent document editor. Not affiliated with Microsoft.' });
    hardenSession(); installIPC(); installMenu();
    await createWindow();
    await updates.start();
    app.on('activate', () => { if (!mainWindow) void createWindow(); });
  });
  app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
}
