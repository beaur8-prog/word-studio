'use strict';

// Real Launch Services requests (the Finder double-click path), with a unique
// bundle ID and disposable user profile. No default file associations change.
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const net = require('node:net');
const { execFileSync } = require('node:child_process');
const { chromium, expect } = require('@playwright/test');
const { Document, Paragraph, Packer } = require('docx');

(async () => {
  if (process.platform !== 'darwin') throw new Error('Launch Services requires macOS.');
  const source = process.argv[2]; if (!source) throw new Error('Pass a packaged .app bundle.');
  const output = path.resolve(process.argv[3] || '.cache/file-open/launch-services'); await fs.mkdir(output, { recursive: true });
  // Launch Services deliberately omits apps in the OS temp directory from
  // Open With. Use a disposable cache bundle so this tests real candidates.
  const temporary = await fs.mkdtemp(path.join(os.homedir(), 'Library/Caches/word-launch-services-'));
  const bundle = path.join(temporary, 'Word File Open QA.app'), profile = path.join(temporary, 'profile');
  const first = path.join(temporary, 'Finder 첫 번째 한글 문서.docx'), second = path.join(temporary, 'Finder 두 번째 문서.txt');
  let browser, pid;
  const lsregister = '/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister';
  const run = (command, args) => execFileSync(command, args, { stdio: 'pipe' });
  try {
    run('/usr/bin/ditto', [source, bundle]);
    // File-provider folders can add Finder metadata to an unpacked build.
    // Remove it only from this disposable copy before giving it a QA identity.
    run('/usr/bin/xattr', ['-cr', bundle]);
    const info = path.join(bundle, 'Contents/Info.plist');
    run('/usr/bin/plutil', ['-replace', 'CFBundleIdentifier', '-string', `local.wordstudio.fileopen.qa${process.pid}`, info]);
    // Electron locates its helper executables using CFBundleName. Keep that
    // name and give only the QA display identity a distinct visible label.
    run('/usr/bin/plutil', ['-replace', 'CFBundleDisplayName', '-string', 'Word File Open QA', info]);
    run('/usr/bin/codesign', ['--force', '--deep', '--sign', '-', '--timestamp=none', bundle]);
    run('/usr/bin/codesign', ['--verify', '--deep', '--strict', bundle]);
    const metadata = JSON.parse(run('/usr/bin/plutil', ['-convert', 'json', '-o', '-', info]));
    expect(metadata.CFBundleDocumentTypes.some(type => type.CFBundleTypeExtensions.includes('docx'))).toBe(true);
    expect(metadata.CFBundleDocumentTypes.every(type => type.LSHandlerRank === 'Alternate')).toBe(true);
    await fs.writeFile(first, await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('Finder 더블클릭 첫 번째 DOCX 본문')] }] })));
    await fs.writeFile(second, 'Finder 실행 중 두 번째 문서 본문');
    const candidates = [];
    for (const ext of require('../electron/document-types.cjs').OPEN_EXTENSIONS) {
      const filename = path.join(temporary, '지원 확장자.' + ext);
      if (ext === 'docx') await fs.copyFile(first, filename);
      else if (['hwp', 'hwpx'].includes(ext)) await fs.copyFile(path.resolve(__dirname, '../tests/fixtures/hwp/reader.' + ext), filename);
      else await fs.writeFile(filename, ext === 'html' || ext === 'htm' ? '<html><body>Document</body></html>' : 'Document');
      candidates.push(filename);
    }
    run(lsregister, ['-f', bundle]);
    const query = path.join(temporary, 'document-handlers');
    run('/usr/bin/xcrun', ['clang', '-fobjc-arc', '-framework', 'AppKit', path.join(__dirname, 'query-mac-document-handlers.m'), '-o', query]);
    const registeredBundle = await fs.realpath(bundle);
    const checkHandlers = async () => { await expect.poll(() => {
      const handlers = JSON.parse(run(query, candidates));
      fs.writeFile(path.join(output, 'open-with-handlers.json'), JSON.stringify(handlers, null, 2)).catch(() => {});
      return candidates.filter(filename => !handlers[filename]?.includes(registeredBundle)).map(filename => path.extname(filename));
    }, { timeout: 10000, message: 'Every supported extension must offer the isolated app in Finder Open With' }).toEqual([]); };
    const server = net.createServer(); await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
    const port = server.address().port; await new Promise(resolve => server.close(resolve));
    run('/usr/bin/open', ['-n', '--stderr', path.join(output, 'app-stderr.log'), '--stdout', path.join(output, 'app-stdout.log'),
      '-a', bundle, first, '--args', '--user-data-dir=' + profile, '--remote-debugging-port=' + port]);
    await expect.poll(async () => {
      try { const response = await fetch(`http://127.0.0.1:${port}/json/version`); return response.status; } catch { return 0; }
    }, { timeout: 30000 }).toBe(200);
    browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`);
    const page = browser.contexts()[0].pages()[0]; page.setDefaultTimeout(15000);
    await expect(page.locator('.document-stack .tiptap')).toHaveText('Finder 더블클릭 첫 번째 DOCX 본문');
    await checkHandlers();
    // Finder returns decomposed Unicode filenames on macOS.
    await expect.poll(async () => (await page.getByRole('textbox', { name: '문서 제목' }).inputValue()).normalize('NFC')).toBe('Finder 첫 번째 한글 문서');
    run('/usr/bin/open', ['-a', bundle, second]);
    await expect(page.locator('.document-stack .tiptap')).toHaveText('Finder 실행 중 두 번째 문서 본문');
    await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
    expect(browser.contexts()[0].pages().length).toBe(1);
    await page.screenshot({ path: path.join(output, 'finder-opened.png') });
    const processes = run('/bin/ps', ['-axo', 'pid,command']).toString().split('\n');
    const processLine = processes.find(line => line.includes(bundle + '/Contents/MacOS/Word Studio') && !line.includes('--type='));
    if (!processLine) throw new Error('Isolated Launch Services process was not found.');
    pid = Number(processLine.trim().split(/\s/)[0]);
    await browser.close(); browser = null;
    const result = { passed: true, isolatedBundle: true, isolatedProfile: true, defaultAssociationsUnchanged: true,
      documentDeclarations: metadata.CFBundleDocumentTypes.length, registeredExtensions: candidates.map(file => path.extname(file).slice(1)),
      finderOpenWith: true, finderColdOpen: true, finderRunningOpen: true, oneEditorWindow: true };
    await fs.writeFile(path.join(output, 'summary.json'), JSON.stringify(result, null, 2) + '\n'); console.log(JSON.stringify(result));
  } catch (error) {
    const diagnostics = { error: String(error), processLines: run('/bin/ps', ['-axo', 'pid,command']).toString().split('\n').filter(line => line.includes(bundle)),
      profileFiles: await fs.readdir(profile).catch(() => []), runAsNode: process.env.ELECTRON_RUN_AS_NODE ?? null };
    await fs.writeFile(path.join(output, 'failure.json'), JSON.stringify(diagnostics, null, 2));
    throw error;
  } finally {
    if (browser) await browser.close().catch(() => {});
    if (!pid) {
      const lines = run('/bin/ps', ['-axo', 'pid,command']).toString().split('\n');
      const line = lines.find(line => line.includes(bundle + '/Contents/MacOS/Word Studio') && !line.includes('--type='));
      if (line) pid = Number(line.trim().split(/\s/)[0]);
    }
    if (pid) {
      try { process.kill(pid, 'SIGTERM'); } catch {}
      await expect.poll(() => { try { process.kill(pid, 0); return true; } catch { return false; } }, { timeout: 10000 }).toBe(false);
    }
    try { run(lsregister, ['-u', bundle]); } catch {}
    await fs.rm(temporary, { recursive: true, force: true });
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
