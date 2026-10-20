'use strict';

// Produces a self-contained Windows folder from a locally cached, verified
// official Electron release. No Microsoft runtime or Office assets are bundled.
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const root = path.resolve(__dirname, '..');
const locations = [root, path.dirname(root), path.join(path.dirname(root), 'slide-studio')];
const dependency = name => require(require.resolve(name, { paths: locations }));
const JSZip = dependency('jszip');
const asar = dependency('@electron/asar');
const ResEdit = dependency('resedit');
const { createWindowsIcon } = require('./create-icon.cjs');
const product = 'Word Studio';
const executable = product + '.exe';

async function hashFile(file) {
  const hash = createHash('sha256');
  for await (const chunk of fsSync.createReadStream(file)) hash.update(chunk);
  return hash.digest('hex');
}

async function runtime() {
  // An ancestor may contain only Electron's npm metadata, while a sibling has
  // the installed runtime. Match each package to an available verified ZIP.
  for (const location of locations) {
    const electronDirectory = path.join(location, 'node_modules', 'electron');
    if (!fsSync.existsSync(path.join(electronDirectory, 'package.json'))) continue;
    const { version } = JSON.parse(await fs.readFile(path.join(electronDirectory, 'package.json'), 'utf8'));
    const filename = `electron-v${version}-win32-x64.zip`;
    let file = locations.map(directory => path.join(directory, '.cache', 'electron', filename)).find(candidate => fsSync.existsSync(candidate));
    if (!file) {
      const { pathToFileURL } = require('node:url');
      const { downloadArtifact } = await import(pathToFileURL(require.resolve('@electron/get', { paths: locations })).href);
      file = await downloadArtifact({ version, artifactName: 'electron', platform: 'win32', arch: 'x64' });
    }
    const checksums = JSON.parse(await fs.readFile(path.join(electronDirectory, 'checksums.json'), 'utf8'));
    const expected = checksums[filename];
    if (!/^[a-f0-9]{64}$/.test(expected || '')) throw new Error('No published runtime checksum for ' + filename);
    if (await hashFile(file) !== expected) throw new Error('Cached Electron checksum mismatch.');
    return { version, file, sha256: expected, source: `https://github.com/electron/electron/releases/download/v${version}/${filename}` };
  }
  throw new Error('No matching official Electron ZIP is available in the local .cache/electron directories.');
}

function verifyPE(bytes) {
  if (bytes.toString('ascii', 0, 2) !== 'MZ') throw new Error('Missing DOS header.');
  const pe = bytes.readUInt32LE(0x3c);
  if (bytes.readUInt32LE(pe) !== 0x4550 || bytes.readUInt16LE(pe + 4) !== 0x8664 || bytes.readUInt16LE(pe + 24) !== 0x20b) {
    throw new Error('Expected a Windows x64 PE32+ executable.');
  }
}

async function rebrand(file, icon, version) {
  const original = await fs.readFile(file);
  verifyPE(original);
  const exe = ResEdit.NtExecutable.from(original, { ignoreCert: true });
  const codeHashes = exe.getAllSections().filter(section => section.info.characteristics & 0x20000000)
    .map(section => ({ name: section.info.name, sha256: createHash('sha256').update(Buffer.from(section.data)).digest('hex') }));
  const resources = ResEdit.NtExecutableResource.from(exe);
  const icons = ResEdit.Data.IconFile.from(await fs.readFile(icon));
  for (const group of ResEdit.Resource.IconGroupEntry.fromEntries(resources.entries)) {
    ResEdit.Resource.IconGroupEntry.replaceIconsForResource(resources.entries, group.id, group.lang, icons.icons.map(item => item.data));
  }
  for (const info of ResEdit.Resource.VersionInfo.fromEntries(resources.entries)) {
    info.setFileVersion(version + '.0'); info.setProductVersion(version + '.0');
    for (const language of info.getAllLanguagesForStringValues()) info.setStringValues(language, {
      CompanyName: product, FileDescription: product, ProductName: product, InternalName: product,
      OriginalFilename: executable, FileVersion: version, ProductVersion: version,
      LegalCopyright: 'Independent document editor. Not affiliated with Microsoft.',
    });
    info.outputToResourceEntries(resources.entries);
  }
  resources.entries = resources.entries.filter(entry => entry.type !== 'INTEGRITY');
  resources.outputResource(exe);
  const output = Buffer.from(exe.generate());
  verifyPE(output);
  const parsed = ResEdit.NtExecutable.from(output);
  for (const before of codeHashes) {
    const after = parsed.getAllSections().find(section => section.info.name === before.name);
    if (!after || createHash('sha256').update(Buffer.from(after.data)).digest('hex') !== before.sha256) throw new Error('Executable code section changed.');
  }
  const verifiedInfo = ResEdit.Resource.VersionInfo.fromEntries(ResEdit.NtExecutableResource.from(parsed).entries);
  if (!verifiedInfo.some(info => info.getAllLanguagesForStringValues().some(lang =>
    info.getStringValues(lang).ProductName === product && info.getStringValues(lang).ProductVersion === version))) {
    throw new Error('Windows product metadata verification failed.');
  }
  await fs.writeFile(path.join(path.dirname(file), executable), output);
  await fs.unlink(file);
}

async function removeGenerated(directory, allowedParent) {
  const resolved = path.resolve(directory);
  const parent = path.resolve(allowedParent);
  if (resolved === parent || !resolved.startsWith(parent + path.sep)) throw new Error('Refusing to remove a path outside the generated output directory.');
  await fs.rm(resolved, { recursive: true, force: true });
}

async function renameGenerated(source, destination) {
  for (let attempt = 0; ; attempt += 1) {
    try { await fs.rename(source, destination); return; }
    catch (error) {
      // Windows antivirus can briefly hold freshly created executable folders.
      if (attempt >= 8 || !['EPERM', 'EACCES', 'EBUSY'].includes(error.code)) throw error;
      await new Promise(resolve => setTimeout(resolve, 200 * (attempt + 1)));
    }
  }
}

async function packageWindows() {
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  if (!fsSync.existsSync(path.join(root, 'dist', 'index.html'))) throw new Error('Run npm run build first.');
  const electron = await runtime();
  const work = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-windows-'));
  try {
    const appFolder = path.join(work, product);
    await fs.mkdir(appFolder);
    const runtimeZip = await JSZip.loadAsync(await fs.readFile(electron.file), { checkCRC32: true });
    for (const [name, entry] of Object.entries(runtimeZip.files)) {
      if (name === 'resources/default_app.asar') continue;
      if (entry.unsafeOriginalName && entry.unsafeOriginalName !== name) throw new Error('Unsafe runtime ZIP entry.');
      const target = path.resolve(appFolder, name);
      if (!target.startsWith(appFolder + path.sep)) throw new Error('Unsafe runtime ZIP path.');
      if (entry.dir) await fs.mkdir(target, { recursive: true });
      else {
        await fs.mkdir(path.dirname(target), { recursive: true });
        await fs.writeFile(target, await entry.async('nodebuffer'));
      }
    }
    const icon = await createWindowsIcon(path.join(root, 'electron'));
    const source = path.join(work, 'app-source');
    await fs.mkdir(source);
    for (const directory of ['dist', 'electron']) await fs.cp(path.join(root, directory), path.join(source, directory), { recursive: true });
    await fs.writeFile(path.join(source, 'package.json'), JSON.stringify({
      name: manifest.name, version: manifest.version, description: manifest.description,
      author: 'Word Studio', main: 'electron/main.cjs',
    }, null, 2));
    const appAsar = path.join(appFolder, 'resources', 'app.asar');
    await asar.createPackage(source, appAsar);
    await rebrand(path.join(appFolder, 'electron.exe'), icon, manifest.version);
    await fs.copyFile(icon, path.join(appFolder, 'WordStudio.ico'));
    await fs.writeFile(path.join(appFolder, '시작하기.txt'), '\ufeff' + [
      `Word Studio ${manifest.version} | Windows 10/11 x64`, '',
      'Word Studio.exe를 실행하세요. Node.js, Microsoft Word, 인터넷 연결이 필요하지 않습니다.',
      'EXE만 이동하지 말고 resources, locales, DLL을 포함한 폴더 전체를 보관하세요.',
      'Ctrl+N 새 문서 / Ctrl+O 열기 / Ctrl+S 저장 / Ctrl+Shift+S 다른 이름으로 저장 / Ctrl+P 인쇄',
      '상단 내보내기에서 Word 문서(.docx), 한글 문서(.hwp/.hwpx), PDF(.pdf)를 선택하세요. 파일 메뉴에서는 HTML과 텍스트도 지원합니다.',
      '메모와 편집 상태를 보관하려면 편집본(.wordstudio)으로 저장하세요. Word와 주고받을 문서는 .docx로 저장하세요.', '',
      '이 앱은 Microsoft와 관련 없는 독립 문서 편집기입니다. Word의 모든 기능과 DOCX 요소를 지원하지는 않습니다.',
      '이 배포 파일에는 코드 서명이 없습니다.', '',
    ].join('\r\n'));
    const required = [executable, 'resources/app.asar', 'locales/ko.pak', 'locales/en-US.pak',
      'icudtl.dat', 'resources.pak', 'v8_context_snapshot.bin', 'ffmpeg.dll', 'libEGL.dll', 'libGLESv2.dll',
      'LICENSE', 'LICENSES.chromium.html'];
    for (const file of required) if (!(await fs.stat(path.join(appFolder, file))).size) throw new Error('Empty runtime file: ' + file);
    const packedManifest = JSON.parse(asar.extractFile(appAsar, 'package.json'));
    if (packedManifest.version !== manifest.version) throw new Error('App version mismatch.');
    for (const file of ['dist/index.html', 'electron/main.cjs', 'electron/preload.cjs', 'electron/WordStudio.ico']) asar.extractFile(appAsar, file);
    const releaseRoot = path.join(root, 'release');
    const release = path.resolve(releaseRoot, process.env.WORD_STUDIO_RELEASE_SUBDIR || 'windows');
    if (!release.startsWith(releaseRoot + path.sep)) throw new Error('Release output must stay inside release/.');
    await fs.mkdir(release, { recursive: true });
    const destination = path.join(release, product);
    const hadPrevious = fsSync.existsSync(destination);
    if (!hadPrevious) {
      // First publication has no existing files to replace. Copying the already
      // verified folder avoids directory-rename restrictions on some Windows hosts.
      await fs.cp(appFolder, destination, { recursive: true, force: false, errorOnExist: true });
    } else {
      const staged = await fs.mkdtemp(path.join(release, '.word-studio-stage-'));
      await fs.cp(appFolder, staged, { recursive: true });
      const backup = path.join(release, '.word-studio-previous-' + Date.now());
      await renameGenerated(destination, backup);
      try { await renameGenerated(staged, destination); }
      catch (error) {
        await renameGenerated(backup, destination);
        throw error;
      }
      await removeGenerated(backup, release);
    }
    const metadata = {
      product, version: manifest.version, platform: 'win32', architecture: 'x64',
      format: 'portable directory', signing: 'unsigned', builtAt: new Date().toISOString(),
      executable: path.relative(root, path.join(destination, executable)),
      appAsarSHA256: await hashFile(path.join(destination, 'resources', 'app.asar')),
      executableSHA256: await hashFile(path.join(destination, executable)),
      electronVersion: electron.version, electronSource: electron.source, electronSHA256: electron.sha256,
      checks: { runtimeChecksum: true, runtimeCRC32: true, peX64: true, productMetadata: true,
        codeSectionsPreserved: true, requiredRuntimeFiles: true, productionAsar: true, nativeWindowsExecution: false },
    };
    await fs.writeFile(path.join(release, 'BUILD-win-x64.json'), JSON.stringify(metadata, null, 2) + '\n');
    console.log(JSON.stringify(metadata, null, 2));
  } finally { await removeGenerated(work, os.tmpdir()); }
}

packageWindows().catch(error => { console.error(error.stack); process.exitCode = 1; });
