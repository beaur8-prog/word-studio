'use strict';

const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { createHash } = require('node:crypto');
const { createMacIcon } = require('./create-icon.cjs');
const root = path.resolve(__dirname, '..');
const product = 'Word Studio';
const appId = 'local.wordstudio.app';
const locations = [root, path.dirname(root), path.join(path.dirname(root), 'slide-studio')];
const run = (command, args) => execFileSync(command, args, { stdio: 'inherit' });
const plist = (file, key, value) => {
  execFileSync('/usr/bin/plutil', ['-replace', key, '-string', value, file]);
};

async function packageMac() {
  if (process.platform !== 'darwin') throw new Error('Build the macOS application on a Mac.');
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  const runtime = locations.map(dir => path.join(dir, 'node_modules', 'electron'))
    .find(dir => fsSync.existsSync(path.join(dir, 'dist', 'Electron.app', 'Contents', 'MacOS', 'Electron'))
      && JSON.parse(fsSync.readFileSync(path.join(dir, 'package.json'), 'utf8')).version === manifest.devDependencies.electron);
  if (!runtime) throw new Error('Install the matching Electron runtime before packaging.');
  if (!fsSync.existsSync(path.join(root, 'dist', 'index.html'))) throw new Error('Run npm run build first.');
  const asar = require(require.resolve('@electron/asar', { paths: locations }));
  const work = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-mac-'));
  // Keep a verified bundle outside file-provider managed Documents directories.
  const output = path.join(os.homedir(), 'Library', 'Caches', 'WordStudioBuild');
  const destination = path.join(output, product + '.app');
  try {
    await createMacIcon(path.join(root, 'electron'));
    const bundle = path.join(work, product + '.app');
    await fs.cp(path.join(runtime, 'dist', 'Electron.app'), bundle, { recursive: true, verbatimSymlinks: true });
    const contents = path.join(bundle, 'Contents');
    const resources = path.join(contents, 'Resources');
    const source = path.join(work, 'source');
    await fs.mkdir(source);
    for (const directory of ['dist', 'electron']) await fs.cp(path.join(root, directory), path.join(source, directory), { recursive: true });
    await fs.writeFile(path.join(source, 'package.json'), JSON.stringify({
      name: manifest.name, version: manifest.version, description: manifest.description,
      author: product, main: 'electron/main.cjs',
    }, null, 2));
    await fs.rm(path.join(resources, 'default_app.asar'));
    const archive = path.join(resources, 'app.asar');
    await asar.createPackage(source, archive);
    await fs.copyFile(path.join(root, 'electron', 'WordStudio.icns'), path.join(resources, 'WordStudio.icns'));
    for (const name of ['LICENSE', 'LICENSES.chromium.html']) await fs.copyFile(path.join(runtime, 'dist', name), path.join(resources, 'Electron-' + name));
    const info = path.join(contents, 'Info.plist');
    await fs.rename(path.join(contents, 'MacOS', 'Electron'), path.join(contents, 'MacOS', product));
    for (const [key, value] of Object.entries({
      CFBundleDisplayName: product, CFBundleName: product, CFBundleExecutable: product,
      CFBundleIdentifier: appId, CFBundleIconFile: 'WordStudio.icns',
      CFBundleShortVersionString: manifest.version, CFBundleVersion: manifest.version,
      LSApplicationCategoryType: 'public.app-category.productivity',
      NSHumanReadableCopyright: 'Independent document editor. Not affiliated with Microsoft.',
    })) plist(info, key, value);
    execFileSync('/usr/bin/plutil', ['-replace', 'CFBundleDocumentTypes', '-json',
      JSON.stringify(require('../electron/document-types.cjs').macDocumentTypes), info]);
    const integrity = { 'Resources/app.asar': { algorithm: 'SHA256',
      hash: createHash('sha256').update(asar.getRawHeader(archive).headerString).digest('hex') } };
    execFileSync('/usr/bin/plutil', ['-replace', 'ElectronAsarIntegrity', '-json', JSON.stringify(integrity), info]);
    const frameworks = path.join(contents, 'Frameworks');
    for (const suffix of ['', ' (GPU)', ' (Plugin)', ' (Renderer)']) {
      const oldName = 'Electron Helper' + suffix, newName = product + ' Helper' + suffix;
      const helper = path.join(frameworks, oldName + '.app');
      const helperContents = path.join(helper, 'Contents');
      await fs.rename(path.join(helperContents, 'MacOS', oldName), path.join(helperContents, 'MacOS', newName));
      const helperInfo = path.join(helperContents, 'Info.plist');
      for (const key of ['CFBundleExecutable', 'CFBundleName', 'CFBundleDisplayName']) plist(helperInfo, key, newName);
      plist(helperInfo, 'CFBundleIdentifier', appId + '.helper' + (suffix ? '.' + suffix.replace(/[ ()]/g, '').toLowerCase() : ''));
      await fs.rename(helper, path.join(frameworks, newName + '.app'));
    }
    run('/usr/bin/codesign', ['--force', '--deep', '--sign', '-', '--timestamp=none', bundle]);
    run('/usr/bin/codesign', ['--verify', '--deep', '--strict', bundle]);
    if (!process.argv.includes('--no-install')) {
      await fs.mkdir(output, { recursive: true });
      const backup = path.join(output, product + '-' + Date.now() + '.previous.app');
      const previous = fsSync.existsSync(destination);
      if (previous) await fs.rename(destination, backup);
      try {
        await fs.cp(bundle, destination, { recursive: true, verbatimSymlinks: true });
        run('/usr/bin/codesign', ['--verify', '--deep', '--strict', destination]);
      } catch (error) {
        await fs.rm(destination, { recursive: true, force: true });
        if (previous) await fs.rename(backup, destination);
        throw error;
      }
      if (previous) await fs.rm(backup, { recursive: true, force: true });
    }
    const release = path.join(root, 'release');
    await fs.mkdir(release, { recursive: true });
    const zip = path.join(release, `Word-Studio-${manifest.version}-mac-${process.arch}.zip`);
    run('/usr/bin/ditto', ['-c', '-k', '--keepParent', bundle, zip]);
    await fs.writeFile(path.join(release, 'BUILD-mac.json'), JSON.stringify({
      product, version: manifest.version, platform: 'darwin', architecture: process.arch,
      bundle: process.argv.includes('--no-install') ? null : destination, archive: zip, electronVersion: manifest.devDependencies.electron,
      signing: 'ad-hoc (local use; not Apple notarized)', builtAt: new Date().toISOString(),
    }, null, 2));
    console.log(`Built ${process.argv.includes("--no-install") ? "release archive without replacing a local app" : destination}\nArchive ${zip}`);
  } finally { await fs.rm(work, { recursive: true, force: true }); }
}
packageMac().catch(error => { console.error(error.stack); process.exitCode = 1; });
