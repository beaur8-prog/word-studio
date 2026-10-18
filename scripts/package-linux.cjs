'use strict';

const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { build, Platform, Arch } = require('electron-builder');
const asar = require('@electron/asar');
const { createLinuxIcon } = require('./create-icon.cjs');
const config = require('../electron-builder.linux.cjs');
const root = path.resolve(__dirname, '..');

async function hash(file) {
  const digest = createHash('sha256');
  for await (const chunk of fsSync.createReadStream(file)) digest.update(chunk);
  return digest.digest('hex');
}

async function run() {
  if (process.platform !== 'linux' || process.arch !== 'x64') throw new Error('Build the Linux x64 package on Linux x64.');
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  await fs.access(path.join(root, 'dist/index.html'));
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-linux-'));
  try {
    const source = path.join(temporary, 'app');
    await fs.mkdir(source);
    for (const directory of ['dist', 'electron']) await fs.cp(path.join(root, directory), path.join(source, directory), { recursive: true });
    await fs.writeFile(path.join(source, 'package.json'), JSON.stringify({
      name: manifest.name, version: manifest.version, description: manifest.description,
      author: 'Word Studio', homepage: 'https://github.com/beaur8-prog/word-studio',
      main: 'electron/main.cjs', desktopName: 'local.wordstudio.app',
    }, null, 2));
    const icon = await createLinuxIcon(path.join(temporary, 'icons'));
    const output = path.join(root, 'release/linux');
    await build({ targets: Platform.LINUX.createTarget(['deb'], Arch.x64), config: {
      ...config, electronVersion: manifest.devDependencies.electron,
      directories: { app: source, output }, linux: { ...config.linux, icon },
    } });
    const app = path.join(output, 'linux-unpacked');
    const archive = path.join(app, 'resources/app.asar');
    const packed = JSON.parse(asar.extractFile(archive, 'package.json'));
    if (packed.version !== manifest.version || packed.desktopName !== 'local.wordstudio.app') throw new Error('Linux app metadata mismatch.');
    for (const name of ['dist/index.html', 'electron/main.cjs', 'electron/preload.cjs', 'electron/WordStudio.png']) asar.extractFile(archive, name);
    const metadata = {
      product: 'Word Studio', version: manifest.version, platform: 'linux', architecture: 'x64',
      format: 'deb', signing: 'unsigned local Debian package', electronVersion: manifest.devDependencies.electron,
      executableSHA256: await hash(path.join(app, 'word-studio')),
      appAsarSHA256: await hash(archive),
      packageSHA256: await hash(path.join(output, `Word-Studio-${manifest.version}-linux-x64.deb`)),
      builtAt: new Date().toISOString(),
    };
    await fs.writeFile(path.join(output, 'BUILD-linux-x64.json'), JSON.stringify(metadata, null, 2) + '\n');
    console.log(JSON.stringify(metadata, null, 2));
  } finally { await fs.rm(temporary, { recursive: true, force: true }); }
}

run().catch(error => { console.error(error.stack); process.exitCode = 1; });
