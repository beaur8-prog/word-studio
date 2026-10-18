'use strict';

const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
const { createHash } = require('node:crypto');
const asar = require('@electron/asar');
const root = path.resolve(__dirname, '..');
const run = (program, args, options = {}) => execFileSync(program, args, { cwd: root, stdio: 'inherit', timeout: 240000, ...options });
async function hash(file) {
  const digest = createHash('sha256');
  for await (const chunk of fsSync.createReadStream(file)) digest.update(chunk);
  return digest.digest('hex');
}

async function test() {
  if (process.platform !== 'linux' || process.env.CI !== 'true') throw new Error('Run Linux install/reinstall/remove tests only on an isolated Linux CI runner.');
  const { version } = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  const directory = path.join(root, 'release/linux');
  const deb = path.join(directory, `Word-Studio-${version}-linux-x64.deb`);
  const metadata = JSON.parse(await fs.readFile(path.join(directory, 'BUILD-linux-x64.json'), 'utf8'));
  assert.equal(await hash(deb), metadata.packageSHA256);
  const report = { version, passed: false, distribution: await fs.readFile('/etc/os-release', 'utf8') };
  const app = '/opt/Word Studio';
  const binary = path.join(app, 'word-studio');
  const desktop = '/usr/share/applications/local.wordstudio.app.desktop';
  const recovery = path.join(process.env.XDG_CONFIG_HOME || path.join(os.homedir(), '.config'), 'Word Studio/recovery/installer-preservation-test.wordstudio');
  const document = path.join(os.homedir(), 'Documents/installer-preservation-test.wordstudio');
  const sentinel = 'PRESERVE WORD STUDIO USER DATA\n';
  assert.equal(fsSync.existsSync(binary), false, 'This runner already has Word Studio installed.');
  const installed = async () => {
    assert.equal(run('dpkg-query', ['-W', '-f=${Version} ${Architecture} ${Status}', 'word-studio'], { encoding: 'utf8', stdio: 'pipe' }), `${version} amd64 install ok installed`);
    assert.equal(await hash(binary), metadata.executableSHA256);
    assert.equal(await hash(path.join(app, 'resources/app.asar')), metadata.appAsarSHA256);
    assert.equal(JSON.parse(asar.extractFile(path.join(app, 'resources/app.asar'), 'package.json')).version, version);
    const entry = await fs.readFile(desktop, 'utf8');
    assert.match(entry, /^Name=Word Studio$/m);
    assert.match(entry, /^StartupWMClass=local\.wordstudio\.app$/m);
    assert.ok(!entry.includes('--no-sandbox'));
    run('desktop-file-validate', [desktop]);
    assert.equal(await fs.realpath('/usr/bin/word-studio'), binary);
  };
  const preserved = async () => {
    for (const file of [recovery, document]) assert.equal(await fs.readFile(file, 'utf8'), sentinel);
  };
  try {
    for (const file of [recovery, document]) { await fs.mkdir(path.dirname(file), { recursive: true }); await fs.writeFile(file, sentinel); }
    run('sudo', ['apt-get', 'install', '-y', deb]);
    await installed(); await preserved();
    report.freshInstall = true; report.desktopEntry = true; report.installedHashes = true;
    run(process.execPath, [path.join(root, 'scripts/desktop-smoke.cjs')], {
      env: { ...process.env, WORD_STUDIO_EXECUTABLE: binary },
    });
    report.installedApplication = true;
    run('sudo', ['apt-get', 'install', '--reinstall', '-y', deb]);
    await installed(); await preserved(); report.reinstall = true;
    run('sudo', ['apt-get', 'remove', '-y', 'word-studio']);
    for (const file of [binary, desktop, '/usr/bin/word-studio']) assert.equal(fsSync.existsSync(file), false, `Uninstall left ${file}`);
    await preserved(); report.uninstall = true; report.userDataPreserved = true; report.passed = true;
  } finally {
    await fs.mkdir(path.join(root, '.cache'), { recursive: true });
    await fs.writeFile(path.join(root, '.cache/linux-package-test.json'), JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify(report, null, 2));
  }
}

test().catch(error => { console.error(error.stack); process.exitCode = 1; });
