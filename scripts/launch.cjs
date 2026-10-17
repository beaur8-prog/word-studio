'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { spawn, spawnSync } = require('node:child_process');
const root = path.resolve(__dirname, '..');
const locations = [root, path.dirname(root), path.join(path.dirname(root), 'slide-studio')];
const suffix = process.platform === 'darwin' ? 'Electron.app/Contents/MacOS/Electron'
  : process.platform === 'win32' ? 'electron.exe' : 'electron';
const executable = locations.map(directory => path.join(directory, 'node_modules', 'electron', 'dist', suffix))
  .find(filename => fs.existsSync(filename));
const isDev = process.argv.includes('--dev');

async function launch() {
  if (!executable) throw new Error('Electron is unavailable. Install Electron locally, or launch the portable Word Studio release.');
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE;
  if (isDev) env.WORD_STUDIO_DEV = '1';
  else delete env.WORD_STUDIO_DEV;

  if (!isDev && !fs.existsSync(path.join(root, 'dist', 'index.html'))) {
    console.log('Building Word Studio…');
    for (const [moduleName, entry, args] of [['typescript', 'bin/tsc', ['--noEmit']], ['vite', 'bin/vite.js', ['build']]]) {
      const script = path.join(path.dirname(require.resolve(moduleName + '/package.json', { paths: [root] })), entry);
      const build = spawnSync(process.execPath, [script, ...args], {
        cwd: root, env, stdio: 'inherit', windowsHide: true,
      });
      if (build.error) throw build.error;
      if (build.status !== 0) throw new Error('The application build did not complete.');
    }
  }
  if (!fs.existsSync(path.join(root, 'electron', 'WordStudio.ico'))) {
    await require('./create-icon.cjs').createWindowsIcon(path.join(root, 'electron'));
  }
  const extraArgs = process.argv.slice(2).filter(argument => argument !== '--dev');
  const child = spawn(executable, [root, ...extraArgs], { cwd: root, stdio: 'inherit', env, windowsHide: true });
  child.on('error', error => { console.error(error.message); process.exitCode = 1; });
  child.on('exit', code => { process.exitCode = code ?? 1; });
}

launch().catch(error => { console.error(error.message); process.exitCode = 1; });
