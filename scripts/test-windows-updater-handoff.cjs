'use strict';
// Adapted from office-studio PR #1 (e3cf073), using the merged Word updater.
// Disposable native Electron + Windows PowerShell checks. The test installers
// only copy a versioned fixture EXE inside RUNNER_TEMP; no Studio app is changed.
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const assert = require('node:assert/strict');
const { spawn, execFileSync } = require('node:child_process');
const { powershellCommand } = require('../electron/studio-updater.cjs');
if (process.platform !== 'win32' || process.env.GITHUB_ACTIONS !== 'true') throw Error('Requires a disposable Windows CI runner');
const root = path.resolve(__dirname, '..');
const report = path.join(root, 'test-results', 'updater-handoff');
const registryKey = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\StudioHandoff-' + process.pid;
const ps = path.join(process.env.SystemRoot, 'System32/WindowsPowerShell/v1.0/powershell.exe');
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
async function poll(check, timeout = 30000) { const end = Date.now() + timeout; while (Date.now() < end) { if (await check()) return; await delay(100); } throw Error('Timed out waiting for fixture'); }
async function exists(file) { return fs.access(file).then(() => true, () => false); }

(async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "studio handoff 한글 O'Brien; & "));
  await fs.mkdir(report, { recursive: true });
  const assembly = version => `using System; using System.IO; using System.Reflection; [assembly: AssemblyVersion("${version}.0")] [assembly: AssemblyInformationalVersion("${version}")] `;
  await fs.writeFile(path.join(directory, 'old.cs'), assembly('0.0.0') + 'public class OldApp { public static void Main() {} }');
  await fs.writeFile(path.join(directory, 'payload.cs'), assembly('1.0.0') + 'public class NewApp { public static void Main() { File.WriteAllText(Path.Combine(Environment.GetEnvironmentVariable("STUDIO_HANDOFF_FIXTURE_ROOT"), "reopened"), "yes"); } }');
  await fs.writeFile(path.join(directory, 'installer.cs'), 'using System; using System.IO; using System.Diagnostics; public class Setup { public static void Main() { try { var parent = Process.GetProcessById(int.Parse(Environment.GetEnvironmentVariable("STUDIO_HANDOFF_PARENT_PID"))); if (!parent.HasExited) throw new Exception("Editor is still running"); } catch (ArgumentException) {} var root = Environment.GetEnvironmentVariable("STUDIO_HANDOFF_FIXTURE_ROOT"); File.Copy(Path.Combine(root, "payload.exe"), Environment.GetEnvironmentVariable("STUDIO_HANDOFF_TARGET_EXE"), true); File.WriteAllText(Path.Combine(root, "installed"), "yes"); } }');
  const compile = `param([string]$JobFile)\n$ErrorActionPreference = 'Stop'\nforeach ($name in @('old','payload','installer')) { Add-Type -Path (Join-Path $JobFile "$name.cs") -OutputAssembly (Join-Path $JobFile "$name.exe") -OutputType WindowsApplication }`;
  execFileSync(ps, ['-NoLogo', '-NoProfile', '-NonInteractive', '-EncodedCommand', powershellCommand(compile, directory)], { stdio: 'inherit', timeout: 60000 });
  const runtime = path.dirname(require('electron'));
  const appDirectory = path.join(runtime, 'resources', 'app');
  await fs.mkdir(appDirectory, { recursive: true });
  await fs.rm(path.join(runtime, 'resources', 'default_app.asar'), { force: true });
  await fs.writeFile(path.join(appDirectory, 'package.json'), JSON.stringify({ name: 'studio-updater-handoff-fixture', version: '0.0.0', main: 'main.cjs' }));
  const nativeMain = `
    const { app, BrowserWindow } = require('electron');
    const fs = require('node:fs'), path = require('node:path'), crypto = require('node:crypto');
    const { spawn } = require('node:child_process');
    const { createStudioUpdater } = require(${JSON.stringify(path.join(root, 'electron/studio-updater.cjs'))});
    const folder = process.env.STUDIO_HANDOFF_FIXTURE_ROOT, id = process.env.STUDIO_HANDOFF_ID, mode = process.env.STUDIO_HANDOFF_MODE;
    process.env.STUDIO_HANDOFF_PARENT_PID = String(process.pid);
    app.setPath('userData', path.join(folder, 'profile'));
    const bytes = fs.readFileSync(path.join(folder, 'installer.exe')), digest = crypto.createHash('sha256').update(bytes).digest('hex');
    const filename = {sheet:'Sheet-Studio-1.0.0-Windows-x64-Setup.exe',slide:'Slide-Studio-Setup-x64.exe',word:'Word-Studio-1.0.0-windows-x64-setup.exe'}[id];
    const item = {tag_name:'v1.0.0',draft:false,prerelease:false,assets:[{name:filename,state:'uploaded',size:bytes.length,digest:'sha256:'+digest,browser_download_url:'https://github.com/beaur8-prog/'+id+'-studio/releases/download/v1.0.0/'+filename}]};
    let win, updater, canceled = false;
    const write = (name, value) => fs.writeFileSync(path.join(folder, name), JSON.stringify(value));
    app.on('window-all-closed', () => app.quit());
    app.on('before-quit', () => { const p = path.join(folder,'profile','updates',digest); write('before-quit.json', {acknowledged:fs.existsSync(p) && fs.readdirSync(p).some(n=>n.endsWith('.ready'))}); });
    app.whenReady().then(async () => {
      win = new BrowserWindow({show:true,webPreferences:{sandbox:true,contextIsolation:true,nodeIntegration:false}});
      await win.loadURL('data:text/html,<h1>Disposable update handoff fixture</h1>');
      if (!app.isPackaged) throw Error('Fixture must use packaged Electron lifecycle');
      updater = createStudioUpdater({app,id,getWindow:()=>win,executable:process.env.STUDIO_HANDOFF_TARGET_EXE,
        dialog:{showMessageBox:async(...args)=>{ const d=args.at(-1); if(d.type==='warning'){write('failure.json',{message:d.message,buttons:d.buttons,editorStillOpen:!win.isDestroyed()});return{response:1};}return{response:0};}},
        shell:{showItemInFolder:()=>{},openExternal:async()=>{}},
        fetchImpl:async url=>url.includes('api.github.com')?Response.json(item):new Response(bytes),
        launch:(exe,args,options)=>{
          const command=Buffer.from(/-EncodedCommand ([A-Za-z0-9+/=]+)$/.exec(args.at(-1))[1],'base64').toString('utf16le');
          write('launch.json',{exe,commandLength:command.length,jobFile:command.slice(command.lastIndexOf('-JobFile '))});
          fs.writeFileSync(path.join(folder,'worker-command.txt'),command);
          if(mode==='tampered'){
            const jobFile=/-JobFile '([^']*(?:''[^']*)*)'$/.exec(command)[1].replaceAll("''", "'");
            fs.appendFileSync(jobFile+'.ps1','\\n# modified after verification');
          }
          const child=mode==='blocked'?spawn(process.env.ComSpec,['/d','/c','exit','9'],options):spawn(exe,args,options);
          child.on('exit',code=>write('worker-exit.json',{code}));return child;
        },
      });
      if(mode==='cancel') win.on('close', event=>{if(!canceled){event.preventDefault();canceled=true;void updater.cancelInstall().then(()=>{write('cancelled.json',{editorStillOpen:!win.isDestroyed()});setTimeout(()=>{win.destroy();app.quit();},1000);});}});
      await updater.check(true);
      if(mode==='blocked'||mode==='tampered'||mode==='unregistered'){write('blocked-complete.json',{phase:updater.getState().phase});win.destroy();app.quit();}
    }).catch(error=>{write('fixture-error.json',{error:String(error.stack)});app.exit(1);});
  `;
  await fs.writeFile(path.join(appDirectory, 'main.cjs'), nativeMain);
  const results = [];
  try {
    for (const [id, mode] of [['sheet', 'install'], ['slide', 'install'], ['word', 'install'], ['word', 'cancel'], ['word', 'blocked'], ['word', 'tampered'], ['word', 'unregistered']]) {
      const name = id + '-' + mode, folder = path.join(directory, name), product = { sheet: 'Sheet Studio', slide: 'Slide Studio', word: 'Word Studio' }[id];
      await fs.mkdir(folder);
      for (const file of ['old.exe', 'payload.exe', 'installer.exe']) await fs.copyFile(path.join(directory, file), path.join(folder, file));
      const target = path.join(folder, product + '.exe');
      await fs.copyFile(path.join(folder, 'old.exe'), target);
      await fs.writeFile(path.join(folder, 'Uninstall ' + product + '.exe'), 'disposable uninstaller marker');
      if (id === 'word') {
        const registration = mode === 'unregistered'
          ? `param($JobFile)\nRemove-Item -LiteralPath '${registryKey}' -Recurse -Force -ErrorAction SilentlyContinue`
          : `param($JobFile)\nNew-Item -Path '${registryKey}' -Force | Out-Null\nNew-ItemProperty -LiteralPath '${registryKey}' -Name DisplayName -Value 'Word Studio' -Force | Out-Null\nNew-ItemProperty -LiteralPath '${registryKey}' -Name UninstallString -Value ('"' + (Join-Path $JobFile 'Uninstall Word Studio.exe') + '" /currentuser') -Force | Out-Null`;
        execFileSync(ps, ['-NoLogo','-NoProfile','-NonInteractive','-EncodedCommand',powershellCommand(registration,folder)], {stdio:'inherit'});
      }
      const binary = path.join(runtime, product + '.exe'); await fs.copyFile(path.join(runtime, 'electron.exe'), binary);
      const env = { ...process.env, STUDIO_HANDOFF_FIXTURE_ROOT: folder, STUDIO_HANDOFF_TARGET_EXE: target, STUDIO_HANDOFF_ID: id, STUDIO_HANDOFF_MODE: mode }; delete env.ELECTRON_RUN_AS_NODE;
      const child = spawn(binary, [], { env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
      let output = ''; child.stdout.on('data', b => { output += b; }); child.stderr.on('data', b => { output += b; });
      let exit; child.on('exit', code => { exit = code; }); child.on('error', error => { output += error.stack; exit = -1; });
      try {
        await poll(async () => exit !== undefined || mode === 'install' && await exists(path.join(folder, 'failure.json')), 45000);
        if (mode === 'install' && await exists(path.join(folder, 'failure.json'))) {
          const startup = await fs.readFile(path.join(folder,'profile/updates/helper-startup.log'),'utf8');
          throw Error('Native helper failed: ' + await fs.readFile(path.join(folder,'failure.json'),'utf8') + '\n' + startup);
        }
        assert.equal(exit, 0, output);
        const fixtureError = await exists(path.join(folder, 'fixture-error.json')); assert.equal(fixtureError, false, fixtureError ? await fs.readFile(path.join(folder, 'fixture-error.json'), 'utf8') : '');
        if (mode === 'install') {
          await poll(() => exists(path.join(folder, 'reopened')));
          assert(await exists(path.join(folder, 'installed')));
          assert.equal(JSON.parse(await fs.readFile(path.join(folder, 'before-quit.json'))).acknowledged, true);
          assert.match(await fs.readFile(path.join(folder, 'profile/updates/last-install.log'), 'utf8'), /Installed /);
          results.push({ id, mode, nativeElectronQuit: true, helperAcknowledgedBeforeQuit: true, installed: true, versionVerified: true, reopened: true });
        } else {
          await delay(1500); assert.equal(await exists(path.join(folder, 'installed')), false);
          const detail = JSON.parse(await fs.readFile(path.join(folder, mode === 'cancel' ? 'cancelled.json' : 'failure.json')));
          assert.equal(detail.editorStillOpen, true);
          results.push({ id, mode, editorKeptOpen: true, noInstallation: true });
        }
      } finally {
        if (exit === undefined) child.kill();
        await fs.writeFile(path.join(folder, 'electron-output.log'), output);
        await fs.cp(folder, path.join(report, name), { recursive: true });
      }
      console.log('Verified native updater handoff:', name);
    }
    await fs.writeFile(path.join(report, 'results.json'), JSON.stringify({ platform: process.platform, passed: true, results }, null, 2) + '\n');
  } finally {
    execFileSync(ps, ['-NoLogo','-NoProfile','-NonInteractive','-EncodedCommand',powershellCommand(`param($JobFile)\nRemove-Item -LiteralPath '${registryKey}' -Recurse -Force -ErrorAction SilentlyContinue`,directory)], {stdio:'inherit'});
    await fs.rm(directory, { recursive: true, force: true });
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
