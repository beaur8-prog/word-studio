'use strict';
// Explicit maintenance command; never invoked by npm install or the application.
const fs = require('node:fs/promises');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { createHash } = require('node:crypto');
const root = path.resolve(__dirname, '..');
const revision = '1a76570e833917d15817415a53c09ad61ab3203f';
const destination = path.join(root, 'vendor/rhwp-core');
const source = path.join(root, '.cache/rhwp-word-studio-rebuild');
const run = (command, args, cwd = source) => execFileSync(command, args, { cwd, stdio: 'inherit' });
async function main() {
  for (const command of ['git', 'rustc', 'cargo', 'wasm-pack', 'wasm-bindgen']) run(command, ['--version'], root);
  if ((await fs.stat(source).catch(() => null))) throw Error('Rebuild checkout already exists; move it aside before rebuilding: ' + source);
  await fs.mkdir(path.dirname(source), { recursive: true });
  run('git', ['clone', '--filter=blob:none', '--no-checkout', 'https://github.com/edwardkim/rhwp.git', source], root);
  run('git', ['sparse-checkout', 'set', 'src', 'crates', 'bindings', 'scripts', 'mydocs/manual', 'mydocs/tech/tools/vendor/saved']);
  run('git', ['checkout', '--detach', revision]);
  run('git', ['apply', '--check', path.join(destination, 'pagination.patch')]);
  run('git', ['apply', path.join(destination, 'pagination.patch')]);
  run('bash', ['scripts/wasm-pack-locked.sh', '--target', 'web', '--out-dir', 'word-studio-pkg', '--mode', 'normal']);
  const files = ['rhwp.js', 'rhwp.d.ts', 'rhwp_bg.wasm', 'rhwp_bg.wasm.d.ts'];
  for (const file of files) await fs.copyFile(path.join(source, 'word-studio-pkg', file), path.join(destination, file));
  const sums = await Promise.all(files.map(async file => createHash('sha256').update(await fs.readFile(path.join(destination, file))).digest('hex') + '  ' + file));
  await fs.writeFile(path.join(destination, 'SHA256SUMS'), sums.join('\n') + '\n');
  console.log('Rebuilt vendor/rhwp-core. Review artifact hashes and run HWP regression/corpus tests before committing.');
}
main().catch(error => { console.error(error); process.exitCode = 1; });
