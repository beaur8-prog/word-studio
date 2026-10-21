'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const { randomUUID } = require('node:crypto');

/** Keep the native working folder out of portable documents and recovery data. */
function createFileDialogDirectory(settingsFile, fallbackDirectory) {
  let directory;
  let pending = fs.readFile(settingsFile, 'utf8').then(contents => {
    const saved = JSON.parse(contents);
    if (saved.version === 1 && typeof saved.directory === 'string' && path.isAbsolute(saved.directory)) directory = saved.directory;
  }).catch(() => {});

  const exists = async candidate => typeof candidate === 'string' && path.isAbsolute(candidate)
    && await fs.stat(candidate).then(stat => stat.isDirectory()).catch(() => false);

  async function current() {
    await pending;
    if (await exists(directory)) return directory;
    if (await exists(fallbackDirectory)) return fallbackDirectory;
    return undefined; // Let the OS choose when neither location is available.
  }

  function rememberFile(filename) {
    const task = pending.then(async () => {
      if (typeof filename !== 'string' || !path.isAbsolute(filename)) return;
      const next = path.dirname(filename);
      if (!await exists(next)) return;
      directory = next;
      const temporary = `${settingsFile}.${randomUUID()}.tmp`;
      try {
        await fs.mkdir(path.dirname(settingsFile), { recursive: true });
        await fs.writeFile(temporary, JSON.stringify({ version: 1, directory }), { flag: 'wx', mode: 0o600 });
        await fs.rename(temporary, settingsFile);
      } catch {
        // A preference write must not turn a completed document save into a
        // reported failure. Keep the chosen folder for this app session.
      } finally { await fs.rm(temporary, { force: true }).catch(() => {}); }
    });
    pending = task.catch(() => {});
    return task;
  }

  return { current, rememberFile };
}

module.exports = { createFileDialogDirectory };
