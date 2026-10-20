'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const MAX = require('./document-limits.json').studioDocumentBytes;
const ID = /^[a-zA-Z0-9_-]{1,128}$/;

function createRecoveryStore(directory) {
  let pending = Promise.resolve();
  async function list() {
    await pending;
    await fs.mkdir(directory, { recursive: true });
    const records = [];
    for (const name of await fs.readdir(directory)) {
      if (!name.endsWith('.json') || !ID.test(name.slice(0, -5))) continue;
      try {
        const file = path.join(directory, name);
        // Two snapshots are JSON strings inside another JSON record. Escaping
        // quotes/backslashes can double each snapshot's encoded size again.
        if ((await fs.stat(file)).size > MAX * 4 + 4096) continue;
        const record = JSON.parse(await fs.readFile(file, 'utf8'));
        if (record && typeof record.current === 'string') records.push(record);
      } catch { /* Keep the damaged file; other recovery documents remain available. */ }
    }
    return records.sort((a, b) => String(b.savedAt).localeCompare(String(a.savedAt))).slice(0, 20);
  }
  function save(serialized) {
    const task = pending.then(async () => {
      if (typeof serialized !== 'string' || Buffer.byteLength(serialized) > MAX) throw new Error(`복구 문서가 ${MAX / 1024 / 1024} MiB를 초과합니다.`);
      const doc = JSON.parse(serialized);
      if (doc.format !== 'word-studio' || !ID.test(doc.id || '')) throw new Error('복구 문서 ID가 올바르지 않습니다.');
      await fs.mkdir(directory, { recursive: true });
      const filename = path.join(directory, doc.id + '.json');
      let previous;
      try { previous = JSON.parse(await fs.readFile(filename, 'utf8')).current; } catch {}
      const temporary = path.join(directory, '.' + randomUUID() + '.tmp');
      try {
        const handle = await fs.open(temporary, 'wx', 0o600);
        try {
          await handle.writeFile(JSON.stringify({ current: serialized, previous, savedAt: new Date().toISOString() }));
          await handle.sync();
        } finally { await handle.close(); }
        await fs.rename(temporary, filename);
      } finally { await fs.rm(temporary, { force: true }).catch(() => {}); }
      return true;
    });
    pending = task.catch(() => {});
    return task;
  }
  return { list, save };
}
module.exports = { createRecoveryStore };
