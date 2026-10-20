'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const MAX = require('./document-limits.json').studioDocumentBytes;
const ID = /^[a-zA-Z0-9_-]{1,128}$/;
const INLINE_MAX = 64 * 1024 * 1024;

function createRecoveryStore(directory) {
  let pending = Promise.resolve();
  const snapshotName = (id, name) => typeof name === 'string' && name.startsWith(id + '.') && /^[a-zA-Z0-9_-]+\.[0-9a-f-]{36}\.snapshot$/.test(name);
  async function readSnapshot(id, name) {
    if (!snapshotName(id, name)) throw new Error('잘못된 복구 파일 경로입니다.');
    const file = path.join(directory, name);
    if ((await fs.stat(file)).size > MAX) throw new Error('복구 파일이 너무 큽니다.');
    return fs.readFile(file, 'utf8');
  }
  async function readRecord(filename) {
    // Legacy records contain two escaped JSON strings; all new large records
    // use small manifests so no single JSON string approaches V8's ceiling.
    if ((await fs.stat(filename)).size > 512 * 1024 * 1024 - 4096) throw new Error('복구 기록이 너무 큽니다.');
    return JSON.parse(await fs.readFile(filename, 'utf8'));
  }
  async function atomicRecord(filename, record) {
    const temporary = path.join(directory, '.' + randomUUID() + '.tmp');
    try {
      const handle = await fs.open(temporary, 'wx', 0o600);
      try { await handle.writeFile(JSON.stringify(record)); await handle.sync(); }
      finally { await handle.close(); }
      await fs.rename(temporary, filename);
    } finally { await fs.rm(temporary, { force: true }).catch(() => {}); }
  }
  async function writeSnapshot(id, serialized) {
    const name = id + '.' + randomUUID() + '.snapshot', filename = path.join(directory, name);
    const handle = await fs.open(filename, 'wx', 0o600);
    let failure;
    try { await handle.writeFile(serialized); await handle.sync(); } catch (error) { failure = error; }
    try { await handle.close(); } catch (error) { failure ||= error; }
    if (failure) { await fs.rm(filename, { force: true }).catch(() => {}); throw failure; }
    return name;
  }
  async function readList() {
    await fs.mkdir(directory, { recursive: true });
    const candidates = [], records = [];
    for (const name of await fs.readdir(directory)) {
      if (!name.endsWith('.json') || !ID.test(name.slice(0, -5))) continue;
      try {
        const record = await readRecord(path.join(directory, name));
        if (!record || (record.version !== 2 && typeof record.current !== 'string')) continue;
        candidates.push({ id: name.slice(0, -5), record });
        candidates.sort((a, b) => String(b.record.savedAt).localeCompare(String(a.record.savedAt)));
        if (candidates.length > 20) candidates.pop();
      } catch { /* Keep the damaged file; other recovery documents remain available. */ }
    }
    for (const { id, record } of candidates) {
      if (record.version !== 2) { records.push(record); continue; }
      const current = await readSnapshot(id, record.currentFile).catch(() => '');
      const previous = await readSnapshot(id, record.previousFile).catch(() => undefined);
      if (current || previous) records.push({ current, previous, savedAt: record.savedAt });
    }
    return records;
  }
  function list() {
    // Keep snapshot cleanup from racing a read on the same store.
    const task = pending.then(readList);
    pending = task.catch(() => {});
    return task;
  }
  function save(serialized) {
    const task = pending.then(async () => {
      if (typeof serialized !== 'string' || Buffer.byteLength(serialized) > MAX) throw new Error(`복구 문서가 ${MAX / 1024 / 1024} MiB를 초과합니다.`);
      const doc = JSON.parse(serialized);
      if (doc.format !== 'word-studio' || !ID.test(doc.id || '')) throw new Error('복구 문서 ID가 올바르지 않습니다.');
      await fs.mkdir(directory, { recursive: true });
      const filename = path.join(directory, doc.id + '.json');
      let prior;
      try { prior = await readRecord(filename); } catch {}
      if (prior?.version !== 2 && Buffer.byteLength(serialized) <= INLINE_MAX &&
          (!prior?.current || typeof prior.current === 'string' && Buffer.byteLength(prior.current) <= INLINE_MAX)) {
        await atomicRecord(filename, { current: serialized, previous: prior?.current, savedAt: new Date().toISOString() });
        return true;
      }
      const created = []; let committed = false;
      try {
        let previousFile;
        if (prior?.version === 2) {
          for (const name of [prior.currentFile, prior.previousFile]) {
            if (await readSnapshot(doc.id, name).then(() => true).catch(() => false)) { previousFile = name; break; }
          }
        } else if (typeof prior?.current === 'string' && Buffer.byteLength(prior.current) <= MAX) {
          previousFile = await writeSnapshot(doc.id, prior.current); created.push(previousFile);
        }
        const currentFile = await writeSnapshot(doc.id, serialized); created.push(currentFile);
        await atomicRecord(filename, { version: 2, currentFile, previousFile, savedAt: new Date().toISOString() }); committed = true;
        // Only obsolete snapshots belonging to this document can be removed.
        if (prior?.version === 2) for (const name of [prior.currentFile, prior.previousFile]) {
          if (snapshotName(doc.id, name) && name !== currentFile && name !== previousFile) await fs.rm(path.join(directory, name), { force: true }).catch(() => {});
        }
      } finally {
        if (!committed) for (const name of created) await fs.rm(path.join(directory, name), { force: true }).catch(() => {});
      }
      return true;
    });
    pending = task.catch(() => {});
    return task;
  }
  return { list, save };
}
module.exports = { createRecoveryStore };
