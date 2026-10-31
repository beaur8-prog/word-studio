import type { StudioDocument } from './types';
import { validateDocument } from './document-io';
type Record = { current: string; previous?: string; savedAt: string };
let history: StudioDocument[] = [];
let previousVersions: StudioDocument[] = [];
let db: IDBDatabase | undefined;
let queue = Promise.resolve();
export let recoveryNotice = '';
export const readHistory = () => history;
export const readPreviousVersions = () => previousVersions;
function parse(raw?: string) {
  if (!raw) return undefined;
  try {
    const doc = validateDocument(JSON.parse(raw));
    return { ...doc, id: doc.id || crypto.randomUUID() };
  } catch { return undefined; }
}
async function database() {
  if (db) return db;
  db = await new Promise<IDBDatabase>((resolve, reject) => {
    const req = indexedDB.open('word-studio-recovery', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('documents');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return db;
}
async function browserRecords(): Promise<Record[]> {
  const databaseHandle = await database();
  return new Promise((resolve, reject) => {
    const request = databaseHandle.transaction('documents').objectStore('documents').getAll();
    request.onsuccess = () => resolve(request.result.sort((a: Record, b: Record) => b.savedAt.localeCompare(a.savedAt)));
    request.onerror = () => reject(request.error);
  });
}
export async function initializeRecovery() {
  let records: Record[] = [];
  try { records = window.wordDesktop ? await window.wordDesktop.readRecovery() : await browserRecords(); }
  catch { recoveryNotice = '복구 저장소를 읽지 못했습니다. 파일로 저장해 주세요.'; }
  for (const record of records) {
    const current = parse(record.current), previous = parse(record.previous);
    if (previous) previousVersions.push(previous);
    if (current) history.push(current);
    else if (previous) {
      history.push(previous);
      recoveryNotice = '최근 복구본을 읽지 못했습니다. 파일 메뉴의 최근 작업한 문서에서 이전 정상 버전을 열 수 있습니다.';
    }
  }
  // One-time migration keeps old local documents available without changing their content.
  try {
    const legacy = localStorage.getItem('word-studio-recovery-migrated') === '1' ? [] : [localStorage.getItem('word-studio-current'),
      ...JSON.parse(localStorage.getItem('word-studio-history') || '[]').map((d: StudioDocument) => JSON.stringify(d))];
    const imported: StudioDocument[] = [];
    for (const raw of legacy) {
      const doc = parse(raw);
      if (doc && !history.some(d => d.id === doc.id || (d.title === doc.title && d.updatedAt === doc.updatedAt))) {
        if (!history.length) history.unshift(doc); else history.push(doc);
        imported.push(doc);
      }
    }
    const preferred = history[0];
    for (const doc of imported.reverse()) {
      // Remove the memory-only migration entry so saveRecovery persists it.
      history = history.filter(d => d.id !== doc.id);
      await saveRecovery(doc);
    }
    if (preferred) history = [preferred, ...history.filter(d => d.id !== preferred.id)];
    if (legacy.length) localStorage.setItem('word-studio-recovery-migrated', '1');
  } catch {}
  history = history.slice(0, 20);
}
export function saveRecovery(doc: StudioDocument): Promise<void> {
  const snapshot = JSON.stringify(validateDocument(doc));
  const task = queue.then(async () => {
    const oldDocument = history.find(d => d.id === doc.id);
    if (oldDocument) {
      const { updatedAt: oldTime, ...oldContent } = oldDocument;
      const { updatedAt: newTime, ...newContent } = doc;
      if (JSON.stringify(oldContent) === JSON.stringify(newContent)) return;
    }
    if (window.wordDesktop) await window.wordDesktop.saveRecovery(snapshot);
    else {
      const databaseHandle = await database();
      await new Promise<void>((resolve, reject) => {
        const tx = databaseHandle.transaction('documents', 'readwrite');
        const store = tx.objectStore('documents');
        const req = store.get(doc.id!);
        req.onsuccess = () => store.put({ current: snapshot, previous: req.result?.current, savedAt: new Date().toISOString() }, doc.id!);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
        tx.onabort = () => reject(tx.error || new Error('복구 저장이 취소됐습니다.'));
      });
    }
    const old = history.find(d => d.id === doc.id);
    if (old) previousVersions = [old, ...previousVersions.filter(d => d.id !== doc.id)].slice(0, 20);
    history = [JSON.parse(snapshot), ...history.filter(d => d.id !== doc.id)].slice(0, 20);
  });
  queue = task.catch(() => {});
  return task;
}
