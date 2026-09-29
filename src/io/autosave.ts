// Crash-recovery autosave in IndexedDB.
import type { SerializedWorkbook } from '../model/workbook';

const DB = 'myexcel';
const STORE = 'autosave';
const KEY = 'current';

export interface AutosaveRecord {
  fileName: string;
  savedAt: number;
  data: SerializedWorkbook;
}

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function writeAutosave(rec: AutosaveRecord): Promise<void> {
  if (typeof indexedDB === 'undefined') return;
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put(rec, KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

export async function readAutosave(): Promise<AutosaveRecord | null> {
  if (typeof indexedDB === 'undefined') return null;
  try {
    const db = await openDb();
    const rec = await new Promise<AutosaveRecord | null>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).get(KEY);
      req.onsuccess = () => resolve((req.result as AutosaveRecord) ?? null);
      req.onerror = () => reject(req.error);
    });
    db.close();
    return rec;
  } catch {
    return null;
  }
}

export async function clearAutosave(): Promise<void> {
  if (typeof indexedDB === 'undefined') return;
  try {
    const db = await openDb();
    await new Promise<void>((resolve) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    });
    db.close();
  } catch {
    /* ignore */
  }
}
