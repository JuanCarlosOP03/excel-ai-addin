// Chat history and undo journal are stored per workbook in IndexedDB (localStorage is too
// small for undo snapshots). The workbook is identified by an id saved in the document's
// add-in settings, so it survives renames and moves of the file.

const DB_NAME = 'excel-ai';
const STORE = 'kv';
const WORKBOOK_ID_SETTING = 'excelAiWorkbookId';

let dbPromise: Promise<IDBDatabase | null> | null = null;
/** Fallback when IndexedDB is unavailable (private mode, blocked storage). */
const memory = new Map<string, unknown>();

const openDb = () => {
  dbPromise ??= new Promise(resolve => {
    try {
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => request.result.createObjectStore(STORE);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => resolve(null);
      request.onblocked = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
  return dbPromise;
};

const run = async <T>(mode: IDBTransactionMode, action: (store: IDBObjectStore) => IDBRequest<T>): Promise<T | undefined> => {
  const db = await openDb();
  if (!db) return undefined;
  return new Promise(resolve => {
    try {
      const request = action(db.transaction(STORE, mode).objectStore(STORE));
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => resolve(undefined);
    } catch {
      resolve(undefined);
    }
  });
};

export const kvGet = async <T>(key: string): Promise<T | undefined> => {
  const value = await run('readonly', store => store.get(key) as IDBRequest<T>);
  return value ?? (memory.get(key) as T | undefined);
};

export const kvSet = async (key: string, value: unknown): Promise<void> => {
  memory.set(key, value);
  await run('readwrite', store => store.put(value, key));
};

export const kvDelete = async (key: string): Promise<void> => {
  memory.delete(key);
  await run('readwrite', store => store.delete(key));
};

let workbookIdPromise: Promise<string> | null = null;

/** Stable id of the open workbook (created and saved in the document on first use). */
export const getWorkbookId = () => {
  workbookIdPromise ??= new Promise(resolve => {
    try {
      const settings = Office.context?.document?.settings;
      if (!settings) return resolve('default');
      const existing = settings.get(WORKBOOK_ID_SETTING);
      if (typeof existing === 'string' && existing) return resolve(existing);
      const id = crypto.randomUUID();
      settings.set(WORKBOOK_ID_SETTING, id);
      // The id is written into the file the next time the user saves it.
      settings.saveAsync(() => resolve(id));
    } catch {
      resolve('default');
    }
  });
  return workbookIdPromise;
};

/** Calls `save` at most once per `delay` ms with the latest value. */
export const debounce = <T>(save: (value: T) => void, delay = 500) => {
  let timer: ReturnType<typeof setTimeout> | undefined;
  let latest: T;
  return (value: T) => {
    latest = value;
    clearTimeout(timer);
    timer = setTimeout(() => save(latest), delay);
  };
};
