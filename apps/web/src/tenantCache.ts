import type { TenantData } from "@intune-preflight/engine";

const DATABASE = "intune-preflight";
const VERSION = 1;
const STORE = "tenant-snapshots";

export interface CacheEntry {
  revision: string;
  data?: TenantData;
}

export interface TenantCache {
  read(key: string): Promise<CacheEntry>;
  write(key: string, revision: string, data: TenantData): Promise<boolean>;
  clear(key: string): Promise<void>;
}

function openCache(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DATABASE, VERSION);
    let blocked = false;
    request.onupgradeneeded = () => request.result.createObjectStore(STORE);
    request.onerror = () => reject(request.error);
    request.onblocked = () => {
      blocked = true;
      reject(new Error("The browser cache is blocked by another tab."));
    };
    request.onsuccess = () => {
      if (blocked) request.result.close();
      else resolve(request.result);
    };
  });
}

async function transact<T>(
  mode: IDBTransactionMode,
  run: (store: IDBObjectStore, finish: (result: T) => void) => void
): Promise<T> {
  const database = await openCache();
  try {
    return await new Promise<T>((resolve, reject) => {
      const transaction = database.transaction(STORE, mode);
      let result: T;
      transaction.oncomplete = () => resolve(result);
      transaction.onabort = () =>
        reject(transaction.error ?? new Error("The browser cache operation was interrupted."));
      run(transaction.objectStore(STORE), (value) => {
        result = value;
      });
    });
  } finally {
    database.close();
  }
}

export const tenantCache: TenantCache = {
  read: (key) =>
    transact("readonly", (store, finish) => {
      const request = store.get(key);
      request.onsuccess = () => {
        const entry: CacheEntry | undefined = request.result;
        if (typeof entry?.revision !== "string") return finish({ revision: "" });
        const data = entry.data;
        if (
          !data?.loadedAt ||
          !Number.isFinite(Date.parse(data.loadedAt)) ||
          !Array.isArray(data.policies) ||
          !Array.isArray(data.groups) ||
          !Array.isArray(data.autopilotProfiles) ||
          !Array.isArray(data.assignmentFilters)
        ) {
          return finish({ revision: entry.revision });
        }
        finish(entry);
      };
    }),
  write: (key, revision, data) =>
    transact("readwrite", (store, finish) => {
      const request = store.get(key);
      request.onsuccess = () => {
        const current: CacheEntry | undefined = request.result;
        if ((current?.revision ?? "") !== revision) {
          finish(false);
          return;
        }
        store.put({ revision, data } satisfies CacheEntry, key);
        finish(true);
      };
    }),
  // A tombstone blocks writes that started before a clear, including from
  // another tab. The transaction removes the previous snapshot in full.
  clear: (key) =>
    transact("readwrite", (store, finish) => {
      store.put({ revision: crypto.randomUUID() } satisfies CacheEntry, key);
      finish(undefined);
    }),
};

export class CacheAccessDenied extends Error {}
