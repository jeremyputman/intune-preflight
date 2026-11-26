import type { TenantData } from "@intune-preflight/engine";

const DATABASE = "intune-preflight";
const VERSION = 2;
const STORE = "tenant-snapshots";

export interface CacheEntry {
  revision: string;
  data?: TenantData;
  reads?: SnapshotRead[];
}

export interface SnapshotRead {
  path: string;
  beta: boolean;
  collection: boolean;
  fingerprint: string;
}

export interface TenantCache {
  read(key: string): Promise<CacheEntry>;
  write(key: string, revision: string, data: TenantData, reads: SnapshotRead[]): Promise<string | undefined>;
  clear(key: string, expectedRevision?: string): Promise<string | undefined>;
}

function openCache(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DATABASE, VERSION);
    let blocked = false;
    request.onupgradeneeded = ({ oldVersion }) => {
      if (oldVersion === 0) request.result.createObjectStore(STORE);
      else request.transaction?.objectStore(STORE).clear();
    };
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
          !Array.isArray(data.assignmentFilters) ||
          !Array.isArray(entry.reads) ||
          !entry.reads.length ||
          !entry.reads.every(
            (read) =>
              typeof read?.path === "string" &&
              /^\/(?:deviceManagement\/|deviceAppManagement\/|groups[/?])[^#\r\n]*$/.test(read.path) &&
              typeof read.beta === "boolean" &&
              typeof read.collection === "boolean" &&
              /^(?:[\da-f]{64}|400|403|404|405|501)$/.test(read.fingerprint)
          )
        ) {
          return finish({ revision: entry.revision });
        }
        finish(entry);
      };
    }),
  write: (key, revision, data, reads) =>
    transact("readwrite", (store, finish) => {
      const request = store.get(key);
      request.onsuccess = () => {
        const current: CacheEntry | undefined = request.result;
        if ((current?.revision ?? "") !== revision) {
          finish(undefined);
          return;
        }
        const nextRevision = crypto.randomUUID();
        store.put({ revision: nextRevision, data, reads } satisfies CacheEntry, key);
        finish(nextRevision);
      };
    }),
  // A tombstone blocks writes that started before a clear, including from
  // another tab. The transaction removes the previous snapshot in full.
  clear: (key, expectedRevision) =>
    transact("readwrite", (store, finish) => {
      const request = store.get(key);
      request.onsuccess = () => {
        const current: CacheEntry | undefined = request.result;
        if (expectedRevision !== undefined && (current?.revision ?? "") !== expectedRevision) {
          finish(undefined);
          return;
        }
        const revision = crypto.randomUUID();
        store.put({ revision } satisfies CacheEntry, key);
        finish(revision);
      };
    }),
};

export class CacheAccessDenied extends Error {}
