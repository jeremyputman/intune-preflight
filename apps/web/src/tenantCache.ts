import type { TenantData } from "@intune-preflight/engine";
import type { DeviceMembership } from "@intune-preflight/shared";

const DATABASE = "intune-preflight";
const VERSION = 2;
const STORE = "tenant-snapshots";

export interface CacheEntry {
  revision: string;
  data?: TenantData;
  reads?: SnapshotRead[];
  retryIncomplete?: boolean;
  discoveredGroups?: DeviceMembership["groups"];
}

export interface SnapshotRead {
  path: string;
  beta: boolean;
  collection: boolean;
  fingerprint: string;
}

export interface TenantCache {
  read(key: string): Promise<CacheEntry>;
  write(
    key: string,
    revision: string,
    data: TenantData,
    reads: SnapshotRead[],
    retryIncomplete: boolean
  ): Promise<string | undefined>;
  addDiscoveredGroups(key: string, revision: string, groups: DeviceMembership["groups"]): Promise<void>;
  clear(key: string, expectedRevision?: string): Promise<string | undefined>;
}

export function cachedGroupIds(entry: CacheEntry): Set<string> {
  return new Set(
    [entry.data?.groups, entry.data?.profileRuleGroups, entry.discoveredGroups].flatMap(
      (groups) => groups?.map(({ id }) => id.toLowerCase()) ?? []
    )
  );
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
  write: (key, revision, data, reads, retryIncomplete) =>
    transact("readwrite", (store, finish) => {
      const request = store.get(key);
      request.onsuccess = () => {
        const current: CacheEntry | undefined = request.result;
        if ((current?.revision ?? "") !== revision) {
          finish(undefined);
          return;
        }
        const nextRevision = crypto.randomUUID();
        store.put({ revision: nextRevision, data, reads, retryIncomplete } satisfies CacheEntry, key);
        finish(nextRevision);
      };
    }),
  addDiscoveredGroups: (key, revision, groups) =>
    transact("readwrite", (store, finish) => {
      const request = store.get(key);
      request.onsuccess = () => {
        const current: CacheEntry | undefined = request.result;
        if (current?.data && current.revision === revision) {
          const known = cachedGroupIds(current);
          const extras = groups.filter(({ id }) => {
            const key = id.toLowerCase();
            if (known.has(key)) return false;
            known.add(key);
            return true;
          });
          if (extras.length) {
            // Keep the snapshot's revision and timestamp. The write records
            // group metadata independently of any device or assignment.
            store.put({ ...current, discoveredGroups: [...(current.discoveredGroups ?? []), ...extras] }, key);
          }
        }
        finish(undefined);
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
