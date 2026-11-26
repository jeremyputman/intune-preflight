import { GraphHttpError, isGraphStatus, type GraphClient } from "@intune-preflight/graph";
import { CacheAccessDenied, type SnapshotRead } from "./tenantCache.ts";

// Hash the response rather than relying on HTTP 200: Intune can return a
// smaller collection when scope tags or delegated permissions change.
function canonical(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value)
        .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
        .map(([key, item]) => [key, canonical(item)])
    );
  }
  return value;
}

async function fingerprint(value: unknown, collection: boolean): Promise<string> {
  const canonicalValue =
    collection && Array.isArray(value) ? value.map((item) => JSON.stringify(canonical(item))).sort() : canonical(value);
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(JSON.stringify(canonicalValue)));
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

interface Observation {
  read: SnapshotRead;
  value?: unknown;
  error?: unknown;
}

export function createSnapshotReader(client: GraphClient) {
  const observations = new Map<string, Promise<Observation>>();
  const used = new Set<Promise<Observation>>();
  let authenticationFailed = false;
  const readKey = (path: string, beta: boolean, collection: boolean) => JSON.stringify([path, beta, collection]);

  function observe(path: string, beta: boolean, collection: boolean): Promise<Observation> {
    const key = readKey(path, beta, collection);
    const existing = observations.get(key);
    if (existing) return existing;
    const pending = (async (): Promise<Observation> => {
      try {
        const value = collection
          ? await client.graphGetCollection(path, beta)
          : await client.graphGetObject(path, beta);
        return { read: { path, beta, collection, fingerprint: await fingerprint(value, collection) }, value };
      } catch (error) {
        if (error instanceof CacheAccessDenied || isGraphStatus(error, 401)) {
          authenticationFailed = true;
          throw new CacheAccessDenied(
            "Microsoft Graph could not authenticate this session. Sign in again to continue."
          );
        }
        // A denied category can be normal for a scoped role. Record its
        // outcome so a changed permission is detected on a later sign-in.
        if (error instanceof GraphHttpError && [400, 403, 404, 405, 501].includes(error.status)) {
          return { read: { path, beta, collection, fingerprint: String(error.status) }, error };
        }
        throw error;
      }
    })();
    observations.set(key, pending);
    return pending;
  }

  async function read<T>(path: string, beta: boolean, collection: boolean): Promise<T> {
    const pending = observe(path, beta, collection);
    used.add(pending);
    const result = await pending;
    if (result.error) throw result.error;
    return result.value as T;
  }

  return {
    // Reuse successful verification reads if a mismatched snapshot needs a
    // fresh load, so those queries aren't issued twice.
    client: {
      ...client,
      graphGetObject: <T>(path: string, beta = false) => read<T>(path, beta, false),
      graphGetCollection: <T>(path: string, beta = false) => read<T[]>(path, beta, true),
    } satisfies GraphClient,
    async matches(previous: SnapshotRead[]): Promise<boolean> {
      const results = await client.mapWithConcurrency(previous, client.concurrency, (check) =>
        observe(check.path, check.beta, check.collection)
      );
      return results.every((result, index) => result.read.fingerprint === previous[index].fingerprint);
    },
    async observedChanges(previous: SnapshotRead[]): Promise<boolean> {
      const results = await Promise.all(
        previous.map(async (check) => {
          const result = await observations.get(readKey(check.path, check.beta, check.collection))?.catch(() => null);
          return result && result.read.fingerprint !== check.fingerprint;
        })
      );
      return results.some(Boolean);
    },
    async manifest(): Promise<{ reads: SnapshotRead[]; incomplete: boolean }> {
      const results = await Promise.allSettled(used);
      if (authenticationFailed) throw new CacheAccessDenied("Sign in again to load tenant data from Microsoft Graph.");
      return {
        reads: results.flatMap((result) => (result.status === "fulfilled" ? [result.value.read] : [])),
        incomplete: results.some((result) => result.status === "rejected"),
      };
    },
  };
}
