import { GraphHttpError, isGraphStatus, type GraphClient } from "@intune-preflight/graph";
import { CacheAccessDenied, type SnapshotRead } from "./tenantCache.ts";

// Pick at most one collection for each API area that returned data in this
// account's last load. Scoped roles may legitimately lack a different area.
const PROBE_AREAS = [
  ["/groups"],
  [
    "/deviceManagement/deviceConfigurations",
    "/deviceManagement/configurationPolicies",
    "/deviceManagement/deviceCompliancePolicies",
    "/deviceManagement/groupPolicyConfigurations",
    "/deviceManagement/assignmentFilters",
    "/deviceManagement/intents",
    "/deviceManagement/templates",
    "/deviceManagement/groupPolicyCategories",
  ],
  ["/deviceManagement/windowsAutopilotDeploymentProfiles"],
  ["/deviceManagement/deviceManagementScripts", "/deviceManagement/deviceShellScripts"],
  ["/deviceAppManagement/mobileApps"],
];

export async function verifySnapshotAccess(client: GraphClient, reads: SnapshotRead[]): Promise<void> {
  const successful = new Map(
    reads
      .filter((read) => read.collection && read.fingerprint.length === 64)
      .map((read) => [read.path.split("?", 1)[0], read.beta])
  );
  const probes = PROBE_AREAS.flatMap((paths) => {
    const path = paths.find((source) => successful.has(source));
    return path ? [{ path, beta: successful.get(path) }] : [];
  });
  if (!probes.length) throw new CacheAccessDenied("No saved Graph access could be verified. Refresh from Intune.");

  const results = await Promise.allSettled(
    probes.map(async ({ path, beta }) => {
      // Only the first page is needed to check the endpoint's current access.
      const page = await client.graphGetObject<{ value?: unknown[] }>(`${path}?$select=id&$top=1`, beta);
      if (!Array.isArray(page?.value)) throw new Error("Graph did not confirm access to the saved tenant data.");
    })
  );
  if (results.some((result) => result.status === "rejected" && isGraphStatus(result.reason, 401, 403))) {
    throw new CacheAccessDenied("Microsoft Graph denied access to a saved data area. Refresh from Intune.");
  }
  const failed = results.find((result) => result.status === "rejected");
  if (failed) throw failed.reason;
}

// On a full refresh, compare responses before keeping or replacing older data.
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
    // Deduplicate identical queries that occur during a full load.
    client: {
      ...client,
      graphGetObject: <T>(path: string, beta = false) => read<T>(path, beta, false),
      graphGetCollection: <T>(path: string, beta = false) => read<T[]>(path, beta, true),
    } satisfies GraphClient,
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
