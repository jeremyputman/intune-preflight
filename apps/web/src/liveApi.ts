/**
 * The third implementation of the API surface: the whole app in a browser tab,
 * reading the signed-in admin's own tenant from Microsoft Graph with no backend
 * at all.
 *
 * It is `demoApi` and `serverApi` joined at the seam `packages/graph` was
 * extracted to create -- the same fetching, paging, retry and v1.0 -> beta
 * fallback the server does (`createGraphClient({ getToken })`), feeding the same
 * pure engine the public demo runs (`computeSimulation`, `buildAssignmentReport`,
 * ...). Nothing here reimplements either side; the only new thing is that the
 * token is delegated and the tenant lives in the tab instead of in a server
 * process.
 *
 * Auth arrives as injected callbacks rather than an import of `./auth.ts`, for
 * the same reason `packages/graph` takes a `getToken`: it keeps the MSAL runtime
 * out of this module, so the shape conformance that `api.ts` type-checks can
 * also be exercised by a test that has no browser.
 */

import type { Platform } from "@intune-preflight/shared";
import {
  assignmentReportToCsv,
  assignmentReportToXlsx,
  buildAssignmentReport,
  computeSimulation,
  listAssignmentFilters,
  listGroupSummaries,
  listProfileNameWarnings,
  listUnassignedPolicies,
  narrowSimulation,
  simulationAppsToCsv,
  simulationAppsToXlsx,
  simulationToCsv,
  simulationToXlsx,
  type TenantData,
  XLSX_MIME,
  type BaselineCriteria,
} from "@intune-preflight/engine";
import {
  createGraphClient,
  loadTenantData,
  readDeviceGroup,
  readDeviceMembership,
  searchDevices,
  type TenantMode,
  TtlCache,
} from "@intune-preflight/graph";
import { DEVICE_LOOKUP_DEMO, type DeviceMembershipResponse, type DeviceSearchResponse } from "@intune-preflight/shared";
import type {
  SimulationExportFormat,
  ManifestExportFormat,
  RefreshResult,
  ServerStatus,
  SimulationInputs,
} from "./api.ts";
import type { Download } from "./download.ts";
import { setLoadProgress } from "./loadProgress.ts";
import { readSession, removeSession, writeSession } from "./storage.ts";
import { CacheAccessDenied, cachedGroupIds, type CacheEntry, type TenantCache } from "./tenantCache.ts";
import { createSnapshotReader, verifySnapshotAccess } from "./snapshotAccess.ts";

/**
 * Graph detail calls in flight at once. Lower than the server's eight: a browser
 * shares its connection pool with the page's own requests, and an admin watching
 * a progress counter would rather the UI stayed responsive than shave two
 * seconds off a twelve-second load.
 */
const BROWSER_CONCURRENCY = 6;

/**
 * The cache TTL. `packages/graph` leaves the lifetime to the caller because a
 * server's is its configured TTL while a browser session's is the TAB -- so this
 * is set past any plausible sitting, and the tenant is dropped by `refresh()`, a
 * mode switch, ending the session, or closing the tab. Re-reading 15MB on a
 * timer nobody asked for would be the wrong default here.
 */
const TAB_LIFETIME_SECONDS = 24 * 60 * 60;

/**
 * Which mode the tab was in, so a sign-in redirect comes back to Connected
 * instead of dropping the admin into Demo after they just signed in. Session-
 * scoped like the token itself.
 */
const MODE_KEY = "preflight.live.mode";
const SNAPSHOT_ACCESS_KEY = "preflight.live.snapshotAccess";

/** How often the progress counter is republished during a load. */
const PROGRESS_TICK_MS = 400;
const CACHE_AUTHORIZATION_MS = 15 * 60_000;
const INCOMPLETE_SNAPSHOT_NOTICE =
  "Some Graph reads failed when this snapshot was saved. Use Refresh from Intune to retry.";

function readSnapshotAccess(key: string, revision: string): number | undefined {
  const saved = readSession(SNAPSHOT_ACCESS_KEY);
  if (!saved) return;
  try {
    const access: { key: string; revision: string; expiresAt: number } | null = JSON.parse(saved);
    const now = Date.now();
    if (
      access?.key === key &&
      access.revision === revision &&
      Number.isFinite(access.expiresAt) &&
      access.expiresAt > now &&
      access.expiresAt <= now + CACHE_AUTHORIZATION_MS
    ) {
      return access.expiresAt;
    }
  } catch {
    // A missing or invalid tab record requires an online check.
  }
}

function rememberSnapshotAccess(key: string, revision: string): number {
  const expiresAt = Date.now() + CACHE_AUTHORIZATION_MS;
  writeSession(SNAPSHOT_ACCESS_KEY, JSON.stringify({ key, revision, expiresAt }));
  return expiresAt;
}

export interface PersistentTenantCache {
  storage: TenantCache;
  key: () => string | null;
  renewToken: () => Promise<string>;
  notifyPurge: (key: string) => void;
}

export interface LiveApiDeps {
  /**
   * Brings auth to a usable state: builds MSAL and consumes a sign-in redirect
   * if this load is the tail of one. Idempotent, and a no-op when no tenant is
   * configured -- an unconfigured live build is still a working demo.
   *
   * Every entry point below awaits it, which is what lets `api.ts` do the
   * wiring and `App.tsx` stay unaware that the live build has an extra step.
   */
  ensureReady: () => Promise<void>;
  /** A delegated Graph token for the signed-in admin; called before every attempt. */
  getToken: () => Promise<string>;
  /** Signs in if nobody is. May navigate the page away, in which case it never resolves. */
  ensureSignedIn: () => Promise<void>;
  /** Whether an admin is signed in to this tab right now. Only valid after `ensureReady`. */
  isSignedIn: () => boolean;
  /** Whether a tenant ID and client ID are configured -- the live build's `hasCredentials`. */
  isConfigured: () => boolean;
  persistence?: PersistentTenantCache;
}

// Stands in for a tenant that has not loaded yet, so the synchronous export-URL
// helpers below can stay synchronous. See the comment on them.
const EMPTY_TENANT: TenantData = { policies: [], groups: [], autopilotProfiles: [], assignmentFilters: [] };

const toInputs = (i: SimulationInputs) => ({
  selectedGroupIds: i.groupIds ?? [],
  platform: i.platform,
  deviceFilterIds: i.deviceFilterIds ?? [],
  unassignedPolicyIds: i.unassignedPolicyIds ?? [],
});

// Same as demoApi: no backend, so the bytes go to `saveDownload` rather than
// onto an anchor as a `data:` URL, which Safari will not save.
const textFile = (mime: string, text: string): Download => ({ kind: "text", text, mime });
/** The xlsx export is a ZIP, so it goes to the Blob as bytes rather than text. */
const xlsxFile = (bytes: Uint8Array): Download => ({ kind: "bytes", bytes, mime: XLSX_MIME });

export function createLiveApi(deps: LiveApiDeps) {
  const client = createGraphClient({ getToken: deps.getToken, concurrency: BROWSER_CONCURRENCY });
  const cache = new TtlCache(TAB_LIFETIME_SECONDS);

  let mode: TenantMode | null = null;
  let loaded: TenantData | null = null;
  let loadedKey: string | null | undefined;
  let warnings: string[] = [];
  let inFlight: Promise<TenantData> | null = null;
  let generation = 0;
  let cachedUntil: number | undefined;
  let cacheMessage: string | undefined;
  const persistence = deps.persistence;

  async function purge(key: string, expectedRevision: string) {
    const revision = await persistence?.storage.clear(key, expectedRevision);
    if (revision) persistence?.notifyPurge(key);
    return revision;
  }

  async function readCache(key: string): Promise<CacheEntry | undefined> {
    try {
      return await persistence?.storage.read(key);
    } catch {
      cacheMessage = "Browser storage is unavailable. Tenant data can still be loaded for this tab.";
      return undefined;
    }
  }

  async function scopedTenant(
    key: string,
    entry: CacheEntry,
    refresh: boolean,
    stillCurrent: () => void,
    setRequestCount: (count: () => number) => void
  ): Promise<TenantData> {
    if (!persistence) throw new Error("Browser storage is unavailable.");
    const { reads } = entry;
    let { revision, data: previous } = entry;
    try {
      if (previous && reads && !refresh) {
        const expiresAt = readSnapshotAccess(key, revision);
        if (expiresAt !== undefined) {
          cachedUntil = expiresAt;
          cacheMessage = entry.retryIncomplete ? INCOMPLETE_SNAPSHOT_NOTICE : undefined;
          return previous;
        }
      }

      const token = previous ? await persistence.renewToken() : undefined;
      stillCurrent();
      const lookupClient = createGraphClient({
        getToken: token ? async () => token : deps.getToken,
        concurrency: BROWSER_CONCURRENCY,
        requestCache: "no-store",
      });
      setRequestCount(() => lookupClient.getGraphStats().requests);
      if (previous && reads && !refresh) {
        await verifySnapshotAccess(lookupClient, reads);
        const current = await readCache(key);
        stillCurrent();
        if (!current || current.revision !== revision) {
          throw new CacheAccessDenied("The local cache changed while it was being checked.");
        }
        cachedUntil = rememberSnapshotAccess(key, revision);
        cacheMessage = entry.retryIncomplete ? INCOMPLETE_SNAPSHOT_NOTICE : undefined;
        return previous;
      }

      const reader = createSnapshotReader(lookupClient);
      const data = await loadTenantData({
        client: reader.client,
        mode: "connected",
        cache: new TtlCache(TAB_LIFETIME_SECONDS),
        reportGroupFailures: true,
      });
      const manifest = await reader.manifest();
      stillCurrent();
      if (key !== persistence.key()) throw new CacheAccessDenied("The signed-in account changed.");
      const complete =
        !manifest.incomplete && (!data.warnings?.length || manifest.reads.some((read) => read.fingerprint === "403"));
      const hasAuthorizedReads = manifest.reads.some((read) => read.fingerprint.length === 64);
      const keepComplete = previous && !entry.retryIncomplete && !complete;
      if (previous && reads) {
        if (await reader.observedChanges(reads)) {
          stillCurrent();
          const currentRevision = await purge(key, revision);
          if (!currentRevision) throw new CacheAccessDenied("The local cache changed during the refresh.");
          revision = currentRevision;
          previous = undefined;
          if (keepComplete) {
            throw new CacheAccessDenied("Graph access or data changed. Refresh to load the current tenant data.");
          }
        } else if (keepComplete) {
          throw new Error("The refresh was incomplete. The last saved tenant snapshot was kept.");
        }
      }
      stillCurrent();
      cachedUntil = undefined;
      if (hasAuthorizedReads) {
        try {
          const savedRevision = await persistence.storage.write(key, revision, data, manifest.reads, !complete);
          if (!savedRevision) {
            throw new CacheAccessDenied("The local cache was cleared during the refresh.");
          }
          stillCurrent();
          if (key !== persistence.key()) throw new CacheAccessDenied("The signed-in account changed.");
          cacheMessage = complete ? undefined : INCOMPLETE_SNAPSHOT_NOTICE;
          cachedUntil = rememberSnapshotAccess(key, savedRevision);
        } catch (error) {
          if (error instanceof CacheAccessDenied) throw error;
          cacheMessage = "The browser could not save the tenant snapshot. The data is available in this tab.";
        }
      } else {
        cacheMessage = "This load could not be saved. The available tenant data is shown for this tab.";
      }
      return data;
    } catch (error) {
      if (previous && error instanceof CacheAccessDenied) await purge(key, revision);
      if (error instanceof CacheAccessDenied) removeSession(SNAPSHOT_ACCESS_KEY);
      throw error;
    }
  }

  /**
   * The current mode, resolved on first use because it depends on whether a
   * sign-in redirect landed on this page load -- which is not known until
   * `ensureReady` has run. A remembered Connected is only honoured if the
   * sign-in actually survived, so a tab whose session has gone comes back in
   * Demo rather than erroring on every read.
   */
  async function ensureMode(): Promise<TenantMode> {
    if (mode) return mode;
    await deps.ensureReady();
    mode = readSession(MODE_KEY) === "connected" && deps.isSignedIn() ? "connected" : "demo";
    return mode;
  }

  /**
   * The loaded tenant, fetching it once per mode. Every read below funnels
   * through here, so the three requests the UI fires on mount share one load
   * rather than starting three.
   */
  function tenant(refresh = false): Promise<TenantData> {
    if (loadedKey && loadedKey !== persistence?.key()) drop();
    if (!refresh && loaded && (cachedUntil === undefined || Date.now() < cachedUntil)) {
      return Promise.resolve(loaded);
    }
    if (inFlight) return inFlight;

    const requestGeneration = generation;
    const stillCurrent = () => {
      if (generation !== requestGeneration) throw new Error("The tenant session changed.");
    };
    inFlight = (async () => {
      const active = await ensureMode();
      let requestCount = () => client.getGraphStats().requests;
      // A demo load is synchronous bundled data -- no counter, no spinner.
      const ticker =
        active === "connected"
          ? setInterval(() => {
              if (generation === requestGeneration) setLoadProgress({ phase: "loading", requests: requestCount() });
            }, PROGRESS_TICK_MS)
          : null;
      if (active === "connected") setLoadProgress({ phase: "loading", requests: 0 });
      try {
        stillCurrent();
        const key = active === "connected" ? persistence?.key() : null;
        if (active === "connected" && persistence && !key) {
          cacheMessage = "Browser storage is unavailable. Tenant data can still be loaded for this tab.";
        }
        const entry = key ? await readCache(key) : undefined;
        stillCurrent();
        if (refresh) cache.clear();
        const data =
          key && entry
            ? await scopedTenant(key, entry, refresh, stillCurrent, (count) => {
                requestCount = count;
              })
            : await loadTenantData({ client, mode: active, cache });
        stillCurrent();
        if (key && key !== persistence?.key()) throw new CacheAccessDenied("The signed-in account changed.");
        loaded = data;
        loadedKey = key;
        warnings = active === "connected" ? (data.warnings ?? []) : [];
        setLoadProgress({ phase: "ready", requests: requestCount() });
        return data;
      } catch (err) {
        if (generation !== requestGeneration) throw err;
        if (persistence && (!refresh || err instanceof CacheAccessDenied)) {
          loaded = EMPTY_TENANT;
          loadedKey = undefined;
          cachedUntil = undefined;
          warnings = [];
          cacheMessage = (err as Error).message;
        }
        setLoadProgress({
          phase: "failed",
          requests: requestCount(),
          error: (err as Error).message,
        });
        throw err;
      } finally {
        if (ticker !== null) clearInterval(ticker);
        if (generation === requestGeneration) inFlight = null;
      }
    })();

    return inFlight;
  }

  /** Forgets the loaded tenant so the next read fetches it again. */
  function drop() {
    generation++;
    inFlight = null;
    cache.clear();
    loaded = null;
    loadedKey = undefined;
    warnings = [];
    cachedUntil = undefined;
    cacheMessage = undefined;
  }

  /**
   * The tenant as it stands, WITHOUT fetching. Only the two export helpers use
   * it, because they are the only synchronous members of the surface -- and
   * they have to be. `serverApi` returns from them synchronously (it just
   * builds an `/api/...` path), and on iOS `saveDownload` can only open the
   * share sheet inside the click's own task, so there is nowhere to put an
   * await. They are reached from a button in a panel that already holds a
   * loaded report, so in practice the tenant is always here; the empty fallback
   * is so a race can't throw out of a click handler.
   */
  const tenantNow = () =>
    loaded &&
    (!loadedKey || loadedKey === persistence?.key()) &&
    (cachedUntil === undefined || Date.now() < cachedUntil)
      ? loaded
      : EMPTY_TENANT;

  async function deviceMembership(id: string): Promise<DeviceMembershipResponse> {
    if ((await ensureMode()) === "demo") return { unavailable: DEVICE_LOOKUP_DEMO };
    const key = persistence?.key();
    if (!key || !persistence) return readDeviceMembership(client, id);

    const lookupGeneration = generation;
    const [result, entry] = await Promise.all([
      readDeviceMembership(client, id),
      persistence.storage.read(key).catch(() => undefined),
    ]);
    const sameSession = () => generation === lookupGeneration && persistence.key() === key;
    if (!entry?.data || !sameSession() || "unavailable" in result) return result;

    const known = cachedGroupIds(entry);
    const candidates = new Map(
      result.membership.groups
        .filter((group) => group.displayName === group.id || !known.has(group.id.toLowerCase()))
        .map((group) => [group.id.toLowerCase(), group])
    );
    if (!candidates.size) return result;

    const groups = await client.mapWithConcurrency([...candidates.values()], client.concurrency, async (group) =>
      group.displayName !== group.id ? group : readDeviceGroup(client, group.id)
    );
    if (!sameSession()) return result;
    const resolved = groups.filter((group) => group !== undefined);
    if (!resolved.length) return result;
    const discovered = resolved.filter((group) => !known.has(group.id.toLowerCase()));
    if (discovered.length) {
      await persistence.storage.addDiscoveredGroups(key, entry.revision, discovered).catch(() => {});
    }

    const byId = new Map(resolved.map((group) => [group.id.toLowerCase(), group]));
    return {
      membership: {
        ...result.membership,
        groups: result.membership.groups.map((group) => byId.get(group.id.toLowerCase()) ?? group),
      },
    };
  }

  return {
    health: async (): Promise<ServerStatus> => {
      // Await a load already in flight so the caller's warnings are the real
      // ones. App.tsx re-polls health a few seconds after mount to catch an
      // incomplete load; a browser load takes longer than that poll, and
      // without this the poll would always report "no warnings".
      const active = await ensureMode();
      if (persistence && active === "connected") await tenant().catch(() => {});
      else if (inFlight) await inFlight.catch(() => {});
      return {
        status: "ok",
        demo: active === "demo",
        hasCredentials: deps.isConfigured(),
        loadWarnings: warnings,
        loadedAt: active === "connected" ? loaded?.loadedAt : undefined,
        ...(cacheMessage ? { cacheMessage } : {}),
      };
    },

    setMode: async (target: "demo" | "connected"): Promise<Omit<ServerStatus, "status">> => {
      await ensureMode();
      if (target === "connected") {
        // Remember the intent BEFORE signing in: `ensureSignedIn` may navigate
        // to Entra, and what we want on the way back is Connected.
        if (!deps.isSignedIn()) removeSession(SNAPSHOT_ACCESS_KEY);
        writeSession(MODE_KEY, "connected");
        setLoadProgress({ phase: "signing-in", requests: 0 });
        try {
          await deps.ensureSignedIn();
        } catch (err) {
          writeSession(MODE_KEY, "demo");
          setLoadProgress({ phase: "failed", requests: 0, error: (err as Error).message });
          throw err;
        }
      } else {
        writeSession(MODE_KEY, "demo");
        setLoadProgress({ phase: "idle", requests: 0 });
      }
      mode = target;
      drop();
      return { demo: target === "demo", hasCredentials: deps.isConfigured(), loadWarnings: [] };
    },

    groups: async () => listGroupSummaries(await tenant()),
    profileNameWarnings: async () => listProfileNameWarnings(await tenant()),
    filters: async () => listAssignmentFilters(await tenant()),
    // Live reads under the signed-in admin's own token, like the tenant load.
    // Not part of the cached tenant: the point is the device as it is now.
    searchDevices: async (
      query: string,
      _signal?: AbortSignal,
      includeInactive = false
    ): Promise<DeviceSearchResponse> =>
      (await ensureMode()) === "demo"
        ? { unavailable: DEVICE_LOOKUP_DEMO }
        : searchDevices(client, query, includeInactive),
    deviceMembership,
    unassignedPolicies: async () => listUnassignedPolicies(await tenant()),
    assignmentReport: async (platform?: Platform) => buildAssignmentReport(await tenant(), platform),
    assignmentReportExport: (platform?: Platform, format: ManifestExportFormat = "csv") => {
      const report = buildAssignmentReport(tenantNow(), platform);
      return format === "xlsx"
        ? xlsxFile(assignmentReportToXlsx(report))
        : textFile("text/csv", assignmentReportToCsv(report));
    },

    refresh: async (): Promise<RefreshResult> => {
      const active = await ensureMode();
      if (!persistence || active === "demo") drop();
      const data = await tenant(true);
      return { warnings, loadedAt: data.loadedAt, ...(cacheMessage ? { cacheMessage } : {}) };
    },

    simulate: async (inputs: SimulationInputs) => computeSimulation(await tenant(), toInputs(inputs)),
    simulateExport: (inputs: SimulationInputs, format: SimulationExportFormat, criteria?: BaselineCriteria) => {
      // Narrowed here rather than on the wire: this build holds the tenant in
      // the tab, so "this view" is already in hand. Same function the panel and
      // the server use, so the three agree about what a filtered export is.
      const sim = narrowSimulation(computeSimulation(tenantNow(), toInputs(inputs)), criteria);
      if (format === "xlsx") return xlsxFile(simulationToXlsx(sim));
      if (format === "apps-xlsx") return xlsxFile(simulationAppsToXlsx(sim) ?? new Uint8Array());
      if (format === "apps-csv") return textFile("text/csv", simulationAppsToCsv(sim) ?? "");
      return format === "csv"
        ? textFile("text/csv", simulationToCsv(sim))
        : textFile("application/json", JSON.stringify(sim, null, 2));
    },
  };
}

/**
 * Forgets the remembered mode, so a tab whose session has been ended comes back
 * in Demo rather than trying to reuse a sign-in that is gone.
 */
export function forgetLiveMode(): void {
  removeSession(MODE_KEY);
  removeSession(SNAPSHOT_ACCESS_KEY);
}

/**
 * Records that this tab wants Connected, for a sign-in started outside the mode
 * toggle -- the first-run form's "Save and sign in", which redirects to Entra
 * before `setMode` would ever run.
 */
export function rememberConnectedMode(): void {
  removeSession(SNAPSHOT_ACCESS_KEY);
  writeSession(MODE_KEY, "connected");
}
