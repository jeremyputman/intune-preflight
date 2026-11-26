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
import { CacheAccessDenied, type CacheEntry, type TenantCache } from "./tenantCache.ts";

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

/** How often the progress counter is republished during a load. */
const PROGRESS_TICK_MS = 400;
const CACHE_AUTHORIZATION_MS = 5 * 60_000;

export interface PersistentTenantCache {
  storage: TenantCache;
  key: () => string | null;
  authorize: (hasApps: boolean) => Promise<boolean>;
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

  async function purge(key: string) {
    if (!persistence) return;
    try {
      await persistence.storage.clear(key);
    } finally {
      persistence.notifyPurge(key);
    }
  }

  async function readCache(key: string): Promise<CacheEntry | undefined> {
    try {
      return await persistence?.storage.read(key);
    } catch {
      cacheMessage = "Browser storage is unavailable. Tenant data can still be loaded for this tab.";
      return undefined;
    }
  }

  async function authorizeCache(key: string, hasApps: boolean, hasSnapshot: boolean): Promise<boolean> {
    try {
      if (await persistence?.authorize(hasApps)) return true;
    } catch (error) {
      if (!(error instanceof CacheAccessDenied)) throw error;
      if (hasSnapshot) await purge(key);
      throw error;
    }
    if (hasSnapshot) await purge(key);
    cacheMessage = "Local storage requires a verified tenant-wide Intune Administrator role.";
    return false;
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
      // A demo load is synchronous bundled data -- no counter, no spinner.
      const ticker =
        active === "connected"
          ? setInterval(() => {
              setLoadProgress({ phase: "loading", requests: client.getGraphStats().requests });
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
        let data: TenantData;
        if (active === "connected" && persistence && !refresh) {
          if (entry?.data && key && (await authorizeCache(key, entry.data.apps !== undefined, true))) {
            const current = await readCache(key);
            data = current?.revision === entry.revision && current.data ? entry.data : EMPTY_TENANT;
            cachedUntil = data === EMPTY_TENANT ? undefined : Date.now() + CACHE_AUTHORIZATION_MS;
          } else {
            data = EMPTY_TENANT;
          }
        } else {
          if (refresh) cache.clear();
          data = await loadTenantData({
            client,
            mode: active,
            cache,
            reportGroupFailures: active === "connected" && Boolean(persistence),
          });
          stillCurrent();
          const complete = !data.warnings?.length && data.appsUnavailable?.reason !== "error";
          const allowed =
            key && entry && (complete || entry.data)
              ? await authorizeCache(
                  key,
                  data.apps !== undefined || entry.data?.apps !== undefined,
                  Boolean(entry.data)
                )
              : false;
          stillCurrent();
          if (active === "connected" && persistence && !complete && (entry?.data ? allowed : loaded?.loadedAt)) {
            throw new Error("The refresh was incomplete. The last complete tenant snapshot was kept.");
          }
          if (key && entry && persistence) {
            if (complete && allowed) {
              if (key !== persistence.key()) throw new Error("The signed-in account changed.");
              try {
                if (!(await persistence.storage.write(key, entry.revision, data))) {
                  throw new CacheAccessDenied("The local cache was cleared during the refresh.");
                }
                cacheMessage = undefined;
                cachedUntil = Date.now() + CACHE_AUTHORIZATION_MS;
              } catch (error) {
                if (error instanceof CacheAccessDenied) throw error;
                cacheMessage = "The browser could not save the tenant snapshot. This tab can still use it.";
                cachedUntil = undefined;
              }
            } else {
              cachedUntil = undefined;
            }
          }
        }
        stillCurrent();
        if (key && key !== persistence?.key()) throw new CacheAccessDenied("The signed-in account changed.");
        loaded = data;
        loadedKey = key;
        warnings = active === "connected" ? (data.warnings ?? []) : [];
        setLoadProgress({ phase: "ready", requests: client.getGraphStats().requests });
        return data;
      } catch (err) {
        if (generation !== requestGeneration) throw err;
        if (persistence && (!refresh || err instanceof CacheAccessDenied)) {
          loaded = refresh ? null : EMPTY_TENANT;
          loadedKey = undefined;
          cachedUntil = undefined;
          warnings = [];
          cacheMessage = (err as Error).message;
        }
        setLoadProgress({
          phase: "failed",
          requests: client.getGraphStats().requests,
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
    deviceMembership: async (id: string): Promise<DeviceMembershipResponse> =>
      (await ensureMode()) === "demo" ? { unavailable: DEVICE_LOOKUP_DEMO } : readDeviceMembership(client, id),
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
      else if (inFlight) await inFlight.catch(() => {});
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
}

/**
 * Records that this tab wants Connected, for a sign-in started outside the mode
 * toggle -- the first-run form's "Save and sign in", which redirects to Entra
 * before `setMode` would ever run.
 */
export function rememberConnectedMode(): void {
  writeSession(MODE_KEY, "connected");
}
