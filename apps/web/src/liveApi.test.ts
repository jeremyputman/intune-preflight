/**
 * `api.ts` already type-checks `liveApi` against the server surface -- that
 * annotation is what has kept `demoApi` from drifting and now does the same job
 * here. These tests cover the two things it cannot: that the object really has
 * every method at runtime (a `typeof` check cannot catch a method lost to a bad
 * merge in a way the compiler notices only where it is called), and that the
 * mode/sign-in sequencing behaves.
 *
 * Everything here runs over the BUNDLED demo tenant. The connected path is the
 * same `loadTenantData` the server's own tests cover; what is new in this file
 * is the wiring around it, so that is what is exercised.
 */

import { afterEach, describe, expect, it, vi } from "vitest";
import { demoApi } from "./demoApi.ts";
import { createLiveApi, type LiveApiDeps, type PersistentTenantCache } from "./liveApi.ts";
import { CacheAccessDenied, type CacheEntry } from "./tenantCache.ts";

function deps(overrides: Partial<LiveApiDeps> = {}): LiveApiDeps {
  return {
    ensureReady: vi.fn(async () => {}),
    getToken: vi.fn(async () => "token"),
    ensureSignedIn: vi.fn(async () => {}),
    isSignedIn: vi.fn(() => false),
    isConfigured: vi.fn(() => true),
    ...overrides,
  };
}

describe("liveApi shape", () => {
  it("has every method the other implementations have, and nothing extra", () => {
    // demoApi stands in for the surface here because it is a real object at
    // runtime, where serverApi's type is all that exists at test time.
    expect(Object.keys(createLiveApi(deps())).sort()).toEqual(Object.keys(demoApi).sort());
  });

  it("exposes them all as functions", () => {
    for (const [name, value] of Object.entries(createLiveApi(deps()))) {
      expect(typeof value, name).toBe("function");
    }
  });
});

describe("starting state", () => {
  it("starts in demo so an unsigned tab is still usable", () => {
    // The live build with nobody signed in is the public demo. Opening it
    // should show something, not an error.
    return expect(createLiveApi(deps()).health()).resolves.toMatchObject({
      status: "ok",
      demo: true,
      hasCredentials: true,
    });
  });

  it("reports hasCredentials from whether a tenant is configured", async () => {
    const api = createLiveApi(deps({ isConfigured: () => false }));
    await expect(api.health()).resolves.toMatchObject({ hasCredentials: false });
  });

  it("does not trust a remembered Connected mode without a surviving sign-in", async () => {
    // sessionStorage is unavailable here, so nothing is remembered -- which is
    // the same outcome as a tab whose session has gone: fall back to demo
    // rather than erroring on every read.
    const api = createLiveApi(deps({ isSignedIn: () => true }));
    await expect(api.health()).resolves.toMatchObject({ demo: true });
  });

  it("reads the bundled sample tenant in demo mode, with no token", async () => {
    const getToken = vi.fn(async () => "token");
    const api = createLiveApi(deps({ getToken }));
    expect(await api.groups()).toEqual(await demoApi.groups());
    expect(await api.filters()).toEqual(await demoApi.filters());
    expect(await api.unassignedPolicies()).toEqual(await demoApi.unassignedPolicies());
    // Nothing was fetched, so nothing needed a token.
    expect(getToken).not.toHaveBeenCalled();
  });
});

describe("setMode", () => {
  it("signs in before switching to connected", async () => {
    const ensureSignedIn = vi.fn(async () => {});
    const api = createLiveApi(deps({ ensureSignedIn }));
    await expect(api.setMode("connected")).resolves.toMatchObject({ demo: false, hasCredentials: true });
    expect(ensureSignedIn).toHaveBeenCalledTimes(1);
  });

  it("stays in demo when the sign-in fails", async () => {
    const failing = createLiveApi(
      deps({
        ensureSignedIn: async () => {
          throw new Error("consent required");
        },
      })
    );
    await expect(failing.setMode("connected")).rejects.toThrow("consent required");
    // A refused consent must not leave the tab in a mode it cannot read.
    await expect(failing.health()).resolves.toMatchObject({ demo: true });
  });

  it("never asks for a sign-in to go back to demo", async () => {
    const ensureSignedIn = vi.fn(async () => {});
    const api = createLiveApi(deps({ ensureSignedIn }));
    await api.setMode("demo");
    expect(ensureSignedIn).not.toHaveBeenCalled();
  });
});

describe("persistent tenant snapshots", () => {
  function cache() {
    const entries = new Map<string, CacheEntry>();
    let nextRevision = 0;
    const persistence = {
      key: vi.fn(() => "tenant:user"),
      renewToken: vi.fn<PersistentTenantCache["renewToken"]>(async () => "renewed delegated token"),
      notifyPurge: vi.fn(),
      storage: {
        read: vi.fn(async (key: string) => entries.get(key) ?? { revision: "" }),
        write: vi.fn<PersistentTenantCache["storage"]["write"]>(async (key, revision, data, reads) => {
          if ((entries.get(key)?.revision ?? "") !== revision) return false;
          entries.set(key, { revision: String(++nextRevision), data, reads });
          return true;
        }),
        clear: vi.fn<PersistentTenantCache["storage"]["clear"]>(async (key, expected) => {
          if (expected !== undefined && (entries.get(key)?.revision ?? "") !== expected) return undefined;
          const revision = String(++nextRevision);
          entries.set(key, { revision });
          return revision;
        }),
      },
    } satisfies PersistentTenantCache;
    return { entries, persistence, api: createLiveApi(deps({ persistence })) };
  }

  function graphResponse(value: unknown[] = [], status = 200) {
    return { ok: status === 200, status, json: async () => ({ value }), text: async () => "Graph response" };
  }

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("loads automatically, shares concurrent requests, and reopens an authorized snapshot", async () => {
    const fetchMock = vi.fn(async (_url: string, _init: RequestInit) => graphResponse());
    vi.stubGlobal("fetch", fetchMock);
    const { api, persistence, entries } = cache();
    await api.setMode("connected");
    const [{ loadedAt }] = await Promise.all([api.health(), api.groups(), api.filters(), api.refresh()]);
    expect(loadedAt).toBeDefined();
    expect(persistence.storage.write).toHaveBeenCalledTimes(1);
    expect(persistence.renewToken).not.toHaveBeenCalled();
    const initialCalls = fetchMock.mock.calls.length;

    fetchMock.mockClear();
    const reopened = createLiveApi(deps({ persistence }));
    await reopened.setMode("connected");
    expect((await reopened.health()).loadedAt).toBe(loadedAt);
    expect(persistence.renewToken).toHaveBeenCalledTimes(1);
    expect(persistence.storage.write).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls).toHaveLength(entries.get("tenant:user")?.reads?.length ?? 0);
    expect(fetchMock.mock.calls.length).toBeLessThanOrEqual(initialCalls);
    expect(fetchMock.mock.calls[0][1].headers).toMatchObject({ Authorization: "Bearer renewed delegated token" });
    expect(fetchMock.mock.calls[0][1].cache).toBe("no-store");
  });

  it("automatically downloads an isolated snapshot for a different signed-in user", async () => {
    vi.stubGlobal("fetch", async () => graphResponse());
    const { api, entries, persistence } = cache();
    await api.setMode("connected");
    await api.health();
    persistence.key.mockReturnValue("tenant:other-user");
    expect((await api.health()).loadedAt).toBeDefined();
    expect(persistence.storage.read).toHaveBeenCalledWith("tenant:other-user");
    expect(persistence.renewToken).not.toHaveBeenCalled();
    expect([...entries.keys()]).toEqual(["tenant:user", "tenant:other-user"]);
  });

  it("purges previously cached data after a denied authorization", async () => {
    vi.stubGlobal("fetch", async () => graphResponse());
    const { api, entries, persistence } = cache();
    await api.setMode("connected");
    await api.health();
    const reopened = createLiveApi(deps({ persistence }));
    persistence.renewToken.mockRejectedValue(new CacheAccessDenied("Access revoked."));
    await reopened.setMode("connected");
    expect(await reopened.health()).toMatchObject({ cacheMessage: "Access revoked." });
    expect((await reopened.health()).loadedAt).toBeUndefined();
    expect(entries.get("tenant:user")?.data).toBeUndefined();
    expect(persistence.notifyPurge).toHaveBeenCalledWith("tenant:user");
  });

  it("locks a restored snapshot when revalidation fails and keeps it on disk after a transient error", async () => {
    vi.stubGlobal("fetch", async () => graphResponse());
    const { api, entries, persistence } = cache();
    const clock = vi.spyOn(Date, "now").mockReturnValue(0);
    await api.setMode("connected");
    const { loadedAt } = await api.health();
    expect(loadedAt).toBeDefined();
    clock.mockReturnValue(16 * 60_000);
    persistence.renewToken.mockRejectedValue(new Error("Graph unavailable."));
    expect((await api.health()).loadedAt).toBeUndefined();
    expect(entries.get("tenant:user")?.data?.loadedAt).toEqual(loadedAt);
    expect(persistence.notifyPurge).not.toHaveBeenCalled();
  });

  it("keeps the previous complete snapshot after a group refresh failure", async () => {
    let unavailable = false;
    const { api, entries, persistence } = cache();
    vi.stubGlobal("fetch", async (url: string) => {
      if (unavailable && url.includes("groupTypes")) throw new Error("Group query failed");
      return graphResponse();
    });
    await api.setMode("connected");
    const { loadedAt } = await api.health();
    unavailable = true;
    await expect(api.refresh()).rejects.toThrow("last saved tenant snapshot was kept");
    expect((await api.health()).loadedAt).toBe(loadedAt);
    expect(entries.get("tenant:user")?.data?.loadedAt).toBe(loadedAt);
    expect(persistence.storage.write).toHaveBeenCalledTimes(1);
  });

  it("cannot restore the cache after it was cleared during authorization", async () => {
    vi.stubGlobal("fetch", async () => graphResponse());
    const { api, persistence } = cache();
    await api.setMode("connected");
    await api.health();
    persistence.renewToken.mockImplementation(async () => {
      await persistence.storage.clear("tenant:user");
      return "renewed token";
    });
    const reopened = createLiveApi(deps({ persistence }));
    await reopened.setMode("connected");
    expect((await reopened.health()).loadedAt).toBeUndefined();
  });

  it("does not display or persist a refresh that overlapped a cache clear", async () => {
    const { api, entries, persistence } = cache();
    let cleared = false;
    vi.stubGlobal("fetch", async () => {
      if (!cleared) {
        cleared = true;
        await persistence.storage.clear("tenant:user");
      }
      return graphResponse();
    });
    await api.setMode("connected");
    await expect(api.refresh()).rejects.toThrow("The local cache was cleared during the refresh.");
    expect(entries.get("tenant:user")?.data).toBeUndefined();
    expect((await api.health()).loadedAt).toBeUndefined();
  });

  it("saves a scoped account's available data and replaces it when access shrinks", async () => {
    const { api, entries, persistence } = cache();
    const policy = (id: string) => ({
      id,
      displayName: id,
      "@odata.type": "#microsoft.graph.windows10GeneralConfiguration",
    });
    let visible = [policy("scope-a"), policy("scope-b")];
    const fetchMock = vi.fn(async (url: string) => {
      if (url.includes("/deviceManagement/deviceConfigurations?")) return graphResponse(visible);
      if (url.includes("/deviceManagement/deviceCompliancePolicies?")) return graphResponse([], 403);
      return graphResponse();
    });
    vi.stubGlobal("fetch", fetchMock);
    await api.setMode("connected");
    expect((await api.health()).loadedAt).toBeDefined();
    expect(entries.get("tenant:user")?.data?.policies.map(({ id }) => id)).toEqual(["scope-a", "scope-b"]);
    expect(entries.get("tenant:user")?.reads?.some((read) => read.fingerprint === "403")).toBe(true);

    fetchMock.mockClear();
    visible = [policy("scope-a")];
    const reopened = createLiveApi(deps({ persistence }));
    await reopened.setMode("connected");
    expect((await reopened.health()).loadedAt).toBeDefined();
    expect(entries.get("tenant:user")?.data?.policies.map(({ id }) => id)).toEqual(["scope-a"]);
    expect(persistence.notifyPurge).toHaveBeenCalledWith("tenant:user");
    expect(fetchMock.mock.calls.filter(([url]) => url.includes("deviceConfigurations?"))).toHaveLength(1);
  });
});

describe("reading a connected tenant", () => {
  /**
   * Every Graph collection answers empty, which is enough to drive a whole load
   * without a tenant: what is under test is the wiring -- that the delegated
   * token reaches the Authorization header, that the tenant is loaded once, and
   * that switching source never serves one tenant's data as the other's.
   */
  function stubGraph() {
    const fetchMock = vi.fn(async () => ({
      ok: true,
      status: 200,
      json: async () => ({ value: [] }),
    }));
    vi.stubGlobal("fetch", fetchMock);
    return fetchMock;
  }

  afterEach(() => vi.unstubAllGlobals());

  it("sends the delegated token as a bearer on every request", async () => {
    const fetchMock = stubGraph();
    const getToken = vi.fn(async () => "delegated-token");
    const api = createLiveApi(deps({ getToken }));
    await api.setMode("connected");
    await api.groups();

    expect(fetchMock.mock.calls.length).toBeGreaterThan(0);
    for (const [, init] of fetchMock.mock.calls as unknown as [string, RequestInit][]) {
      expect((init.headers as Record<string, string>).Authorization).toBe("Bearer delegated-token");
    }
    expect(getToken).toHaveBeenCalled();
  });

  // A 500-policy tenant moves ~65MB, so the UI's three concurrent requests
  // causing three loads would not be a subtle waste. Both this module and the
  // `TtlCache` it is handed collapse concurrent callers; what is asserted is the
  // property, not which layer provides it.
  it("serves the requests the UI fires together from a single load", async () => {
    const oneRead = stubGraph();
    const a = createLiveApi(deps());
    await a.setMode("connected");
    await a.groups();
    const callsForOneRead = oneRead.mock.calls.length;
    vi.unstubAllGlobals();

    const threeReads = stubGraph();
    const b = createLiveApi(deps());
    await b.setMode("connected");
    await Promise.all([b.groups(), b.filters(), b.unassignedPolicies()]);
    expect(threeReads.mock.calls.length).toBe(callsForOneRead);
  });

  it("drops the connected tenant when switching back to demo", async () => {
    // Showing one tenant's policies while the header says the other is the
    // worst failure this tool has. The stubbed Graph answers empty, so the two
    // sources are trivially distinguishable.
    stubGraph();
    const api = createLiveApi(deps());
    await api.setMode("connected");
    expect(await api.groups()).toEqual([]);

    await api.setMode("demo");
    expect(await api.groups()).toEqual(await demoApi.groups());

    await api.setMode("connected");
    expect(await api.groups()).toEqual([]);
  });

  it("re-reads from Graph after a refresh", async () => {
    const fetchMock = stubGraph();
    const api = createLiveApi(deps());
    await api.setMode("connected");
    await api.groups();
    const afterFirst = fetchMock.mock.calls.length;
    await api.refresh();
    expect(fetchMock.mock.calls.length).toBeGreaterThan(afterFirst);
  });

  it("passes the inactive-device option through to the live Graph search", async () => {
    const fetchMock = vi.fn(async () => ({
      ok: true,
      status: 200,
      json: async () => ({
        value: [
          {
            id: "stale",
            displayName: "KIOSK-STALE",
            isManaged: null,
            mdmAppId: "0000000a-0000-0000-c000-000000000000",
            approximateLastSignInDateTime: "2000-01-01T00:00:00Z",
          },
        ],
      }),
    }));
    vi.stubGlobal("fetch", fetchMock);
    const api = createLiveApi(deps());
    await api.setMode("connected");

    expect(await api.searchDevices("kiosk")).toEqual({ devices: [], hidden: 1 });
    expect(await api.searchDevices("kiosk", undefined, true)).toEqual({
      devices: [{ id: "stale", displayName: "KIOSK-STALE" }],
    });
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  // "Last refresh" in the header: the refresh reports when its read finished,
  // and health reports the same time afterwards, so the two cannot disagree.
  it("reports when the connected tenant was read, and only for connected", async () => {
    stubGraph();
    const api = createLiveApi(deps());
    await api.groups();
    expect((await api.health()).loadedAt).toBeUndefined();

    await api.setMode("connected");
    const before = Date.now();
    const { loadedAt } = await api.refresh();
    expect(loadedAt).toBeDefined();
    const readAt = Date.parse(loadedAt!);
    expect(readAt).toBeGreaterThanOrEqual(before - 1000);
    expect(readAt).toBeLessThanOrEqual(Date.now());
    expect((await api.health()).loadedAt).toBe(loadedAt);

    await api.setMode("demo");
    expect((await api.health()).loadedAt).toBeUndefined();
  });
});

describe("refresh", () => {
  it("re-reads and reports no warnings for the sample tenant", async () => {
    const api = createLiveApi(deps());
    await api.groups();
    await expect(api.refresh()).resolves.toEqual({ warnings: [] });
  });
});

describe("exports", () => {
  it("carry the bytes, not a URL, so a download needs no backend", async () => {
    const api = createLiveApi(deps());
    await api.assignmentReport();
    expect(api.assignmentReportExport()).toMatchObject({ kind: "text", mime: "text/csv" });
    expect(api.simulateExport({ groupIds: [] }, "csv")).toMatchObject({ kind: "text", mime: "text/csv" });
    expect(api.simulateExport({ groupIds: [] }, "json")).toMatchObject({ kind: "text", mime: "application/json" });
  });

  it("are never a data: URL", async () => {
    // The whole point of the descriptor. Safari does not honour an anchor's
    // `download` attribute on a data: URL, so a data: export is one that
    // silently fails to save on every Apple device -- see download.ts.
    const api = createLiveApi(deps());
    await api.assignmentReport();
    for (const d of [api.assignmentReportExport(), api.simulateExport({ groupIds: [] }, "json")]) {
      expect(d.kind).toBe("text");
      if (d.kind === "text") expect(d.text.startsWith("data:")).toBe(false);
    }
  });

  it("do not throw before anything has loaded", () => {
    // They are reached from a click handler, where a throw would take out the
    // panel. The panels that render the buttons already hold loaded data, but a
    // race must not be fatal.
    const api = createLiveApi(deps());
    expect(() => api.assignmentReportExport()).not.toThrow();
    expect(() => api.simulateExport({ groupIds: [] }, "csv")).not.toThrow();
  });

  it("carry the loaded tenant once there is one", async () => {
    const api = createLiveApi(deps());
    const groups = await api.groups();
    expect(groups.length).toBeGreaterThan(0);
    const report = api.assignmentReportExport();
    expect(report.kind).toBe("text");
    if (report.kind !== "text") return;
    expect(report.text.length).toBeGreaterThan(0);
    // Sanity: it is the sample tenant's report, not an empty one.
    expect(report.text.split("\n").length).toBeGreaterThan(1);
  });
});
