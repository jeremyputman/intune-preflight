/**
 * The live build's session: the one place that knows about the tenant/client
 * config, MSAL, and the fact that a page load might be the tail of a sign-in
 * redirect. `api.ts` wires `liveApi` to these functions; the header chrome reads
 * them to say who is signed in.
 *
 * It exists so `liveApi.ts` can stay free of the MSAL runtime (see its header)
 * and so `App.tsx` never has to sequence an init step -- everything that needs
 * auth awaits `ensureReady()`, which is idempotent.
 */

import { acquireGraphToken, currentAccount, endSession, initAuth, renewTenantCacheToken, signIn } from "./auth.ts";
import { forgetLiveMode, rememberConnectedMode, type PersistentTenantCache } from "./liveApi.ts";
import { type LiveConfig, readLiveConfig, resolveLiveConfig, saveLiveConfig } from "./liveConfig.ts";
import { clearSession, readSession, removeSession, writeSession } from "./storage.ts";
import { tenantCache } from "./tenantCache.ts";

export const PERSIST_TENANT_CACHE = import.meta.env.VITE_PERSIST_TENANT_CACHE === "1";

function cacheKey(): string | null {
  const account = currentAccount();
  const clientId = getLiveConfig()?.clientId;
  if (!account?.tenantId || !account.localAccountId || !clientId || !cacheChannel) return null;
  return JSON.stringify([account.tenantId.toLowerCase(), account.localAccountId.toLowerCase(), clientId.toLowerCase()]);
}

let cacheChannel: BroadcastChannel | undefined;

function setupCacheChannel() {
  if (!PERSIST_TENANT_CACHE || cacheChannel || typeof BroadcastChannel === "undefined") return;
  cacheChannel = new BroadcastChannel("intune-preflight-tenant-cache");
  cacheChannel.onmessage = ({ data }: MessageEvent<{ key?: string; end?: boolean; clear?: boolean }>) => {
    if (!data?.key || data.key !== cacheKey()) return;
    if (data.end) {
      forgetLiveMode();
      void endSession()
        .catch(() => {})
        .finally(() => window.location.reload());
    } else {
      if (data.clear) forgetLiveMode();
      window.location.reload();
    }
  };
}

export const liveCachePersistence: PersistentTenantCache = {
  storage: tenantCache,
  key: cacheKey,
  renewToken: renewTenantCacheToken,
  notifyPurge: (key) => cacheChannel?.postMessage({ key }),
};

async function purgeCache(key: string, end: boolean): Promise<void> {
  try {
    await tenantCache.clear(key);
  } finally {
    cacheChannel?.postMessage({ key, end, clear: !end });
  }
}

export async function clearLocalTenantCache(): Promise<void> {
  const key = PERSIST_TENANT_CACHE ? cacheKey() : null;
  if (!key) throw new Error("Local tenant storage is unavailable in this browser.");
  await purgeCache(key, false);
  forgetLiveMode();
  window.location.reload();
}

// Resolved once per page load: `resolveLiveConfig` persists whatever the query
// string supplied, and re-running it on every caller would mean re-writing
// localStorage over and over for no reason.
let config: LiveConfig | null | undefined;

/**
 * Set while the admin is changing which tenant this build points at. It lives
 * in the tab's session, not in a component, because the change is made across
 * a reload: MSAL is built once per page load against one tenant's authority,
 * so the way to stop using a tenant is to load the page without it.
 */
const CHANGING_KEY = "preflight.live.changingTenant";

export function isChangingTenant(): boolean {
  return readSession(CHANGING_KEY) === "1";
}

export function getLiveConfig(): LiveConfig | null {
  // While the tenant is being changed the build behaves as unconfigured, which
  // is what brings the setup form back -- with the saved IDs filled in, since
  // the form reads those from storage itself.
  if (config === undefined) config = isChangingTenant() ? null : resolveLiveConfig(window.location.search);
  return config;
}

/**
 * Reopens the setup form so the tenant ID and client ID can be corrected or
 * replaced. Before this there was no way back to that form once two IDs were
 * saved: a mistyped ID failed at Microsoft's sign-in page every time, with
 * nothing in the app to fix it.
 *
 * The session is left alone until new IDs are actually submitted, so Cancel
 * returns to exactly where the admin was. The address is reloaded without its
 * query string, because a `?tenant=&client=` bookmark would otherwise put the
 * old IDs straight back.
 */
export function changeLiveTenant(): void {
  writeSession(CHANGING_KEY, "1");
  window.location.replace(window.location.pathname);
}

export function cancelChangeTenant(): void {
  removeSession(CHANGING_KEY);
  window.location.reload();
}

export function isConfigured(): boolean {
  return getLiveConfig() !== null;
}

let readying: Promise<void> | null = null;

export function ensureReady(): Promise<void> {
  if (!readying) {
    readying = (async () => {
      const cfg = getLiveConfig();
      // Nothing configured yet: there is no authority to build MSAL against,
      // and Demo mode needs none. Staying silent here is what makes an
      // unconfigured live build behave exactly like the public demo.
      if (!cfg) return;
      await initAuth(cfg);
      setupCacheChannel();
    })();
  }
  return readying;
}

/**
 * Adopts a config from the first-run form and goes straight to sign-in, rather
 * than saving it and making the admin click Connected as a second step.
 *
 * Both memos are reset because they were populated by an UNCONFIGURED load:
 * `ensureReady` resolved without building MSAL, and there is no authority for it
 * to have used. Clearing them is what lets the new tenant take effect without a
 * page reload.
 */
export async function configureAndSignIn(next: LiveConfig): Promise<void> {
  // Replacing a tenant: drop everything this tab held for the old one first.
  // MSAL keeps its accounts and tokens in session storage, and a new sign-in
  // must not find the previous tenant's account waiting there.
  if (isChangingTenant()) clearSession();
  saveLiveConfig(next);
  config = next;
  readying = null;
  // `signIn` navigates to Entra, so the mode has to be remembered first.
  rememberConnectedMode();
  await ensureSignedIn();
}

export async function ensureSignedIn(): Promise<void> {
  const cfg = getLiveConfig();
  if (!cfg) {
    throw new Error("No tenant configured. Add a tenant ID and client ID to sign in.");
  }
  await ensureReady();
  if (currentAccount()) return;
  // Navigates to Entra; this does not return on the success path.
  await signIn();
}

export function isSignedIn(): boolean {
  return currentAccount() !== null;
}

export async function getToken(): Promise<string> {
  await ensureReady();
  return acquireGraphToken();
}

export interface LiveSessionIdentity {
  /** UPN of the signed-in admin. */
  username: string;
  /** Display name, when the token carried one. */
  name?: string;
  /** The tenant the token is for, as Entra reported it. */
  tenantId: string;
}

export function signedInIdentity(): LiveSessionIdentity | null {
  const account = currentAccount();
  if (!account) return null;
  return {
    username: account.username,
    name: account.name,
    // Prefer the tenant the TOKEN is for over the one we asked for: if an admin
    // configured a domain name, this is the GUID it resolved to, and if the two
    // ever disagree the token is the truth.
    tenantId: account.tenantId || (getLiveConfig()?.tenantId ?? ""),
  };
}

/**
 * Ends the session: drops the tokens, forgets the remembered mode, and reloads
 * so the in-memory tenant data goes with the page. A reload is the only way to
 * be sure 15MB of someone's tenant is actually gone from the tab rather than
 * still reachable from a closure somewhere.
 *
 * The tenant/client config is deliberately KEPT -- it is not sensitive, and the
 * admin would otherwise have to retype it to sign in again.
 */
export async function endLiveSession(): Promise<void> {
  const key = PERSIST_TENANT_CACHE ? cacheKey() : null;
  forgetLiveMode();
  const results = await Promise.allSettled([endSession(), key ? purgeCache(key, true) : undefined]);
  if (results.some((result) => result.status === "rejected")) {
    window.alert(
      "The browser could not finish clearing this session. Clear this site's data in your browser settings."
    );
  }
  window.location.reload();
}

/** The config as stored, ignoring the query string -- for the setup form's initial values. */
export { readLiveConfig };
