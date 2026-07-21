/**
 * OneDrive client — Microsoft Graph with the device-code flow.
 *
 * Two transports, deliberately split:
 *  - login.microsoftonline.com (device-code + token endpoints) is NOT
 *    CORS-open for public clients → those calls go through the platform
 *    `http` envelope (server-side, host-allowlisted). No secret exists:
 *    device-code is a public-client flow, only the client id is needed.
 *  - graph.microsoft.com IS CORS-open → interactive browsing calls it
 *    directly from the browser with a short-lived access token
 *    (declared in `externalHosts`).
 *
 * Connect flow (sandboxed iframes can't host redirects or open popups):
 * request a device code, show the user code + microsoft.com/devicelogin
 * link, poll the token endpoint until approval. The refresh token lands in
 * app storage — and unlike Dropbox, Microsoft ROTATES refresh tokens
 * (~90-day rolling window), so every refresh persists the new one; the
 * scheduled sync task does the same server-side.
 */

import { platform } from "./cortex";
import { storageDelete, storageGet, storagePut } from "./platform";

export const REFRESH_TOKEN_KEY = "auth/refresh_token";
const AUTHORITY = "https://login.microsoftonline.com/common";
const SCOPE = "Files.Read offline_access User.Read";

export interface DriveEntry {
  id: string;
  name: string;
  /** Graph parent path, e.g. "/drive/root:/Work" */
  parentPath: string;
  size?: number;
  isFolder: boolean;
}

/** The admin-configured Entra application (client) id, or null. */
export async function getClientId(): Promise<string | null> {
  const res = await platform("config");
  if (!res.ok) return null;
  const data = await res.json();
  return data.values?.ENTRA_CLIENT_ID ?? null;
}

/** Form-POST to the login authority via the platform envelope (no CORS). */
async function loginPost(path: string, form: Record<string, string>): Promise<any> {
  const res = await platform("http", {
    method: "POST",
    body: JSON.stringify({
      method: "POST",
      url: `${AUTHORITY}${path}`,
      body: new URLSearchParams(form).toString(),
      content_type: "application/x-www-form-urlencoded",
    }),
  });
  return { status: res.status, body: await res.json().catch(() => null) };
}

// ---------------------------------------------------------------------------
// Device-code connect flow
// ---------------------------------------------------------------------------

export interface DeviceCode {
  userCode: string;
  verificationUri: string;
  deviceCode: string;
  intervalMs: number;
  expiresAt: number;
}

export async function beginConnect(clientId: string): Promise<DeviceCode> {
  const { status, body } = await loginPost("/oauth2/v2.0/devicecode", {
    client_id: clientId,
    scope: SCOPE,
  });
  if (status !== 200 || !body?.device_code) {
    throw new Error(
      body?.error === "invalid_client" || body?.error === "unauthorized_client"
        ? "Microsoft rejected the client id — check ENTRA_CLIENT_ID and make " +
          'sure the app registration has "Allow public client flows" enabled.'
        : `Device code request failed: ${body?.error_description ?? status}`,
    );
  }
  return {
    userCode: body.user_code,
    verificationUri: body.verification_uri ?? "https://microsoft.com/devicelogin",
    deviceCode: body.device_code,
    intervalMs: (Number(body.interval) || 5) * 1000,
    expiresAt: Date.now() + (Number(body.expires_in) || 900) * 1000,
  };
}

/** One poll tick. Returns true when connected, false while pending; throws
 * on terminal errors (expired, declined). */
export async function pollConnect(clientId: string, device: DeviceCode): Promise<boolean> {
  if (Date.now() > device.expiresAt) {
    throw new Error("The code expired — generate a new one.");
  }
  const { status, body } = await loginPost("/oauth2/v2.0/token", {
    client_id: clientId,
    grant_type: "urn:ietf:params:oauth:grant-type:device_code",
    device_code: device.deviceCode,
  });
  if (status === 200 && body?.refresh_token) {
    await storagePut(REFRESH_TOKEN_KEY, body.refresh_token);
    cachedAccess = {
      token: body.access_token,
      expiresAt: Date.now() + (Number(body.expires_in) || 3600) * 1000,
    };
    return true;
  }
  const error = body?.error ?? "";
  if (error === "authorization_pending" || error === "slow_down") return false;
  throw new Error(
    error === "expired_token"
      ? "The code expired — generate a new one."
      : `Connection failed: ${body?.error_description ?? error ?? status}`,
  );
}

export async function isConnected(): Promise<boolean> {
  return (await storageGet<string>(REFRESH_TOKEN_KEY)) !== null;
}

export async function disconnect(): Promise<void> {
  await storageDelete(REFRESH_TOKEN_KEY);
  cachedAccess = null;
}

// ---------------------------------------------------------------------------
// Access tokens (browser side, for the interactive folder browser)
// ---------------------------------------------------------------------------

let cachedAccess: { token: string; expiresAt: number } | null = null;

async function accessToken(clientId: string): Promise<string> {
  if (cachedAccess && Date.now() < cachedAccess.expiresAt - 60_000) {
    return cachedAccess.token;
  }
  const refresh = await storageGet<string>(REFRESH_TOKEN_KEY);
  if (!refresh) throw new Error("Not connected to OneDrive.");
  const { status, body } = await loginPost("/oauth2/v2.0/token", {
    client_id: clientId,
    grant_type: "refresh_token",
    refresh_token: refresh,
    scope: SCOPE,
  });
  if (status !== 200 || !body?.access_token) {
    if (body?.error === "invalid_grant") {
      await disconnect();
      throw new Error("The OneDrive session expired or was revoked — connect again.");
    }
    throw new Error(`Token refresh failed: ${body?.error_description ?? status}`);
  }
  // Microsoft rotates refresh tokens — persist the new one or the session
  // dies when the old token ages out of its rolling window.
  if (body.refresh_token) await storagePut(REFRESH_TOKEN_KEY, body.refresh_token);
  cachedAccess = {
    token: body.access_token,
    expiresAt: Date.now() + (Number(body.expires_in) || 3600) * 1000,
  };
  return cachedAccess.token;
}

async function graphGet(clientId: string, path: string): Promise<any> {
  const token = await accessToken(clientId);
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const detail = await res.json().catch(() => null);
    throw new Error(`Graph ${path} failed: ${detail?.error?.message ?? res.status}`);
  }
  return res.json();
}

export async function currentAccount(clientId: string): Promise<string> {
  const me = await graphGet(clientId, "/me");
  return me?.displayName ?? me?.userPrincipalName ?? "connected";
}

// ---------------------------------------------------------------------------
// Folder browsing (interactive)
// ---------------------------------------------------------------------------

function childrenPath(folderPath: string): string {
  if (!folderPath) return "/me/drive/root/children";
  const encoded = folderPath.split("/").filter(Boolean).map(encodeURIComponent).join("/");
  return `/me/drive/root:/${encoded}:/children`;
}

/** Immediate children of a path ("" = root): sub-folders (selectable) and
 * files (context). Paths are plain human paths like "Work/Reports". */
export async function listEntries(
  clientId: string,
  folderPath: string,
): Promise<{ folders: DriveEntry[]; files: DriveEntry[] }> {
  const folders: DriveEntry[] = [];
  const files: DriveEntry[] = [];
  let url = `${childrenPath(folderPath)}?$select=id,name,size,folder,file,parentReference&$top=500`;
  for (;;) {
    const data = await graphGet(clientId, url);
    for (const entry of data.value ?? []) {
      const item: DriveEntry = {
        id: entry.id,
        name: entry.name,
        parentPath: entry.parentReference?.path ?? "",
        size: entry.size,
        isFolder: Boolean(entry.folder),
      };
      (item.isFolder ? folders : files).push(item);
    }
    const next: string | undefined = data["@odata.nextLink"];
    if (!next) break;
    url = next.replace("https://graph.microsoft.com/v1.0", "");
  }
  const byName = (a: DriveEntry, b: DriveEntry) => a.name.localeCompare(b.name);
  return { folders: folders.sort(byName), files: files.sort(byName) };
}

export interface SelectionSurvey {
  fileCount: number;
  truncated: boolean;
}

/** Project how many files the selected roots would sync: BFS over children
 * (Graph offers no folder-scoped recursive listing), folder-visit capped. */
export async function surveySelection(
  clientId: string,
  roots: string[],
  keep: (entry: DriveEntry) => boolean,
  folderCap = 40,
): Promise<SelectionSurvey> {
  const seen = new Set<string>();
  const queue = [...roots.map((r) => r.replace(/\/+$/, ""))];
  let visited = 0;
  let fileCount = 0;
  while (queue.length > 0) {
    const current = queue.shift()!;
    if (seen.has(current)) continue;
    seen.add(current);
    if (visited >= folderCap) return { fileCount, truncated: true };
    visited += 1;
    const entries = await listEntries(clientId, current);
    fileCount += entries.files.filter(keep).length;
    for (const sub of entries.folders) {
      queue.push(current ? `${current}/${sub.name}` : sub.name);
    }
  }
  return { fileCount, truncated: false };
}
