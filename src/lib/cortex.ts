/**
 * Cortex app client — the only place your app talks to the outside world.
 *
 * Contract (do not work around it):
 *  - All requests are RELATIVE (`./api/…`). Hosted apps live under
 *    /apps/{slug}/ and absolute paths break there. `npm run validate`
 *    rejects absolute API paths.
 *  - Your app never holds a Cortex API key. In production the hosting proxy
 *    attaches the app's server-side key; the browser only ever sees a
 *    short-lived app token, delivered by the launcher via postMessage.
 *  - In `npm run dev`, the Vite proxy plays the hosting proxy's role using
 *    CORTEX_DEV_URL / CORTEX_DEV_KEY from your .env (see vite.config.ts).
 */

// ---------------------------------------------------------------------------
// App token handshake (production: sandboxed iframe inside the launcher)
// ---------------------------------------------------------------------------

const API_BASE = new URL("./api/", document.baseURI);
const inLauncher = window.parent !== window;

let appToken: string | null = null;
let waiters: Array<() => void> = [];

if (inLauncher) {
  window.addEventListener("message", (event: MessageEvent) => {
    const d = event.data;
    if (d && typeof d === "object" && d.type === "cortex:token" && typeof d.token === "string") {
      appToken = d.token;
      for (const w of waiters) w();
      waiters = [];
    }
  });
  window.parent.postMessage({ type: "cortex:ready" }, "*");
}

function awaitToken(timeoutMs = 3000): Promise<void> {
  if (!inLauncher || appToken) return Promise.resolve();
  return new Promise((resolve) => {
    waiters.push(resolve);
    setTimeout(resolve, timeoutMs); // never deadlock in unexpected hosts
  });
}

function requestRenewal(): void {
  appToken = null;
  if (inLauncher) window.parent.postMessage({ type: "cortex:token:renew" }, "*");
}

// ---------------------------------------------------------------------------
// Fetch helpers
// ---------------------------------------------------------------------------

async function appFetch(path: string, init: RequestInit): Promise<Response> {
  await awaitToken();
  // FormData bodies must NOT get a manual Content-Type — the browser sets
  // the multipart boundary itself.
  const isForm = typeof FormData !== "undefined" && init.body instanceof FormData;
  const doFetch = () =>
    fetch(new URL(path, API_BASE), {
      ...init,
      headers: {
        ...(isForm ? {} : { "Content-Type": "application/json" }),
        ...(appToken ? { Authorization: `Bearer ${appToken}` } : {}),
        ...(init.headers ?? {}),
      },
    });

  let res = await doFetch();
  if (res.status === 401 && inLauncher) {
    // token expired — ask the launcher for a fresh one and retry once
    requestRenewal();
    await awaitToken();
    res = await doFetch();
  }
  return res;
}

/** Raw access to allowlisted Cortex API endpoints (path relative to /api/). */
export function cortex(path: string, init: RequestInit = {}): Promise<Response> {
  return appFetch(`cortex/${path}`, init);
}

/** Platform capabilities (type: "platform" apps only — declare them in app.json). */
export function platform(path: string, init: RequestInit = {}): Promise<Response> {
  return appFetch(`platform/${path}`, init);
}

// ---------------------------------------------------------------------------
// High-level API
// ---------------------------------------------------------------------------

export interface CortexCollection {
  id: string;
  name?: string;
  [k: string]: unknown;
}

export async function listCollections(): Promise<CortexCollection[]> {
  const res = await cortex("collections");
  if (!res.ok) throw new Error(`collections failed: ${res.status}`);
  const data = await res.json();
  return Array.isArray(data) ? data : (data.collections ?? []);
}
