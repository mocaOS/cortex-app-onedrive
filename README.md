# OneDrive Sync — a Cortex App

Sync selected [OneDrive](https://onedrive.com/) folders into the
[Cortex](https://github.com/mocaOS/cortex-app) knowledge graph — on demand
or on a schedule, fully server-side via Microsoft Graph's delta feed.

Connect with a device code (no redirect server, no client secret), pick
folders in a built-in browser, and Cortex tracks exactly what changed
through Graph's own change feed. New and updated documents land in your
knowledge graph automatically; edited files replace their previous document
instead of duplicating.

Built from [cortex-app-template](https://github.com/mocaOS/cortex-app-template).
Type: `platform` — the sync itself runs **inside Cortex** as a declarative
platform task, not in your browser.

> Requires a Cortex instance with dynamic task auth, binary passthrough
> upload with redirect following, and paginate cursor support (cortex-app
> ≥ 2.1 / July 2026 platform extensions, wave 2).

## Install

1. Create a free **Microsoft Entra app registration** at
   [portal.azure.com](https://portal.azure.com) (App registrations → New):
   supported account types must include **personal Microsoft accounts**
   (pick the multitenant + personal option), and under Authentication
   enable **"Allow public client flows"** (the device-code sign-in needs
   it). No client secret, no redirect URI, no API permissions beyond the
   default delegated `Files.Read` / `offline_access` / `User.Read`
   (requested at sign-in).

   **Registrations must live in an Entra directory.** Sign in with a
   work/Microsoft 365 account that has one, or give a personal account its
   own directory via the free [Azure signup](https://azure.microsoft.com/free)
   — a bare personal Microsoft account has no directory and fails with
   "user does not exist in tenant" errors. The directory only *hosts* the
   registration; any OneDrive account can connect through it. This
   registration is deliberately **per deployment** — Cortex ships no shared
   client id, so your instance talks to Microsoft as *your* app, not ours.
2. `npm install && npm run package` → `onedrive-sync-{version}.zip`
3. Cortex admin → **Settings → Apps → Install**, upload the zip.
   (Requires `ENABLE_APPS=true` on the instance.)
4. Configure `ENTRA_CLIENT_ID` (Settings → Apps → Configure).
5. Open the app (`/apps` → OneDrive Sync) and connect: generate a sign-in
   code, enter it at **microsoft.com/devicelogin** in another tab, approve —
   the app connects by itself.
6. Select folders, pick a schedule, hit **Sync**.

## How the sync works

- **setup** — refreshes the access token and **persists the rotated refresh
  token** (Microsoft rotates them on every refresh; the token lives in the
  app's private storage inside your instance). First run: full enumeration
  via `/me/drive/root/delta`; afterwards the stored `@odata` cursor returns
  exactly what changed. Backlogs drain across runs (~2000 entries per run).
- **per file** — Graph has no folder-scoped delta, so the drive-wide feed is
  filtered per item: deleted entries, folders, items outside your selected
  folders, non-document extensions (optional) and files over ~19 MB are
  skipped with visible reasons. Then a **binary passthrough upload**: Cortex
  GETs `/items/{id}/content` server-side and follows Graph's one redirect to
  the pre-authenticated download host (variable per drive — the platform
  validates the hop and forwards no credentials to it).
  Edited files **replace their Cortex document in place**: the upload
  carries a stable `source_key` (`onedrive:{driveId}:{itemId}`) plus the previous document id
  (`docmap/{id}`), so the document keeps its id and Cortex re-extracts
  only the parts that changed. On a Cortex backend without in-place
  replace, the app falls back to the old behaviour (new document, then
  the previous one is deleted).
- **finally** — persists the next cursor, only when nothing failed.

### Token custody

- The refresh token lives in the app's private storage inside your Cortex
  instance; access tokens are minted per run server-side and per browser
  session for the folder picker. Sign-in and token calls go through the
  instance (Microsoft's login endpoints aren't CORS-open); Graph browsing
  calls run directly from the browser.
- Revoke anytime at account.microsoft.com → Privacy → App access.

### v1 limitations

- Deleting a file in OneDrive does **not** delete it in Cortex.
- OneNote notebooks are skipped (they're packages, not files).
- Files larger than ~19 MB are skipped (platform transfer cap).
- Work/school tenants may require admin consent depending on tenant policy;
  personal accounts need none.

## Manifest notes

- `keyScope: read_write`; `endpoints: ["upload", "documents", "collections"]`
- `capabilities: http` (hosts `login.microsoftonline.com` +
  `graph.microsoft.com`), `tasks`, `storage`
- `externalHosts: ["https://graph.microsoft.com"]` — browser-direct browsing

## Dev loop

```
cp .env.example .env   # CORTEX_DEV_URL + a cortex_rw_… key
npm run dev
```
