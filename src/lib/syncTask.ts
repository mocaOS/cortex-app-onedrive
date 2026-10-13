/**
 * The sync, as a Cortex platform task.
 *
 * Instead of a browser-side loop, the whole OneDrive → Cortex pipeline is a
 * declarative step-queue that Cortex executes server-side:
 *
 *   setup    refresh the access token (and PERSIST the rotated refresh
 *            token — Microsoft rotates them; losing one kills the session
 *            within its 90-day rolling window) → read the drive-wide delta
 *            cursor → first run: full enumeration via /me/drive/root/delta;
 *            afterwards: continue from the stored @odata link — Graph's own
 *            change feed returns exactly what changed
 *   items    the delta feed, filtered per item: deleted entries, folders,
 *            items outside the selected folders, non-document extensions
 *            (optional) and oversized files are skipped with visible reasons
 *   steps    binary passthrough upload that REPLACES the file's existing
 *            Cortex document in place (source_key + replace_document_id
 *            from docmap/{id} — same id, only changes re-extracted): Cortex GETs
 *            graph /items/{id}/content server-side and follows Graph's one
 *            302 to the pre-authenticated download host (followRedirect —
 *            the target host varies per drive and can't be allowlisted)
 *   finally  persist the next cursor (@odata.nextLink while a backlog is
 *            draining, @odata.deltaLink once caught up) — only when nothing
 *            failed, so failed files are re-listed by the next run
 *
 * Graph has no folder-scoped delta, so ONE drive-wide cursor feeds the task
 * and the selected folders act as a filter (parentReference.path prefix).
 */

import { storageGet, storagePut } from "./platform";

export const SYNC_TASK_NAME = "onedrive-sync";

/** Max selected root folders — each adds one `contains` clause to the
 * per-item filter; the ceiling is UX, not the platform. */
export const MAX_ROOTS = 8;

/** Delta page budget per run: 10 pages × $top=200 stays at the platform's
 * 2000-item task cap. Backlogs drain cursor-wise across runs. */
const PAGE_SIZE = 200;
const MAX_PAGES = 10;

/** Files above this are skipped — the platform's server-side fetch caps
 * responses at 20 MB; leave headroom. */
const MAX_FILE_BYTES = 19_000_000;

const DOC_EXTS_LIST = [
  "pdf", "docx", "xlsx", "pptx", "doc", "md", "markdown", "txt", "text",
  "html", "htm", "csv", "epub", "png", "jpg", "jpeg", "tiff", "webp",
];
const DOC_EXTS = ` ${DOC_EXTS_LIST.join(" ")} `;

/** Client-side mirror of the task's per-item skip rules — drives the "~N
 * files will sync" projection so it matches what a run would actually do. */
export function wouldSync(
  entry: { name: string; size?: number },
  docsOnly: boolean,
): boolean {
  if ((entry.size ?? 0) > MAX_FILE_BYTES) return false;
  if (!docsOnly) return true;
  const name = entry.name.replace(/^\.+/, "");
  const ext = name.includes(".") ? name.split(".").pop()!.toLowerCase() : "";
  return DOC_EXTS_LIST.includes(ext);
}

export interface SyncFolder {
  /** human path relative to the drive root, "" = entire OneDrive */
  path: string;
}

export interface SyncTaskOptions {
  folders: SyncFolder[];
  collectionId?: string;
  everyMinutes?: number;
  docsOnly?: boolean;
}

/** Graph parentReference.path for a human path: "/drive/root:" + "/Work" */
function graphNeedle(path: string): string {
  const clean = path.replace(/^\/+|\/+$/g, "");
  return `/drive/root:${clean ? `/${clean}` : ""}/`;
}

const DELTA_SELECT =
  "$select=id,name,size,file,folder,parentReference,deleted,eTag";

/** Old backends (no in-place replace) answer an upload without a "replaced"
 * field and mint a new document — then remove the superseded one, exactly
 * like the pre-replace delete-then-upload flow did. */
const REPLACE_FALLBACK_STEPS = [
  {
    cortex: { method: "DELETE", path: "documents/{prev.value.docId}" },
    when: {
      and: [
        { notEmpty: "$prev.value.docId" },
        { eq: ["{up.body.replaced}", ""] },
        { neq: ["{up.body.document_id}", "{prev.value.docId}"] },
      ],
    },
  },
];

export function buildSyncTask(opts: SyncTaskOptions): Record<string, unknown> {
  const uploadParams = new URLSearchParams({
    start_processing: "true",
    source: "onedrive-sync",
  });
  if (opts.collectionId) uploadParams.set("collection_id", opts.collectionId);
  // item ids are unique per drive — key on drive id + item id
  const uploadPath =
    `upload?${uploadParams.toString()}` +
    `&source_key=onedrive%3A{vars.driveId|urlencode}%3A{vars.id|urlencode}` +
    `&replace_document_id={prev.value.docId|urlencode}`;

  const auth = { bearer: "$tok.body.access_token" };
  const paginate = { items: "value", next: "@odata.nextLink", maxPages: MAX_PAGES };
  const roots = opts.folders.slice(0, MAX_ROOTS).map((f) => f.path);
  const inSelected = {
    or: roots.map((path) => ({
      contains: ["{vars.path}/", graphNeedle(path)],
    })),
  };

  const cursorWrites = [];
  for (const variant of ["da", "db"]) {
    for (const link of ["@odata.nextLink", "@odata.deltaLink"]) {
      cursorWrites.push({
        store: { put: "delta/cursor", value: `$${variant}.last.${link}` },
        when: {
          and: [
            { eq: ["$run.failedCount", 0] },
            { notEmpty: `$${variant}.last.${link}` },
          ],
        },
      });
    }
  }

  return {
    name: SYNC_TASK_NAME,
    ...(opts.everyMinutes ? { schedule: { everyMinutes: opts.everyMinutes } } : {}),
    concurrency: 2,
    setup: [
      { id: "rt", store: { get: "auth/refresh_token" } },
      {
        id: "tok",
        http: {
          method: "POST",
          url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
          contentType: "application/x-www-form-urlencoded",
          body:
            "client_id={config.ENTRA_CLIENT_ID|urlencode}" +
            "&grant_type=refresh_token" +
            "&refresh_token={rt.value|urlencode}" +
            "&scope=Files.Read%20offline_access",
          responseType: "json",
        },
      },
      // Microsoft ROTATES refresh tokens — persist the new one immediately,
      // before anything that could fail, or the session dies mid-window
      {
        store: { put: "auth/refresh_token", value: "$tok.body.refresh_token" },
        when: { notEmpty: "$tok.body.refresh_token" },
      },
      { id: "cur", store: { get: "delta/cursor" } },
      {
        id: "da",
        when: { empty: "$cur.value" },
        http: {
          method: "GET",
          url:
            "https://graph.microsoft.com/v1.0/me/drive/root/delta" +
            `?${DELTA_SELECT}&$top=${PAGE_SIZE}`,
          auth,
          paginate,
        },
      },
      {
        id: "db",
        when: { notEmpty: "$cur.value" },
        http: { method: "GET", url: "{cur.value}", auth, paginate },
      },
    ],
    items: {
      fromEach: ["$da.items", "$db.items"],
      vars: {
        id: "{item.id}",
        name: "{item.name}",
        path: "{item.parentReference.path}",
        driveId: "{item.parentReference.driveId}",
        etag: "{item.eTag}",
        size: "{item.size}",
        mime: "{item.file.mimeType}",
        deleted: "{item.deleted.state}",
      },
      skipIfStored: "synced/{item.id}:{item.eTag|slug}",
    },
    steps: [
      {
        skipItem: {
          when: { notEmpty: "$vars.deleted" },
          reason: "deleted on OneDrive (deletion sync not supported)",
        },
      },
      {
        skipItem: {
          when: { empty: "$vars.mime" },
          reason: "folder or non-file entry",
        },
      },
      {
        skipItem: {
          when: { not: inSelected },
          reason: "outside the selected folders",
        },
      },
      ...(opts.docsOnly !== false
        ? [
            {
              skipItem: {
                when: { not: { contains: [DOC_EXTS, " {vars.name|ext} "] } },
                reason: "not a document type (.{vars.name|ext})",
              },
            },
          ]
        : []),
      {
        skipItem: {
          when: { gt: ["{vars.size}", MAX_FILE_BYTES] },
          reason: "larger than the 19 MB transfer cap",
        },
      },
      { id: "prev", store: { get: "docmap/{vars.id}" } },
      // binary passthrough: Graph /content 302s to a pre-authenticated URL
      // on a variable host — followRedirect handles the one hop server-side.
      // A changed file REPLACES its Cortex document in place (same id — only
      // the changed parts are re-extracted): source_key finds it,
      // replace_document_id adopts documents synced before source keys
      // existed (empty when new).
      {
        id: "up",
        cortex: {
          method: "POST",
          path: uploadPath,
          multipart: {
            fromUrl: "https://graph.microsoft.com/v1.0/me/drive/items/{vars.id}/content",
            auth,
            followRedirect: true,
            filename: "onedrive-{vars.name}",
          },
        },
      },
      // compatibility: a Cortex backend without in-place replace ignores the
      // replace params and creates a new document (its response has no
      // "replaced" field) — delete the previous one as older app versions did
      ...REPLACE_FALLBACK_STEPS,
      {
        store: {
          put: "docmap/{vars.id}",
          value: {
            etag: "{vars.etag}",
            docId: "{up.body.document_id}",
            name: "{vars.name}",
            path: "{vars.path}",
            at: "{run.startedAt}",
          },
        },
      },
      {
        store: {
          put: "synced/{vars.id}:{vars.etag|slug}",
          value: { at: "{run.startedAt}" },
        },
      },
    ],
    finally: cursorWrites,
  };
}

// UI preferences, persisted server-side so every device sees the same setup.
export interface SyncPrefs {
  folders?: SyncFolder[];
  collectionId?: string;
  everyMinutes?: number;
  docsOnly?: boolean;
}

const PREFS_KEY = "settings/prefs";

export async function loadPrefs(): Promise<SyncPrefs> {
  return (await storageGet<SyncPrefs>(PREFS_KEY)) ?? {};
}

export async function savePrefs(prefs: SyncPrefs): Promise<void> {
  await storagePut(PREFS_KEY, prefs);
}
