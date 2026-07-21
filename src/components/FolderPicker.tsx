import { useCallback, useEffect, useState } from "react";
import { listEntries, type DriveEntry } from "../lib/onedrive";
import { MAX_ROOTS, type SyncFolder } from "../lib/syncTask";

function formatSize(bytes?: number): string {
  if (!bytes) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * Minimal folder browser: breadcrumb navigation + checkbox selection.
 * Selecting a folder covers all its sub-folders (the sync filters Graph's
 * drive-wide delta feed by path prefix). Files are listed too — read-only
 * context, so the browser reflects what the sync will actually pick up.
 */
export default function FolderPicker({
  clientId,
  selected,
  onChange,
  disabled,
}: {
  clientId: string;
  selected: SyncFolder[];
  onChange: (folders: SyncFolder[]) => void;
  disabled: boolean;
}) {
  const [path, setPath] = useState("");
  const [folders, setFolders] = useState<DriveEntry[]>([]);
  const [files, setFiles] = useState<DriveEntry[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async (folderPath: string) => {
    setLoading(true);
    setError(null);
    try {
      const entries = await listEntries(clientId, folderPath);
      setFolders(entries.folders);
      setFiles(entries.files);
      setPath(folderPath);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  }, [clientId]);

  useEffect(() => {
    void load("");
  }, [load]);

  const paths = selected.map((f) => f.path);
  const coveredBy = (folderPath: string) =>
    paths.find((root) => folderPath === root || root === "" || folderPath.startsWith(`${root}/`));

  function toggle(folderPath: string) {
    if (paths.includes(folderPath)) {
      onChange(selected.filter((f) => f.path !== folderPath));
    } else {
      if (selected.length >= MAX_ROOTS) {
        setError(`At most ${MAX_ROOTS} folders — sub-folders are always included.`);
        return;
      }
      onChange([
        ...selected.filter((f) => !f.path.startsWith(`${folderPath}/`)),
        { path: folderPath },
      ]);
    }
  }

  const crumbs = path.split("/").filter(Boolean);
  const childPath = (name: string) => (path ? `${path}/${name}` : name);

  return (
    <div className="border border-[var(--color-line)] bg-[var(--color-card)]">
      <div className="flex flex-wrap items-center gap-2 border-b border-[var(--color-line)] px-4 py-3">
        <span className="font-mono text-xs uppercase tracking-[0.2em] text-[var(--color-muted)]">
          Folders
        </span>
        <button
          onClick={() => void load("")}
          className="font-mono text-xs text-[var(--color-accent)] hover:underline"
        >
          root
        </button>
        {crumbs.map((crumb, i) => (
          <span key={i} className="flex items-center gap-2 font-mono text-xs">
            <span className="text-[var(--color-faint)]">/</span>
            <button
              onClick={() => void load(crumbs.slice(0, i + 1).join("/"))}
              className="hover:text-[var(--color-accent)]"
            >
              {crumb}
            </button>
          </span>
        ))}
      </div>

      {error && <p className="px-4 py-3 font-mono text-xs text-red-400">{error}</p>}
      {loading && (
        <p className="px-4 py-3 font-mono text-xs uppercase tracking-[0.1em] text-[var(--color-faint)]">
          loading…
        </p>
      )}

      {!loading && !error && (
        <ul className="max-h-72 overflow-y-auto">
          {folders.length === 0 && files.length === 0 && (
            <li className="px-4 py-3 font-mono text-xs text-[var(--color-faint)]">
              this folder is empty
            </li>
          )}
          {folders.map((folder) => {
            const full = childPath(folder.name);
            const root = coveredBy(full);
            const checked = root !== undefined;
            const inherited = checked && root !== full;
            return (
              <li
                key={folder.id}
                className="flex items-center gap-3 border-t border-[var(--color-line)] px-4 py-2"
              >
                <input
                  type="checkbox"
                  checked={checked}
                  disabled={disabled || inherited}
                  onChange={() => toggle(full)}
                  className="accent-[var(--color-accent)]"
                  title={inherited ? `covered by "${root || "entire OneDrive"}"` : undefined}
                />
                <button
                  onClick={() => void load(full)}
                  className="min-w-0 flex-1 truncate text-left text-sm hover:text-[var(--color-accent)]"
                >
                  {folder.name}
                  <span className="ml-2 font-mono text-xs text-[var(--color-faint)]">›</span>
                </button>
                {inherited && (
                  <span className="font-mono text-[10px] uppercase tracking-[0.1em] text-[var(--color-faint)]">
                    inherited
                  </span>
                )}
              </li>
            );
          })}
          {files.map((file) => (
            <li
              key={file.id}
              className="flex items-center gap-3 border-t border-[var(--color-line)] px-4 py-2 opacity-60"
            >
              <span className="w-[13px] shrink-0 text-center font-mono text-xs text-[var(--color-faint)]">
                ·
              </span>
              <span className="min-w-0 flex-1 truncate text-sm text-[var(--color-muted)]">
                {file.name}
              </span>
              <span className="shrink-0 font-mono text-[10px] text-[var(--color-faint)]">
                {formatSize(file.size)}
              </span>
            </li>
          ))}
        </ul>
      )}

      {selected.length > 0 && (
        <div className="flex flex-wrap gap-2 border-t border-[var(--color-line)] px-4 py-3">
          {selected.map((folder) => (
            <span
              key={folder.path}
              className="flex items-center gap-2 border border-[var(--color-line)] px-2 py-1 font-mono text-xs"
            >
              {folder.path || "(entire OneDrive)"}
              <button
                onClick={() => toggle(folder.path)}
                disabled={disabled}
                className="text-[var(--color-faint)] hover:text-red-400"
                aria-label={`remove ${folder.path}`}
              >
                ×
              </button>
            </span>
          ))}
        </div>
      )}
    </div>
  );
}
