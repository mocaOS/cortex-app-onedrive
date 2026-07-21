import { useEffect, useRef, useState } from "react";
import SyncPanel from "./components/SyncPanel";
import { copyText } from "./lib/copy";
import {
  beginConnect,
  currentAccount,
  disconnect,
  getClientId,
  isConnected,
  pollConnect,
  type DeviceCode,
} from "./lib/onedrive";

type Status =
  | { kind: "checking" }
  | { kind: "unconfigured" }
  | { kind: "disconnected"; clientId: string }
  | { kind: "error"; message: string }
  | { kind: "ready"; clientId: string; account: string };

/** Device-code connect flow — made for exactly this situation: no redirect
 * server, no popups. Show a short code, the user enters it on
 * microsoft.com/devicelogin in another tab, and we poll until approved. */
function ConnectFlow({
  clientId,
  onConnected,
}: {
  clientId: string;
  onConnected: () => void;
}) {
  const [device, setDevice] = useState<DeviceCode | null>(null);
  const [busy, setBusy] = useState(false);
  const [copied, setCopied] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const pollRef = useRef<number | null>(null);

  async function begin() {
    setBusy(true);
    setError(null);
    try {
      setDevice(await beginConnect(clientId));
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  // poll while a device code is pending
  useEffect(() => {
    if (!device) return;
    let cancelled = false;
    const tick = async () => {
      if (cancelled) return;
      try {
        if (await pollConnect(clientId, device)) {
          onConnected();
          return;
        }
        pollRef.current = window.setTimeout(tick, device.intervalMs);
      } catch (err) {
        if (!cancelled) {
          setError(err instanceof Error ? err.message : String(err));
          setDevice(null);
        }
      }
    };
    pollRef.current = window.setTimeout(tick, device.intervalMs);
    return () => {
      cancelled = true;
      if (pollRef.current) window.clearTimeout(pollRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [device]);

  return (
    <div className="border border-[var(--color-line)] bg-[var(--color-card)] p-6">
      <h2 className="font-mono text-xs font-semibold uppercase tracking-[0.2em] text-[var(--color-accent)]">
        Connect your OneDrive
      </h2>
      <ol className="mt-4 max-w-xl list-decimal space-y-3 pl-5 text-sm leading-relaxed text-[var(--color-muted)]">
        <li>Generate a sign-in code below.</li>
        <li>
          Open{" "}
          <span className="text-[var(--color-fg)]">microsoft.com/devicelogin</span>{" "}
          in a new browser tab and enter the code (this sandboxed app cannot
          open tabs itself).
        </li>
        <li>
          Sign in and approve <em>Files.Read</em> — this page connects by
          itself once you have.
        </li>
      </ol>

      <div className="mt-5 space-y-4">
        {!device ? (
          <button
            onClick={() => void begin()}
            disabled={busy}
            className="border border-[var(--color-accent)] px-5 py-2 font-mono text-xs font-semibold uppercase tracking-[0.1em] text-[var(--color-accent)] transition-colors hover:bg-[var(--color-accent)] hover:text-[var(--color-bg)] disabled:opacity-40"
          >
            Generate sign-in code
          </button>
        ) : (
          <div className="flex flex-wrap items-center gap-4">
            <span className="border border-[var(--color-line)] bg-[var(--color-bg)] px-4 py-2 font-mono text-xl tracking-[0.3em] text-[var(--color-fg)]">
              {device.userCode}
            </span>
            <button
              onClick={() => {
                void copyText(device.userCode).then((ok) => {
                  setCopied(ok);
                  setTimeout(() => setCopied(false), 2000);
                });
              }}
              className="border border-[var(--color-accent)] px-4 py-2 font-mono text-xs font-semibold uppercase tracking-[0.1em] text-[var(--color-accent)] transition-colors hover:bg-[var(--color-accent)] hover:text-[var(--color-bg)]"
            >
              {copied ? "Copied" : "Copy code"}
            </button>
            <span className="font-mono text-xs uppercase tracking-[0.1em] text-[var(--color-faint)]">
              waiting for approval…
            </span>
          </div>
        )}
        {error && <p className="font-mono text-xs text-red-400">{error}</p>}
      </div>
    </div>
  );
}

export default function App() {
  const [status, setStatus] = useState<Status>({ kind: "checking" });

  async function check() {
    setStatus({ kind: "checking" });
    try {
      const clientId = await getClientId();
      if (!clientId) {
        setStatus({ kind: "unconfigured" });
        return;
      }
      if (!(await isConnected())) {
        setStatus({ kind: "disconnected", clientId });
        return;
      }
      const account = await currentAccount(clientId);
      setStatus({ kind: "ready", clientId, account });
    } catch (err) {
      setStatus({
        kind: "error",
        message: err instanceof Error ? err.message : String(err),
      });
    }
  }

  useEffect(() => {
    void check();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div className="mx-auto max-w-4xl px-6 py-12">
      <header className="mb-10">
        <p className="font-mono text-xs font-semibold uppercase tracking-[0.2em] text-[var(--color-accent)]">
          Cortex App
        </p>
        <h1 className="mt-3 text-5xl font-bold tracking-[-0.04em] md:text-6xl">
          OneDrive Sync
        </h1>
        <p className="mt-4 max-w-xl text-[var(--color-muted)]">
          Pick folders in your OneDrive and Cortex keeps them in the knowledge
          graph: new and edited files flow in through Microsoft Graph&apos;s
          change feed on a schedule &mdash; server-side, even with this tab
          closed.
        </p>
      </header>

      <div className="space-y-6">
        {status.kind === "checking" && (
          <p className="font-mono text-xs uppercase tracking-[0.1em] text-[var(--color-faint)]">
            checking connection…
          </p>
        )}

        {status.kind === "unconfigured" && (
          <div className="border border-[var(--color-line)] bg-[var(--color-card)] p-6">
            <h2 className="font-mono text-xs font-semibold uppercase tracking-[0.2em] text-[var(--color-accent)]">
              Setup required
            </h2>
            <ol className="mt-3 max-w-xl list-decimal space-y-3 pl-5 text-sm leading-relaxed text-[var(--color-muted)]">
              <li>
                Create a free app registration at{" "}
                <code className="font-mono text-xs">portal.azure.com</code>{" "}
                (App registrations → New). Under supported account types,
                include <em>personal Microsoft accounts</em>.{" "}
                <span className="text-[var(--color-fg)]">
                  Registrations must live in a directory:
                </span>{" "}
                sign in with a work/Microsoft&nbsp;365 account that has one, or
                give a personal account its own via the free Azure signup — a
                bare personal account hits &quot;does not exist in tenant&quot;
                errors. Which directory hosts the registration doesn&apos;t
                matter: any OneDrive account can connect to it.
              </li>
              <li>
                <span className="text-[var(--color-fg)]">
                  Under Authentication, enable &quot;Allow public client
                  flows&quot;
                </span>{" "}
                — the device-code sign-in needs it, and no client secret is
                required anywhere.
              </li>
              <li>
                Paste the Application (client) ID into{" "}
                <code className="font-mono text-xs">ENTRA_CLIENT_ID</code> under{" "}
                <span className="text-[var(--color-fg)]">
                  Settings → Apps → OneDrive Sync → Configure
                </span>
                .
              </li>
            </ol>
            <button
              onClick={() => void check()}
              className="mt-5 border border-[var(--color-line)] px-5 py-2 font-mono text-xs font-semibold uppercase tracking-[0.1em] transition-colors hover:border-[var(--color-accent)] hover:text-[var(--color-accent)]"
            >
              Re-check
            </button>
          </div>
        )}

        {status.kind === "disconnected" && (
          <ConnectFlow clientId={status.clientId} onConnected={() => void check()} />
        )}

        {status.kind === "error" && (
          <div className="border border-[var(--color-line)] bg-[var(--color-card)] p-6">
            <p className="font-mono text-xs text-red-400">{status.message}</p>
            <button
              onClick={() => void check()}
              className="mt-5 border border-[var(--color-line)] px-5 py-2 font-mono text-xs font-semibold uppercase tracking-[0.1em] transition-colors hover:border-[var(--color-accent)] hover:text-[var(--color-accent)]"
            >
              Retry
            </button>
          </div>
        )}

        {status.kind === "ready" && (
          <>
            <div className="flex flex-wrap items-baseline gap-x-6 gap-y-2 border border-[var(--color-line)] bg-[var(--color-card)] p-4">
              <span className="font-mono text-xs uppercase tracking-[0.2em] text-[var(--color-muted)]">
                Connected
              </span>
              <span className="font-mono text-sm">{status.account}</span>
              <button
                onClick={() => {
                  void disconnect().then(() => void check());
                }}
                className="ml-auto font-mono text-xs text-[var(--color-faint)] transition-colors hover:text-red-400"
              >
                disconnect
              </button>
            </div>
            <SyncPanel clientId={status.clientId} />
          </>
        )}
      </div>

      <footer className="mt-10 font-mono text-xs text-[var(--color-faint)]">
        POWERED BY CORTEX — SYNC RUNS SERVER-SIDE, EVEN WITH THIS TAB CLOSED
      </footer>
    </div>
  );
}
