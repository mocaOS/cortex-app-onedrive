/**
 * Clipboard helper for sandboxed-iframe apps.
 *
 * The async Clipboard API is permission-gated and the app's iframe has an
 * opaque origin — hosts that don't delegate `clipboard-write` (older Cortex
 * launchers, odd embeds) reject `navigator.clipboard.writeText`. Fall back
 * to the legacy execCommand path, which only needs a user gesture. Returns
 * false when both fail so the UI can tell the user to select manually.
 */
export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    let ok = false;
    try {
      ok = document.execCommand("copy");
    } catch {
      ok = false;
    }
    textarea.remove();
    return ok;
  }
}
