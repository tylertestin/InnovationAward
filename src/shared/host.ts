import { HostApp } from "./models";

/**
 * Detect whether we are running:
 * - inside an Office host (OneNote / PowerPoint / Outlook), or
 * - as a standalone web app.
 */
export function detectHost(): HostApp {
  try {
    const host = (globalThis as any).Office?.context?.host;
    if (!host) return "Web";
    if (host === (globalThis as any).Office.HostType.Outlook) return "Outlook";
    if (host === (globalThis as any).Office.HostType.OneNote) return "OneNote";
    if (host === (globalThis as any).Office.HostType.PowerPoint) return "PowerPoint";
  } catch {
    // ignore
  }
  return "Web";
}
