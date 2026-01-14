import { HostApp } from "./models";

/**
 * Detect whether we are running:
 * - inside an Office host (OneNote / PowerPoint / Outlook), or
 * - as a standalone web app.
 */
export function detectHost(): HostApp {
  try {
    const office = (globalThis as any).Office;
    const host = office?.context?.host;
    if (host) {
      if (host === office.HostType.Outlook) return "Outlook";
      if (host === office.HostType.OneNote) return "OneNote";
      if (host === office.HostType.PowerPoint) return "PowerPoint";
    }

    if ((globalThis as any).PowerPoint?.run) return "PowerPoint";
    if ((globalThis as any).OneNote?.run) return "OneNote";

    const params = new URLSearchParams(globalThis.location?.search ?? "");
    const hintedHost = (params.get("host") || params.get("officeHost") || params.get("app") || "").toLowerCase();
    if (hintedHost.includes("powerpoint")) return "PowerPoint";
    if (hintedHost.includes("onenote")) return "OneNote";
    if (hintedHost.includes("outlook")) return "Outlook";
  } catch {
    // ignore
  }
  return "Web";
}
