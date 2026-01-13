import Papa from "papaparse";

export type OutlookEmail = {
  id: string;
  subject?: string;
  from?: { name?: string; address?: string };
  to?: Array<{ name?: string; address?: string }>;
  cc?: Array<{ name?: string; address?: string }>;
  receivedDateTime?: string;
  bodyPreview?: string;
};

export type OutlookEvent = {
  id: string;
  subject?: string;
  start?: string;
  end?: string;
  organizer?: { name?: string; address?: string };
  attendees?: Array<{ name?: string; address?: string }>;
  bodyPreview?: string;
};

function normHeader(h: string) {
  return String(h || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function splitEmails(v: any): string[] {
  const s = String(v || "").trim();
  if (!s) return [];
  // Outlook CSV exports commonly separate recipients with semicolons.
  return s
    .split(/;|,|\n/)
    .map((x) => x.trim())
    .filter(Boolean)
    .map((x) => {
      // Strip quotes and angle brackets if present: "Name <email@x.com>"
      const m = x.match(/<([^>]+)>/);
      return (m?.[1] || x).replace(/^"|"$/g, "");
    });
}

function pick(row: Record<string, any>, headers: Record<string, string>, candidates: string[]): any {
  for (const c of candidates) {
    const key = headers[normHeader(c)];
    if (key && row[key] != null && String(row[key]).trim() !== "") return row[key];
  }
  // also try by normalized equality
  const normalized = Object.keys(row).reduce<Record<string, string>>((acc, k) => {
    acc[normHeader(k)] = k;
    return acc;
  }, {});
  for (const c of candidates) {
    const k = normalized[normHeader(c)];
    if (k && row[k] != null && String(row[k]).trim() !== "") return row[k];
  }
  return undefined;
}

async function parseCsvFile(file: File): Promise<Array<Record<string, any>>> {
  const text = await file.text();
  const parsed = Papa.parse<Record<string, any>>(text, {
    header: true,
    skipEmptyLines: true,
    transformHeader: (h) => String(h || "").trim(),
  });
  if (parsed.errors?.length) {
    // Best-effort: still return rows if we have them.
    // eslint-disable-next-line no-console
    console.warn("CSV parse warnings", parsed.errors);
  }
  return (parsed.data || []).filter((r) => r && Object.keys(r).length > 0);
}

/**
 * Imports an Outlook-exported CSV of Inbox (or another mail folder).
 * Expected columns vary by Outlook version; we use heuristics.
 */
export async function importOutlookEmailCsv(file: File): Promise<OutlookEmail[]> {
  const rows = await parseCsvFile(file);
  if (rows.length === 0) return [];

  // Build a mapping from normalized header -> actual header
  const headerMap = Object.keys(rows[0] || {}).reduce<Record<string, string>>((acc, k) => {
    acc[normHeader(k)] = k;
    return acc;
  }, {});

  return rows.map((row, idx) => {
    const subject = pick(row, headerMap, ["Subject", "subject"]);
    const from = pick(row, headerMap, ["From", "Sender", "From (Name)", "From: (Name)"]); // varies
    const fromEmail = pick(row, headerMap, ["From (Address)", "From: (Address)", "Sender Address", "From Address", "E-mail Address", "Email Address"]);
    const to = pick(row, headerMap, ["To", "To: (Name)", "To (Name)", "To Recipients"]);
    const cc = pick(row, headerMap, ["Cc", "CC", "Cc: (Name)", "Cc (Name)"]);
    const received = pick(row, headerMap, ["Received", "Received Time", "Received Date", "Date/Time Sent", "Sent", "Sent Time"]);
    const body = pick(row, headerMap, ["Body", "Body Preview", "BodyPreview", "Preview"]);

    return {
      id: `email-${idx + 1}`,
      subject: subject ? String(subject) : undefined,
      receivedDateTime: received ? new Date(String(received)).toISOString() : undefined,
      bodyPreview: body ? String(body).slice(0, 500) : undefined,
      from: (fromEmail || from) ? { name: from ? String(from) : undefined, address: fromEmail ? String(fromEmail) : undefined } : undefined,
      to: splitEmails(to).map((address) => ({ address })),
      cc: splitEmails(cc).map((address) => ({ address })),
    } as OutlookEmail;
  });
}

/**
 * Imports an Outlook-exported CSV of Calendar.
 * We read subject/start/end and attendee fields where possible.
 */
export async function importOutlookCalendarCsv(file: File): Promise<OutlookEvent[]> {
  const rows = await parseCsvFile(file);
  if (rows.length === 0) return [];

  const headerMap = Object.keys(rows[0] || {}).reduce<Record<string, string>>((acc, k) => {
    acc[normHeader(k)] = k;
    return acc;
  }, {});

  return rows.map((row, idx) => {
    const subject = pick(row, headerMap, ["Subject", "subject", "Title"]);
    const start = pick(row, headerMap, ["Start Date", "Start", "Start Time", "Start Date/Time", "Begin"]);
    const end = pick(row, headerMap, ["End Date", "End", "End Time", "End Date/Time", "Finish"]);
    const organizer = pick(row, headerMap, ["Organizer", "Meeting Organizer", "From"]);
    const attendees = pick(row, headerMap, ["Required Attendees", "Attendees", "Invitees", "To", "Optional Attendees"]);
    const body = pick(row, headerMap, ["Description", "Body", "Notes"]);

    const startIso = start ? new Date(String(start)).toISOString() : undefined;
    const endIso = end ? new Date(String(end)).toISOString() : undefined;

    return {
      id: `event-${idx + 1}`,
      subject: subject ? String(subject) : undefined,
      start: startIso,
      end: endIso,
      organizer: organizer ? { name: String(organizer), address: splitEmails(organizer)[0] } : undefined,
      attendees: splitEmails(attendees).map((address) => ({ address })),
      bodyPreview: body ? String(body).slice(0, 500) : undefined,
    } as OutlookEvent;
  });
}

export function isExternalEmail(email: string, internalDomains: string[] = ["bcg.com"]): boolean {
  const e = String(email || "").trim().toLowerCase();
  if (!e || !e.includes("@")) return false;
  const domain = e.split("@").pop() || "";
  return !internalDomains.some((d) => domain === d || domain.endsWith(`.${d}`));
}
