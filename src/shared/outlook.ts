import { HostApp } from "./models";

export type OutlookParticipant = { email?: string; name?: string };

function getAsyncProm<T>(getter: any): Promise<T | undefined> {
  return new Promise((resolve) => {
    try {
      getter((result: Office.AsyncResult<T>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else resolve(undefined);
      });
    } catch {
      resolve(undefined);
    }
  });
}

async function getSubject(item: any): Promise<string> {
  try {
    if (typeof item.subject === "string") return item.subject;
    if (item.subject?.getAsync) {
      const v = await getAsyncProm<string>(item.subject.getAsync.bind(item.subject));
      return v ?? "(no subject)";
    }
  } catch {}
  return "(no subject)";
}

async function getAttendeesList(prop: any): Promise<any[]> {
  try {
    if (Array.isArray(prop)) return prop;
    if (prop?.getAsync) {
      const v = await getAsyncProm<any[]>(prop.getAsync.bind(prop));
      return v ?? [];
    }
  } catch {}
  return [];
}

function normalizeEmail(s: string | undefined): string | undefined {
  const e = (s ?? "").trim();
  return e ? e.toLowerCase() : undefined;
}

export async function captureFromOutlookItem(): Promise<{
  host: HostApp;
  itemType: "message" | "appointment" | "unknown";
  itemId?: string;
  title: string;
  participants: OutlookParticipant[];
}> {
  const mailbox = (Office as any)?.context?.mailbox;
  const item = mailbox?.item;
  if (!item) {
    return {
      host: "Outlook",
      itemType: "unknown",
      title: "No item available",
      participants: [],
    };
  }

  const itemType = (item.itemType as any) || "unknown";
  const title = await getSubject(item);

  const participants: OutlookParticipant[] = [];

  if (itemType === "message") {
    const from = item.from;
    if (from) participants.push({ email: normalizeEmail(from.emailAddress), name: from.displayName });

    const toList = await getAttendeesList(item.to);
    toList.forEach((p: any) =>
      participants.push({ email: normalizeEmail(p.emailAddress), name: p.displayName })
    );

    const ccList = await getAttendeesList(item.cc);
    ccList.forEach((p: any) =>
      participants.push({ email: normalizeEmail(p.emailAddress), name: p.displayName })
    );
  } else if (itemType === "appointment") {
    const organizer = item.organizer;
    if (organizer)
      participants.push({ email: normalizeEmail(organizer.emailAddress), name: organizer.displayName });

    const req = await getAttendeesList(item.requiredAttendees);
    req.forEach((p: any) =>
      participants.push({ email: normalizeEmail(p.emailAddress), name: p.displayName })
    );

    const opt = await getAttendeesList(item.optionalAttendees);
    opt.forEach((p: any) =>
      participants.push({ email: normalizeEmail(p.emailAddress), name: p.displayName })
    );
  }

  // Deduplicate by email where possible.
  const seen = new Set<string>();
  const deduped: OutlookParticipant[] = [];
  for (const p of participants) {
    const key = p.email ?? `${p.name ?? "unknown"}-${Math.random()}`;
    if (p.email) {
      if (seen.has(p.email)) continue;
      seen.add(p.email);
    }
    deduped.push(p);
  }

  const itemId = item.itemId;

  return {
    host: "Outlook",
    itemType,
    itemId,
    title,
    participants: deduped,
  };
}
