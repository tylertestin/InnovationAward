import * as React from "react";
import { createRoot } from "react-dom/client";
import "./taskpane.css";

import { detectHost } from "../shared/host";
import { AppState, HostApp, Stakeholder } from "../shared/models";
import {
  addInteraction,
  addNote,
  exportState,
  importState,
  loadState,
  saveState,
  upsertStakeholderByEmail,
} from "../shared/storage";
import { captureFromOneNotePage } from "../shared/onenote";
import { captureFromPowerPointSlide, StakeholderImpact } from "../shared/powerpoint";
import { API_BASE_URL } from "../shared/config";
import {
  OutlookEmail,
  OutlookEvent,
  importOutlookCalendarCsv,
  importOutlookEmailCsv,
  isExternalEmail,
} from "../shared/outlookExport";

function sortStakeholders(state: AppState) {
  return [...state.stakeholders].sort((a, b) => {
    const at = a.lastInteractionAt ?? a.updatedAt;
    const bt = b.lastInteractionAt ?? b.updatedAt;
    return bt.localeCompare(at);
  });
}

type ImpactRow = StakeholderImpact & { rationale?: string };

async function readFileAsText(file: File): Promise<string> {
  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error);
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.readAsText(file);
  });
}

async function analyzeStakeholderImpact(params: {
  stakeholders: Stakeholder[];
  slideText: string;
  emails: OutlookEmail[];
}): Promise<{
  impacts: Array<{ stakeholderId: string; reaction: "green" | "red"; rationale?: string }>;
}> {
  const res = await fetch(`${API_BASE_URL}/api/openai/stakeholder-impact`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(params),
  });
  if (!res.ok) throw new Error(`API error ${res.status}: ${await res.text().catch(() => "")}`);
  return (await res.json()) as any;
}

function App() {
  const [host] = React.useState<HostApp>(() => detectHost());
  const [state, setState] = React.useState<AppState>(() => loadState());
  const [selectedStakeholderId, setSelectedStakeholderId] = React.useState<string | null>(null);

  // Outlook (email + calendar) imported from user-exported CSV files.
  const [inbox, setInbox] = React.useState<OutlookEmail[]>([]);
  const [events, setEvents] = React.useState<OutlookEvent[]>([]);
  const [outlookError, setOutlookError] = React.useState<string | null>(null);

  // PowerPoint
  const [slideText, setSlideText] = React.useState<string>("");
  const [impacts, setImpacts] = React.useState<ImpactRow[]>([]);
  const [pptError, setPptError] = React.useState<string | null>(null);

  const stakeholdersSorted = React.useMemo(() => sortStakeholders(state), [state]);

  React.useEffect(() => {
    saveState(state);
  }, [state]);

  const selected = React.useMemo(
    () => state.stakeholders.find((s) => s.id === selectedStakeholderId) ?? null,
    [state, selectedStakeholderId]
  );

  async function captureFromOneNote() {
    const capture = await captureFromOneNotePage();

    let working = state;

    // Treat OneNote-captured emails as potential stakeholders; keep external only (client).
    const participantIds: string[] = [];
    for (const email of capture.extractedEmails || []) {
      if (!isExternalEmail(email)) continue;
      const { state: s2, stakeholder } = upsertStakeholderByEmail(working, email, email);
      working = s2;
      participantIds.push(stakeholder.id);
    }

    const summary =
      (capture.extractedTextSample || "").trim() ||
      `Captured ${participantIds.length} external emails from OneNote page.`;

    working = addInteraction(working, {
      type: "note",
      at: new Date().toISOString(),
      title: capture.title || "OneNote page",
      summary,
      participantIds,
      source: { host, onenotePageId: capture.pageId },
      suggestedNextActions: ["Capture decisions", "Assign owners", "Schedule follow-ups"],
    });

    setState(working);
  }

  async function importOutlookEmailExport(file: File) {
    try {
      setOutlookError(null);
      const emails = await importOutlookEmailCsv(file);
      setInbox(emails);

      let working = state;

      for (const e of emails.slice(0, 200)) {
        const participantEmails = [
          e.from?.address,
          ...(e.to || []).map((x) => x.address),
          ...(e.cc || []).map((x) => x.address),
        ]
          .filter(Boolean)
          .map((x) => String(x));

        const participantIds: string[] = [];
        for (const em of participantEmails) {
          if (!isExternalEmail(em)) continue;
          const { state: s2, stakeholder } = upsertStakeholderByEmail(working, em, em);
          working = s2;
          participantIds.push(stakeholder.id);
        }

        if (!e.subject && !e.bodyPreview && participantIds.length === 0) continue;

        working = addInteraction(working, {
          type: "email",
          at: e.receivedDateTime || new Date().toISOString(),
          title: e.subject || "Email",
          summary: e.bodyPreview,
          participantIds,
          source: { host, outlookItemId: e.id },
          suggestedNextActions: ["Send follow-up", "Confirm next steps", "Schedule check-in"],
        });
      }

      working = addInteraction(working, {
        type: "note",
        at: new Date().toISOString(),
        title: "Outlook Inbox export imported",
        summary: `Imported ${emails.length} emails from CSV.`,
        participantIds: [],
        source: { host },
      });

      setState(working);
    } catch (e: any) {
      setOutlookError(String(e?.message || e));
    }
  }

  async function importOutlookCalendarExport(file: File) {
    try {
      setOutlookError(null);
      const evs = await importOutlookCalendarCsv(file);
      setEvents(evs);

      let working = state;
      for (const ev of evs.slice(0, 200)) {
        const participantEmails = [ev.organizer?.address, ...(ev.attendees || []).map((x) => x.address)]
          .filter(Boolean)
          .map((x) => String(x));

        const participantIds: string[] = [];
        for (const em of participantEmails) {
          if (!isExternalEmail(em)) continue;
          const { state: s2, stakeholder } = upsertStakeholderByEmail(working, em, em);
          working = s2;
          participantIds.push(stakeholder.id);
        }

        if (!ev.subject && !ev.bodyPreview && participantIds.length === 0) continue;

        working = addInteraction(working, {
          type: "meeting",
          at: ev.start || new Date().toISOString(),
          title: ev.subject || "Meeting",
          summary: ev.bodyPreview,
          participantIds,
          source: { host, outlookItemId: ev.id },
          suggestedNextActions: ["Confirm decisions", "Send recap", "Assign owners"],
        });
      }

      working = addInteraction(working, {
        type: "note",
        at: new Date().toISOString(),
        title: "Outlook Calendar export imported",
        summary: `Imported ${evs.length} events from CSV.`,
        participantIds: [],
        source: { host },
      });

      setState(working);
    } catch (e: any) {
      setOutlookError(String(e?.message || e));
    }
  }

  async function captureFromPowerPoint() {
    try {
      setPptError(null);
      const cap = await captureFromPowerPointSlide();
      setSlideText(cap.slideText);

      const working = addInteraction(state, {
        type: "note",
        at: new Date().toISOString(),
        title: "PowerPoint slide captured",
        summary: cap.slideText.slice(0, 500),
        participantIds: [],
        source: { host, powerpointSlideId: cap.slideId },
        suggestedNextActions: ["Validate stakeholder reactions", "Adjust storyline", "Prep pre-wires"],
      });

      setState(working);
    } catch (e: any) {
      setPptError(String(e?.message || e));
    }
  }

  async function runImpactAnalysisAndRender() {
    try {
      setPptError(null);
      if (!slideText.trim()) throw new Error("No slide text captured yet. Click 'Read current slide' first.");

      const result = await analyzeStakeholderImpact({
        stakeholders: state.stakeholders,
        slideText,
        emails: inbox,
      });

      const rows: ImpactRow[] = (result.impacts ?? [])
        .map((r: any) => {
          const s = state.stakeholders.find((x) => x.id === r.stakeholderId);
          if (!s) return null;
          return {
            stakeholderId: s.id,
            displayName: s.displayName,
            email: s.email,
            reaction: r.reaction === "red" ? "red" : "green",
            rationale: r.rationale,
          } as ImpactRow;
        })
        .filter(Boolean) as any;

      setImpacts(rows);
    } catch (e: any) {
      setPptError(String(e?.message || e));
    }
  }

  function onAddNote(text: string) {
    if (!selected) return;
    const newState = addNote(state, selected.id, text);
    setState(newState);
  }

  async function handleImportJsonFile(file: File) {
    try {
      const txt = await readFileAsText(file);
      const imported = importState(txt);
      if (imported) setState(imported);
      else alert("Import failed (invalid JSON).");
    } catch (e: any) {
      alert(`Import failed: ${String(e?.message || e)}`);
    }
  }

  const hostLabel =
    host === "Web" ? "Web app" : host === "OneNote" ? "OneNote" : host === "PowerPoint" ? "PowerPoint" : host;

  return (
    <div className="container">
      <header className="header">
        <div>
          <div className="title">Stakeholder CRM</div>
          <div className="subtitle">Running in: {hostLabel}</div>
        </div>
        <div className="headerActions">
          <button
            className="btn"
            onClick={() => {
              const blob = exportState(state);
              const url = URL.createObjectURL(blob);
              const a = document.createElement("a");
              a.href = url;
              a.download = "stakeholder-crm.json";
              a.click();
              URL.revokeObjectURL(url);
            }}
          >
            Export
          </button>

          <label className="btn btnSecondary">
            Import
            <input
              type="file"
              accept="application/json"
              style={{ display: "none" }}
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) handleImportJsonFile(f);
                // reset so selecting the same file twice triggers onChange
                e.currentTarget.value = "";
              }}
            />
          </label>
        </div>
      </header>

      <section className="controls">
        {host === "OneNote" && (
          <button className="btnPrimary" onClick={captureFromOneNote}>
            Capture from OneNote page
          </button>
        )}

        {host === "PowerPoint" && (
          <div className="row">
            <button className="btnPrimary" onClick={captureFromPowerPoint}>
              Read current slide
            </button>
            <button className="btn" onClick={runImpactAnalysisAndRender}>
              Analyze stakeholder impact
            </button>
          </div>
        )}

        {(host === "Web" || host === "PowerPoint" || host === "OneNote") && (
          <div className="row">
            <label className="btn">
              Import Inbox CSV
              <input
                type="file"
                accept=".csv,text/csv"
                style={{ display: "none" }}
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) importOutlookEmailExport(f);
                  e.currentTarget.value = "";
                }}
              />
            </label>

            <label className="btn">
              Import Calendar CSV
              <input
                type="file"
                accept=".csv,text/csv"
                style={{ display: "none" }}
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) importOutlookCalendarExport(f);
                  e.currentTarget.value = "";
                }}
              />
            </label>

            {(inbox.length > 0 || events.length > 0) && (
              <span className="muted">
                {inbox.length} emails, {events.length} events loaded
              </span>
            )}
          </div>
        )}

        {outlookError && <div className="error">Outlook import: {outlookError}</div>}
        {pptError && <div className="error">PowerPoint error: {pptError}</div>}
      </section>

      {host === "PowerPoint" && slideText.trim().length > 0 && (
        <section className="panel">
          <div className="panelTitle">Current slide text (preview)</div>
          <div className="monoBox">
            {slideText.slice(0, 1200)}
            {slideText.length > 1200 ? "…" : ""}
          </div>

          {impacts.length > 0 && (
            <div className="panel">
              <div className="panelTitle">Stakeholders impacted</div>
              {impacts.map((r) => (
                <div key={r.stakeholderId} className="impactRow">
                  <span className={r.reaction === "red" ? "pill pillRed" : "pill pillGreen"}>
                    {r.reaction.toUpperCase()}
                  </span>
                  <span className="impactName">{r.displayName}</span>
                  <span className="muted">{r.rationale ?? ""}</span>
                </div>
              ))}
            </div>
          )}
        </section>
      )}

      <main className="main">
        <section className="left">
          <div className="sectionTitle">Stakeholders</div>
          <div className="list">
            {stakeholdersSorted.map((s) => (
              <button
                key={s.id}
                className={s.id === selectedStakeholderId ? "listItem selected" : "listItem"}
                onClick={() => setSelectedStakeholderId(s.id)}
              >
                <div className="listItemTitle">{s.displayName}</div>
                <div className="listItemMeta">{s.email ?? ""}</div>
                <div className="listItemMeta muted">Last touch: {(s.lastInteractionAt ?? s.updatedAt).slice(0, 10)}</div>
              </button>
            ))}
            {stakeholdersSorted.length === 0 && (
              <div className="muted">No stakeholders yet. Capture from OneNote, or import CSV/JSON.</div>
            )}
          </div>
        </section>

        <section className="right">
          <div className="sectionTitle">Details</div>
          {selected ? (
            <div className="details">
              <div className="detailsHeader">
                <div className="detailsName">{selected.displayName}</div>
                <div className="muted">{selected.email ?? ""}</div>
              </div>

              <div className="detailsBlock">
                <div className="detailsLabel">Notes</div>
                <ul className="notes">
                  {selected.notes.map((n, idx) => (
                    <li key={idx}>
                      <span className="muted">{n.at.slice(0, 10)}:</span> {n.text}
                    </li>
                  ))}
                  {selected.notes.length === 0 && <li className="muted">No notes yet.</li>}
                </ul>
                <AddNote onAdd={onAddNote} />
              </div>
            </div>
          ) : (
            <div className="muted">Select a stakeholder to view details and add notes.</div>
          )}
        </section>
      </main>

      <section className="panel">
        <div className="panelTitle">Recent interactions</div>
        <div className="interactions">
          {state.interactions.slice(0, 15).map((i) => (
            <div key={i.id} className="interaction">
              <div className="interactionTop">
                <span className="pill">{i.type}</span>
                <span className="muted">{i.at.slice(0, 16).replace("T", " ")}</span>
                <span className="interactionTitle">{i.title}</span>
              </div>
              {i.summary && <div className="muted">{i.summary}</div>}
            </div>
          ))}
          {state.interactions.length === 0 && <div className="muted">No interactions logged yet.</div>}
        </div>
      </section>

      <footer className="footer muted">
        MVP: Web + OneNote + PowerPoint. Outlook data is ingested via user-exported CSV (Inbox + Calendar). Slide impact uses the local OpenAI-backed API.
      </footer>
    </div>
  );
}

function AddNote({ onAdd }: { onAdd: (text: string) => void }) {
  const [text, setText] = React.useState("");
  return (
    <div className="addNote">
      <input className="input" placeholder="Add a quick note..." value={text} onChange={(e) => setText(e.target.value)} />
      <button
        className="btn"
        onClick={() => {
          const t = text.trim();
          if (!t) return;
          onAdd(t);
          setText("");
        }}
      >
        Add
      </button>
    </div>
  );
}

const root = createRoot(document.getElementById("root")!);
root.render(<App />);
