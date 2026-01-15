import { v4 as uuidv4 } from "uuid";
import { API_BASE_URL } from "./config";
import { AppState, Interaction, Stakeholder } from "./models";

const STORAGE_KEY = "stakeholderCrm.v1.state";
let memoryState: AppState | null = null;

function nowIso(): string {
  return new Date().toISOString();
}

function parseIso(timestamp?: string): number {
  if (!timestamp) return 0;
  const parsed = Date.parse(timestamp);
  return Number.isNaN(parsed) ? 0 : parsed;
}

export function getStateTimestamp(state: AppState): number {
  const stateUpdatedAt = parseIso(state.updatedAt);
  const stakeholderTimes = state.stakeholders.map((s) => Math.max(parseIso(s.updatedAt), parseIso(s.createdAt)));
  const interactionTimes = state.interactions.map((i) => parseIso(i.at));
  return Math.max(0, stateUpdatedAt, ...stakeholderTimes, ...interactionTimes);
}

async function pushStateToServer(state: AppState): Promise<void> {
  try {
    await fetch(`${API_BASE_URL}/api/state`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ state }),
    });
  } catch {
    // ignore sync errors (offline or server unavailable)
  }
}

export async function loadStateFromServer(): Promise<AppState | null> {
  try {
    const res = await fetch(`${API_BASE_URL}/api/state`, { method: "GET" });
    if (!res.ok) return null;
    const data = (await res.json()) as { state?: AppState };
    if (!data?.state) return null;
    return {
      stakeholders: data.state.stakeholders ?? [],
      interactions: data.state.interactions ?? [],
      updatedAt: data.state.updatedAt,
    };
  } catch {
    return null;
  }
}

export function loadState(): AppState {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as AppState;
      const normalized = {
        stakeholders: parsed.stakeholders ?? [],
        interactions: parsed.interactions ?? [],
        updatedAt: parsed.updatedAt,
      };
      memoryState = normalized;
      return normalized;
    }
  } catch {
    // ignore storage errors (e.g., blocked in Office webviews)
  }
  return memoryState ?? { stakeholders: [], interactions: [] };
}

export function saveState(state: AppState): AppState {
  const stamped = { ...state, updatedAt: nowIso() };
  memoryState = stamped;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(stamped));
  } catch {
    // ignore storage errors (e.g., blocked in Office webviews)
  }
  void pushStateToServer(stamped);
  return stamped;
}

export function exportState(): string {
  return JSON.stringify(loadState(), null, 2);
}

export function importState(jsonText: string): AppState {
  const parsed = JSON.parse(jsonText) as AppState;
  const normalized: AppState = {
    stakeholders: parsed.stakeholders ?? [],
    interactions: parsed.interactions ?? [],
    updatedAt: parsed.updatedAt,
  };
  return saveState(normalized);
}

export function resetState(): AppState {
  const cleared: AppState = { stakeholders: [], interactions: [], updatedAt: nowIso() };
  return saveState(cleared);
}

export function upsertStakeholderByEmail(
  state: AppState,
  email: string | undefined,
  displayName: string | undefined
): { state: AppState; stakeholder: Stakeholder } {
  const cleanEmail = (email ?? "").trim().toLowerCase();
  const cleanName = (displayName ?? "").trim();

  if (!cleanEmail) {
    // Fall back to creating a "nameless" stakeholder.
    const s: Stakeholder = {
      id: uuidv4(),
      displayName: cleanName || "Unknown Stakeholder",
      tags: [],
      createdAt: nowIso(),
      updatedAt: nowIso(),
      notes: [],
    };
    const newState: AppState = {
      ...state,
      stakeholders: [s, ...state.stakeholders],
    };
    return { state: newState, stakeholder: s };
  }

  const existing = state.stakeholders.find((s) => (s.email ?? "").toLowerCase() === cleanEmail);

  if (existing) {
    const updated: Stakeholder = {
      ...existing,
      displayName: cleanName || existing.displayName,
      updatedAt: nowIso(),
    };
    const stakeholders = state.stakeholders.map((s) => (s.id === existing.id ? updated : s));
    return { state: { ...state, stakeholders }, stakeholder: updated };
  }

  const s: Stakeholder = {
    id: uuidv4(),
    displayName: cleanName || cleanEmail,
    email: cleanEmail,
    tags: [],
    createdAt: nowIso(),
    updatedAt: nowIso(),
    notes: [],
  };
  const newState: AppState = {
    ...state,
    stakeholders: [s, ...state.stakeholders],
  };
  return { state: newState, stakeholder: s };
}

export function addNote(state: AppState, stakeholderId: string, noteText: string): AppState {
  const clean = noteText.trim();
  if (!clean) return state;

  const stakeholders = state.stakeholders.map((s) => {
    if (s.id !== stakeholderId) return s;
    return {
      ...s,
      updatedAt: nowIso(),
      notes: [{ at: nowIso(), text: clean }, ...s.notes],
    };
  });

  const newState = { ...state, stakeholders };
  return saveState(newState);
}

export function addInteraction(state: AppState, interaction: Omit<Interaction, "id">): AppState {
  const full: Interaction = { ...interaction, id: uuidv4() };

  // Update lastInteractionAt for participants.
  const at = full.at;
  const stakeholders = state.stakeholders.map((s) => {
    if (!full.participantIds.includes(s.id)) return s;
    return { ...s, lastInteractionAt: at, updatedAt: nowIso() };
  });

  const newState: AppState = {
    ...state,
    stakeholders,
    interactions: [full, ...state.interactions],
  };

  return saveState(newState);
}
