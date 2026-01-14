import { v4 as uuidv4 } from "uuid";
import { AppState, Interaction, Stakeholder } from "./models";

const STORAGE_KEY = "stakeholderCrm.v1.state";
let memoryState: AppState | null = null;

function nowIso(): string {
  return new Date().toISOString();
}

export function loadState(): AppState {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as AppState;
      const normalized = {
        stakeholders: parsed.stakeholders ?? [],
        interactions: parsed.interactions ?? [],
      };
      memoryState = normalized;
      return normalized;
    }
  } catch {
    // ignore storage errors (e.g., blocked in Office webviews)
  }
  return memoryState ?? { stakeholders: [], interactions: [] };
}

export function saveState(state: AppState): void {
  memoryState = state;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    // ignore storage errors (e.g., blocked in Office webviews)
  }
}

export function exportState(): string {
  return JSON.stringify(loadState(), null, 2);
}

export function importState(jsonText: string): AppState {
  const parsed = JSON.parse(jsonText) as AppState;
  const normalized: AppState = {
    stakeholders: parsed.stakeholders ?? [],
    interactions: parsed.interactions ?? [],
  };
  saveState(normalized);
  return normalized;
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
  saveState(newState);
  return newState;
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
    stakeholders,
    interactions: [full, ...state.interactions],
  };

  saveState(newState);
  return newState;
}
