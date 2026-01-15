export type HostApp = "Web" | "Outlook" | "OneNote" | "PowerPoint" | "Unknown";

export type InteractionType = "email" | "meeting" | "note";

export interface Stakeholder {
  id: string;
  displayName: string;
  email?: string;
  // Optional fields you can extend later.
  title?: string;
  company?: string;
  tags: string[];
  createdAt: string; // ISO
  updatedAt: string; // ISO
  lastInteractionAt?: string; // ISO
  notes: Array<{ at: string; text: string }>;
}

export interface Interaction {
  id: string;
  type: InteractionType;
  at: string; // ISO
  title: string; // e.g., subject/page title
  summary?: string;
  participantIds: string[];
  source: {
    host: HostApp;
    outlookItemId?: string;
    onenotePageId?: string;
    powerpointSlideId?: string;
    powerpointShapeIds?: string[];
    webContextId?: string;
  };
  suggestedNextActions?: string[];
}

export interface AppState {
  stakeholders: Stakeholder[];
  interactions: Interaction[];
  updatedAt?: string;
}
