# Repo-specific Copilot instructions — Stakeholder CRM

Be concise and prefer small, focused edits. This file highlights patterns, important files, and local workflows that make an AI agent productive quickly.

- **Big picture**: A frontend React taskpane served by Webpack (HTTPS dev server on `https://localhost:3000`) and a local Node API proxy (Express, `http://localhost:3001`) that holds the OpenAI API key. See `package.json` scripts (`npm run dev` runs both). Key UX targets: standalone web, OneNote (web) and PowerPoint (desktop/web) taskpanes.

- **Run / build**: Use `npm install`, copy `.env.example` → `.env` and set `OPENAI_API_KEY`, then `npm run dev`. Production build: `npm run build`. API dev uses `ts-node-dev` (`server/index.ts`). See `webpack.config.js` for HTTPS dev-certs and static output `dist/`.

- **Where to change runtime config**: Frontend base URL is in `src/shared/config.ts` (`API_BASE_URL` reads `window.__API_BASE_URL__` or falls back to `https://localhost:3001`). Server-side OpenAI config is in `server/index.ts` (env: `OPENAI_API_KEY`, optional `OPENAI_MODEL`). Never move the key to client-side code.

- **Office integration points**:
  - Host detection: `src/shared/host.ts` (`detectHost()`), used throughout to toggle OneNote/PowerPoint flows.
  - PowerPoint capture: `src/shared/powerpoint.ts` uses `PowerPoint.run(...)` and expects Office.js API shape — edits here must respect async `context.sync()` patterns.
  - OneNote capture: `src/shared/onenote.ts` (used by `taskpane.tsx`).
  - Manifests live in `manifest/*.xml` and are copied into `dist/` by Webpack (`CopyWebpackPlugin`).

- **Data flow & persistence**:
  - In-memory + browser persistence: `src/shared/storage.ts` implements `loadState()`, `saveState()`, `addInteraction()`, `upsertStakeholderByEmail()` and persists to `localStorage` under key `stakeholderCrm.v1.state`.
  - Frontend uses storage helpers that return new AppState instances; prefer using those helpers rather than mutating state manually.

- **Outlook ingestion**: No Graph/OAuth. Users export CSVs and the frontend imports them via `src/shared/outlookExport.ts`. That module uses heuristics (`normHeader`, `pick`, `splitEmails`) to map diverse CSV headers. Default internal domain is `bcg.com` in `isExternalEmail()` — be cautious when adjusting internal/external rules.

- **OpenAI usage**: The server endpoint `POST /api/openai/stakeholder-impact` (in `server/index.ts`) constructs a chat-style prompt and uses `openai.chat.completions.create(...)` with `response_format: { type: "json_object" }`. The frontend calls it from `src/taskpane/taskpane.tsx` via `API_BASE_URL + '/api/openai/stakeholder-impact'`. Keep prompt and parsing changes server-side.

- **Patterns & conventions**:
  - Single source for runtime checks: `detectHost()` and `API_BASE_URL` control branching and endpoints.
  - Small pure helpers: Most logic for storage/parsing is in `src/shared/*` modules. Prefer adding helpers there and keeping `taskpane.tsx` focused on UI flows.
  - No tests present: avoid adding test scaffolding unless requested; keep changes minimal and well-scoped.

- **How to add a new API route**:
  1. Add Express route to `server/index.ts`.
  2. Use `ts-node-dev` in dev (`npm run dev-api`) — server restarts on change.
  3. Update `src/shared/config.ts` or `window.__API_BASE_URL__` if a different base is required for the frontend.

- **Files to inspect for common tasks**:
  - Frontend entry: `src/taskpane/taskpane.tsx` (UI + host workflows)
  - CSV import: `src/shared/outlookExport.ts`
  - Storage model: `src/shared/storage.ts` and `src/shared/models.ts`
  - PowerPoint/OneNote: `src/shared/powerpoint.ts`, `src/shared/onenote.ts`
  - Local API: `server/index.ts`
  - Dev server config: `webpack.config.js`

- **Quick examples**:
  - To debug OpenAI server errors: run `npm run dev`, then check `http://localhost:3001/health` and server console logs (`server/index.ts` prints startup and warnings about missing `OPENAI_API_KEY`).
  - To test PowerPoint capture locally: run the dev server with dev certs (`npm run dev`), sideload `manifest/manifest-powerpoint.xml` into PowerPoint, open the taskpane, and use the `Read current slide` flow.

- **Do not**:
  - Put the OpenAI key in client code or check it into the repo.
  - Assume CSV headers — rely on `outlookExport` heuristics and add tolerant parsing there if needed.

If anything above is unclear or you want additional examples (CI, packaging, or a CONTRIBUTING section), tell me which area to expand. I'll iterate quickly.
