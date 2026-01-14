# CaseForce — Web + OneNote + PowerPoint

This repo is a **consultant-grade stakeholder CRM** focused on day-to-day relationship management, branded as CaseForce:
- Runs as a **standalone web app** (`https://localhost:3000/taskpane.html`)
- Runs as an **Office taskpane** in **OneNote (web)** and **PowerPoint (desktop/web)** via sideloaded manifests
- Imports **email + calendar** data from **Outlook-exported CSV files** (Inbox export + Calendar export).
- Uses **OpenAI** (via a local server proxy) to predict red/green reactions to the current PowerPoint slide (displayed in the taskpane).

> Note: This repo does **not** rely on an Outlook add-in (many enterprise tenants disable sideloading).
> Outlook ingestion is handled via user exports to CSV to avoid tenant app-registration requirements.

## Prerequisites

- Node.js (LTS) + npm
- PowerPoint (desktop or web) and/or OneNote (web)
- An OpenAI API key (kept server-side)

## Install

```bash
npm install
cp .env.example .env
# edit .env and set OPENAI_API_KEY
# optional: adjust API_PORT (defaults to 3001) and WEB_PORT (defaults to 3000)
npm run dev
```

This starts:
- Webpack dev server: `https://localhost:3000` (or `WEB_PORT` if set)
- Local API server: `http://localhost:3001` (or `API_PORT` if set)

## Export from Outlook (temporary ingestion path)

You will create **two CSV exports**:
1) **Inbox export** (emails)
2) **Calendar export** (events)

### Outlook for Windows (classic)
1. **File → Open & Export → Import/Export**
2. Choose **Export to a file**
3. Choose **Comma Separated Values (CSV)**
4. Select **Inbox** (for emails), export
5. Repeat and select **Calendar** (for events), export

### Outlook for Mac (classic)
Export options vary by Outlook version. If CSV export is unavailable in your build, two alternatives work:
- Export calendar to **.ics** (not supported by this MVP yet), or
- Do the CSV export from Outlook for Windows (or Outlook web + Power Automate) and then import the CSVs here.

### Import into the CRM
In the taskpane/web app, use:
- **Import Inbox CSV**
- **Import Calendar CSV**

The importer is heuristic: it looks for common Outlook CSV column names (Subject/From/To/Cc/Received, Start/End, Attendees). If your export uses different headers, open an issue and paste the header row.

## Run as standalone web app

Open:
- `https://localhost:3000/taskpane.html`

Click:
- **Import Inbox CSV** / **Import Calendar CSV**

## Sideload into OneNote (web)

Use:
- `manifest/manifest-onenote.xml`

## Sideload into PowerPoint

Use:
- `manifest/manifest-powerpoint.xml`

In PowerPoint:
- Insert → Add-ins → **Upload My Add-in** → select the manifest XML
- Open the taskpane, click **Read current slide**
- Click **Analyze stakeholder impact**

## Key files

- `src/taskpane/taskpane.tsx` — UI and host-specific workflows
- `src/shared/powerpoint.ts` — reads selected slide
- `server/index.ts` — local API proxy for OpenAI (keeps API key off the client)
- `src/shared/outlookExport.ts` — Outlook CSV importer (Inbox + Calendar)
- `src/shared/storage.ts` — local storage persistence + data model

## Security notes (MVP)

- The OpenAI API key is used **only on the local API server** via `.env`.
- Outlook ingestion is via **user-exported CSV** only (no Graph access, no OAuth client registration required).
