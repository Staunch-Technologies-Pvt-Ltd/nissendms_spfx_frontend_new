# Migration Assistant module

This folder is a self-contained SPFx module that embeds the **SharePoint AI
Migration Assistant** (originally a standalone Vite/React/Tailwind app at
`C:\llm\llm\SharePoint-AI-Migration-Assistant`) as a new "Migration
Assistant" tab inside the Vessel DMS web part.

## What changed, and why it's a separate module

The migration assistant is a fully independent tool: its own FastAPI
backend, its own SQLite/Postgres database, its own Microsoft Entra app
registration (`Sites.Selected`), its own classification logic (deterministic
keyword matching — no LLM). It does not share code, config, or a database
with the Vessel DMS backend, and per its own `CLAUDE.md` was deliberately
built that way. Merging its backend into `backend/app/main.py` would break
that independence and the "ask before changing auth / DB schema" rule this
project's CLAUDE-CONTEXT.md sets for the main DMS.

So the integration is:

- **Frontend**: this folder (`components/migrationAssistant/`) — a plain
  React 17 / Fluent UI 8 port of the original Tailwind UI (Tailwind isn't
  available in this SPFx build, see the project's tsconfig/package.json), it
  wired into a new "Migration Assistant" tab in `Sidebar.tsx` / `AppLayout`
  and view `'migration'` in `VesselEmail.tsx`. It is entirely self-contained
  — no import from the host app's state, and the host app imports only
  `MigrationAssistantModule` from `./MigrationAssistantModule`.
- **Backend**: unchanged, still runs as its own process (see below). This
  module just points the frontend at wherever it's running.

## Running it

**1. Migration backend** (unchanged from the standalone project — run from
`C:\llm\llm\SharePoint-AI-Migration-Assistant`):

```
cd SharePoint-AI-Migration-Assistant
.venv\Scripts\python.exe -m uvicorn backend.app.main:app --reload --port 8020
```

Make sure `SharePoint-AI-Migration-Assistant\.env` has `ALLOWED_ORIGINS`
including the SharePoint site that hosts this web part (and
`http://localhost:4321` / your workbench origin for local dev), so the
migration backend's own CORS accepts requests from the SPFx page.

**2. This web part**: build/serve as usual (`gulp serve` / `npm run
build`). Then in the web part's property pane, set:

- **Migration API Base URL** → the migration backend's URL from step 1
  (e.g. `http://localhost:8020` for local dev, or wherever it's deployed —
  this backend has no single documented production URL, unlike the main DMS
  API, since it's deployed independently).

Left blank, the module falls back to `http://127.0.0.1:8020` only when
served from the local SPFx workbench; otherwise it shows a short
"configure the URL" notice instead of erroring.

## Files

- `types.ts` — response/request types, ported from the standalone project's
  `frontend/src/api.ts`.
- `api.ts` — fetch-based client (the original used axios, which isn't a
  dependency of this SPFx project, so this is a from-scratch fetch wrapper
  with the same method surface).
- `fileUtils.ts` — file-type icon/size helpers (Fluent UI icon names instead
  of the original's lucide-react).
- `styles.ts` — shared inline-style tokens themed off this app's existing
  `clayTheme.ts` (no Tailwind/CSS-in-JS available here).
- `MigrationScanTab.tsx` — the scan → subfolders → vessel → review → confirm
  workflow (the main feature).
- `SiteToSiteTab.tsx` — copy-only migration between two SharePoint sites.
- `TagExistingFilesTab.tsx` — Managed Metadata tagging for files already in
  their correct destination folder.
- `VesselExcelExportTab.tsx` — read-only per-vessel Excel export.
- `MigrationAssistantModule.tsx` — the tab shell that ties the four above
  together and is the one thing the host app imports.

## Notes

- The migration backend's `/api/*` routes have no session auth of their own
  (see its README's "Future enhancements" — `X-User-Email` is only used for
  its audit trail, not access control). This module sends the SPFx-resolved
  signed-in user's email in that header on every call, same as the original
  standalone app's login screen did.
- Nothing here touches the main DMS's `X-Session-ID` / Graph-token auth —
  the migration backend does its own Graph app-only auth independently.
