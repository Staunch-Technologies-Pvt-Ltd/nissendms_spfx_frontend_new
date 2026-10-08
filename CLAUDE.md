# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A SharePoint Framework (SPFx 1.23.2, React 17, Fluent UI 8, packaged with `@rushstack/heft`) solution that renders the Vessel Document Management System inside SharePoint Online / Teams. It calls the [Vessel DMS backend](../nissendms_spfx_backend) (FastAPI) over REST for business logic, and calls Microsoft Graph directly for SharePoint folder listing, uploads, and delta sync.

## Commands (run from this directory, in the integrated terminal)

```bash
# Install deps (Node >=22.14.0 <23.0.0 required — see package.json "engines")
npm install

# Serve against the hosted workbench for local dev (heft build-watch --serve)
npm start
# then open https://nk-dms-dev.sg-nissenkaiun.com/_layouts/15/workbench.aspx,
# add the "Vessel Document List" web part, and set its apiBaseUrl property
# (defaults to https://nk-dms-dev.sg-nissenkaiun.com; falls back to
# http://localhost:8000 automatically if that remote is unreachable)

# Full test + lint + production bundle + package solution
npm run build

# Just clean build output
npm run clean
```

There is no separate `npm test`/`npm run lint` script — `heft test` (invoked as part of `npm run build`, or run directly as `heft test`) runs both linting (`.eslintrc.js`, extending `@microsoft/eslint-config-spfx`) and any Jest specs together. No `*.test.*`/`*.spec.*` files are committed, so in practice `heft test` is lint + type-check (`jest-output/` is untracked generated output; `@typescript-eslint/no-explicit-any` and `no-unused-vars` are turned off in `.eslintrc.js`). Run `npx heft test --clean` directly during development instead of the full production build when you just want lint/test feedback.

Packaged output lands at `sharepoint/solution/vessel-email-spfx-v-2.sppkg` (prod) or the `-dev` variant, per the `config/package-solution*.json` / `config/serve*.json` dev/prod pairs.

## Architecture

**Two components in one solution:**
- `src/webparts/vesselEmail/` — the main web part (`VesselEmailWebPart.ts`), rendering `components/VesselEmail.tsx`. `VesselEmail.tsx` alone is ~12,500 lines and holds most app state, auth bootstrap and Graph calls — grep it before assuming logic lives in a page file. This is essentially the whole application: vessel registry, document browser/uploads, approvals, recycle bin, archive, sessions, settings, reports, users, and the AI BANTO email dashboard.
- `src/extensions/uploadAndTag/` — a ListView Command Set extension (`UploadAndTagCommandSet.ts`) that adds upload/tag actions directly to the SharePoint document library's command bar, independent of the web part.

**Page components live flat in `components/pages/`** (one file per screen — `DashboardPage`, `DocumentsPage`, `VesselsPage`, `ApprovalsPage`, `RecycleBinPage`, `ArchivePage`, `BentoEmailDashboardPage`, `SettingsPage` + its `*ManagementSection`/`*Section` sub-panels, etc.), routed/switched inside `AppLayout.tsx` + `Sidebar.tsx` rather than through a router library. When adding a new screen, follow this pattern: a new file in `pages/`, wired into `AppLayout`/`Sidebar`.

**Auth and session bootstrap (`VesselEmail.tsx` `onInit`/mount path):**
1. Resolve the signed-in user's email via Graph `/me`, and the current site's `siteId`/document-library `driveId`.
2. Call backend `POST /api/auth/bypass-login` with that email to mint a server-side `session_id`.
3. Send `session_id` on every subsequent backend call as `X-Session-ID` / `Authorization: Bearer`.
4. Load the vessel list via the backend's `/api/vessels/flat-tree` (single DB query, not a live Graph walk) for the initial Documents view.

**Uploads bypass the backend.** Files are `PUT`-ed directly to SharePoint via Graph (`…/items/{folderId}:/{file}:/content`), not proxied through the FastAPI backend. Gated folders instead get staged behind an approval request (a backend call after the Graph upload), so upload code paths touch both Graph directly and the backend's approvals API.

**`deltaSync.ts` polls the Graph drive delta endpoint** on an interval and merges changes into an in-memory folder map, which is what makes uploads/moves from other users or other tabs show up without a manual reload — check this module before assuming the folder tree state is only ever mutated by local user actions.

**Supporting modules to know about in `components/`:**
- `graphFolderService.ts` — folder creation against the Graph drive.
- `vesselFolderTemplate.ts` — the canonical `Documents/Vessels/{vessel}/{main}/...` tree shape used both when provisioning and when rendering.
- `vesselImagePool.ts` / `assets/vessel_images/` — per-vessel thumbnail assignment.
- `migrationAssistant/` — a self-contained module (scan / site-to-site / tag-existing-files / Excel export tabs) for one-time data migration tooling; has its own `README.md` worth reading before touching it.
- `copilot/CopilotSearchPanel.tsx` — floating bottom-right chat button ("Documents Copilot") mounted once in the app shell (`<CopilotSearchPanel host={this} />` at the end of `VesselEmail.tsx` render), so it shows on every page; calls the backend's `app/copilot/` API (`/status`, `/query`) via `host._base()` / `host._headers()`. Clicking a file result calls `host._openDocumentFile`; folders aren't clickable. Shows "basic mode" when the backend has no Azure OpenAI key.
- `types/` (`rows.ts`, `bento.ts`, `ui.ts`, `view.ts`) — shared TypeScript types; check here before redefining a shape that likely already exists.

## Key configuration

- `apiBaseUrl` (web part property pane, default in `VesselEmailWebPart.manifest.json`) — the backend origin. Dev default is `https://nk-dms-dev.sg-nissenkaiun.com`; falls back to `http://localhost:8000` if that's unreachable.
- `VESSEL_ROOT` / `MAIN_FOLDER_NAMES` in `VesselEmail.tsx` define the SharePoint tree layout (`Documents/Vessels/{vessel}/{main}`) — keep in sync with the backend's own folder-template assumptions if you change either side.
- Document library is resolved by name (**"Documents"**) at runtime via Graph, not hardcoded by ID.
- `config/serve.json` / `config/serve.dev.json` / `config/serve.prod.json` and `config/package-solution*.json` are the dev/prod pairs swapped depending on which environment you're building for.
