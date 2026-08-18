# Vessel DMS — Frontend (SPFx Web Part)

The SharePoint Framework solution that renders the Vessel Document Management
System inside SharePoint Online (and Microsoft Teams). Built with the SharePoint
Framework **1.23.2**, **React 17**, and **Fluent UI 8**, packaged with
`@rushstack/heft`.

It talks to the [Vessel DMS backend](../backend/README.md) (FastAPI) over REST,
and talks to the SharePoint **Documents** library directly through the **Microsoft
Graph** API for folder listing, uploads, and near-real-time delta sync.

## Solution contents

| Component | Type | What it does |
|---|---|---|
| **VesselEmail** web part (`src/webparts/vesselEmail/`) | Web part | The main application: vessel registry, document browser, uploads, approvals, recycle bin, archive, sessions, settings, and the AI BANTO email dashboard. |
| **UploadAndTag** (`src/extensions/uploadAndTag/`) | ListView Command Set | A list-view command bar for the SharePoint document list — upload & tag documents directly from the library. |

### UI pages (`components/pages/`)

- **Dashboard** — landing overview with quick stats and recent activity.
- **Documents** — flattened per-vessel document list (`Vessel > Main Folder > Category > Sub-Category`)
  with search, filters, folder/list view, and preview.
- **Vessels** — create / edit / delete vessels and provision their folder trees.
- **Templates** — document templates per category.
- **Approvals** — pending / approved / rejected admin decisions on uploads & actions.
- **Notifications** — feed derived from approval events.
- **Reports / Users / Settings** — reporting, user directory, and configuration.
- **Bento Email** (`BentoEmailDashboardPage`) — AI BANTO email log, compose modal,
  and auto-tag document upload.
- **Recycle Bin / Archive** — soft-deleted and archived items with restore /
  permanent delete.

Supporting modules in `components/`: `graphFolderService.ts` (folder creation),
`deltaSync.ts` (Graph drive delta polling), `vesselFolderTemplate.ts`,
`vesselImagePool.ts`, and `types/` (rows, bento, ui, view).

## Prerequisites

- **Node** `>=22.14.0 < 23.0.0` (see `engines` in `package.json`).
- A SharePoint Online developer/tenant site with a **Documents** library.
- The Vessel DMS backend reachable at the URL you configure (default for dev is
  `https://nk-dms-dev.sg-nissenkaiun.com`).
- Azure AD app registration / Graph permissions for the web part (see backend README).

## Getting started

```bash
npm install
```

### Run against the hosted workbench (dev)

```bash
npm start        # heft start --clean  (gulp serve equivalent)
```

Open the workbench (default dev: `https://nk-dms-dev.sg-nissenkaiun.com/_layouts/15/workbench.aspx`),
add the **Vessel Document List** web part, and in its property pane set:

- **API Base URL** — the backend origin, e.g. `https://nk-dms-dev.sg-nissenkaiun.com`
  (or `http://localhost:8000` when running the backend locally).
- **API Token** — optional.

You can also set the default in `src/webparts/vesselEmail/VesselEmailWebPart.manifest.json`
under `properties.apiBaseUrl`.

> Note: if the configured remote backend is unreachable (network/TLS failure), the
> web part automatically falls back to `http://localhost:8000` so local backend
> development keeps working.

## Build & package

```bash
npm run build          # heft test --clean --production && heft package-solution --production
# or, per DEPLOY.md:
gulp bundle --ship
gulp package-solution --ship
```

Output: `sharepoint/solution/vessel-email-spfx-v-2.sppkg` (prod) /
`vessel-email-spfx-v-2-dev.sppkg` (dev). See the repo-root [`DEPLOY.md`](../DEPLOY.md)
for the exact config-file swaps (`serve.json`, `package-solution.json`) and App
Catalog deployment steps per environment.

Other build options are listed via `heft --help`.

## How the web part works

1. On `onInit`, it resolves the signed-in user's email via Graph `/me` and the
   site's `siteId` / `driveId` (the document library drive).
2. It calls the backend `POST /api/auth/bypass-login` with the SharePoint user's
   email to get a server-side `session_id`, which it then sends on every request as
   `X-Session-ID` / `Authorization: Bearer`.
3. `VesselEmail` loads the vessel list and the backend's fast `/api/vessels/flat-tree`
   endpoint (a single DB query) for the initial Documents view.
4. A **delta-sync scheduler** polls the Graph drive for folder/file changes and
   merges them into an in-memory folder map, so uploads appear live without a reload.
5. Uploads go **directly to SharePoint via Graph** (`PUT …/items/{folderId}:/{file}:/content`)
   and (for gated folders) are staged behind an approval request on the backend.

## Key configuration

| Setting | Where | Notes |
|---|---|---|
| `apiBaseUrl` | Web part property pane / manifest default | Backend origin; default `https://nk-dms-dev.sg-nissenkaiun.com` |
| `apiToken` | Web part property pane | Optional token, passed through as-is |
| `VESSEL_ROOT` / `MAIN_FOLDER_NAMES` | `VesselEmail.tsx` | SharePoint tree layout: `Documents/Vessels/{vessel}/{main}` |
| Document library drive | Graph resolution on init | Looks up the library named **Documents** in the current site |

## Version history

| Version | Date | Comments |
|---|---|---|
| 1.0.x | — | Current development line (SPFx 1.23.2, Heft, React 17) |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR
IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE,
MERCHANTABILITY, OR NON-INFRINGEMENT.**

## References

- [SharePoint Framework](https://aka.ms/spfx)
- [Building for Microsoft Teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Heft Documentation](https://heft.rushstack.io/)
- [Backend README](../backend/README.md) — the REST API this web part calls
- [DEPLOY.md](../DEPLOY.md) — dev/prod deployment runbook
