# Migration Assistant module (frontend)

The SharePoint AI Migration Assistant UI, shown as the "Migration Assistant"
tab of the Vessel DMS web part. It is a plain React 17 / Fluent UI 8 port of
the original Tailwind app.

## Backend: merged into the main DMS API

The migration backend is no longer a separate application. It is part of the
DMS backend (`backend/app/migration_assistant/`), served under
`/api/migration-assistant/*` and authenticated with the normal DMS session
(`Authorization` / `X-Session-ID`). Start the DMS backend and the web part as
usual and the tab works; nothing else to run.

The web part's **API Base URL** is used. The "Migration API Base URL" property
is now only an optional override — leave it blank.

Backend configuration (Entra app, sites, destination): `backend/.env.migration`,
see `backend/migration.env.example` and `backend/app/migration_assistant/README.md`.

## Files

- `types.ts` — response/request types.
- `api.ts` — fetch client (`MigrationApi`), adds the DMS session headers.
- `fileUtils.ts`, `styles.ts` — icon/size helpers and shared style tokens.
- `MigrationScanTab.tsx` — scan → subfolders → vessel → review → confirm.
- `SiteToSiteTab.tsx` — copy-only migration between two SharePoint sites: site
  search / URL lookup, copy options, live progress over the `/stream` NDJSON
  feed (pause / resume / cancel), verification summary, Excel report, job history.
- `TagExistingFilesTab.tsx` — Managed Metadata tagging of already-placed files.
- `VesselExcelExportTab.tsx` — read-only per-vessel Excel export.
- `MigrationAssistantModule.tsx` — tab shell; the only file the host app imports.
