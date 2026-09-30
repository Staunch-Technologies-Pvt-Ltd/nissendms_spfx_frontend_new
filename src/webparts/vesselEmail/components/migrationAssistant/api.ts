// Fetch client for the Migration Assistant API, which now lives INSIDE the
// main Vessel DMS backend (backend/app/migration_assistant/, routes under
// /api/migration-assistant/*) — no separate migration process any more. It
// authenticates with the same DMS session as the rest of the app. No axios
// dependency (not installed in this SPFx project) — plain fetch.
//
// This module is self-contained: it never imports from the host app's
// api/state, and the host app never imports from it except the one
// <MigrationAssistantModule /> entry point.

import type {
  ConfirmSummary,
  ExistingFileTagApply,
  ExistingFileTagScan,
  HierarchyNode,
  MigrationItem,
  MigrationScanJob,
  S2SConfirmSummary,
  S2SDrive,
  S2SFile,
  S2SFolder,
  S2SItem,
  S2SJob,
  S2SSite,
  SourceFile,
  SourceFolder,
  TaggedScanJob,
} from './types';

export class MigrationApiError extends Error {
  public status: number;
  constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}

function headers(actingEmail: string, sessionId: string): Record<string, string> {
  const h: Record<string, string> = { 'Content-Type': 'application/json', 'X-User-Email': actingEmail };
  if (sessionId) { h['Authorization'] = `Bearer ${sessionId}`; h['X-Session-ID'] = sessionId; }
  return h;
}

async function request<T>(
  baseUrl: string,
  path: string,
  actingEmail: string,
  sessionId: string,
  init?: RequestInit
): Promise<T> {
  const url = `${baseUrl.replace(/\/$/, '')}${path}`;
  const res = await fetch(url, {
    ...init,
    headers: { ...headers(actingEmail, sessionId), ...(init?.headers || {}) },
  });
  if (!res.ok) {
    let detail = `Request failed (${res.status})`;
    try {
      const body = await res.json();
      detail = body?.detail || detail;
    } catch {
      // ignore — non-JSON error body
    }
    throw new MigrationApiError(res.status, detail);
  }
  if (res.status === 204) return undefined as unknown as T;
  return (await res.json()) as T;
}

export function errDetail(e: unknown, fallback: string): string {
  return e instanceof MigrationApiError ? e.message : fallback;
}

export class MigrationApi {
  constructor(private baseUrl: string, private actingEmail: string, private sessionId: string = '') {}

  /** Auth headers for raw fetch() calls (file preview / Excel download). */
  authHeaders(): Record<string, string> {
    const h: Record<string, string> = { 'X-User-Email': this.actingEmail };
    if (this.sessionId) { h['Authorization'] = `Bearer ${this.sessionId}`; h['X-Session-ID'] = this.sessionId; }
    return h;
  }

  health(): Promise<{ status: string; configured: boolean }> {
    return request(this.baseUrl, '/api/migration-assistant/health', this.actingEmail, this.sessionId);
  }

  listSourceFolders(path?: string): Promise<{ path: string; folders: SourceFolder[]; files: SourceFile[] }> {
    const qs = path ? `?path=${encodeURIComponent(path)}` : '';
    return request(this.baseUrl, `/api/migration-assistant/migration/source-folders${qs}`, this.actingEmail, this.sessionId);
  }

  listVessels(): Promise<{ destination_root: string; vessels: SourceFolder[] }> {
    return request(this.baseUrl, '/api/migration-assistant/migration/vessels', this.actingEmail, this.sessionId);
  }

  listMigrationJobs(): Promise<MigrationScanJob[]> {
    return request(this.baseUrl, '/api/migration-assistant/migration/jobs', this.actingEmail, this.sessionId);
  }

  startMigrationScan(
    sourceFolder: string,
    subfolders: string[],
    vesselPath: string,
    files: string[] = []
  ): Promise<MigrationScanJob> {
    return request(this.baseUrl, '/api/migration-assistant/migration/scan', this.actingEmail, this.sessionId, {
      method: 'POST',
      body: JSON.stringify({ source_folder: sourceFolder, subfolders, vessel_path: vesselPath, files }),
    });
  }

  getMigrationScanJob(jobId: string): Promise<MigrationScanJob> {
    return request(this.baseUrl, `/api/migration-assistant/migration/scan/${jobId}`, this.actingEmail, this.sessionId);
  }

  getJobItems(jobId: string): Promise<{ job: MigrationScanJob; items: MigrationItem[] }> {
    return request(this.baseUrl, `/api/migration-assistant/migration/jobs/${jobId}/items`, this.actingEmail, this.sessionId);
  }

  getJobHierarchy(jobId: string): Promise<{ tree: HierarchyNode[]; paths: string[] }> {
    return request(this.baseUrl, `/api/migration-assistant/migration/jobs/${jobId}/hierarchy`, this.actingEmail, this.sessionId);
  }

  overrideJobItem(jobId: string, itemId: string, targetPath: string): Promise<MigrationItem> {
    return request(this.baseUrl, `/api/migration-assistant/migration/jobs/${jobId}/items/${itemId}/override`, this.actingEmail, this.sessionId, {
      method: 'POST',
      body: JSON.stringify({ target_path: targetPath }),
    });
  }

  reclassifyJobItem(jobId: string, itemId: string): Promise<MigrationItem> {
    return request(this.baseUrl, `/api/migration-assistant/migration/jobs/${jobId}/items/${itemId}/classify`, this.actingEmail, this.sessionId, {
      method: 'POST',
      body: '{}',
    });
  }

  jobItemPreviewUrl(jobId: string, itemId: string): string {
    return `${this.baseUrl.replace(/\/$/, '')}/api/migration/jobs/${jobId}/items/${itemId}/preview`;
  }

  confirmJob(jobId: string): Promise<ConfirmSummary> {
    return request(this.baseUrl, `/api/migration-assistant/migration/jobs/${jobId}/confirm`, this.actingEmail, this.sessionId, {
      method: 'POST',
      body: '{}',
      headers: this.authHeaders(),
    });
  }

  vesselExcelExportUrl(vesselPath: string): string {
    return `${this.baseUrl.replace(/\/$/, '')}/api/vessel-export/excel?vessel_path=${encodeURIComponent(vesselPath)}`;
  }

  scanExistingFiles(rootPath: string): Promise<ExistingFileTagScan> {
    return request(this.baseUrl, '/api/migration-assistant/tag-existing/scan', this.actingEmail, this.sessionId, {
      method: 'POST',
      body: JSON.stringify({ root_path: rootPath }),
    });
  }

  applyExistingFileTags(rootPath: string, fileIds: string[]): Promise<ExistingFileTagApply> {
    return request(this.baseUrl, '/api/migration-assistant/tag-existing/apply', this.actingEmail, this.sessionId, {
      method: 'POST',
      body: JSON.stringify({ root_path: rootPath, file_ids: fileIds }),
    });
  }

  listTaggedScanJobs(): Promise<TaggedScanJob[]> {
    return request(this.baseUrl, '/api/migration-assistant/tag-existing/jobs', this.actingEmail, this.sessionId);
  }

  getTaggedScanJob(jobId: string): Promise<ExistingFileTagScan> {
    return request(this.baseUrl, `/api/migration-assistant/tag-existing/jobs/${jobId}`, this.actingEmail, this.sessionId);
  }

  // --- Site-to-Site migration ---

  listS2SSites(): Promise<S2SSite[]> {
    return request(this.baseUrl, '/api/migration-assistant/site-to-site/sites', this.actingEmail, this.sessionId);
  }

  listS2SSiteDrives(siteKey: string): Promise<S2SDrive[]> {
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/sites/${siteKey}/drives`, this.actingEmail, this.sessionId);
  }

  browseS2SFolder(siteKey: string, driveId: string, path?: string): Promise<{ path: string; folders: S2SFolder[]; files: S2SFile[] }> {
    const params = new URLSearchParams({ site_key: siteKey, drive_id: driveId });
    if (path) params.set('path', path);
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/browse?${params.toString()}`, this.actingEmail, this.sessionId);
  }

  listS2SJobs(): Promise<S2SJob[]> {
    return request(this.baseUrl, '/api/migration-assistant/site-to-site/jobs', this.actingEmail, this.sessionId);
  }

  startS2SScan(args: {
    sourceSiteKey: string;
    sourceDriveId: string;
    sourceFolderPath: string;
    selectedFolders: string[];
    selectedFiles: string[];
    destSiteKey: string;
    destDriveId: string;
    destFolderPath: string;
  }): Promise<S2SJob> {
    return request(this.baseUrl, '/api/migration-assistant/site-to-site/scan', this.actingEmail, this.sessionId, {
      method: 'POST',
      body: JSON.stringify({
        source_site_key: args.sourceSiteKey,
        source_drive_id: args.sourceDriveId,
        source_folder_path: args.sourceFolderPath,
        selected_folders: args.selectedFolders,
        selected_files: args.selectedFiles,
        dest_site_key: args.destSiteKey,
        dest_drive_id: args.destDriveId,
        dest_folder_path: args.destFolderPath,
      }),
    });
  }

  getS2SJobItems(jobId: string): Promise<{ job: S2SJob; items: S2SItem[] }> {
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/jobs/${jobId}/items`, this.actingEmail, this.sessionId);
  }

  confirmS2SJob(jobId: string): Promise<S2SConfirmSummary> {
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/jobs/${jobId}/confirm`, this.actingEmail, this.sessionId, {
      method: 'POST',
      body: '{}',
    });
  }
}
