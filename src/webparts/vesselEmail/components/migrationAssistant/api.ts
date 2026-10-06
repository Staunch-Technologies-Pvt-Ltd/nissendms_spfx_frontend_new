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
  S2SDrive,
  S2SFile,
  S2SFolder,
  S2SItem,
  S2SCopyOptions,
  S2SJob,
  S2SSite,
  S2SStreamLine,
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

  /** Folders/files under `path` in the picked source library (`source`),
   *  or in the configured default source when none is picked. */
  listSourceFolders(path?: string, source?: { siteKey: string; driveId: string } | null): Promise<{ path: string; folders: SourceFolder[]; files: SourceFile[] }> {
    const params = new URLSearchParams();
    if (path) params.set('path', path);
    if (source) { params.set('site_key', source.siteKey); params.set('drive_id', source.driveId); }
    const qs = params.toString() ? `?${params.toString()}` : '';
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
    files: string[] = [],
    source?: { siteKey: string; driveId: string } | null
  ): Promise<MigrationScanJob> {
    return request(this.baseUrl, '/api/migration-assistant/migration/scan', this.actingEmail, this.sessionId, {
      method: 'POST',
      body: JSON.stringify({
        source_folder: sourceFolder, subfolders, vessel_path: vesselPath, files,
        source_site_key: source?.siteKey ?? null, source_drive_id: source?.driveId ?? null,
      }),
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
    return `${this.baseUrl.replace(/\/$/, '')}/api/migration-assistant/migration/jobs/${jobId}/items/${itemId}/preview`;
  }

  confirmJob(jobId: string): Promise<ConfirmSummary> {
    return request(this.baseUrl, `/api/migration-assistant/migration/jobs/${jobId}/confirm`, this.actingEmail, this.sessionId, {
      method: 'POST',
      body: '{}',
      headers: this.authHeaders(),
    });
  }

  vesselExcelExportUrl(vesselPath: string): string {
    return `${this.baseUrl.replace(/\/$/, '')}/api/migration-assistant/vessel-export/excel?vessel_path=${encodeURIComponent(vesselPath)}`;
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

  /** Tenant-wide site search (needs Sites.Read.All on the app; the error
   *  message says so when it isn't granted). */
  searchS2SSites(query: string): Promise<S2SSite[]> {
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/sites/search?q=${encodeURIComponent(query)}`, this.actingEmail, this.sessionId);
  }

  /** Look up a site from any URL inside it and check the app can read it. */
  resolveS2SSite(url: string): Promise<S2SSite> {
    return request(this.baseUrl, '/api/migration-assistant/site-to-site/sites/resolve', this.actingEmail, this.sessionId, {
      method: 'POST',
      body: JSON.stringify({ url }),
    });
  }

  listS2SSiteDrives(siteKey: string): Promise<S2SDrive[]> {
    // Query-string form: URL-based site keys contain slashes.
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/drives?site_key=${encodeURIComponent(siteKey)}`, this.actingEmail, this.sessionId);
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

  /** Starts the copy in the background; returns the job right away. Also
   *  retries a finished job's failed / remaining items. */
  confirmS2SJob(jobId: string, options: S2SCopyOptions): Promise<S2SJob> {
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/jobs/${jobId}/confirm`, this.actingEmail, this.sessionId, {
      method: 'POST',
      body: JSON.stringify({
        conflict_policy: options.conflictPolicy,
        copy_permissions: options.copyPermissions,
        copy_versions: options.copyVersions,
      }),
    });
  }

  getS2SJob(jobId: string): Promise<S2SJob> {
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/jobs/${jobId}`, this.actingEmail, this.sessionId);
  }

  pauseS2SJob(jobId: string): Promise<S2SJob> {
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/jobs/${jobId}/pause`, this.actingEmail, this.sessionId, { method: 'POST', body: '{}' });
  }

  resumeS2SJob(jobId: string): Promise<S2SJob> {
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/jobs/${jobId}/resume`, this.actingEmail, this.sessionId, { method: 'POST', body: '{}' });
  }

  cancelS2SJob(jobId: string): Promise<S2SJob> {
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/jobs/${jobId}/cancel`, this.actingEmail, this.sessionId, { method: 'POST', body: '{}' });
  }

  verifyS2SJob(jobId: string): Promise<S2SJob> {
    return request(this.baseUrl, `/api/migration-assistant/site-to-site/jobs/${jobId}/verify`, this.actingEmail, this.sessionId, { method: 'POST', body: '{}' });
  }

  s2sReportUrl(jobId: string): string {
    return `${this.baseUrl.replace(/\/$/, '')}/api/migration-assistant/site-to-site/jobs/${jobId}/report`;
  }

  /**
   * Follow a copy job's live NDJSON progress feed, calling `onLine` for each
   * line. Resolves when the server sends the final "done" line; rejects if
   * the connection drops (the caller reconnects) or `signal` aborts.
   */
  async streamS2SJob(jobId: string, onLine: (line: S2SStreamLine) => void, signal: AbortSignal): Promise<void> {
    const url = `${this.baseUrl.replace(/\/$/, '')}/api/migration-assistant/site-to-site/jobs/${jobId}/stream`;
    const res = await fetch(url, { headers: this.authHeaders(), signal });
    if (!res.ok || !res.body) {
      let detail = `Request failed (${res.status})`;
      try { detail = (await res.json())?.detail || detail; } catch { /* non-JSON body */ }
      throw new MigrationApiError(res.status, detail);
    }
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      let nl = buffer.indexOf('\n');
      while (nl >= 0) {
        const raw = buffer.slice(0, nl).trim();
        buffer = buffer.slice(nl + 1);
        if (raw) {
          const line = JSON.parse(raw) as S2SStreamLine;
          onLine(line);
          if (line.type === 'done') { void reader.cancel(); return; }
        }
        nl = buffer.indexOf('\n');
      }
    }
    throw new MigrationApiError(0, 'Progress stream closed before the copy finished');
  }
}
