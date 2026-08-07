/**
 * graphFolderService.ts
 *
 * Idempotent, recursive Graph API folder-creation service.
 *
 * Responsibilities:
 *  - Ensure the top-level "Vessel Management" root exists.
 *  - For each MainFolder (Folder-1/2/3):
 *      • Create the MainFolder itself if absent.
 *      • Create all "common" sub-trees once (skip if already present).
 *      • Create the {VesselName} sub-folder and its full per-vessel tree.
 *  - Return a per-path success/failure log so the caller can retry only
 *    the paths that failed (fully idempotent re-runs).
 *
 * Graph endpoint used:
 *   POST /sites/{siteId}/drives/{driveId}/root:/{parentPath}:/children
 *   { name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }
 *
 * A 409 Conflict response means the folder already exists → treated as success.
 */

import { MSGraphClientV3 } from '@microsoft/sp-http';
import {
  VESSEL_MANAGEMENT_ROOT,
  MAIN_FOLDERS,
  FolderNode,
  MainFolder,
} from './vesselFolderTemplate';

// ── Types ─────────────────────────────────────────────────────────────────────

export interface FolderResult {
  path: string;
  id?: string;
  status: 'created' | 'existed' | 'failed';
  error?: string;
}

export interface VesselFolderCreationResult {
  vesselName: string;
  siteId: string;
  driveId: string;
  results: FolderResult[];
  success: boolean;
}

// ── Internal helpers ──────────────────────────────────────────────────────────

/**
 * Encode each path segment individually (preserves "/" separators).
 * Graph requires each segment to be URL-encoded but the "/" must stay literal.
 */
function encodePath(path: string): string {
  return path
    .split('/')
    .map(seg => encodeURIComponent(seg))
    .join('/');
}

/**
 * Attempt to create a single folder under parentPath.
 * Returns the folder item on success/already-exists, throws on real errors.
 */
async function createFolder(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  parentPath: string,
  folderName: string,
): Promise<{ id: string; existed: boolean }> {
  // When parentPath is empty, the folder goes at the drive root.
  // Graph API root-level children endpoint is /root/children (no path segment).
  // Using /root:/<empty>:/children produces a double-slash and a 400/404.
  const url = parentPath
    ? `/sites/${siteId}/drives/${driveId}/root:/${encodePath(parentPath)}:/children`
    : `/sites/${siteId}/drives/${driveId}/root/children`;

  console.log(`[VesselDMS] createFolder → POST ${url} name="${folderName}"`);

  try {
    const response = await client
      .api(url)
      .post({
        name: folderName,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'fail',
      });
    console.log(`[VesselDMS] createFolder ✓ created id=${response.id} "${folderName}"`);
    return { id: response.id as string, existed: false };
  } catch (err: any) {
    // 409 = folder already exists — treat as success
    const status: number =
      err?.statusCode ?? err?.response?.status ?? err?.code ?? 0;
    const isConflict =
      status === 409 ||
      (err?.message ?? '').toLowerCase().includes('namealreadyexists');

    if (isConflict) {
      console.log(`[VesselDMS] createFolder ↩ existed "${folderName}"`);
      // Fetch the existing folder's ID so callers have it
      try {
        const fullPath = parentPath ? `${parentPath}/${folderName}` : folderName;
        const existing = await client
          .api(`/sites/${siteId}/drives/${driveId}/root:/${encodePath(fullPath)}`)
          .get();
        return { id: existing.id as string, existed: true };
      } catch {
        return { id: '', existed: true };
      }
    }
    console.error(`[VesselDMS] createFolder ✗ FAILED "${folderName}" status=${status}`, err?.message ?? err);
    throw err;
  }
}

/**
 * Recursively walk a FolderNode tree, creating each node under its parent path.
 * Pushes a FolderResult entry for every node attempted.
 */
async function createTree(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  parentPath: string,
  nodes: FolderNode[],
  log: FolderResult[],
): Promise<void> {
  for (const node of nodes) {
    const fullPath = `${parentPath}/${node.name}`;
    try {
      const { id, existed } = await createFolder(
        client,
        siteId,
        driveId,
        parentPath,
        node.name,
      );
      log.push({ path: fullPath, id, status: existed ? 'existed' : 'created' });

      if (node.children && node.children.length > 0) {
        await createTree(client, siteId, driveId, fullPath, node.children, log);
      }
    } catch (err: any) {
      log.push({
        path: fullPath,
        status: 'failed',
        error: err?.message ?? String(err),
      });
      // Do NOT recurse into children if the parent failed — they would also fail.
    }
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Ensure the "Vessel Management" root and all three MainFolder roots exist.
 * Safe to call on every app load — all operations are idempotent.
 */
export async function ensureRootStructure(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
): Promise<FolderResult[]> {
  const log: FolderResult[] = [];
  console.log(`[VesselDMS] ensureRootStructure siteId=${siteId} driveId=${driveId}`);

  if (!siteId || !driveId) {
    console.error('[VesselDMS] ensureRootStructure aborted — siteId or driveId is empty!');
    return log;
  }

  // 1. Vessel Management root
  try {
    const { id, existed } = await createFolder(
      client, siteId, driveId, '', VESSEL_MANAGEMENT_ROOT,
    );
    log.push({
      path: VESSEL_MANAGEMENT_ROOT,
      id,
      status: existed ? 'existed' : 'created',
    });
  } catch (err: any) {
    log.push({
      path: VESSEL_MANAGEMENT_ROOT,
      status: 'failed',
      error: err?.message ?? String(err),
    });
    return log; // Can't continue without the root
  }

  // 2. Each MainFolder + its common sub-tree
  for (const mf of MAIN_FOLDERS) {
    const mfPath = `${VESSEL_MANAGEMENT_ROOT}/${mf.name}`;
    try {
      const { id, existed } = await createFolder(
        client, siteId, driveId, VESSEL_MANAGEMENT_ROOT, mf.name,
      );
      log.push({ path: mfPath, id, status: existed ? 'existed' : 'created' });
    } catch (err: any) {
      log.push({ path: mfPath, status: 'failed', error: err?.message ?? String(err) });
      continue;
    }

    // Common sub-tree (created once, idempotent)
    if (mf.commonTree.length > 0) {
      await createTree(client, siteId, driveId, mfPath, mf.commonTree, log);
    }
  }

  return log;
}

/**
 * Create the full per-vessel folder tree for a new vessel.
 *
 * @param client      MSGraphClientV3 from SPFx context
 * @param siteId      SharePoint site ID
 * @param driveId     Document library drive ID
 * @param vesselName  Exact vessel name (used as the folder name)
 *
 * Returns a VesselFolderCreationResult with per-path success/failure entries.
 * Re-running for an existing vessel is safe — already-present folders are skipped.
 */
export async function createVesselFolders(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  vesselName: string,
): Promise<VesselFolderCreationResult> {
  const log: FolderResult[] = [];
  console.log(`[VesselDMS] createVesselFolders START vessel="${vesselName}" siteId=${siteId} driveId=${driveId}`);

  if (!siteId || !driveId) {
    console.error('[VesselDMS] createVesselFolders aborted — siteId or driveId is empty!');
    return { vesselName, siteId, driveId, results: log, success: false };
  }

  // Ensure root + common folders exist first
  const rootLog = await ensureRootStructure(client, siteId, driveId);
  log.push(...rootLog);

  const rootFailed = rootLog.some(r => r.path === VESSEL_MANAGEMENT_ROOT && r.status === 'failed');
  if (rootFailed) {
    return { vesselName, siteId, driveId, results: log, success: false };
  }

  // Per-vessel tree under each MainFolder
  for (const mf of MAIN_FOLDERS) {
    // Skip MainFolders that failed during root setup
    const mfPath = `${VESSEL_MANAGEMENT_ROOT}/${mf.name}`;
    const mfFailed = log.some(r => r.path === mfPath && r.status === 'failed');
    if (mfFailed) continue;

    if (mf.perVesselTree.length === 0) continue;

    // Create {VesselName} folder under this MainFolder
    const vesselFolderPath = `${mfPath}/${vesselName}`;
    try {
      const { id, existed } = await createFolder(
        client, siteId, driveId, mfPath, vesselName,
      );
      log.push({
        path: vesselFolderPath,
        id,
        status: existed ? 'existed' : 'created',
      });
    } catch (err: any) {
      log.push({
        path: vesselFolderPath,
        status: 'failed',
        error: err?.message ?? String(err),
      });
      continue; // Can't create children without the vessel root
    }

    // Recursively create the per-vessel sub-tree
    await createTree(
      client, siteId, driveId, vesselFolderPath, mf.perVesselTree, log,
    );
  }

  const success = log.every(r => r.status !== 'failed');
  console.log(`[VesselDMS] createVesselFolders DONE vessel="${vesselName}" success=${success}`, log);
  return { vesselName, siteId, driveId, results: log, success };
}

/**
 * Retry only the paths that previously failed.
 * Pass the results array from a previous createVesselFolders call.
 * Returns a merged result with updated statuses.
 */
export async function retryFailedFolders(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  previousResults: FolderResult[],
): Promise<FolderResult[]> {
  const failed = previousResults.filter(r => r.status === 'failed');
  const updated = [...previousResults];

  for (const entry of failed) {
    // Derive parent path and folder name from the full path
    const lastSlash = entry.path.lastIndexOf('/');
    const parentPath = lastSlash > 0 ? entry.path.substring(0, lastSlash) : '';
    const folderName = entry.path.substring(lastSlash + 1);

    try {
      const { id, existed } = await createFolder(
        client, siteId, driveId, parentPath, folderName,
      );
      const idx = updated.findIndex(r => r.path === entry.path);
      if (idx !== -1) {
        updated[idx] = { path: entry.path, id, status: existed ? 'existed' : 'created' };
      }
    } catch (err: any) {
      // Leave as failed with updated error message
      const idx = updated.findIndex(r => r.path === entry.path);
      if (idx !== -1) {
        updated[idx] = { ...updated[idx], error: err?.message ?? String(err) };
      }
    }
  }

  return updated;
}
