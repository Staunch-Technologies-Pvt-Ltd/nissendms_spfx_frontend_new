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
  SPECIFIC_VESSELS_ROOT,
  COMMON_SHIPS_ROOT,
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

/** Called once for every folder attempted (created / existed / failed), in order. */
export type FolderProgressCallback = (result: FolderResult) => void;
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

function withTimeout<T>(promise: Promise<T>, ms: number, label: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error(`Timed out after ${Math.round(ms / 1000)}s waiting for: ${label}`));
    }, ms);
    promise.then(
      value => { clearTimeout(timer); resolve(value); },
      err => { clearTimeout(timer); reject(err); },
    );
  });
}

const GRAPH_CALL_TIMEOUT_MS = 25000;

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
  const response = await withTimeout(
      client.api(url).post({
        name: folderName,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'fail',
      }),
      GRAPH_CALL_TIMEOUT_MS,
      `create "${folderName}"`,
    );
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
       const existing = await withTimeout(
          client.api(`/sites/${siteId}/drives/${driveId}/root:/${encodePath(fullPath)}`).get(),
          GRAPH_CALL_TIMEOUT_MS,
          `lookup existing "${folderName}"`,
        );
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
  onProgress?: (result: FolderResult) => void,
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
      const entry: FolderResult = { path: fullPath, id, status: existed ? 'existed' : 'created' };
      log.push(entry);
      onProgress?.(entry);

      if (node.children && node.children.length > 0) {
        await createTree(client, siteId, driveId, fullPath, node.children, log, onProgress);
      }
    } catch (err: any) {
      const entry: FolderResult = {
        path: fullPath,
        status: 'failed',
        error: err?.message ?? String(err),
      };
      log.push(entry);
      onProgress?.(entry);
      // Do NOT recurse into children if the parent failed — they would also fail.
    }
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Ensure the base structure exists:
 *   Documents/Vessels/
 *   Documents/Vessels/Specific Vessels/
 *   Documents/Vessels/Common for all ships/{MainFolder}/{commonTree}
 *   Documents/Kaizen - Knowledge Bank/  (sibling of Vessels)
 * Safe to call on every app load — all operations are idempotent.
 */
export async function ensureRootStructure(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  onProgress?: FolderProgressCallback,
): Promise<FolderResult[]> {
  const log: FolderResult[] = [];
  if (!siteId || !driveId) return log;

  const record = (result: FolderResult): void => { log.push(result); onProgress?.(result); };

  // 1. Documents/Vessels
  await createFolder(client, siteId, driveId, '', VESSEL_MANAGEMENT_ROOT)
    .then(({ id, existed }) => record({ path: VESSEL_MANAGEMENT_ROOT, id, status: existed ? 'existed' : 'created' }))
    .catch(err => record({ path: VESSEL_MANAGEMENT_ROOT, status: 'failed', error: err?.message }));

  // 2. Documents/Vessels/Specific Vessels
  const specificPath = `${VESSEL_MANAGEMENT_ROOT}/${SPECIFIC_VESSELS_ROOT}`;
  await createFolder(client, siteId, driveId, VESSEL_MANAGEMENT_ROOT, SPECIFIC_VESSELS_ROOT)
    .then(({ id, existed }) => record({ path: specificPath, id, status: existed ? 'existed' : 'created' }))
    .catch(err => record({ path: specificPath, status: 'failed', error: err?.message }));

  // 3. Documents/Vessels/Common for all ships
  const commonPath = `${VESSEL_MANAGEMENT_ROOT}/${COMMON_SHIPS_ROOT}`;
  await createFolder(client, siteId, driveId, VESSEL_MANAGEMENT_ROOT, COMMON_SHIPS_ROOT)
    .then(({ id, existed }) => record({ path: commonPath, id, status: existed ? 'existed' : 'created' }))
    .catch(err => record({ path: commonPath, status: 'failed', error: err?.message }));

  // 4. Documents/Vessels/Common for all ships/{MainFolder}/{commonTree}
  for (const mf of MAIN_FOLDERS) {
    const mfPath = `${commonPath}/${mf.name}`;
    try {
      const { id, existed } = await createFolder(client, siteId, driveId, commonPath, mf.name);
      record({ path: mfPath, id, status: existed ? 'existed' : 'created' });
      if (mf.commonTree.length > 0) {
        await createTree(client, siteId, driveId, mfPath, mf.commonTree, log, onProgress);
      }
    } catch (err: any) {
      record({ path: mfPath, status: 'failed', error: err?.message });
    }
  }

  // 5. Documents/Kaizen - Knowledge Bank (sibling of Vessels, at drive root)
  const kaizenName = 'Kaizen - Knowledge Bank';
  await createFolder(client, siteId, driveId, '', kaizenName)
    .then(({ id, existed }) => record({ path: kaizenName, id, status: existed ? 'existed' : 'created' }))
    .catch(err => record({ path: kaizenName, status: 'failed', error: err?.message }));

  return log;
}
/**
 * Create the full per-vessel folder tree for a new vessel.
 * Structure: Documents/Vessels/Specific Vessels/{VesselName}/{MainFolder}/...
 */
export async function createVesselFolders(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  vesselName: string,
  onProgress?: FolderProgressCallback,
  skipRootStructure?: boolean,
): Promise<VesselFolderCreationResult> {
  const log: FolderResult[] = [];
  console.log(`[VesselDMS] createVesselFolders START vessel="${vesselName}"`);

  if (!siteId || !driveId) {
    return { vesselName, siteId, driveId, results: log, success: false };
  }

  // Ensure Vessels/, Vessels/Specific Vessels/, Vessels/Common for all ships/, Kaizen/
  if (!skipRootStructure) {
    const rootLog = await ensureRootStructure(client, siteId, driveId, onProgress);
    log.push(...rootLog);
  }

  // Documents/Vessels/Specific Vessels/{VesselName}
  const specificVesselsPath = `${VESSEL_MANAGEMENT_ROOT}/${SPECIFIC_VESSELS_ROOT}`;
  const vesselFolderPath = `${specificVesselsPath}/${vesselName}`;
  try {
    const { id, existed } = await createFolder(client, siteId, driveId, specificVesselsPath, vesselName);
    const result: FolderResult = { path: vesselFolderPath, id, status: existed ? 'existed' : 'created' };
    log.push(result);
    onProgress?.(result);
  } catch (err: any) {
    const result: FolderResult = { path: vesselFolderPath, status: 'failed', error: err?.message ?? String(err) };
    log.push(result);
    onProgress?.(result);
    return { vesselName, siteId, driveId, results: log, success: false };
  }

  // Documents/Vessels/Specific Vessels/{VesselName}/{MainFolder}/{perVesselTree}
  for (const mf of MAIN_FOLDERS) {
    const mainFolderPath = `${vesselFolderPath}/${mf.name}`;
    try {
      const { id, existed } = await createFolder(client, siteId, driveId, vesselFolderPath, mf.name);
      const result: FolderResult = { path: mainFolderPath, id, status: existed ? 'existed' : 'created' };
      log.push(result);
      onProgress?.(result);
      await createTree(client, siteId, driveId, mainFolderPath, mf.perVesselTree, log, onProgress);
    } catch (err: any) {
      const result: FolderResult = { path: mainFolderPath, status: 'failed', error: err?.message ?? String(err) };
      log.push(result);
      onProgress?.(result);
    }
  }

  const success = log.every(r => r.status !== 'failed');
  console.log(`[VesselDMS] createVesselFolders DONE vessel="${vesselName}" success=${success}`);
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
