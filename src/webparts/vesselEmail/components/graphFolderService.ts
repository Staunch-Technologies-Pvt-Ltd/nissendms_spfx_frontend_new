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

function sanitizeFolderName(name: string): string {
  return (name || '')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
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
  const safeFolderName = sanitizeFolderName(folderName);

  // When parentPath is empty, the folder goes at the drive root.
  // Graph API root-level children endpoint is /root/children (no path segment).
  // Using /root:/<empty>:/children produces a double-slash and a 400/404.
  const url = parentPath
    ? `/sites/${siteId}/drives/${driveId}/root:/${encodePath(parentPath)}:/children`
    : `/sites/${siteId}/drives/${driveId}/root/children`;

  console.log(`[VesselDMS] createFolder → POST ${url} name="${safeFolderName}"`);

  try {
  const response = await withTimeout(
      client.api(url).post({
        name: safeFolderName,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'fail',
      }),
      GRAPH_CALL_TIMEOUT_MS,
      `create "${safeFolderName}"`,
    );
    console.log(`[VesselDMS] createFolder ✓ created id=${response.id} "${safeFolderName}"`);
    return { id: response.id as string, existed: false };
  } catch (err: any) {
    // 409 = folder already exists — treat as success
    const status: number =
      err?.statusCode ?? err?.response?.status ?? err?.code ?? 0;
    const isConflict =
      status === 409 ||
      (err?.message ?? '').toLowerCase().includes('namealreadyexists');

    if (isConflict) {
      console.log(`[VesselDMS] createFolder ↩ existed "${safeFolderName}"`);
      // Fetch the existing folder's ID so callers have it
      try {
        const fullPath = parentPath ? `${parentPath}/${safeFolderName}` : safeFolderName;
       const existing = await withTimeout(
          client.api(`/sites/${siteId}/drives/${driveId}/root:/${encodePath(fullPath)}`).get(),
          GRAPH_CALL_TIMEOUT_MS,
          `lookup existing "${safeFolderName}"`,
        );
        return { id: existing.id as string, existed: true };
      } catch {
        return { id: '', existed: true };
      }
    }
    console.error(`[VesselDMS] createFolder ✗ FAILED "${safeFolderName}" status=${status}`, err?.message ?? err);
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
    const safeName = sanitizeFolderName(node.name);
    const fullPath = `${parentPath}/${safeName}`;
    try {
      const { id, existed } = await createFolder(
        client,
        siteId,
        driveId,
        parentPath,
        safeName,
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
 *   Documents/{MainFolder}/
 *   Documents/{MainFolder}/Common for all ships/{commonTree}
 *   Documents/Kaizen - Knowledge Bank/  (at drive root)
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

  // 1. Documents/{MainFolder} & Documents/{MainFolder}/{commonFolderName}
  for (const mf of MAIN_FOLDERS) {
    try {
      const { id: mfId, existed: mfExisted } = await createFolder(client, siteId, driveId, '', mf.name);
      record({ path: mf.name, id: mfId, status: mfExisted ? 'existed' : 'created' });

      const commonName = (mf as any).commonFolderName ?? COMMON_SHIPS_ROOT;
      const commonPath = `${mf.name}/${commonName}`;
      const { id: cId, existed: cExisted } = await createFolder(client, siteId, driveId, mf.name, commonName);
      record({ path: commonPath, id: cId, status: cExisted ? 'existed' : 'created' });

      if (mf.commonTree.length > 0) {
        await createTree(client, siteId, driveId, commonPath, mf.commonTree, log, onProgress);
      }
    } catch (err: any) {
      record({ path: mf.name, status: 'failed', error: err?.message });
    }
  }

  // 2. Documents/Kaizen - Knowledge Bank (at drive root)
  const kaizenName = 'Kaizen - Knowledge Bank';
  await createFolder(client, siteId, driveId, '', kaizenName)
    .then(({ id, existed }) => record({ path: kaizenName, id, status: existed ? 'existed' : 'created' }))
    .catch(err => record({ path: kaizenName, status: 'failed', error: err?.message }));

  return log;
}

/**
 * Create the full per-vessel folder tree for a new vessel.
 * Structure: Documents/{MainFolder}/{VesselName}/{perVesselTree}
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
  const safeVesselName = sanitizeFolderName(vesselName);
  console.log(`[VesselDMS] createVesselFolders START vessel="${safeVesselName}"`);

  if (!siteId || !driveId) {
    return { vesselName: safeVesselName, siteId, driveId, results: log, success: false };
  }

  // Ensure main folders & common trees exist
  if (!skipRootStructure) {
    const rootLog = await ensureRootStructure(client, siteId, driveId, onProgress);
    log.push(...rootLog);
  }

  // For each main folder, create {MainFolder}/{VesselName}/{perVesselTree}
  for (const mf of MAIN_FOLDERS) {
    const vesselFolderPath = `${mf.name}/${safeVesselName}`;
    try {
      const { id, existed } = await createFolder(client, siteId, driveId, mf.name, safeVesselName);
      const result: FolderResult = { path: vesselFolderPath, id, status: existed ? 'existed' : 'created' };
      log.push(result);
      onProgress?.(result);
      await createTree(client, siteId, driveId, vesselFolderPath, mf.perVesselTree, log, onProgress);
    } catch (err: any) {
      const result: FolderResult = { path: vesselFolderPath, status: 'failed', error: err?.message ?? String(err) };
      log.push(result);
      onProgress?.(result);
    }
  }

  const success = log.every(r => r.status !== 'failed');
  console.log(`[VesselDMS] createVesselFolders DONE vessel="${safeVesselName}" success=${success}`);
  return { vesselName: safeVesselName, siteId, driveId, results: log, success };
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

/**
 * Automatically retry failed folders in a loop until all succeed or the
 * attempt budget is exhausted.
 *
 * @param client         - MSGraph client
 * @param siteId         - SharePoint site ID
 * @param driveId        - Drive ID
 * @param previousResults - Full result array from a previous createVesselFolders / retryFailedFolders call
 * @param maxAttempts    - Maximum number of retry rounds (default 3)
 * @param delayMs        - Milliseconds to wait between rounds (default 3 000 ms)
 * @param onProgress     - Optional callback fired after each round with (attemptNumber, updatedResults)
 * @returns Final merged FolderResult[] — may still contain failures if all attempts exhausted
 */
export async function retryUntilComplete(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  previousResults: FolderResult[],
  maxAttempts = 3,
  delayMs = 3000,
  onProgress?: (attempt: number, totalAttempts: number, results: FolderResult[]) => void,
): Promise<FolderResult[]> {
  let results = [...previousResults];

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const stillFailed = results.filter(r => r.status === 'failed');
    if (stillFailed.length === 0) break; // all resolved — stop early

    // Wait between rounds (no wait before the very first retry so it feels snappy)
    if (attempt > 1) {
      await new Promise<void>(resolve => setTimeout(resolve, delayMs));
    }

    console.log(
      `[VesselDMS] retryUntilComplete attempt ${attempt}/${maxAttempts} — retrying ${stillFailed.length} failed folder(s)`,
    );

    results = await retryFailedFolders(client, siteId, driveId, results);
    onProgress?.(attempt, maxAttempts, results);
  }

  const remaining = results.filter(r => r.status === 'failed').length;
  console.log(
    `[VesselDMS] retryUntilComplete finished — ${remaining} folder(s) still failed after ${maxAttempts} attempt(s)`,
  );

  return results;
}

/**
 * Delete a vessel's root folders across all main departments in SharePoint Online.
 * Calling DELETE on a folder via Graph API automatically moves it into the
 * SharePoint Site Recycle Bin (accessible via Site Contents > Recycle Bin).
 *
 * For each department in MAIN_FOLDERS:
 *   DELETE /sites/{siteId}/drives/{driveId}/root:/{MainFolder}/{vesselName}
 */
export async function deleteVesselFolders(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  vesselName: string,
): Promise<{ success: boolean; deletedPaths: string[]; errors: string[] }> {
  const safeName = sanitizeFolderName(vesselName);
  const deletedPaths: string[] = [];
  const errors: string[] = [];

  for (const mf of MAIN_FOLDERS) {
    // Check both standard name and sanitized name
    const pathsToTry = [
      `${mf.name}/${safeName}`,
      ...(safeName !== vesselName ? [`${mf.name}/${vesselName}`] : []),
    ];

    for (const vesselFolderPath of pathsToTry) {
      const url = `/sites/${siteId}/drives/${driveId}/root:/${encodePath(vesselFolderPath)}`;
      try {
        console.log(`[VesselDMS] deleteVesselFolders → DELETE ${url}`);
        await withTimeout(
          client.api(url).delete(),
          GRAPH_CALL_TIMEOUT_MS,
          `delete folder "${vesselFolderPath}"`,
        );
        deletedPaths.push(vesselFolderPath);
        console.log(`[VesselDMS] deleteVesselFolders ✓ moved to SharePoint Recycle Bin: "${vesselFolderPath}"`);
        break; // Successfully deleted for this main folder
      } catch (err: any) {
        const status: number =
          err?.statusCode ?? err?.response?.status ?? err?.code ?? 0;
        const msg: string = (err?.message || '').toLowerCase();
        // 404 or itemNotFound means it doesn't exist under this path — continue checking or skip
        if (status === 404 || msg.includes('itemnotfound') || msg.includes('resource not found')) {
          console.log(`[VesselDMS] deleteVesselFolders ℹ not found at "${vesselFolderPath}"`);
        } else {
          console.error(`[VesselDMS] deleteVesselFolders ✗ failed "${vesselFolderPath}" status=${status}`, err);
          errors.push(`${vesselFolderPath}: ${err?.message ?? String(err)}`);
        }
      }
    }
  }

  return {
    success: errors.length === 0,
    deletedPaths,
    errors,
  };
}


