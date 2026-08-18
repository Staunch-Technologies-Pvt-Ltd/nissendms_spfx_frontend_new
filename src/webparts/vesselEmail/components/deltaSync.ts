/**
 * deltaSync.ts
 *
 * Graph delta-query sync service.
 *
 * Detects folders created, renamed, moved, or deleted directly in SharePoint
 * Online (outside the SPFx app) and merges them into the correct position in
 * the app's in-memory folder tree — treating SPO as the source of truth.
 *
 * Key design decisions:
 *  - Uses /drives/{driveId}/root/delta so changes anywhere in the drive are
 *    captured, not just under a specific known parent.
 *  - Resolves each changed item's true parent via parentReference.path.
 *  - If a changed item's parent branch hasn't been loaded yet, triggers a
 *    targeted re-fetch of that path rather than guessing its position.
 *  - Persists the deltaLink in localStorage (keyed by driveId) so each poll
 *    only fetches what changed since the last sync.
 *  - Designed to be called on a timer (every 30–60 s) and/or on demand.
 */

import { MSGraphClientV3 } from '@microsoft/sp-http';

// ── Types ─────────────────────────────────────────────────────────────────────

export interface SpoFolderNode {
  id: string;
  name: string;
  parentId: string | null;
  /** True for folders; false for files emitted by the Graph delta feed. */
  isFolder: boolean;
  /** Full server-relative path, e.g. "/sites/MySite/Shared Documents/Vessel Management/..." */
  serverRelativePath: string;
  children: SpoFolderNode[];
  deleted?: boolean;
}

export interface DeltaSyncResult {
  added: SpoFolderNode[];
  updated: SpoFolderNode[];
  deleted: string[];   // item IDs
  newDeltaLink: string;
  /** True when this result is from a full baseline scan (not an incremental delta). */
  isBaseline?: boolean;
}

// ── Delta link persistence (localStorage) ────────────────────────────────────

const DELTA_LINK_KEY_PREFIX = 'vesselDMS_deltaLink_';

function loadDeltaLink(driveId: string): string | null {
  try {
    return localStorage.getItem(`${DELTA_LINK_KEY_PREFIX}${driveId}`);
  } catch {
    return null;
  }
}

function saveDeltaLink(driveId: string, link: string): void {
  try {
    localStorage.setItem(`${DELTA_LINK_KEY_PREFIX}${driveId}`, link);
  } catch { /* storage unavailable — non-fatal */ }
}

// ── Graph helpers ─────────────────────────────────────────────────────────────

/**
 * Walk all pages of a delta/list response, collecting every item.
 * Returns { items, nextDeltaLink }.
 */
async function drainPages(
  client: MSGraphClientV3,
  firstUrl: string,
): Promise<{ items: any[]; deltaLink: string }> {
  const items: any[] = [];
  let url: string = firstUrl;
  let deltaLink = '';

  while (url) {
    // MSGraphClientV3 .api() only accepts relative paths — strip the base if present
    const relativeUrl = url.startsWith('https://')
      ? url.replace('https://graph.microsoft.com/v1.0', '')
      : url;

    const page: any = await client.api(relativeUrl).get();
    if (page.value) items.push(...page.value);

    if (page['@odata.deltaLink']) {
      deltaLink = page['@odata.deltaLink'];
      break;
    }
    url = page['@odata.nextLink'] ?? '';
  }

  return { items, deltaLink };
}

/**
 * Fetch children of a specific path to rebuild an unloaded branch.
 */
export async function fetchFolderChildren(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  folderPath: string,
): Promise<SpoFolderNode[]> {
  // Strip leading slash and any drive-relative prefix like /drives/{id}/root:
  const cleanPath = folderPath
    .replace(/^\/+/, '')
    .replace(/^drives\/[^/]+\/root:\/?/i, '');

  const encoded = cleanPath
    .split('/')
    .map(s => encodeURIComponent(s))
    .join('/');

  try {
    const url = `/sites/${siteId}/drives/${driveId}/root:/${encoded}:/children?$select=id,name,parentReference,folder`;
    const result: any = await client.api(url).get();
    return (result.value ?? []).filter((item: any) => !!item.folder).map((item: any) => graphItemToNode(item));
  } catch {
    return [];
  }
}

// ── Item → SpoFolderNode mapping ──────────────────────────────────────────────

/**
 * Extract the server-relative path from a Graph item's parentReference.
 * parentReference.path looks like:
 *   "/drives/{driveId}/root:/Vessel Management/Folder-1 Technical & Crewing"
 * We strip the "/drives/{driveId}/root:" prefix to get the clean path.
 */
function resolveServerPath(item: any): string {
  const raw: string = item.parentReference?.path ?? '';
  const rootMarker = '/root:';
  const idx = raw.indexOf(rootMarker);
  const parentPath = idx !== -1 ? raw.substring(idx + rootMarker.length) : '';
  // Decode URI components (Graph encodes spaces as %20 etc.)
  const decodedParent = decodeURIComponent(parentPath);
  return decodedParent ? `${decodedParent}/${item.name}` : `/${item.name}`;
}

function graphItemToNode(item: any): SpoFolderNode {
  return {
    id: item.id,
    name: item.name,
    parentId: item.parentReference?.id ?? null,
    isFolder: !!item.folder,
    serverRelativePath: resolveServerPath(item),
    children: [],
    deleted: !!item.deleted,
  };
}

// ── Main sync function ────────────────────────────────────────────────────────

/**
 * Poll Graph for changes since the last sync.
 *
 * On first call (no stored deltaLink) it performs a full drive scan to build
 * the initial deltaLink baseline — this may be slow for large drives but only
 * happens once per browser session.
 *
 * @param client   MSGraphClientV3
 * @param siteId   SharePoint site ID
 * @param driveId  Document library drive ID
 * @returns DeltaSyncResult with added/updated/deleted items and the new deltaLink
 */
export async function pollDelta(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  forceBaseline: boolean = false,
): Promise<DeltaSyncResult> {
  // The delta link survives a browser reload, but the in-memory folder map
  // does not. A new web part instance must therefore build a full baseline
  // before it can safely use the persisted incremental token.
  const storedLink = forceBaseline ? null : loadDeltaLink(driveId);

  // Use stored deltaLink for incremental sync, or start a fresh delta scan
  const startUrl = storedLink
    ? storedLink
    : `/sites/${siteId}/drives/${driveId}/root/delta?$select=id,name,parentReference,folder,file,deleted`;

  const { items, deltaLink } = await drainPages(client, startUrl);

  if (deltaLink) saveDeltaLink(driveId, deltaLink);

  const added: SpoFolderNode[] = [];
  const updated: SpoFolderNode[] = [];
  const deleted: string[] = [];

  for (const item of items) {
    // Track folders AND files — file changes from SPO-direct uploads must trigger a reload
    if (item.deleted) {
      deleted.push(item.id);
    } else if (item.folder) {
      const node = graphItemToNode(item);
      added.push(node);
    } else if (item.file) {
      // File added/modified in SPO directly — emit as a synthetic node so the
      // caller's debounced _loadData fires and the list view refreshes.
      const node = graphItemToNode(item);
      added.push(node);
    }
  }

  return { added, updated, deleted, newDeltaLink: deltaLink };
}

// ── Tree merge helpers (used by the React component) ─────────────────────────

/**
 * Insert or update a node at the correct position in a flat id→node map.
 * The map is keyed by folder ID for O(1) lookups.
 *
 * If the node's parent is not yet in the map (branch not loaded), returns
 * the parentId so the caller can trigger a targeted re-fetch.
 */
export function mergeNodeIntoMap(
  map: Map<string, SpoFolderNode>,
  node: SpoFolderNode,
): { missingParentId: string | null } {
  // Update or insert the node itself
  const existing = map.get(node.id);
  if (existing) {
    // Update name/path in place, preserve children
    existing.name = node.name;
    existing.serverRelativePath = node.serverRelativePath;
    existing.parentId = node.parentId;
    existing.isFolder = node.isFolder;
  } else {
    map.set(node.id, { ...node, children: [] });
  }

  // Wire into parent's children array
  if (node.parentId) {
    const parent = map.get(node.parentId);
    if (!parent) {
      // Parent branch not loaded yet — caller must re-fetch
      return { missingParentId: node.parentId };
    }
    const alreadyChild = parent.children.some(c => c.id === node.id);
    if (!alreadyChild) {
      parent.children.push(map.get(node.id)!);
    }
  }

  return { missingParentId: null };
}

/**
 * Remove a deleted node (and all its descendants) from the map.
 */
export function removeNodeFromMap(
  map: Map<string, SpoFolderNode>,
  deletedId: string,
): void {
  const node = map.get(deletedId);
  if (!node) return;

  // Recursively remove children first
  for (const child of node.children) {
    removeNodeFromMap(map, child.id);
  }

  // Detach from parent
  if (node.parentId) {
    const parent = map.get(node.parentId);
    if (parent) {
      parent.children = parent.children.filter(c => c.id !== deletedId);
    }
  }

  map.delete(deletedId);
}

// ── Periodic sync scheduler ───────────────────────────────────────────────────

export interface SyncScheduler {
  start: () => void;
  stop: () => void;
  triggerNow: () => Promise<DeltaSyncResult | null>;
}

/**
 * Create a scheduler that calls pollDelta on a fixed interval.
 *
 * @param client        MSGraphClientV3
 * @param siteId        SharePoint site ID
 * @param driveId       Document library drive ID
 * @param intervalMs    Poll interval in milliseconds (default 45 000 = 45 s)
 * @param onResult      Callback invoked with each DeltaSyncResult
 * @param onError       Optional error callback
 */
export function createSyncScheduler(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  onResult: (result: DeltaSyncResult) => void,
  intervalMs: number = 45_000,
  onError?: (err: unknown) => void,
): SyncScheduler {
  let timerId: ReturnType<typeof setInterval> | null = null;
  let hasBaseline = false;

  const run = async (): Promise<DeltaSyncResult | null> => {
    try {
      // isBaseline = true on the very first run of this scheduler instance,
      // regardless of whether a stored deltaLink exists. The in-memory folder
      // map is always empty on a fresh page load, so we treat the first result
      // as a baseline (populate the map) and skip new-vessel alerts to avoid
      // false positives before the vessels list has loaded from the DB.
      const isFirstRun = !hasBaseline;
      const result = await pollDelta(client, siteId, driveId, false);
      result.isBaseline = isFirstRun;
      hasBaseline = true;
      onResult(result);
      return result;
    } catch (err) {
      onError?.(err);
      return null;
    }
  };

  return {
    start() {
      if (timerId !== null) return;
      timerId = setInterval(() => { run().catch(() => undefined); }, intervalMs);
    },
    stop() {
      if (timerId !== null) {
        clearInterval(timerId);
        timerId = null;
      }
    },
    triggerNow: run,
  };
}
