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
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
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

export function clearDeltaLink(driveId: string): void {
  try {
    localStorage.removeItem(`${DELTA_LINK_KEY_PREFIX}${driveId}`);
  } catch { /* storage unavailable — non-fatal */ }
}

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

// Global timestamp until which delta queries must wait due to HTTP 429 throttling
let globalThrottledUntil = 0;

export function getDeltaSyncThrottledSeconds(): number {
  const remaining = Math.ceil((globalThrottledUntil - Date.now()) / 1000);
  return remaining > 0 ? remaining : 0;
}

function extractRetryAfterSeconds(err: any): number {
  if (!err) return 0;
  const sec =
    err.retryAfterSeconds ??
    err.body?.error?.retryAfterSeconds ??
    err.body?.error?.innerError?.retryAfterSeconds ??
    err.response?.headers?.get?.('retry-after') ??
    err.headers?.['retry-after'];
  if (typeof sec === 'number' && sec > 0) return sec;
  if (typeof sec === 'string') {
    const parsed = parseInt(sec, 10);
    if (!isNaN(parsed) && parsed > 0) return parsed;
  }
  const is429 =
    err.statusCode === 429 ||
    err.status === 429 ||
    err.code === 'activityLimitReached' ||
    err.body?.error?.code === 'activityLimitReached' ||
    (typeof err.message === 'string' && (err.message.includes('throttled') || err.message.includes('activityLimitReached')));
  return is429 ? 180 : 0;
}

// ── Graph helpers ─────────────────────────────────────────────────────────────

/**
 * Walk all pages of a delta/list response, collecting every item.
 * Returns { items, nextDeltaLink }.
 *
 * Throttling (Throttle.htm / HTTP 429): these calls run with the signed-in
 * user's token, so they count against that USER's SharePoint quota — the
 * same quota their SharePoint pages use. To keep a full-drive baseline from
 * exhausting it:
 *  - pages are requested with $top=1000 (≈5x fewer calls than the default
 *    page size; retried without $top if Graph rejects it),
 *  - pages are paced (PAGE_PAUSE_MS) instead of fired back-to-back,
 *  - a page that fails (e.g. 429) is remembered with the items collected so
 *    far, and the next run RESUMES from that page instead of restarting the
 *    whole drive scan (which used to re-trigger the throttle every time).
 */
const PAGE_PAUSE_MS = 250;
const _pendingDrain: { [key: string]: { url: string; items: any[] } } = {};

function clearPendingDrain(key: string): void {
  delete _pendingDrain[key];
}

function withTop(url: string): string {
  if (/[?&]\$top=/i.test(url) || /[?&]token=/i.test(url) || /[?&]\$skiptoken=/i.test(url)) return url;
  return `${url}${url.indexOf('?') >= 0 ? '&' : '?'}$top=1000`;
}

function isBadRequest(err: any): boolean {
  return err?.statusCode === 400 || err?.status === 400;
}

async function drainPages(
  client: MSGraphClientV3,
  firstUrl: string,
  resumeKey?: string,
): Promise<{ items: any[]; deltaLink: string }> {
  const pending = resumeKey ? _pendingDrain[resumeKey] : undefined;
  const items: any[] = pending ? pending.items : [];
  let url: string = pending ? pending.url : withTop(firstUrl);
  let deltaLink = '';
  let pageNo = 0;

  while (url) {
    // MSGraphClientV3 .api() only accepts relative paths — strip the base if present
    const relativeUrl = url.startsWith('https://')
      ? url.replace('https://graph.microsoft.com/v1.0', '')
      : url;

    let page: any;
    try {
      page = await client.api(relativeUrl).get();
    } catch (err) {
      if (pageNo === 0 && !pending && isBadRequest(err) && url !== firstUrl) {
        url = firstUrl; // $top not accepted here — retry with the default page size
        continue;
      }
      if (resumeKey) _pendingDrain[resumeKey] = { url, items };
      throw err;
    }
    pageNo++;
    if (page.value) items.push(...page.value);

    if (page['@odata.deltaLink']) {
      deltaLink = page['@odata.deltaLink'];
      break;
    }
    url = page['@odata.nextLink'] ?? '';
    if (url) {
      if (resumeKey) _pendingDrain[resumeKey] = { url, items };
      await new Promise<void>(resolve => setTimeout(resolve, PAGE_PAUSE_MS));
    }
  }

  if (resumeKey) clearPendingDrain(resumeKey);
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
    const url = `/sites/${siteId}/drives/${driveId}/root:/${encoded}:/children?$select=id,name,parentReference,folder,size,createdDateTime,lastModifiedDateTime`;
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
    size: typeof item.size === 'number' ? item.size : undefined,
    createdDateTime: item.createdDateTime,
    lastModifiedDateTime: item.lastModifiedDateTime,
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
    : `/sites/${siteId}/drives/${driveId}/root/delta?$select=id,name,parentReference,folder,file,size,createdDateTime,lastModifiedDateTime,deleted`;

  try {
    const { items, deltaLink } = await drainPages(client, startUrl, driveId);

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
  } catch (err: any) {
    // If Graph returns resyncRequired (e.g. token expired), reset to baseline
    const errCode = err?.code || err?.body?.error?.code || '';
    const errMsg = typeof err?.message === 'string' ? err.message : '';
    if ((errCode === 'resyncRequired' || errMsg.includes('resyncRequired')) && storedLink) {
      console.warn('[VesselDMS] Graph delta token expired/resyncRequired. Clearing deltaLink and rebuilding baseline.');
      clearDeltaLink(driveId);
      clearPendingDrain(driveId);
      return pollDelta(client, siteId, driveId, true);
    }
    throw err;
  }
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
 * Create a scheduler that calls pollDelta on an adaptive interval, with built-in
 * HTTP 429 throttling backoff and page visibility awareness.
 *
 * @param client        MSGraphClientV3
 * @param siteId        SharePoint site ID
 * @param driveId       Document library drive ID
 * @param intervalMs    Poll interval in milliseconds (default 60 000 = 60 s)
 * @param onResult      Callback invoked with each DeltaSyncResult
 * @param onError       Optional error callback
 */
export function createSyncScheduler(
  client: MSGraphClientV3,
  siteId: string,
  driveId: string,
  onResult: (result: DeltaSyncResult) => void,
  intervalMs: number = 60_000,
  onError?: (err: unknown) => void,
): SyncScheduler {
  let timerId: ReturnType<typeof setTimeout> | null = null;
  let isRunning = false;
  let isStopped = false;
  let hasBaseline = false;
  const effectiveIntervalMs = Math.max(intervalMs, 30_000);

  const scheduleNext = (delayMs: number): void => {
    if (isStopped) return;
    if (timerId !== null) {
      clearTimeout(timerId);
    }
    timerId = setTimeout(() => {
      void run();
    }, delayMs);
  };

  const run = async (): Promise<DeltaSyncResult | null> => {
    if (isStopped) return null;

    // Concurrency guard: avoid overlapping delta runs
    if (isRunning) return null;

    // Check throttle cooldown
    const now = Date.now();
    if (now < globalThrottledUntil) {
      const waitRemainingSec = Math.ceil((globalThrottledUntil - now) / 1000);
      console.warn(`[VesselDMS] Graph delta sync throttled (429). Cooldown active for ${waitRemainingSec}s more.`);
      scheduleNext(waitRemainingSec * 1000 + 2000);
      return null;
    }

    // Tab visibility guard: skip background polling if user is not actively viewing the tab
    if (typeof document !== 'undefined' && document.hidden) {
      scheduleNext(effectiveIntervalMs);
      return null;
    }

    isRunning = true;
    try {
      const isFirstRun = !hasBaseline;
      const result = await pollDelta(client, siteId, driveId, false);
      result.isBaseline = isFirstRun;
      hasBaseline = true;
      onResult(result);
      scheduleNext(effectiveIntervalMs);
      return result;
    } catch (err: any) {
      const retrySec = extractRetryAfterSeconds(err);
      if (retrySec > 0) {
        globalThrottledUntil = Date.now() + retrySec * 1000;
        console.warn(
          `[VesselDMS] Graph delta sync throttled (429). Pausing delta polling for ${retrySec}s (until ${new Date(globalThrottledUntil).toLocaleTimeString()}).`
        );
        scheduleNext(retrySec * 1000 + 3000);
      } else {
        scheduleNext(effectiveIntervalMs);
      }
      onError?.(err);
      return null;
    } finally {
      isRunning = false;
    }
  };

  // Listen for tab visibility changes: when user comes back to the tab, check if sync is needed
  const onVisibilityChange = (): void => {
    if (typeof document !== 'undefined' && !document.hidden && !isStopped && !isRunning) {
      const now = Date.now();
      if (now >= globalThrottledUntil) {
        void run();
      }
    }
  };

  if (typeof document !== 'undefined' && document.addEventListener) {
    document.addEventListener('visibilitychange', onVisibilityChange);
  }

  return {
    start() {
      if (!isStopped && timerId !== null) return;
      isStopped = false;
      scheduleNext(effectiveIntervalMs);
    },
    stop() {
      isStopped = true;
      if (timerId !== null) {
        clearTimeout(timerId);
        timerId = null;
      }
      if (typeof document !== 'undefined' && document.removeEventListener) {
        document.removeEventListener('visibilitychange', onVisibilityChange);
      }
    },
    triggerNow: run,
  };
}
