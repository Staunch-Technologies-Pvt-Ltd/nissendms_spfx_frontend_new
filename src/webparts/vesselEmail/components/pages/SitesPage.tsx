/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable @typescript-eslint/no-unused-expressions */
import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { Icon } from '@fluentui/react/lib/Icon';
import { clay } from '../clayTheme';
import { SiteManagementSection } from './SiteManagementSection';
import { DmsPageHeader, dmsControlStyle, dmsBtn, dmsRowBtn, dmsTone, DMS_ON_ACCENT } from '../dmsDesignSystem';

 type Site = { id: string; site_key?: string; name?: string; display_name: string; web_url?: string; description?: string; thumbnail?: string; is_default?: boolean };
type Drive = { id: string; name: string; web_url?: string; is_system?: boolean; item_count?: number | null };
type FolderCounts = {
  direct_subfolders: number;
  direct_files: number;
  total_subfolders: number;
  total_files: number;
  is_estimated?: boolean;
};
type SummaryCounts = {
  direct_folders: number;
  direct_files: number;
  total_folders: number;
  total_files: number;
};
type Item = {
  id: string;
  name: string;
  folder?: { childCount?: number } | null;
  folder_counts?: FolderCounts | null;
  file?: object;
  size?: number;
  web_url?: string;
  download_url?: string;
  tags?: Record<string, string>;
};
type FolderCacheEntry = {
  items: Item[];
  summaryCounts: SummaryCounts | null;
  detectedVessel: string;
  detectedTags: Tags;
  timestamp: number;
};
type Context = { site: Site; drive: Drive };
type Tags = { department: string; vessel: string; group: string; category: string };
type ScanResult = {
  filename?: string;
  item_id: string;
  parent_path?: string;
  subfolder_name?: string;
  confidence: number;
  error?: string;
  status?: string;
  current_tags?: Tags;
  current_values?: Tags;
  proposed_tags?: Tags;
  ocr_suggestion?: Record<string, { value: string; confidence: number }>;
  path_suggestion?: Record<string, { value: string; label: string }>;
  /** True when vessel was matched from filename alias (e.g. N-2119 → Bow Fighter) but is absent from file text. */
  vessel_in_filename_only?: boolean;
};

type TagFeedEntry = {
  name: string;
  status: 'pending' | 'ok' | 'failed';
  error?: string;
};

type TaggingModal = {
  title: string;
  feed: TagFeedEntry[];
  total: number;
  finished: boolean;
  summary: string;
} | null;
type TagFailure = {
  file_id: string;
  filename: string;
  parent_path: string;
  error_reason: string;
  attempt_count: number;
  last_attempted_at?: string | null;
};



const emptyTags: Tags = { department: '', vessel: '', group: '', category: '' };
const tagNames: Array<[keyof Tags, string]> = [
  ['department', 'Department'],
  ['vessel', 'Vessel'],
  ['group', 'Group'],
  ['category', 'Category'],
];

/** Placeholder values that mean "not really set" for Department/Vessel/Group. */
const CORE_TAG_PLACEHOLDERS = new Set(['', 'to be classified', 'unknown', 'n/a']);

/**
 * True when a single tag field is missing. Mirrors the backend's
 * `_tags_need_attention` policy: Department/Vessel/Group are missing when
 * blank or a non-value placeholder; Category is only missing when
 * completely blank ("To Be Classified" is an intentional already-reviewed
 * state and does NOT count as missing).
 */
function isTagFieldMissing(item: Item, field: keyof Tags): boolean {
  const value = (item.tags?.[field] || '').trim().toLowerCase();
  if (field === 'category') return value === '';
  return CORE_TAG_PLACEHOLDERS.has(value);
}

/** Auto-tag review targets only files without a vessel tag. */
function hasAnyMissingTag(item: Item): boolean {
  return isTagFieldMissing(item, 'vessel');
}

function getInitialSite(host: VesselEmail): Site[] {
  const siteId = host.props.siteId;
  if (!siteId) return [];
  const siteUrl = host.props.siteUrl || '';
  const siteName = siteUrl.split('/').filter(Boolean).pop() || siteId;
  return [{
    id: siteId,
    display_name: siteName,
    web_url: siteUrl,
    description: 'Current SharePoint site',
  }];
}

function getSiteName(site: Site): string {
  const candidates = [site.name, site.display_name].filter(Boolean) as string[];
  const readable = candidates.find(value => !/^https?:\/\//i.test(value) && value !== site.id);
  if (readable) return readable;

  try {
    const url = new URL(site.web_url || '');
    const lastSegment = url.pathname.split('/').filter(Boolean).pop();
    if (lastSegment) return decodeURIComponent(lastSegment).replace(/[-_]+/g, ' ');
    return url.hostname;
  } catch {
    return site.display_name || site.id;
  }
}

export function SitesPage({ host }: { host: VesselEmail }): React.ReactElement {
  const [sites, setSites] = React.useState<Site[]>([]);
  const [query, setQuery] = React.useState('');
  const [expanded, setExpanded] = React.useState<string | null>(null);
  const [drives, setDrives] = React.useState<Record<string, Drive[]>>({});
  // SharePoint auto-provisions "system" libraries (Site Assets, Style
  // Library, Form Templates, ...) on every site alongside the real content
  // library — a site with versioning/retention/publishing features on can
  // easily show 4-6 of these. They're never where a user's files live, so
  // they're hidden by default per-site and only shown on request via this
  // toggle, keyed by site id.
  const [showSystemLibraries, setShowSystemLibraries] = React.useState<Record<string, boolean>>({});
  const [context, setContext] = React.useState<Context | null>(null);
  const [items, setItems] = React.useState<Item[]>([]);
  const [crumbs, setCrumbs] = React.useState<Array<{ id: string; name: string }>>([]);
  const [loading, setLoading] = React.useState(false);
  const [sitesLoading, setSitesLoading] = React.useState(true);
    const [selectedSiteKey, setSelectedSiteKey] = React.useState('');
  const [message, setMessage] = React.useState('');
  const [selected, setSelected] = React.useState<Set<string>>(new Set());
  const [editing, setEditing] = React.useState<string | null>(null);
  const [tagDraft, setTagDraft] = React.useState<Tags>(emptyTags);
  const [recursive, setRecursive] = React.useState(false);
  const [rescanAll, setRescanAll] = React.useState(false);
  const [tagFailures, setTagFailures] = React.useState<TagFailure[]>([]);
  const [scanResults, setScanResults] = React.useState<ScanResult[]>([]);
  const [fieldChoices, setFieldChoices] = React.useState<Record<string, Record<keyof Tags, string>>>({});
  const [vesselOptions, setVesselOptions] = React.useState<string[]>([]);
  const [bulkVessel, setBulkVessel] = React.useState('');
  const [detectedVessel, setDetectedVessel] = React.useState('');
  const [detectedTags, setDetectedTags] = React.useState<Tags>(emptyTags);
  const [taggingModal, setTaggingModal] = React.useState<TaggingModal>(null);
  const [confirmOverwrite, setConfirmOverwrite] = React.useState<{ count: number; vessel: string } | null>(null);
  const [confirmScanModal, setConfirmScanModal] = React.useState<{
    mode: 'autoTagAndReview' | 'scan' | 'autoTagFromPath';
    targetItemIds: string[];
    isSelection: boolean;
    count: number;
    missingCount?: number;
    names: string[];
    isRecursive: boolean;
  } | null>(null);
  const [summaryCounts, setSummaryCounts] = React.useState<SummaryCounts | null>(null);
  const [countsLoading, setCountsLoading] = React.useState(false);

  // In-memory client cache for folder views: enables 0ms instant folder switching
  const folderCacheRef = React.useRef<Map<string, FolderCacheEntry>>(new Map());

  // Tracks the in-flight loadFolder/fetchSubfolderCounts request, if any.
  // Neither used to be cancelled when a newer one started, so clicking
  // through several sites/drives/folders in a row (exactly what "navigate
  // to other sites" does) left every earlier click's /children and
  // /subfolder-counts requests running to completion in the background, all
  // competing with the latest click's own request for the browser's small
  // per-origin connection pool — and adding to the real SharePoint Online
  // request volume, which is what can tip the tenant into Graph/SPO
  // throttling (the "Something's not right" Throttle.htm page). Aborting the
  // previous one before starting a new one is the same fix already applied
  // to the Documents module's site switcher (VesselEmail.tsx
  // _switchDocumentSite / _siteSwitchAbort).
  const loadFolderAbortRef = React.useRef<AbortController | null>(null);
  const subfolderCountsAbortRef = React.useRef<AbortController | null>(null);

  const api = host._base();
  const headers = host._headers();

  const loadTagFailures = React.useCallback(async (): Promise<void> => {
    if (!context) return;
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/tag-failures?status=needs_retry`, { headers });
      if (response.ok) setTagFailures((await response.json()).failures || []);
    } catch {
      // Retry queue is supplementary; keep the main folder view usable.
    }
  }, [api, context?.site.id, context?.drive.id]);

  React.useEffect(() => { void loadTagFailures(); }, [loadTagFailures]);

  const retryFailures = async (fileIds: string[] = []): Promise<void> => {
    if (!context) return;
    setLoading(true);
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/retry-tag-failures`, {
        method: 'POST', headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify({ file_ids: fileIds }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Retry failed');
      setMessage(`Retry completed for ${data.updated_count || 0} file(s).`);
      await loadTagFailures();
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Retry failed');
    } finally { setLoading(false); }
  };

  const dismissFailures = async (fileIds: string[]): Promise<void> => {
    if (!context) return;
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/dismiss-tag-failures`, {
        method: 'POST', headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify({ file_ids: fileIds, reason: 'Dismissed from Needs Attention' }),
      });
      if (!response.ok) throw new Error('Could not dismiss failed files');
      await loadTagFailures();
    } catch (error) { setMessage(error instanceof Error ? error.message : 'Could not dismiss failed files'); }
  };

  /** Async backfill for recursive counts so cold-cache folder views never block initial rendering */
  const fetchSubfolderCounts = React.useCallback(async (ctx: Context, folderId: string): Promise<void> => {
    subfolderCountsAbortRef.current?.abort();
    const controller = new AbortController();
    subfolderCountsAbortRef.current = controller;
    setCountsLoading(true);
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(ctx.site.id)}/drives/${encodeURIComponent(ctx.drive.id)}/folders/${encodeURIComponent(folderId)}/subfolder-counts`, { headers, signal: controller.signal });
      if (controller.signal.aborted) return;
      if (response.ok) {
        const data = await response.json();
        const countsMap: Record<string, FolderCounts> = data.counts || {};
        const cacheKey = `${ctx.drive.id}:${folderId}`;
        const cached = folderCacheRef.current.get(cacheKey);
        if (cached) {
          cached.items = cached.items.map(it => {
            if (it.folder && countsMap[it.id]) {
              return { ...it, folder_counts: countsMap[it.id] };
            }
            return it;
          });
          if (data.summary_counts) {
            cached.summaryCounts = data.summary_counts;
          }
        }
        setItems(previous => previous.map(it => {
          if (it.folder && countsMap[it.id]) {
            return { ...it, folder_counts: countsMap[it.id] };
          }
          return it;
        }));
        if (data.summary_counts) {
          setSummaryCounts(data.summary_counts);
        }
      }
    } catch {
      // Non-fatal (including an abort from a newer call superseding this
      // one): UI gracefully falls back to direct properties
    } finally {
      // Only the still-current call should clear the spinner — an aborted,
      // superseded call finishing its cleanup must not flip countsLoading
      // to false out from under the newer call that's still in flight.
      if (subfolderCountsAbortRef.current === controller) {
        setCountsLoading(false);
      }
    }
  }, [api, headers]);

  /**
   * Computes the number of files inside a folder item:
   * 1. Checks host.state.spoFolderMap (delta sync cache) to count files (recursively across subfolders).
   * 2. If not in spoFolderMap or count is 0, falls back to Graph childCount property.
   * 3. Always defaults to 0 so "0 files" is clearly displayed when empty (matching Documents module).
   */
  const getFolderFileCount = React.useCallback((item: Item): number => {
    if (host?.state?.spoFolderMap && item.id) {
      const node = host.state.spoFolderMap.get(item.id);
      if (node) {
        let totalFiles = 0;
        const visited = new Set<string>();
        const countFiles = (n: { id?: string; children?: Array<{ id?: string; isFolder?: boolean }> }): void => {
          if (n.id) {
            if (visited.has(n.id)) return;
            visited.add(n.id);
          }
          for (const c of (n.children || [])) {
            if (c.id && host._appDeletedItemIds && host._appDeletedItemIds.has(c.id)) continue;
            if (!c.isFolder) {
              totalFiles++;
            } else if (c.id && host.state.spoFolderMap.has(c.id)) {
              countFiles(host.state.spoFolderMap.get(c.id)!);
            }
          }
        };
        countFiles(node);
        if (totalFiles > 0) return totalFiles;
      }
    }
    const cc = (item.folder as any)?.childCount;
    return typeof cc === 'number' ? cc : 0;
  }, [host]);

  const loadSites = React.useCallback(async (): Promise<void> => {
    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), 25000);
    setSitesLoading(true);
    try {
      const response = await fetch(`${api}/api/sites?limit=500`, { headers, signal: controller.signal });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Could not load sites');
      const discoveredSites: Site[] = data.sites || [];
      setSites(discoveredSites);
      const activeSite = host.state.activeDocumentSite || '';
      const normalizeSiteValue = (value: string | null | undefined): string =>
        String(value || '').trim().toLowerCase().replace(/[_-]+/g, ' ').replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
      const defaultSite = discoveredSites.find(site => {
        const candidates = [site.site_key, site.id, site.name, site.display_name, site.web_url];
        return candidates.some(value => normalizeSiteValue(value) === normalizeSiteValue(activeSite));
      }) || discoveredSites.find(site => site.is_default) || discoveredSites[0];
      const nextDefaultValue = defaultSite?.site_key || defaultSite?.id || defaultSite?.name || defaultSite?.display_name || '';
      setSelectedSiteKey(nextDefaultValue);
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') {
        setMessage('Site discovery is taking longer than expected. The current site is still available.');
      } else {
        setMessage(error instanceof Error ? error.message : 'Could not load sites');
      }
    } finally {
      window.clearTimeout(timeoutId);
      setSitesLoading(false);
    }
  }, [api]);

  React.useEffect(() => { void loadSites(); }, [loadSites]);

  const loadVesselOptions = React.useCallback(async (siteId: string): Promise<void> => {
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(siteId)}/term-store-vessels`, { headers });
      if (response.ok) {
        const data = await response.json();
        setVesselOptions(data.vessels || []);
      }
    } catch {
      // silently ignore — vessel list is a convenience feature
    }
  }, [api]);

  React.useEffect(() => {
    if (context?.site.id) { void loadVesselOptions(context.site.id); }
  }, [context?.site.id, loadVesselOptions]);

  const loadDrives = async (site: Site): Promise<void> => {
    if (expanded === site.id) { setExpanded(null); return; }
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(site.id)}/drives`, { headers });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Could not load libraries');
      setDrives(previous => ({ ...previous, [site.id]: data.drives || [] }));
      setExpanded(site.id);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Could not load libraries');
    }
  };

  /** Navigate into a folder and reset scan state — uses client cache for 0ms instant opening */
  const loadFolder = async (
    nextContext: Context,
    folderId: string,
    name: string,
    nextCrumbs: Array<{ id: string; name: string }>,
    forceRefresh: boolean = false
  ): Promise<void> => {
    const cacheKey = `${nextContext.drive.id}:${folderId}`;
    const cached = folderCacheRef.current.get(cacheKey);
    const isFresh = cached && (Date.now() - cached.timestamp < 120_000);

    if (!forceRefresh && cached) {
      // Instant render from client cache — zero delay
      setContext(nextContext);
      setItems(cached.items);
      setSummaryCounts(cached.summaryCounts);
      setCrumbs(nextCrumbs.length ? nextCrumbs : [{ id: 'root', name }]);
      setSelected(new Set());
      setDetectedVessel(cached.detectedVessel);
      if (cached.detectedVessel) {
        setBulkVessel(cached.detectedVessel);
      }
      setDetectedTags(cached.detectedTags);
      setMessage('');
      setLoading(false);

      if (isFresh) {
        return;
      }
      // Revalidate in background without wiping UI or blocking user
    } else {
      setLoading(true);
      setMessage('');
      setScanResults([]);
      setFieldChoices({});
    }

    // Abort whatever folder load is still in flight before starting this
    // one — clicking through several sites/drives/folders in a row (see the
    // note on loadFolderAbortRef above) used to leave every earlier click's
    // request running uncancelled.
    loadFolderAbortRef.current?.abort();
    const controller = new AbortController();
    loadFolderAbortRef.current = controller;
    const timeoutId = window.setTimeout(() => controller.abort(), 15000);
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(nextContext.site.id)}/drives/${encodeURIComponent(nextContext.drive.id)}/folders/${encodeURIComponent(folderId)}/children`, { headers, signal: controller.signal });
      const data = await response.json();
      // Superseded by a newer loadFolder call while this one was in flight —
      // drop the (now stale) result instead of applying it over whatever
      // the newer call has already rendered.
      if (loadFolderAbortRef.current !== controller) return;
      if (!response.ok) throw new Error(data.detail || 'Could not load folder');
      setContext(nextContext);
      const childItems: Item[] = data.items || [];
      const sumCounts = data.summary_counts || null;
      const detV = data.detected_vessel || '';
      const detTags = data.detected_tags || emptyTags;

      // Update client cache
      folderCacheRef.current.set(cacheKey, {
        items: childItems,
        summaryCounts: sumCounts,
        detectedVessel: detV,
        detectedTags: detTags,
        timestamp: Date.now(),
      });

      setItems(childItems);
      setSummaryCounts(sumCounts);
      setCrumbs(nextCrumbs.length ? nextCrumbs : [{ id: 'root', name }]);
      setSelected(new Set());
      setDetectedVessel(detV);
      if (detV) {
        setBulkVessel(detV);
      }
      setDetectedTags(detTags);

      // Trigger non-blocking async count resolution only if needed
      const hasUncached = childItems.some(i => i.folder && (!i.folder_counts || i.folder_counts.is_estimated));
      if (hasUncached) {
        void fetchSubfolderCounts(nextContext, folderId);
      }

      // NOTE: this used to also fire a speculative /children fetch for every
      // direct subfolder (up to 6) 120ms after opening this folder, purely so
      // the *next* click would render instantly from cache. Combined with
      // fetchSubfolderCounts() above (which already fans out a bounded but
      // real burst of Graph calls for recursive counts) it meant every folder
      // open could kick off a dozen-plus concurrent requests — visibly
      // queuing up in the Network tab and contributing to Graph 429
      // throttling, which made the *current* folder's own contents take a
      // long time to appear. Removed: the 120s client cache below still makes
      // a re-visited folder instant; a folder you haven't opened yet just
      // fetches normally on click instead of speculatively in the background.
    } catch (error) {
      // A superseded call's abort should be silent, not surfaced as a
      // "timed out" error over whatever the newer call is now showing.
      if (loadFolderAbortRef.current !== controller) return;
      if (!cached) {
        setMessage(error instanceof DOMException && error.name === 'AbortError'
          ? 'SharePoint folder loading timed out. Please retry.'
          : error instanceof Error ? error.message : 'Could not load folder');
      }
    } finally {
      window.clearTimeout(timeoutId);
      // Only the still-current call should clear the spinner/cache-revalidate
      // state — a superseded call's cleanup must not flip `loading` to false
      // out from under the newer call that's still in flight.
      if (loadFolderAbortRef.current === controller) {
        setLoading(false);
      }
    }
  };

  /** Reload folder items without clearing scanResults so the review panel stays visible after scan */
  const refreshFolder = async (ctx: Context, folderId: string, currentCrumbs: Array<{ id: string; name: string }>): Promise<void> => {
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(ctx.site.id)}/drives/${encodeURIComponent(ctx.drive.id)}/folders/${encodeURIComponent(folderId)}/children`, { headers });
      const data = await response.json();
      if (response.ok) {
        const childItems: Item[] = data.items || [];
        const sumCounts = data.summary_counts || null;
        const detV = data.detected_vessel || '';
        const detTags = data.detected_tags || emptyTags;
        const cacheKey = `${ctx.drive.id}:${folderId}`;
        folderCacheRef.current.set(cacheKey, {
          items: childItems,
          summaryCounts: sumCounts,
          detectedVessel: detV,
          detectedTags: detTags,
          timestamp: Date.now(),
        });
        setItems(childItems);
        if (data.summary_counts) {
          setSummaryCounts(data.summary_counts);
        }
        setCrumbs(currentCrumbs);
        setSelected(new Set());
        if (data.detected_vessel) {
          setDetectedVessel(data.detected_vessel);
        }
        if (data.detected_tags) {
          setDetectedTags(data.detected_tags);
        }

        const hasUncached = childItems.some(i => i.folder && (!i.folder_counts || i.folder_counts.is_estimated));
        if (hasUncached) {
          void fetchSubfolderCounts(ctx, folderId);
        }
      }
    } catch {
      // Keep existing items if refresh fails
    }
  };

  const chooseDrive = (site: Site, drive: Drive): void => {
    void loadFolder({ site, drive }, 'root', drive.name, [{ id: 'root', name: drive.name }]);
  };

  /** Returns the direct SharePoint Online URL for the current breadcrumb folder. */
  const getCurrentFolderSharePointUrl = (): string => {
    if (!context) return '';
    const driveBase = context.drive.web_url || '';
    if (!driveBase) return '';
    if (crumbs.length <= 1) return driveBase;
    try {
      const driveUrl = new URL(driveBase);
      const basePath = driveUrl.pathname.replace(/\/+$/, '');
      // Append each breadcrumb folder name (skip the root crumb which is the drive name)
      const relativeSegments = crumbs.slice(1).map(c => encodeURIComponent(c.name)).join('/');
      return `${driveUrl.origin}${basePath}/${relativeSegments}`;
    } catch {
      return driveBase;
    }
  };

  const openSiteFile = (item: Item): void => {
    if (item.web_url) { window.open(item.web_url, '_blank'); return; }
    if (item.download_url) { window.open(item.download_url, '_blank'); return; }
    if (!context) return;
    const url = `${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/items/${encodeURIComponent(item.id)}/content`;
    window.open(url, '_blank');
  };

  const openSharePointLocation = (item: Item): void => {
    if (item.folder) {
      // Folders: item.web_url from Graph is already the correct direct SharePoint Online folder URL
      const targetUrl = item.web_url || getCurrentFolderSharePointUrl();
      if (targetUrl) { window.open(targetUrl, '_blank'); return; }
    } else {
      // Files: navigate to the parent folder in SharePoint Online
      if (item.web_url) {
        try {
          const fileUrl = new URL(item.web_url);
          const lastSlash = fileUrl.pathname.lastIndexOf('/');
          const parentPath = lastSlash > 0 ? fileUrl.pathname.substring(0, lastSlash) : '';
          if (parentPath) {
            window.open(`${fileUrl.origin}${parentPath}`, '_blank');
            return;
          }
        } catch {
          // fall through
        }
      }
      // Fallback: open the current breadcrumb folder
      const parentUrl = getCurrentFolderSharePointUrl();
      if (parentUrl) { window.open(parentUrl, '_blank'); return; }
    }
    if (context?.drive.web_url) {
      window.open(context.drive.web_url, '_blank');
    } else {
      setMessage(`SharePoint URL not available for ${item.name}`);
    }
  };

  const saveTags = async (item: Item): Promise<void> => {
    if (!context) return;
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/items/${encodeURIComponent(item.id)}/tags`, {
        method: 'PATCH',
        headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify(tagDraft),
      });
      const data = await response.json();
      if (!response.ok) { setMessage(data.detail || 'Could not save tags'); return; }
      if (data.ok === false) { setMessage(data.metadata_patch?.error || 'SharePoint did not save the tags'); return; }
      setItems(previous => previous.map(current => current.id === item.id ? { ...current, tags: data.tags } : current));
      if (crumbs.length > 0) {
        folderCacheRef.current.delete(`${context.drive.id}:${crumbs[crumbs.length - 1].id}`);
      }
      setEditing(null);
      setMessage(`Tags saved for ${item.name}`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Could not save tags');
    }
  };

  const promptScan = (targetItemIds?: string[]): void => {
    if (!context) return;
    const ids = targetItemIds && targetItemIds.length > 0
      ? targetItemIds
      : (selected.size > 0 ? Array.from(selected) : items.map(i => i.id));
    if (ids.length === 0) {
      setMessage('No items to scan.');
      return;
    }

    const isSelection = (targetItemIds && targetItemIds.length > 0) || selected.size > 0;
    const targetItems = ids.map(id => items.find(it => it.id === id)).filter(Boolean) as Item[];
    const hasFolder = targetItems.some(it => it.folder);
    const isRecursive = hasFolder ? true : (!isSelection && recursive);
    const directFiles = targetItems.filter(item => !item.folder);
    const missingCount = directFiles.filter(item => hasAnyMissingTag(item)).length;
    const isSingleFile = isSelection && !hasFolder && ids.length === 1;

    // If 1 single file selected and it already has a vessel tag, scan it directly
    if (isSingleFile && missingCount === 0) {
      void scan(ids, 'all');
      return;
    }
    // If multiple files and none missing tags and no folders, scan directly
    if (!isSingleFile && missingCount === 0 && !hasFolder) {
      void scan(ids, 'all');
      return;
    }

    let fileCount = 0;
    if (isSelection) {
      if (!hasFolder) {
        fileCount = targetItems.length;
      } else {
        fileCount = summaryCounts?.total_files ?? ids.length;
      }
    } else {
      fileCount = summaryCounts?.total_files ?? items.filter(it => !it.folder).length;
    }

    setConfirmScanModal({
      mode: 'scan',
      targetItemIds: ids,
      isSelection,
      count: fileCount,
      missingCount,
      names: targetItems.map(it => it.name),
      isRecursive,
    });
  };

  const scan = async (targetItemIds?: string[], scope: 'missing_only' | 'all' = 'missing_only'): Promise<void> => {
    if (!context) return;
    // If nothing is checked, scan all visible items (folders + files) at current level
    const idsToScan = targetItemIds && targetItemIds.length > 0
      ? targetItemIds
      : (selected.size > 0 ? Array.from(selected) : items.map(i => i.id));
    if (idsToScan.length === 0) {
      setMessage('No items to scan.');
      return;
    }
    const isSelection = (targetItemIds && targetItemIds.length > 0) || selected.size > 0;
    const targetItems = idsToScan.map(id => items.find(it => it.id === id)).filter(Boolean) as Item[];
    const hasFolder = targetItems.some(id => id.folder);
    const scanRecursive = hasFolder ? true : (!isSelection && recursive);
    const isSingleFile = isSelection && !hasFolder && idsToScan.length === 1;

    setLoading(true);
    setFieldChoices({});
    setMessage(isSingleFile
      ? `Scanning "${targetItems[0]?.name || ''}" with OCR and AI classification...`
      : `Scanning ${idsToScan.length} item(s) with OCR and AI classification...`);
    const abortCtrl = new AbortController();
    const timeoutId = setTimeout(() => abortCtrl.abort(), 90000);
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/scan-tags`, {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        signal: abortCtrl.signal,
        body: JSON.stringify({
          item_ids: idsToScan,
          recursive: scanRecursive,
          scope: rescanAll ? 'all' : scope,
        }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Scan failed');
      const results: ScanResult[] = data.results || [];

      // Initialize default field choices: prefer 'ocr' if confident, else 'path' if present, else 'manual' or 'skip'
      const initialChoices: Record<string, Record<keyof Tags, string>> = {};
      results.forEach(res => {
        initialChoices[res.item_id] = { ...emptyTags };
        (['department', 'vessel', 'group', 'category'] as Array<keyof Tags>).forEach(field => {
          const ocrVal = res.ocr_suggestion?.[field]?.value;
          const ocrConf = res.ocr_suggestion?.[field]?.confidence || 0;
          const pathVal = res.path_suggestion?.[field]?.value;
          // For vessel and group: strongly prefer path when OCR is not highly confident,
          // because the file's folder location is the authoritative source.
          const isVesselOrGroup = field === 'vessel' || field === 'group';
          const ocrThreshold = isVesselOrGroup ? 0.8 : 0.5;
          // Category folders are the authoritative taxonomy location; OCR can read
          // document content from a different drawing family (for example Hull).
          if (field === 'category' && pathVal) {
            initialChoices[res.item_id][field] = 'path';
          } else if (ocrVal && ocrConf >= ocrThreshold) {
            initialChoices[res.item_id][field] = 'ocr';
          } else if (pathVal) {
            initialChoices[res.item_id][field] = 'path';
          } else if (ocrVal && ocrConf >= 0.4) {
            initialChoices[res.item_id][field] = 'ocr';
          } else if (res.proposed_tags?.[field]) {
            initialChoices[res.item_id][field] = 'manual';
          } else {
            initialChoices[res.item_id][field] = 'skip';
          }
        });
      });
      setFieldChoices(initialChoices);
      setScanResults(previous => {
        const merged = new Map(previous.map(result => [result.item_id, result]));
        results.forEach(result => merged.set(result.item_id, result));
        return Array.from(merged.values());
      });

      if (crumbs.length > 0) {
        const crumb = crumbs[crumbs.length - 1];
        await refreshFolder(context, crumb.id, crumbs);
      }
      setMessage(`Scanned ${data.scanned || results.length} item(s). Review suggestions and confirm below.`);
    } catch (error) {
      const errorMessage = error instanceof Error && error.name === 'AbortError'
        ? 'Scan request timed out after 90 seconds. Try scanning fewer files or specific sub-folders.'
        : (error instanceof Error ? error.message : 'Scan failed');
      setMessage(errorMessage);
    } finally {
      clearTimeout(timeoutId);
      setLoading(false);
    }
  };

  const updateProposedTag = (itemId: string, key: keyof Tags, value: string): void => {
    setScanResults(previous => previous.map(result => result.item_id === itemId
      ? { ...result, proposed_tags: { ...emptyTags, ...(result.proposed_tags || {}), [key]: value } }
      : result));
  };

  const confirmTags = async (): Promise<void> => {
    if (!context) return;
    const reviewable = scanResults.filter(result => result.status === 'needs_selection' && !result.error);
    if (!reviewable.length) { setMessage('There are no tag changes waiting for confirmation.'); return; }

    const initialFeed: TagFeedEntry[] = reviewable.map(r => ({ name: r.filename || r.item_id, status: 'pending' as const }));
    setTaggingModal({ title: 'Confirming & Applying OCR Tags', feed: initialFeed, total: reviewable.length, finished: false, summary: '' });
    setLoading(true);

    const feedState = [...initialFeed];
    let okCount = 0;
    let failCount = 0;
    const confirmedResults: Array<{ result: ScanResult; data: any }> = [];

    try {
      let nextIndex = 0;
      const confirmOne = async (): Promise<void> => {
        for (let _safetyBreaker = 0; _safetyBreaker < reviewable.length + 1; _safetyBreaker++) {
          const i = nextIndex++;
          if (i >= reviewable.length) return;
          const result = reviewable[i];
          try {
            const choices = fieldChoices[result.item_id] || {};
            const values = result.proposed_tags || emptyTags;
            const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/items/${encodeURIComponent(result.item_id)}/resolve-tags`, {
              method: 'POST',
              headers: { ...headers, 'Content-Type': 'application/json' },
              body: JSON.stringify({
                choices,
                values: {
                  ...values,
                  ocr_tags: Object.fromEntries(Object.entries(result.ocr_suggestion || {}).map(([k, v]) => [k, v.value])),
                  path_tags: Object.fromEntries(Object.entries(result.path_suggestion || {}).map(([k, v]) => [k, v.value])),
                },
              }),
            });
            const data = await response.json();
            if (!response.ok) throw new Error(data.detail || 'Could not confirm tags');
            feedState[i] = { ...feedState[i], status: 'ok' };
            okCount++;
            confirmedResults.push({ result, data });
          } catch (fileErr) {
            feedState[i] = { ...feedState[i], status: 'failed', error: fileErr instanceof Error ? fileErr.message : 'Failed' };
            failCount++;
          }
          setTaggingModal(prev => prev ? { ...prev, feed: [...feedState] } : prev);
        }
      };
      await Promise.all(Array.from({ length: Math.min(8, reviewable.length) }, () => confirmOne()));

      setScanResults(previous => previous.map(result => {
        const applied = confirmedResults.find(entry => entry.result.item_id === result.item_id);
        return applied ? { ...result, status: 'confirmed', current_tags: applied.data.tags } : result;
      }));

      setItems(previous => previous.map(item => {
        const match = confirmedResults.find(entry => entry.result.item_id === item.id);
        return match ? { ...item, tags: match.data.tags } : item;
      }));

      if (crumbs.length > 0) {
        const crumb = crumbs[crumbs.length - 1];
        await refreshFolder(context, crumb.id, crumbs);
      }

      const summary = failCount > 0
        ? `✓ ${okCount} file(s) tagged. ✗ ${failCount} file(s) failed.`
        : `✓ All ${okCount} file(s) confirmed and saved to SharePoint!`;
      setTaggingModal(prev => prev ? { ...prev, finished: true, summary } : prev);
      setMessage(summary);
    } catch (error) {
      const errMsg = error instanceof Error ? error.message : 'Could not confirm tags';
      setTaggingModal(prev => prev ? { ...prev, finished: true, summary: `✗ Error: ${errMsg}` } : prev);
      setMessage(errMsg);
    } finally {
      setLoading(false);
      await loadTagFailures();
    }
  };

  /** Update vessel on selected files/folders directly in SharePoint WITHOUT scanning */
  const applyBulkVessel = async (
    targetItemIds?: string[],
    specificVessel?: string,
    options?: { onlyMissing?: boolean; overwriteExisting?: boolean; confirmed?: boolean }
  ): Promise<void> => {
    if (!context) return;
    const vesselToApply = (specificVessel !== undefined ? specificVessel : bulkVessel).trim();
    if (!vesselToApply) {
      setMessage('Please select or enter a vessel name.');
      return;
    }

    const overwriteExisting = options?.overwriteExisting === true;
    const onlyMissing = !overwriteExisting;

    let candidateIds: string[] = [];
    if (targetItemIds && targetItemIds.length > 0) {
      candidateIds = targetItemIds;
    } else if (selected.size > 0) {
      candidateIds = Array.from(selected);
    } else {
      candidateIds = items.map(i => i.id);
    }

    if (candidateIds.length === 0) {
      setMessage('No items available for vessel update.');
      return;
    }

    // Inspect candidate items known in memory
    const candidateItems = candidateIds.map(id => items.find(it => it.id === id)).filter(Boolean) as Item[];
    const directFiles = candidateItems.filter(it => !it.folder);
    const missingDirectFiles = directFiles.filter(it => !it.tags?.vessel);
    const alreadyTaggedDirectFiles = directFiles.filter(it => Boolean(it.tags?.vessel));
    const folderItems = candidateItems.filter(it => it.folder);
    const hasFolders = folderItems.length > 0 || (candidateIds.length === items.length && items.some(i => i.folder));

    // If overwrite is requested but not confirmed, show in-app confirmation modal (no alert box)
    if (overwriteExisting && !options?.confirmed) {
      const taggedCount = alreadyTaggedDirectFiles.length;
      if (taggedCount > 0) {
        setConfirmOverwrite({ count: taggedCount, vessel: vesselToApply });
        return;
      }
    }

    // Resolve IDs to send
    let idsToUpdate: string[] = [];
    let expectedCount = 0;

    if (onlyMissing) {
      if (missingDirectFiles.length > 0) {
        // Direct files missing vessel tag: update ONLY these files (no subfolders, no alert box)
        idsToUpdate = missingDirectFiles.map(it => it.id);
        expectedCount = missingDirectFiles.length;
      } else if (hasFolders) {
        // No direct files missing: update missing across subfolders
        idsToUpdate = folderItems.length > 0 ? folderItems.map(it => it.id) : items.filter(i => i.folder).map(i => i.id);
        expectedCount = summaryCounts?.total_files ?? candidateIds.length;
      } else {
        setMessage('All files in the current view already have a vessel tag.');
        return;
      }
    } else {
      idsToUpdate = candidateIds;
      expectedCount = summaryCounts?.total_files ?? candidateIds.length;
    }

    setLoading(true);
    // Directly open the in-app progress popup modal!
    setTaggingModal({
      title: onlyMissing ? `Updating Missing Vessel: "${vesselToApply}"` : `Updating Vessel Tag: "${vesselToApply}"`,
      feed: [],
      total: expectedCount || idsToUpdate.length,
      finished: false,
      summary: '',
    });

    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/bulk-update-tags/stream`, {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          item_ids: idsToUpdate,
          vessel: vesselToApply,
          recursive: idsToUpdate.some(id => items.find(it => it.id === id)?.folder),
          skip_if_tagged: !overwriteExisting,
          skip_if_any_vessel_set: !overwriteExisting && Boolean(onlyMissing),
        }),
      });

      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        throw new Error((errData as { detail?: string }).detail || 'Could not update vessel');
      }
      if (!response.body) throw new Error('No response body');

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      let updatedCount = 0;
      let skippedCount = 0;
      let failedCount = 0;
      let totalFromServer = expectedCount || idsToUpdate.length;
      let isTruncated = false;
      const updatedMap = new Map<string, Record<string, string>>();

      // eslint-disable-next-line no-constant-condition
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });

        // Process all complete newline-delimited JSON lines
        const lines = buffer.split('\n');
        buffer = lines.pop() ?? ''; // last (possibly incomplete) line stays in buffer

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed) continue;
          let event: Record<string, unknown>;
          try { event = JSON.parse(trimmed); } catch { continue; }

          if (event.type === 'start') {
            totalFromServer = (event.total as number) || totalFromServer;
            isTruncated = !!(event.truncated);
            setTaggingModal(prev => prev ? { ...prev, total: totalFromServer } : prev);

          } else if (event.type === 'progress') {
            const r = event.result as { item_id: string; filename?: string; parent_path?: string; ok: boolean; error?: string; tags?: Record<string, string>; skipped?: boolean };
            if (r.ok && r.tags) updatedMap.set(r.item_id, r.tags);
            updatedCount = (event.completed as number) ?? updatedCount;
            skippedCount = (event.skipped as number) ?? skippedCount;
            const progressTotal = (event.total as number) || totalFromServer;

            const folderName = r.parent_path ? r.parent_path.split('/').pop() || '' : '';
            const tagLabel = r.tags?.vessel ? ` → ${r.tags.vessel}` : ` → ${vesselToApply}`;
            let formattedError = r.error;
            if (formattedError) {
              if (formattedError.includes('2147018884') || formattedError.toLowerCase().includes('lock')) {
                formattedError = 'File is locked (open in Excel/Office). Close open file tab & retry.';
              }
            }
            const feedEntry: TagFeedEntry = {
              name: r.filename
                ? `${r.filename}${folderName ? ` (${folderName})` : ''}${tagLabel}`
                : r.item_id,
              status: r.ok ? 'ok' : 'failed',
              error: formattedError,
            };
            // Append the new result to the live feed
            setTaggingModal(prev => {
              if (!prev) return prev;
              return { ...prev, feed: [...prev.feed, feedEntry], total: progressTotal };
            });

          } else if (event.type === 'done') {
            updatedCount = (event.updated_count as number) ?? updatedCount;
            skippedCount = (event.skipped_count as number) ?? skippedCount;
            failedCount = (event.failed_count as number) ?? failedCount;
          }
        }
      }

      // Update items in list with new tags
      setItems(prev => prev.map(item => {
        const updated = updatedMap.get(item.id);
        return updated ? { ...item, tags: updated } : item;
      }));
      setSelected(new Set());
      folderCacheRef.current.clear();

      let summary = failedCount > 0
        ? `✓ ${updatedCount} file(s) updated. ✗ ${failedCount} file(s) failed. (Please close any open Excel tabs for failed files & retry.)`
        : `✓ All ${updatedCount} file(s) successfully updated with vessel "${vesselToApply}"!`;
      if (skippedCount > 0) {
        summary += ` (${skippedCount} already tagged — skipped.)`;
      }
      if (isTruncated) {
        summary += ` (Large folder: select a narrower range and run again for remaining files.)`;
      }
      if (updatedCount === 0 && failedCount === 0 && skippedCount > 0) {
        summary = `All ${skippedCount} file(s) already have vessel tags — none needed updating.`;
      }
      if (updatedCount === 0 && failedCount === 0 && skippedCount === 0) {
        summary = 'No files found to update in the selected location.';
      }

      setTaggingModal(prev => prev ? {
        ...prev,
        finished: true,
        summary,
        total: Math.max(prev.feed.length, 1),
      } : prev);
      setMessage(summary);

      if (crumbs.length > 0) {
        const crumb = crumbs[crumbs.length - 1];
        await refreshFolder(context, crumb.id, crumbs);
      }
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : 'Could not update vessel';
      setTaggingModal(prev => prev ? { ...prev, finished: true, summary: `✗ Error: ${errMsg}` } : prev);
      setMessage(errMsg);
    } finally {
      setLoading(false);
    }
  };

  const promptAutoTagFromPath = (targetItemIds?: string[]): void => {
    if (!context) return;
    let ids: string[] = [];
    if (targetItemIds && targetItemIds.length > 0) {
      ids = targetItemIds;
    } else if (selected.size > 0) {
      ids = Array.from(selected);
    } else {
      ids = items.map(i => i.id);
    }
    if (ids.length === 0) {
      setMessage('No items to tag.');
      return;
    }

    const isSelection = (targetItemIds && targetItemIds.length > 0) || selected.size > 0;
    const targetItems = ids.map(id => items.find(it => it.id === id)).filter(Boolean) as Item[];
    const hasFolder = targetItems.some(it => it.folder);
    const fileCount = (isSelection && !hasFolder)
      ? targetItems.length
      : (summaryCounts?.total_files ?? items.filter(it => !it.folder).length);

    setConfirmScanModal({
      mode: 'autoTagFromPath',
      targetItemIds: ids,
      isSelection,
      count: fileCount,
      names: targetItems.map(it => it.name),
      isRecursive: hasFolder || (!isSelection && recursive),
    });
  };

  /** Automatically tag files from folder path hierarchy (Department, Vessel, Group, Category) WITHOUT scanning */
  const autoTagFromPath = async (targetItemIds?: string[]): Promise<void> => {
    if (!context) return;
    let idsToUpdate: string[] = [];
    if (targetItemIds && targetItemIds.length > 0) {
      idsToUpdate = targetItemIds;
    } else if (selected.size > 0) {
      idsToUpdate = Array.from(selected);
    } else {
      idsToUpdate = items.map(i => i.id);
    }

    if (idsToUpdate.length === 0) {
      setMessage('No items to tag.');
      return;
    }

    const isSelection = (targetItemIds && targetItemIds.length > 0) || selected.size > 0;
    const targetItems = idsToUpdate.map(id => items.find(i => i.id === id)).filter(Boolean) as Item[];
    const hasFolder = targetItems.some(it => it.folder);
    const isSingleFile = isSelection && !hasFolder && idsToUpdate.length === 1;
    const estimatedFiles = (isSelection && !hasFolder)
      ? idsToUpdate.length
      : (summaryCounts?.total_files ?? items.length);

    setLoading(true);
    const initialFeed: TagFeedEntry[] = idsToUpdate.slice(0, 50).map(id => {
      const it = items.find(i => i.id === id);
      return {
        name: it ? (it.folder ? `📁 ${it.name} (discovering files...)` : `📄 ${it.name}`) : id,
        status: 'pending' as const,
      };
    });
    if (idsToUpdate.length > 50) {
      initialFeed.push({
        name: `... and ${idsToUpdate.length - 50} more items across sub-folders`,
        status: 'pending' as const,
      });
    }
    setTaggingModal({
      title: isSingleFile ? `Auto-Tagging: ${targetItems[0]?.name || 'Selected File'}` : 'Auto-Tagging from Folder Path',
      feed: initialFeed,
      total: Math.max(estimatedFiles, idsToUpdate.length),
      finished: false,
      summary: '',
    });

    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/bulk-update-tags`, {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          item_ids: idsToUpdate,
          auto_from_path: true,
          recursive: hasFolder ? true : (!isSelection && recursive),
          scope: rescanAll ? 'all' : 'missing_only',
        }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Could not auto-tag files');

      const results = data.results || [];
      const okCount = results.filter((r: { ok: boolean }) => r.ok).length;
      const failCount = results.filter((r: { ok: boolean }) => !r.ok).length;

      const feed: TagFeedEntry[] = results.map((r: { filename?: string; parent_path?: string; item_id: string; ok: boolean; error?: string; tags?: Record<string, string> }) => {
        const folderName = r.parent_path ? r.parent_path.split('/').pop() || '' : '';
        const tagParts = [];
        if (r.tags?.vessel) tagParts.push(r.tags.vessel);
        if (r.tags?.group) tagParts.push(r.tags.group);
        if (r.tags?.category) tagParts.push(r.tags.category);
        const tagLabel = tagParts.length > 0 ? ` → ${tagParts.join(' / ')}` : '';
        return {
          name: r.filename
            ? `${r.filename}${folderName ? ` (${folderName})` : ''}${tagLabel}`
            : r.item_id,
          status: r.ok ? ('ok' as const) : ('failed' as const),
          error: r.error,
        };
      });

      if (results.length === 0) {
        feed.push({
          name: 'No files found to tag in the selected location.',
          status: 'failed' as const,
          error: '0 files discovered',
        });
      }

      const updatedMap = new Map<string, Record<string, string>>();
      results.forEach((r: { item_id: string; ok: boolean; tags?: Record<string, string> }) => {
        if (r.ok && r.tags) updatedMap.set(r.item_id, r.tags);
      });
      setItems(prev => prev.map(item => {
        const updated = updatedMap.get(item.id);
        return updated ? { ...item, tags: updated } : item;
      }));
      setSelected(new Set());

      let summary = failCount > 0
        ? `✓ ${okCount} file(s) tagged. ✗ ${failCount} file(s) failed.`
        : `✓ All ${okCount} file(s) successfully tagged from folder path!`;
      if (data.truncated) {
        summary += ` (The service limit was reached; narrow the folder selection to continue.)`;
      }
      setTaggingModal({
        title: isSingleFile ? `Auto-Tagging: ${targetItems[0]?.name || 'Selected File'}` : 'Auto-Tagging from Folder Path',
        feed,
        total: Math.max(results.length, 1),
        finished: true,
        summary,
      });
      setMessage(summary);

      if (crumbs.length > 0) {
        const crumb = crumbs[crumbs.length - 1];
        await refreshFolder(context, crumb.id, crumbs);
      }
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : 'Could not auto-tag files';
      setTaggingModal(prev => prev ? { ...prev, finished: true, summary: `✗ Error: ${errMsg}` } : prev);
      setMessage(errMsg);
    } finally {
      setLoading(false);
    }
  };

  const promptAutoTagAndReview = (targetItemIds?: string[]): void => {
    if (!context) return;
    let ids: string[] = [];
    if (targetItemIds && targetItemIds.length > 0) {
      ids = targetItemIds;
    } else if (selected.size > 0) {
      ids = Array.from(selected);
    } else {
      ids = items.map(i => i.id);
    }
    if (ids.length === 0) {
      setMessage('No items to scan and review.');
      return;
    }

    const isSelection = (targetItemIds && targetItemIds.length > 0) || selected.size > 0;
    const targetItems = ids.map(id => items.find(it => it.id === id)).filter(Boolean) as Item[];
    const hasFolder = targetItems.some(it => it.folder);
    const isRecursive = hasFolder ? true : (!isSelection && recursive);

    let fileCount = 0;
    if (isSelection) {
      if (!hasFolder) {
        fileCount = targetItems.length;
      } else {
        fileCount = summaryCounts?.total_files ?? ids.length;
      }
    } else {
      fileCount = summaryCounts?.total_files ?? items.filter(it => !it.folder).length;
    }

    // Count direct files (non-folders) missing ANY core tag (department/vessel/group/category)
    const directFiles = targetItems.filter(it => !it.folder);
    const missingCount = directFiles.filter(it => hasAnyMissingTag(it)).length;
    const isSingleFile = isSelection && !hasFolder && ids.length === 1;

    // A single file has no scope choice to make: start OCR review immediately.
    if (isSingleFile) {
      void autoTagAndReview(ids, rescanAll ? 'all' : 'missing_only');
      return;
    }

    // If all direct files already have vessel tags and no folders are selected,
    // skip the 2-option modal and go straight to a full scan
    if (missingCount === 0 && !hasFolder) {
      void autoTagAndReview(ids, 'all');
      return;
    }

    setConfirmScanModal({
      mode: 'autoTagAndReview',
      targetItemIds: ids,
      isSelection,
      count: fileCount,
      missingCount,
      names: targetItems.map(it => it.name),
      isRecursive,
    });
  };

  /** Simultaneous Auto-Tag & Review across sub-folders */
  const autoTagAndReview = async (targetItemIds?: string[], scope: 'missing_only' | 'all' = 'missing_only'): Promise<void> => {
    if (!context) return;
    const effectiveScope = rescanAll ? 'all' : scope;
    let idsToScan: string[] = [];
    if (targetItemIds && targetItemIds.length > 0) {
      idsToScan = targetItemIds;
    } else if (selected.size > 0) {
      idsToScan = Array.from(selected);
    } else {
      idsToScan = items.map(i => i.id);
    }
    if (idsToScan.length === 0) {
      setMessage('No items to scan and review.');
      return;
    }

    const isSelection = (targetItemIds && targetItemIds.length > 0) || selected.size > 0;
    const targetItems = idsToScan.map(id => items.find(it => it.id === id)).filter(Boolean) as Item[];
    const hasFolder = targetItems.some(it => it.folder);
    const scanRecursive = hasFolder ? true : (!isSelection && recursive);
    const isSingleFile = isSelection && !hasFolder && idsToScan.length === 1;

    const initialFeed: TagFeedEntry[] = effectiveScope === 'all' ? idsToScan.slice(0, 50).map(id => {
      const item = items.find(current => current.id === id);
      return {
        name: item ? (item.folder ? `📁 ${item.name} (scanning...)` : `📄 ${item.name}`) : id,
        status: 'pending' as const,
      };
    }) : [{ name: 'Checking files for missing tags...', status: 'pending' as const }];
    if (effectiveScope === 'all' && idsToScan.length > 50) {
      initialFeed.push({
        name: `... and ${idsToScan.length - 50} more items across sub-folders`,
        status: 'pending',
      });
    }

    setLoading(true);
    setScanResults([]);
    setFieldChoices({});
    setTaggingModal({
      title: isSingleFile ? `Scanning: ${targetItems[0]?.name || 'Selected File'}` : 'Scanning & Preparing Tag Review',
      feed: initialFeed,
      total: effectiveScope === 'all' ? idsToScan.length : 1,
      finished: false,
      summary: '',
    });
    const scanTitle = isSingleFile ? `Scanning: ${targetItems[0]?.name || 'Selected File'}` : 'Scanning & Preparing Tag Review';
    host.setState({
      scanProgress: { status: 'running', completed: 0, total: 0, title: scanTitle, recentFiles: [] },
    });
    setMessage(isSingleFile
      ? `Auto-tagging and scanning selected file "${targetItems[0]?.name || ''}" with OCR...`
      : (isSelection && !hasFolder
        ? `Auto-tagging and scanning ${idsToScan.length} selected file(s) with OCR...`
        : (effectiveScope === 'all'
          ? 'Auto-tagging and scanning all files with OCR & folder taxonomy...'
          : 'Finding files missing tags before OCR...')));
    const abortCtrl = new AbortController();
    const timeoutId = window.setTimeout(() => abortCtrl.abort(), 10 * 60 * 1000);
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/scan-tags/stream`, {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        signal: abortCtrl.signal,
        body: JSON.stringify({
          item_ids: idsToScan,
          recursive: scanRecursive,
          scope: effectiveScope,
        }),
      });
      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.detail || 'Scan failed');
      }
      if (!response.body) throw new Error('Scan stream was unavailable');

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      let total = 0;
      let completed = 0;
      const results: ScanResult[] = [];
      const addChoices = (res: ScanResult): Record<keyof Tags, string> => {
        const choices = { ...emptyTags };
        (['department', 'vessel', 'group', 'category'] as Array<keyof Tags>).forEach(field => {
          const ocrVal = res.ocr_suggestion?.[field]?.value;
          const ocrConf = res.ocr_suggestion?.[field]?.confidence || 0;
          const pathVal = res.path_suggestion?.[field]?.value;
          const isVesselOrGroup = field === 'vessel' || field === 'group';
          const ocrThreshold = isVesselOrGroup ? 0.85 : 0.60;
          if (field === 'category' && pathVal) {
            choices[field] = 'path';
          } else if (ocrVal && ocrConf >= ocrThreshold) {
            choices[field] = 'ocr';
          } else if (pathVal) {
            choices[field] = 'path';
          } else if (ocrVal && ocrConf >= 0.40) {
            choices[field] = 'ocr';
          } else if (res.proposed_tags?.[field]) {
            choices[field] = 'manual';
          } else {
            choices[field] = 'skip';
          }
        });
        return choices;
      };
      const handleEvent = (event: { type: string; total?: number; completed?: number; result?: ScanResult; scanned?: number; scope?: string; total_discovered?: number }): void => {
        if (event.type === 'start') {
          total = event.total || 0;
          setTaggingModal(prev => prev ? { ...prev, total: Math.max(total, 1), feed: [] } : prev);
          host.setState(prev => ({ scanProgress: { ...prev.scanProgress, status: 'running', total, title: scanTitle } }));
          return;
        }
        if (event.type === 'progress' && event.result) {
          const result = event.result;
          results.unshift(result);
          completed = event.completed || completed + 1;
          setFieldChoices(prev => ({ ...prev, [result.item_id]: addChoices(result) }));
          setScanResults(prev => [result, ...prev.filter(item => item.item_id !== result.item_id)]);
          setTaggingModal(prev => prev ? {
            ...prev,
            total: Math.max(total, 1),
            feed: [{ name: result.filename || result.item_id, status: result.error ? 'failed' : 'ok', error: result.error }, ...prev.feed],
          } : prev);
          host.setState(prev => ({ scanProgress: {
            ...prev.scanProgress,
            status: 'running', completed, total,
            recentFiles: [result.filename || result.item_id, ...prev.scanProgress.recentFiles].slice(0, 5),
          } }));
        }
      };

      let _reading = true;
      while (_reading) {
        const chunk = await reader.read();
        if (chunk.done) { _reading = false; break; }
        buffer += decoder.decode(chunk.value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';
        lines.filter(line => line.trim()).forEach(line => {
          try {
            handleEvent(JSON.parse(line));
          } catch (jsonErr) {
            console.warn('Failed to parse scan stream event:', line, jsonErr);
          }
        });
      }
      if (buffer.trim()) {
        try {
          handleEvent(JSON.parse(buffer));
        } catch (jsonErr) {
          console.warn('Failed to parse trailing scan stream event:', buffer, jsonErr);
        }
      }

      if (crumbs.length > 0) {
        const crumb = crumbs[crumbs.length - 1];
        await refreshFolder(context, crumb.id, crumbs);
      }
      const countMsg = effectiveScope === 'missing_only'
        ? `⚡ Scanned ${results.length} file(s) missing a tag. Review suggestions below and confirm!`
        : `⚡ Scanned all ${results.length} file(s) across sub-folders. Review suggestions below and confirm!`;
      setTaggingModal(prev => prev ? {
        ...prev,
        feed: prev.feed.length > 0 ? prev.feed : [{ name: 'All selected files were already scanned.', status: 'ok' }],
        total: Math.max(results.length, 1),
        finished: true,
        summary: countMsg,
      } : prev);
      host.setState(prev => ({ scanProgress: { ...prev.scanProgress, status: 'completed', completed: results.length, total: Math.max(total, results.length), title: countMsg } }));
      setMessage(countMsg);
    } catch (error) {
      const errorMessage = error instanceof DOMException && error.name === 'AbortError'
        ? 'Scan and review timed out. Only files missing tags are included; try a smaller folder.'
        : (error instanceof Error ? error.message : 'Scan and review failed');
      setTaggingModal(prev => prev ? { ...prev, finished: true, summary: `✗ ${errorMessage}` } : prev);
      host.setState(prev => ({ scanProgress: { ...prev.scanProgress, status: 'failed', error: errorMessage, title: 'OCR scan failed' } }));
      setMessage(errorMessage);
    } finally {
      window.clearTimeout(timeoutId);
      setLoading(false);
    }
  };

  const normalizeSiteSelectionValue = (value: string | null | undefined): string => {
    return String(value || '').trim().toLowerCase().replace(/[_-]+/g, ' ').replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  };

  const filteredSites = sites.filter(site => `${site.display_name} ${site.description || ''}`.toLowerCase().includes(query.toLowerCase()));
  const selectedSites = (() => {
    const key = selectedSiteKey.trim();
    if (!key) return filteredSites;
    const matches = filteredSites.filter(site => {
      const candidates = [site.site_key, site.id, site.name, site.display_name, site.web_url];
      return candidates.some(value => normalizeSiteSelectionValue(value) === normalizeSiteSelectionValue(key));
    });
    return matches.length > 0 ? matches : filteredSites;
  })();

  const renderTags = (item: Item): React.ReactElement => (
    <span style={{ display: 'flex', flexWrap: 'wrap', gap: 4, alignItems: 'center' }}>
      {tagNames.filter(([key]) => item.tags?.[key]).map(([key, label]) => (
        <span key={key} style={{ background: dmsTone('accent').bg, color: dmsTone('accent').fg, borderRadius: 12, padding: '3px 8px', fontSize: 11, fontWeight: 500 }}>
          {label}: {item.tags?.[key]}
        </span>
      ))}
      {!item.folder && !item.tags?.vessel && (
        <button
          type="button"
          onClick={() => void applyBulkVessel([item.id], detectedVessel || bulkVessel)}
          title={`Click to set vessel "${detectedVessel || bulkVessel || 'vessel'}" without scanning`}
          style={{
            background: dmsTone('warning').bg, color: dmsTone('warning').fg, border: `1px dashed ${clay.pillWarnText}`,
            borderRadius: 12, padding: '2px 8px', fontSize: 11, fontWeight: 600,
            cursor: (detectedVessel || bulkVessel) ? 'pointer' : 'default',
            display: 'inline-flex', alignItems: 'center', gap: 3,
          }}
        >
          + Vessel{detectedVessel ? `: ${detectedVessel}` : (bulkVessel ? `: ${bulkVessel}` : '')}
        </button>
      )}
      {!tagNames.some(([key]) => item.tags?.[key]) && item.folder && (() => {
        const fc = item.folder_counts;
        const childCount = (item.folder as { childCount?: number } | null)?.childCount ?? 0;
        // Prefer direct subfolder count; if there are subfolders, show "X folders"
        const subfolderCount = fc ? fc.direct_subfolders : 0;
        const fileCount = fc ? fc.direct_files : childCount;
        const hasSubfolders = subfolderCount > 0;
        const count = hasSubfolders ? subfolderCount : fileCount;
        const label = hasSubfolders
          ? `${count} folder${count === 1 ? '' : 's'}`
          : `${count} file${count === 1 ? '' : 's'}`;
        return (
          <span style={{ color: 'var(--vdms-text-muted)', fontSize: 11, fontStyle: 'italic', display: 'inline-flex', alignItems: 'center', gap: 4 }}>
            <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 11 }} /> Sub-folder ({label})
          </span>
        );
      })()}
    </span>
  );

  const reviewPanel = scanResults.length > 0 ? (
    <div style={{ marginTop: 16, marginBottom: 20, background: dmsTone('warning').bg, border: `1px solid ${clay.pillWarnText}33`, padding: 16, borderRadius: 8, boxShadow: clay.shadowRaised }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
        <strong style={{ fontSize: 15, color: dmsTone('warning').fg }}><Icon iconName="Search" aria-hidden="true" style={{ fontSize: 15 }} /> Review OCR & AI Tag Suggestions ({scanResults.length} item{scanResults.length > 1 ? 's' : ''})</strong>
        <button
          onClick={() => setScanResults([])}
          style={{ border: 0, background: 'transparent', cursor: 'pointer', fontSize: 16, color: dmsTone('warning').fg }}
          title="Close review panel"
        ><Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 16 }} /></button>
      </div>

      {/* Quick vessel apply bar */}
      <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 12, padding: '8px 12px', background: 'var(--vdms-surface-alt)', border: '1px solid var(--vdms-line)', borderRadius: 6 }}>
        <span style={{ fontSize: 12, fontWeight: 600, color: dmsTone('warning').fg, whiteSpace: 'nowrap' }}><Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 12 }} /> Apply vessel to all:</span>
        <select
          value={bulkVessel}
          onChange={e => setBulkVessel(e.target.value)}
          style={{ ...dmsControlStyle(), flex: 1, maxWidth: 260, padding: '5px 8px' }}
        >
          <option value="">— Select vessel —</option>
          {vesselOptions.map(v => <option key={v} value={v}>{v}</option>)}
        </select>
        <button
          disabled={!bulkVessel}
          onClick={() => {
            setScanResults(prev => prev.map(r => ({
              ...r,
              proposed_tags: { ...emptyTags, ...(r.proposed_tags || {}), vessel: bulkVessel },
            })));
            setFieldChoices(prev => {
              const next = { ...prev };
              scanResults.forEach(r => {
                next[r.item_id] = { ...emptyTags, ...(prev[r.item_id] || {}), vessel: 'manual' };
              });
              return next;
            });
          }}
          style={{ ...dmsBtn('primary', Boolean(bulkVessel)), height: 'auto', padding: '5px 14px' }}
        >
          Apply to all
        </button>
      </div>

      {scanResults.map(result => (
        <div key={result.item_id} style={{ padding: '14px 0', borderBottom: `1px solid ${clay.pillWarnText}33` }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              <span style={{ fontWeight: 700, fontSize: 13, color: 'var(--vdms-text)' }}>
                <Icon iconName="Page" aria-hidden="true" style={{ fontSize: 13 }} /> {result.filename || result.item_id}
              </span>
              {result.subfolder_name && (
                <span style={{ background: 'var(--vdms-surface-alt)', border: '1px solid var(--vdms-border)', color: 'var(--vdms-text-secondary)', fontSize: 11, padding: '2px 8px', borderRadius: 12, fontWeight: 600 }}>
                  <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 11 }} /> {result.subfolder_name}
                </span>
              )}
              {Boolean(result.vessel_in_filename_only || (result.ocr_suggestion?.vessel as any)?.vessel_in_filename_only) && (
                <span style={{ background: dmsTone('warning').bg, border: `1px solid ${clay.pillWarnText}55`, color: dmsTone('warning').fg, fontSize: 11, padding: '2px 8px', borderRadius: 12, fontWeight: 600 }}>
                  <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 11 }} /> Vessel name not in file
                </span>
              )}
            </div>
            <span>
              {result.status === 'confirmed'
                ? <span style={{ color: dmsTone('success').fg, fontWeight: 600 }}><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 13 }} /> Confirmed and saved to SharePoint</span>
                : result.error
                  ? <span style={{ color: dmsTone('danger').fg, fontWeight: 600 }}><Icon iconName="ErrorBadge" aria-hidden="true" style={{ fontSize: 13 }} /> {result.error}</span>
                  : <span style={{ color: dmsTone('warning').fg, fontWeight: 600 }}><Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 13 }} /> Ready for review & confirmation</span>
              }
            </span>
          </div>

          {result.status === 'needs_selection' && (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: 10, marginTop: 12 }}>
              {tagNames.map(([field, fieldLabel]) => {
                const ocr = result.ocr_suggestion?.[field];
                const path = result.path_suggestion?.[field];
                const choice = fieldChoices[result.item_id]?.[field] || 'skip';
                // For vessel field, show a Term Store dropdown when choice is 'manual'
                const isVessel = field === 'vessel';
                const isVesselNotInFile = isVessel && Boolean(result.vessel_in_filename_only || (ocr as any)?.vessel_in_filename_only);
                return (
                  <div key={field} style={{ background: 'var(--vdms-surface)', border: `1px solid ${clay.pillWarnText}33`, borderRadius: 6, padding: '8px 10px' }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: 'var(--vdms-text-secondary)', marginBottom: 4 }}>
                      {fieldLabel}
                    </div>
                    <select
                      value={choice}
                      onChange={event => setFieldChoices(previous => ({
                        ...previous,
                        [result.item_id]: { ...emptyTags, ...(previous[result.item_id] || {}), [field]: event.target.value },
                      }))}
                      style={{ ...dmsControlStyle(), width: '100%', padding: '6px 8px' }}
                    >
                      <option value="skip">Skip / Keep unchanged</option>
                      <option value="ocr" disabled={!ocr?.value}>
                        OCR: {ocr?.value || 'None'}{ocr ? ` (${Math.round(ocr.confidence * 100)}%)` : ''}{isVesselNotInFile ? ' - vessel name not in the file' : ''}
                      </option>
                      <option value="path" disabled={!path?.value}>
                        Path: {path?.value || 'None'}
                      </option>
                      <option value="manual">Enter manually</option>
                    </select>
                    {isVesselNotInFile && (
                      <div style={{ marginTop: 4, fontSize: 11, color: dmsTone('warning').fg, display: 'flex', alignItems: 'center', gap: 4, fontWeight: 500 }}>
                        <span><Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 11 }} /> Vessel name not in file (matched from term store)</span>
                      </div>
                    )}
                    {choice === 'manual' && (
                      isVessel && vesselOptions.length > 0 ? (
                        <select
                          value={result.proposed_tags?.vessel || ''}
                          onChange={event => updateProposedTag(result.item_id, 'vessel', event.target.value)}
                          style={{ ...dmsControlStyle(), width: '100%', boxSizing: 'border-box', marginTop: 6, padding: 6 }}
                        >
                          <option value="">— Select vessel —</option>
                          {vesselOptions.map(v => <option key={v} value={v}>{v}</option>)}
                        </select>
                      ) : (
                        <input
                          aria-label={`${result.filename || result.item_id} ${field}`}
                          value={result.proposed_tags?.[field] || ''}
                          onChange={event => updateProposedTag(result.item_id, field, event.target.value)}
                          placeholder={`Enter ${fieldLabel}`}
                          style={{ ...dmsControlStyle(), width: '100%', boxSizing: 'border-box', marginTop: 6, padding: 6 }}
                        />
                      )
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      ))}

      <div style={{ marginTop: 14, display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
        <button
          type="button"
          disabled={loading || !scanResults.some(result => result.status === 'needs_selection')}
          onClick={() => {
            setFieldChoices(prev => {
              const next = { ...prev };
              scanResults.forEach(r => {
                next[r.item_id] = { ...emptyTags };
                (['department', 'vessel', 'group', 'category'] as Array<keyof Tags>).forEach(f => {
                  const ocrConf = r.ocr_suggestion?.[f]?.confidence || 0;
                  const isVesselOrGroup = f === 'vessel' || f === 'group';
                  if (r.path_suggestion?.[f]?.value) {
                    next[r.item_id][f] = 'path';
                  } else if (r.ocr_suggestion?.[f]?.value && ocrConf >= (isVesselOrGroup ? 0.8 : 0.6)) {
                    next[r.item_id][f] = 'ocr';
                  } else if (r.proposed_tags?.[f]) {
                    next[r.item_id][f] = 'manual';
                  } else {
                    next[r.item_id][f] = 'skip';
                  }
                });
              });
              return next;
            });
          }}
          style={{ ...dmsBtn('secondary', !(loading || !scanResults.some(result => result.status === 'needs_selection'))), height: 'auto', padding: '10px 16px' }}
        >
          <Icon iconName="LightningBolt" aria-hidden="true" style={{ fontSize: 13 }} /> Auto-Select All High Confidence
        </button>

        <button
          disabled={loading || !scanResults.some(result => result.status === 'needs_selection')}
          onClick={() => void confirmTags()}
          style={{ ...dmsBtn('primary', !(loading || !scanResults.some(result => result.status === 'needs_selection'))), height: 'auto', padding: '10px 18px' }}
        >
          <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 13 }} /> Confirm & Apply Selected Tags
        </button>
      </div>
    </div>
  ) : null;

  // ── Tagging Progress Modal ──────────────────────────────────────────────────
  const renderTaggingModal = (): React.ReactElement | null => {
    if (!taggingModal) return null;
    const doneCount = taggingModal.feed.filter(f => f.status !== 'pending').length;
    const pct = taggingModal.total > 0 ? Math.round((doneCount / taggingModal.total) * 100) : 0;
    return (
      <div style={{
        position: 'fixed', inset: 0, zIndex: 9999,
        background: 'rgba(15,23,42,0.55)', display: 'flex', alignItems: 'center', justifyContent: 'center',
      }}>
        <div style={{
          background: 'var(--vdms-surface)', borderRadius: 16, padding: '28px 32px', width: '100%', maxWidth: 540,
          boxShadow: '0 20px 60px rgba(0,0,0,0.25)', display: 'flex', flexDirection: 'column', gap: 16,
        }}>
          {/* Header */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: clay.accentSoft, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}><Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 20 }} /></div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 16, color: 'var(--vdms-text)' }}>{taggingModal.title}</div>
              <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2 }}>
                {taggingModal.finished ? `Completed (${doneCount} file${doneCount === 1 ? '' : 's'} updated)` : `Processing ${doneCount} of ${taggingModal.total} file(s)…`}
              </div>
            </div>
          </div>

          {/* Progress bar */}
          <div style={{ height: 6, background: 'var(--vdms-surface-alt)', borderRadius: 99, overflow: 'hidden' }}>
            <div style={{
              height: '100%', borderRadius: 99,
              background: taggingModal.finished && taggingModal.feed.some(f => f.status === 'failed') ? clay.pillDangerText : clay.accent,
              width: `${pct}%`, transition: 'width 0.3s ease',
            }} />
          </div>

          {/* File feed list */}
          <div style={{
            maxHeight: 320, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 4,
            background: 'var(--vdms-surface-alt)', borderRadius: 10, padding: '10px 12px',
            border: '1px solid var(--vdms-border)',
          }}>
            {taggingModal.feed.map((entry, idx) => (
              <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, padding: '3px 0' }}>
                <span style={{ fontSize: 15, flexShrink: 0 }}>
                  {entry.status === 'pending' ? <Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 15 }} /> : entry.status === 'ok' ? <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 15 }} /> : <Icon iconName="ErrorBadge" aria-hidden="true" style={{ fontSize: 15 }} />}
                </span>
                <span style={{
                  flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                  color: entry.status === 'pending' ? 'var(--vdms-text-muted)' : entry.status === 'ok' ? dmsTone('success').fg : dmsTone('danger').fg,
                  fontWeight: entry.status !== 'pending' ? 600 : 400,
                }}>
                  {entry.name}
                </span>
                {entry.status === 'failed' && entry.error && (
                  <span
                    style={{
                      fontSize: 11, color: dmsTone('danger').fg, background: dmsTone('danger').bg, padding: '2px 8px',
                      borderRadius: 4, whiteSpace: 'nowrap', maxWidth: 320, overflow: 'hidden',
                      textOverflow: 'ellipsis', flexShrink: 0, fontWeight: 600,
                    }}
                    title={entry.error}
                  >
                    <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 11 }} /> {entry.error}
                  </span>
                )}
              </div>
            ))}
          </div>

          {/* Summary / close */}
          {taggingModal.finished && (
            <>
              <div style={{
                padding: '12px 16px', borderRadius: 10,
                background: taggingModal.feed.every(f => f.status !== 'failed') ? dmsTone('success').bg : dmsTone('warning').bg,
                color: taggingModal.feed.every(f => f.status !== 'failed') ? dmsTone('success').fg : dmsTone('warning').fg,
                fontWeight: 700, fontSize: 14, textAlign: 'center',
              }}>
                {taggingModal.summary}
              </div>
              <button
                type="button"
                onClick={() => setTaggingModal(null)}
                style={{ ...dmsBtn('primary'), height: 'auto', padding: '10px 20px', fontSize: 14, alignSelf: 'flex-end' }}
              >
                <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 14 }} /> Close
              </button>
            </>
          )}
        </div>
      </div>
    );
  };

  // ── In-App Overwrite Confirmation Modal (No native alert box) ───────────────
  const renderConfirmOverwriteModal = (): React.ReactElement | null => {
    if (!confirmOverwrite) return null;
    return (
      <div style={{
        position: 'fixed', inset: 0, zIndex: 9999,
        background: 'rgba(15,23,42,0.55)', display: 'flex', alignItems: 'center', justifyContent: 'center',
      }}>
        <div style={{
          background: 'var(--vdms-surface)', borderRadius: 16, padding: '24px 28px', width: '100%', maxWidth: 460,
          boxShadow: '0 20px 60px rgba(0,0,0,0.25)', display: 'flex', flexDirection: 'column', gap: 16,
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{ width: 36, height: 36, borderRadius: 8, background: dmsTone('warning').bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18 }}><Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 18 }} /></div>
            <div style={{ fontWeight: 700, fontSize: 16, color: 'var(--vdms-text)' }}>Overwrite Existing Vessel Tags?</div>
          </div>
          <div style={{ fontSize: 13, color: 'var(--vdms-text-secondary)', lineHeight: 1.5 }}>
            This will overwrite existing vessel tags on <b>{confirmOverwrite.count}</b> file(s) with vessel <b>"{confirmOverwrite.vessel}"</b>.
          </div>
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 8 }}>
            <button
              type="button"
              onClick={() => setConfirmOverwrite(null)}
              style={{ ...dmsBtn('secondary'), height: 'auto', padding: '8px 16px' }}
            >
              Cancel
            </button>
            <button
              type="button"
              onClick={() => {
                setConfirmOverwrite(null);
                void applyBulkVessel(undefined, undefined, { overwriteExisting: true, confirmed: true });
              }}
              style={{ ...dmsBtn('primary'), height: 'auto', padding: '8px 16px' }}
            >
              <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 13 }} /> Confirm Overwrite
            </button>
          </div>
        </div>
      </div>
    );
  };

  // ── In-App Scan & Auto-Tag Confirmation Modal (No native alert box) ────────
  const renderConfirmScanModal = (): React.ReactElement | null => {
    if (!confirmScanModal) return null;
    const { mode, isSelection, count, missingCount, names, isRecursive } = confirmScanModal;
    const isSingle = isSelection && count === 1;

    let modalTitle = 'Auto-Tag & Review OCR';
    let icon: React.ReactNode = <><Icon iconName="LightningBolt" aria-hidden="true" style={{ fontSize: 18 }} /><Icon iconName="Search" aria-hidden="true" style={{ fontSize: 18 }} /></>;
    let iconBg = clay.accentGradient;
    if (mode === 'scan') {
      modalTitle = 'Scan & Review OCR';
      icon = <Icon iconName="Search" aria-hidden="true" style={{ fontSize: 18 }} />;
      iconBg = clay.accentGradient;
    } else if (mode === 'autoTagFromPath') {
      modalTitle = 'Auto-Tag from Folder Path';
      icon = <Icon iconName="LightningBolt" aria-hidden="true" style={{ fontSize: 18 }} />;
      iconBg = clay.accentGradient;
    }

    return (
      <div style={{
        position: 'fixed', inset: 0, zIndex: 9999,
        background: 'rgba(15,23,42,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center',
        padding: 16,
      }}>
        <div style={{
          background: 'var(--vdms-surface)', borderRadius: 16, padding: '24px 28px', width: '100%', maxWidth: 480,
          boxShadow: '0 25px 50px -12px rgba(0,0,0,0.25)', display: 'flex', flexDirection: 'column', gap: 16,
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <div style={{
              width: 42, height: 42, borderRadius: 10, background: iconBg,
              display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, color: DMS_ON_ACCENT,
              flexShrink: 0,
            }}>
              {icon}
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 17, color: 'var(--vdms-text)' }}>{modalTitle}</div>
              <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2 }}>
                {isSingle ? '1 Selected File' : (isSelection ? `${count} Selected Files` : `${count} Files (Folder Scope)`)}
              </div>
            </div>
          </div>

          <div style={{ fontSize: 13, color: 'var(--vdms-text-secondary)', lineHeight: 1.5, display: 'flex', flexDirection: 'column', gap: 10 }}>
            {isSingle ? (
              <>
                <div>You have selected <b>1 file</b> to process:</div>
                <div style={{
                  background: 'var(--vdms-surface-alt)', border: '1px solid var(--vdms-border)', borderRadius: 8,
                  padding: '10px 14px', fontSize: 13, fontWeight: 600, color: 'var(--vdms-text)',
                  display: 'flex', alignItems: 'center', gap: 8, wordBreak: 'break-all',
                }}>
                  <Icon iconName="Page" aria-hidden="true" style={{ fontSize: 13 }} /> {names[0] || 'Selected file'}
                </div>
                <div style={{ color: 'var(--vdms-text-muted)', fontSize: 12 }}>
                  {mode === 'autoTagFromPath'
                    ? 'Only this selected file will be tagged using folder path taxonomy.'
                    : 'This selected file alone will be scanned with OCR and AI classification for review. Other files will not be touched.'}
                </div>
              </>
            ) : isSelection ? (
              <>
                <div>You have selected <b>{count} files</b> to process:</div>
                <div style={{
                  background: 'var(--vdms-surface-alt)', border: '1px solid var(--vdms-border)', borderRadius: 8,
                  padding: '8px 12px', fontSize: 12, color: 'var(--vdms-text)', maxHeight: 110,
                  overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 4,
                }}>
                  {names.slice(0, 10).map((name, idx) => (
                    <div key={idx} style={{ wordBreak: 'break-all' }}><Icon iconName="Page" aria-hidden="true" style={{ fontSize: 12 }} /> {name}</div>
                  ))}
                  {names.length > 10 && (
                    <div style={{ color: 'var(--vdms-text-muted)', fontStyle: 'italic' }}>
                      ... and {names.length - 10} more files
                    </div>
                  )}
                </div>
                <div style={{ color: 'var(--vdms-text-muted)', fontSize: 12 }}>
                  Only these {count} selected files will be processed.
                </div>
              </>
            ) : (
              <>
                <div>
                  No specific file selected. This will scan and auto-tag all <b>{count} file(s)</b>{isRecursive ? ' across all sub-folders' : ' in this folder'} for review.
                </div>
                <div style={{ color: 'var(--vdms-text-muted)', fontSize: 12 }}>
                  OCR text extraction and folder path taxonomy will be evaluated for each file. You can review suggestions before changes are saved.
                </div>
              </>
            )}
          </div>

          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 6 }}>
            <button
              type="button"
              onClick={() => setConfirmScanModal(null)}
              style={{ ...dmsBtn('secondary'), height: 'auto', padding: '8px 16px' }}
            >
              Cancel
            </button>
            {mode === 'autoTagAndReview' && ((missingCount || 0) > 0 || isRecursive) && !isSingle ? (
              <>
                <button
                  type="button"
                  onClick={() => { const target = confirmScanModal.targetItemIds; setConfirmScanModal(null); void autoTagAndReview(target, 'missing_only'); }}
                  style={{ ...dmsBtn('primary'), height: 'auto', padding: '8px 14px' }}
                >
                  <Icon iconName="LightningBolt" aria-hidden="true" style={{ fontSize: 13 }} /> Tag Missing Files Only {(missingCount || 0) > 0 ? `(${missingCount})` : ''}
                </button>
                <button
                  type="button"
                  onClick={() => { const target = confirmScanModal.targetItemIds; setConfirmScanModal(null); void autoTagAndReview(target, 'all'); }}
                  style={{ ...dmsBtn('secondary'), height: 'auto', padding: '8px 14px' }}
                >
                  <Icon iconName="Search" aria-hidden="true" style={{ fontSize: 13 }} /> Scan All Files ({count})
                </button>
              </>
            ) : mode === 'scan' && ((missingCount || 0) > 0 || isRecursive) && !isSingle ? (
              <>
                <button
                  type="button"
                  onClick={() => { const target = confirmScanModal.targetItemIds; setConfirmScanModal(null); void scan(target, 'missing_only'); }}
                  style={{ ...dmsBtn('primary'), height: 'auto', padding: '8px 14px' }}
                >
                  <Icon iconName="Search" aria-hidden="true" style={{ fontSize: 13 }} /> Scan Missing Files Only {(missingCount || 0) > 0 ? `(${missingCount})` : ''}
                </button>
                <button
                  type="button"
                  onClick={() => { const target = confirmScanModal.targetItemIds; setConfirmScanModal(null); void scan(target, 'all'); }}
                  style={{ ...dmsBtn('secondary'), height: 'auto', padding: '8px 14px' }}
                >
                  Scan All Files ({count})
                </button>
              </>
            ) : (
              <button
                type="button"
                onClick={() => {
                  const target = confirmScanModal.targetItemIds;
                  const m = confirmScanModal.mode;
                  setConfirmScanModal(null);
                  if (m === 'scan') void scan(target, 'all');
                  else if (m === 'autoTagAndReview') void autoTagAndReview(target, 'all');
                  else void autoTagFromPath(target);
                }}
                style={{ ...dmsBtn('primary'), height: 'auto', padding: '8px 18px' }}
              >
                {isSingle
                  ? (mode === 'autoTagFromPath' ? 'Tag Selected File' : <><Icon iconName="LightningBolt" aria-hidden="true" style={{ fontSize: 13 }} /><Icon iconName="Search" aria-hidden="true" style={{ fontSize: 13 }} /> Start OCR on Selected File</>)
                  : mode === 'autoTagAndReview'
                    ? <><Icon iconName="LightningBolt" aria-hidden="true" style={{ fontSize: 13 }} /><Icon iconName="Search" aria-hidden="true" style={{ fontSize: 13 }} /> Scan All Files ({count})</>
                    : `Proceed (${count} file${count === 1 ? '' : 's'})`}
              </button>
            )}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div style={{ background: 'transparent', minHeight: '100%' }}>
      {renderTaggingModal()}
      {renderConfirmOverwriteModal()}
      {renderConfirmScanModal()}
      <div style={{ marginBottom: 20 }}>
        <DmsPageHeader title="Sites" subtitle="Browse libraries and tag documents across your tenant.">
          {sites.length > 1 && (
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: 'var(--vdms-text)', fontWeight: 600 }}>
              SharePoint site
              <select
                aria-label="Select SharePoint site"
                value={selectedSiteKey}
                onChange={event => {
                  setSelectedSiteKey(event.target.value || '');
                  setContext(null);
                  setItems([]);
                  setCrumbs([]);
                  setExpanded(null);
                  setMessage('');
                }}
                style={{ ...dmsControlStyle(), minWidth: 230 }}
              >
                {sites.map(site => (
                  <option key={site.site_key || site.id} value={site.site_key || site.id}>
                    {site.display_name || getSiteName(site)}
                  </option>
                ))}
              </select>
            </label>
          )}
          <input
            aria-label="Search sites"
            value={query}
            onChange={event => setQuery(event.target.value)}
            placeholder="Search sites"
            style={{ ...dmsControlStyle(), width: 260 }}
          />
        </DmsPageHeader>
      </div>

      {message && (
        <div style={{ padding: 12, marginBottom: 16, background: dmsTone('warning').bg, color: dmsTone('warning').fg, borderRadius: 8, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <span>{message}</span>
          <button onClick={() => setMessage('')} style={{ border: 0, background: 'transparent', cursor: 'pointer', fontSize: 16, color: dmsTone('warning').fg }}><Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 16 }} /></button>
        </div>
      )}

      {reviewPanel}

      {/* Sites grid */}
      {!context && (
        <>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(260px,1fr))', gap: 16 }}>
          {sitesLoading && <div style={{ color: clay.textMuted }}>Loading sites...</div>}
          {selectedSites.map(site => (
            <div key={site.id} style={{ background: clay.surface, border: 'none', borderRadius: clay.radiusCard, padding: 18, boxShadow: clay.shadowRaised, fontWeight: 600 }}>
              <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                <div style={{ width: 42, height: 42, background: clay.iconBgGradient, color: clay.accentDark, display: 'grid', placeItems: 'center', borderRadius: clay.radiusIcon, fontWeight: 700, boxShadow: clay.shadowIcon }}>SP</div>
                <div style={{ minWidth: 0 }}>
                  <h2 style={{ margin: 0, fontSize: 17, fontWeight: 800, color: clay.text, overflowWrap: 'anywhere' }}>{getSiteName(site)}</h2>
                  <small style={{ color: clay.text, fontWeight: 600, display: 'block', overflowWrap: 'anywhere' }}>{site.web_url || site.id}</small>
                </div>
              </div>
              <p style={{ color: clay.text, fontWeight: 600, fontSize: 13, minHeight: 34 }}>{site.description || 'SharePoint site'}</p>
              <button
                onClick={() => void loadDrives(site)}
                style={dmsBtn('secondary')}
              >
                {expanded === site.id ? 'Hide libraries' : 'Show libraries'}
              </button>
              {expanded === site.id && (() => {
                const allDrives = drives[site.id] || [];
                const systemCount = allDrives.filter(d => d.is_system).length;
                const showSystem = !!showSystemLibraries[site.id];
                const visibleDrives = showSystem ? allDrives : allDrives.filter(d => !d.is_system);
                return (
                  <div style={{ marginTop: 12, display: 'grid', gap: 6 }}>
                    {visibleDrives.length === 0 && (
                      <div style={{ color: clay.textMuted, fontSize: 13, padding: '4px 2px' }}>
                        No document libraries here — only SharePoint's built-in ones, hidden below.
                      </div>
                    )}
                    {visibleDrives.map(drive => {
                      const hasCount = typeof drive.item_count === 'number';
                      const isEmpty = hasCount && drive.item_count === 0;
                      return (
                        <button
                          key={drive.id}
                          onClick={() => chooseDrive(site, drive)}
                          style={{ textAlign: 'left', padding: 10, border: 'none', background: clay.bg, color: clay.text, fontWeight: 700, borderRadius: 6, cursor: 'pointer', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8 }}
                        >
                          <span>
                            {drive.is_system ? <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 13 }} /> : <Icon iconName="Library" aria-hidden="true" style={{ fontSize: 13 }} />} {drive.name}
                            {drive.is_system && <span style={{ marginLeft: 6, fontSize: 11, color: clay.textMuted }}>(system library)</span>}
                          </span>
                          {hasCount && (
                            <span style={{ fontSize: 12, color: isEmpty ? clay.textMuted : clay.accentDark, whiteSpace: 'nowrap' }}>
                              {isEmpty ? 'empty' : `${drive.item_count} item${drive.item_count === 1 ? '' : 's'}`}
                            </span>
                          )}
                        </button>
                      );
                    })}
                    {systemCount > 0 && (
                      <button
                        onClick={() => setShowSystemLibraries(previous => ({ ...previous, [site.id]: !showSystem }))}
                        style={{ textAlign: 'left', border: 'none', background: 'transparent', color: clay.accentDark, cursor: 'pointer', fontSize: 12, padding: '4px 2px' }}
                      >
                        {showSystem ? `Hide ${systemCount} system librar${systemCount === 1 ? 'y' : 'ies'}` : `Show ${systemCount} more (SharePoint system librar${systemCount === 1 ? 'y' : 'ies'})`}
                      </button>
                    )}
                  </div>
                );
              })()}
            </div>
          ))}
        </div>

        {/* Site Management — moved here from Settings → Site Management;
            same component/state/handlers, just relocated + given a card
            that matches this page instead of Settings' left-nav shell. */}
        <SiteManagementSection host={host} />
        </>
      )}

      {/* Folder / file browser */}
      {context && (
        <div>
          <button
            onClick={() => { setContext(null); setScanResults([]); setMessage(''); }}
            style={{ border: 0, background: 'transparent', color: clay.accentDark, cursor: 'pointer', padding: 0, marginBottom: 16 }}
          >
            ← All sites
          </button>

          <h2 style={{ color: clay.text, margin: '0 0 10px', fontWeight: 800 }}>
            {context.site.display_name} / {context.drive.name}
          </h2>

          {/* Breadcrumb + action toolbar */}
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap', marginBottom: 16 }}>
            {crumbs.map((crumb, index) => (
              <React.Fragment key={crumb.id}>
                <button
                  onClick={() => void loadFolder(context, crumb.id, crumb.name, crumbs.slice(0, index + 1))}
                  style={{ border: 0, background: 'transparent', color: clay.accentDark, cursor: 'pointer', fontWeight: index === crumbs.length - 1 ? 700 : 400 }}
                >
                  {crumb.name}
                </button>
                {index < crumbs.length - 1 && <span style={{ color: clay.textMuted }}>/</span>}
              </React.Fragment>
            ))}

            <div style={{ marginLeft: 'auto', display: 'flex', gap: 8, alignItems: 'center' }}>
              <button
                type="button"
                onClick={() => {
                  const url = getCurrentFolderSharePointUrl();
                  if (url) window.open(url, '_blank');
                  else if (context?.drive.web_url) window.open(context.drive.web_url, '_blank');
                }}
                title="Open current folder in SharePoint Online"
                aria-label="Open current folder in SharePoint Online"
                style={{ ...dmsBtn('secondary'), height: 'auto', padding: '9px 13px' }}
              >
                ↗ Open in SharePoint
              </button>
              <label style={{ fontSize: 13, display: 'flex', gap: 4, alignItems: 'center', cursor: 'pointer' }}>
                <input type="checkbox" checked={recursive} onChange={event => setRecursive(event.target.checked)} />
                Recursive
              </label>
              <label style={{ fontSize: 13, display: 'flex', gap: 4, alignItems: 'center', cursor: 'pointer', color: rescanAll ? dmsTone('warning').fg : 'var(--vdms-text-secondary)' }} title="By default only files missing a tag (Department, Vessel, Group, or Category) are processed">
                <input type="checkbox" checked={rescanAll} onChange={event => setRescanAll(event.target.checked)} />
                Re-scan all files
              </label>
              <button
                disabled={loading}
                onClick={() => void promptAutoTagAndReview()}
                style={{ ...dmsBtn('primary', !loading), height: 'auto', padding: '9px 13px' }}
                title="Automatically tag from folder path & AI, and open interactive review at the same time"
              >
                <Icon iconName="LightningBolt" aria-hidden="true" style={{ fontSize: 13 }} /><Icon iconName="Search" aria-hidden="true" style={{ fontSize: 13 }} /> Auto-Tag & Review{selected.size > 0 ? ` (${selected.size})` : ''}
              </button>
              <button
                disabled={selected.size === 0 || loading}
                onClick={() => void promptScan()}
                style={{ ...dmsBtn('secondary', !(selected.size === 0 || loading)), height: 'auto', padding: '9px 13px' }}
              >
                {loading ? 'Scanning...' : <><Icon iconName="Search" aria-hidden="true" style={{ fontSize: 13 }} /> Scan & Review{selected.size > 0 ? ` (${selected.size})` : ''}</>}
              </button>
              <button
                disabled={loading || !scanResults.some(result => result.status === 'needs_selection')}
                onClick={() => void confirmTags()}
                style={{ ...dmsBtn('primary', scanResults.some(result => result.status === 'needs_selection') && !loading), height: 'auto', padding: '9px 13px' }}
              >
                <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 13 }} /> Confirm Tags
              </button>
            </div>
          </div>

          {tagFailures.length > 0 && (
            <div style={{ marginBottom: 16, padding: 14, background: dmsTone('warning').bg, border: `1px solid ${clay.pillWarnText}33`, borderRadius: 8 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                <strong style={{ color: dmsTone('warning').fg }}>Needs Attention ({tagFailures.length})</strong>
                <button type="button" onClick={() => void retryFailures()} disabled={loading} style={{ ...dmsBtn('danger', !loading), height: 'auto', padding: '6px 10px', marginLeft: 'auto' }}>Retry All</button>
              </div>
              <div style={{ display: 'grid', gap: 6, maxHeight: 180, overflowY: 'auto' }}>
                {tagFailures.map(failure => (
                  <div key={failure.file_id} style={{ display: 'flex', gap: 8, alignItems: 'center', fontSize: 12, color: dmsTone('warning').fg }}>
                    <span style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={failure.error_reason}>{failure.filename} <span style={{ color: dmsTone('warning').fg, opacity: 0.8 }}>({failure.attempt_count} attempts)</span></span>
                    <button type="button" onClick={() => void retryFailures([failure.file_id])} disabled={loading} style={{ ...dmsRowBtn('plain'), height: 'auto', padding: '4px 8px' }}>Retry</button>
                    <button type="button" onClick={() => void dismissFailures([failure.file_id])} style={{ ...dmsRowBtn('plain'), height: 'auto', padding: '4px 8px' }}>Dismiss</button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Folder and file count summary (matching Documents module with recursive accuracy) */}
          <div style={{ fontSize: 13, color: 'var(--vdms-text-secondary)', marginBottom: 14, display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
            <span style={{ fontWeight: 700, color: 'var(--vdms-text)' }}>
              <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 13 }} /> {summaryCounts ? summaryCounts.direct_folders : items.filter(i => i.folder).length} {(summaryCounts ? summaryCounts.direct_folders : items.filter(i => i.folder).length) === 1 ? 'folder' : 'folders'}
            </span>
            <span>•</span>
            <span style={{ fontWeight: 700, color: 'var(--vdms-text)', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              <Icon iconName="Page" aria-hidden="true" style={{ fontSize: 13 }} /> {summaryCounts ? `${summaryCounts.total_files} total files` : `${items.filter(i => !i.folder).length} files`}
              {summaryCounts && summaryCounts.direct_files === 0 && summaryCounts.direct_folders > 0 && (
                <span style={{ fontWeight: 400, color: 'var(--vdms-text-muted)' }}>(0 at this level)</span>
              )}
              {summaryCounts && summaryCounts.direct_files > 0 && summaryCounts.direct_folders > 0 && (
                <span style={{ fontWeight: 400, color: 'var(--vdms-text-muted)' }}>({summaryCounts.direct_files} at this level)</span>
              )}
              {countsLoading && (
                <span style={{ fontSize: 11, color: clay.accent, fontWeight: 500 }}>
                  (calculating subfolders...)
                </span>
              )}
            </span>
          </div>

          {/* Quick Vessel & Tags Action Bar (Without Scanning) */}
          {items.length > 0 && (
            <div style={{
              display: 'flex', flexWrap: 'wrap', gap: 10, alignItems: 'center',
              padding: '10px 14px', background: 'var(--vdms-surface-alt)', border: '1px solid var(--vdms-border)',
              borderRadius: 8, marginBottom: 14,
            }}>
              <span style={{ fontSize: 13, fontWeight: 700, color: 'var(--vdms-text)', display: 'flex', alignItems: 'center', gap: 6 }}>
                <Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 13 }} /> Vessel:
              </span>

              <select
                value={bulkVessel}
                onChange={e => setBulkVessel(e.target.value)}
                style={{ ...dmsControlStyle(), padding: '6px 10px', fontSize: 13, minWidth: 180, fontWeight: 500 }}
              >
                <option value="">-- Choose vessel --</option>
                {detectedVessel && (
                  <option value={detectedVessel}>⚡ {detectedVessel} (from folder)</option>
                )}
                {vesselOptions.filter(v => v !== detectedVessel).map(v => (
                  <option key={v} value={v}>{v}</option>
                ))}
              </select>

              <input
                type="text"
                placeholder="Or type vessel name..."
                value={bulkVessel}
                onChange={e => setBulkVessel(e.target.value)}
                style={{ ...dmsControlStyle(), padding: '6px 10px', fontSize: 13, width: 170 }}
              />

              {(() => {
                const targetFiles = selected.size > 0
                  ? items.filter(i => selected.has(i.id) && !i.folder)
                  : items.filter(i => !i.folder);
                const missingCount = targetFiles.filter(i => !i.tags?.vessel).length;
                const taggedCount = targetFiles.filter(i => Boolean(i.tags?.vessel)).length;
                const hasFolders = selected.size > 0
                  ? items.some(i => selected.has(i.id) && i.folder)
                  : items.some(i => i.folder);
                const totalTargetCount = selected.size > 0 ? selected.size : (summaryCounts ? summaryCounts.total_files : items.length);

                return (
                  <>
                    {/* Primary Button: Updates ONLY missing vessel files */}
                    {(missingCount > 0 || hasFolders) && (
                      <button
                        type="button"
                        disabled={!bulkVessel.trim() || loading}
                        onClick={() => void applyBulkVessel(undefined, undefined, { onlyMissing: true })}
                        style={{ ...dmsBtn('primary', Boolean(bulkVessel.trim()) && !loading), height: 'auto', padding: '7px 14px' }}
                        title={missingCount > 0
                          ? `Update vessel tag "${bulkVessel.trim()}" directly on the ${missingCount} file(s) missing a vessel tag`
                          : 'Update vessel tag on missing files across sub-folders'}
                      >
                        <Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 13 }} /> Tag Missing Vessels {missingCount > 0 ? `(${missingCount})` : (hasFolders ? '(missing only)' : '')}
                      </button>
                    )}

                    {/* When all files in this view already have a vessel tag */}
                    {missingCount === 0 && !hasFolders && (
                      <button
                        type="button"
                        disabled={!bulkVessel.trim() || loading}
                        onClick={() => void applyBulkVessel(undefined, undefined, { overwriteExisting: true })}
                        style={{ background: dmsTone('success').bg, color: dmsTone('success').fg, border: '1px solid transparent', borderRadius: 6, fontWeight: 600, fontSize: 13, height: 'auto', padding: '7px 14px', cursor: (!bulkVessel.trim() || loading) ? 'not-allowed' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6, opacity: (!bulkVessel.trim() || loading) ? 0.5 : 1 }}
                        title={`All ${totalTargetCount} file(s) already have a vessel tag. Click to overwrite (confirmation required).`}
                      >
                        <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 13 }} /> All Tagged — Overwrite ({totalTargetCount})
                      </button>
                    )}

                    {/* Secondary Overwrite Button: requires explicit confirmation before updating existing tags */}
                    {taggedCount > 0 && (missingCount > 0 || hasFolders) && (
                      <button
                        type="button"
                        disabled={!bulkVessel.trim() || loading}
                        onClick={() => void applyBulkVessel(undefined, undefined, { overwriteExisting: true })}
                        style={{ ...dmsBtn('secondary', Boolean(bulkVessel.trim()) && !loading), height: 'auto', padding: '7px 12px', fontWeight: 500, fontSize: 12 }}
                        title={`Overwrite vessel tag on all ${totalTargetCount} file(s) including the ${taggedCount} already-tagged file(s) (requires confirmation)`}
                      >
                        <Icon iconName="Refresh" aria-hidden="true" style={{ fontSize: 12 }} /> Overwrite All ({totalTargetCount})
                      </button>
                    )}
                  </>
                );
              })()}

              <button
                type="button"
                disabled={loading}
                onClick={() => void promptAutoTagFromPath()}
                style={{ background: dmsTone('success').bg, color: dmsTone('success').fg, border: '1px solid transparent', borderRadius: 6, fontWeight: 600, fontSize: 12, height: 'auto', padding: '7px 12px', cursor: loading ? 'not-allowed' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 5, opacity: loading ? 0.6 : 1 }}
                title={selected.size > 0
                  ? `Automatically tag ${selected.size} selected item(s) from folder path and taxonomy`
                  : 'Automatically tag all files from folder path and taxonomy across selected items and sub-folders'}
              >
                <Icon iconName="LightningBolt" aria-hidden="true" style={{ fontSize: 12 }} /> Auto-Tag from Folder Path {selected.size > 0 ? `(${selected.size} selected)` : (detectedVessel ? `(${detectedVessel})` : '')}
              </button>

              {/* Direct OCR run on exactly the files missing a tag — no modal, since scope is already known here */}
              {(() => {
                const missingTagFiles = items.filter(i => !i.folder && hasAnyMissingTag(i));
                if (missingTagFiles.length === 0) return null;
                return (
                  <button
                    type="button"
                    disabled={loading}
                    onClick={() => void autoTagAndReview(missingTagFiles.map(i => i.id), 'missing_only')}
                    style={{ background: dmsTone('accent').bg, color: dmsTone('accent').fg, border: '1px solid transparent', borderRadius: 6, fontWeight: 600, fontSize: 12, height: 'auto', padding: '7px 12px', cursor: loading ? 'not-allowed' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 5, opacity: loading ? 0.6 : 1 }}
                    title={`Run OCR & AI classification directly on the ${missingTagFiles.length} file(s) missing a vessel tag. Files with an existing vessel tag are skipped.`}
                  >
                    <Icon iconName="Search" aria-hidden="true" style={{ fontSize: 12 }} /> Run OCR on Missing Tags ({missingTagFiles.length})
                  </button>
                );
              })()}

              <span style={{ marginLeft: 'auto', fontSize: 12 }}>
                {items.filter(i => !i.folder && !i.tags?.vessel).length > 0 ? (
                  <button
                    type="button"
                    disabled={loading || !bulkVessel.trim()}
                    onClick={() => void applyBulkVessel(undefined, undefined, { onlyMissing: true })}
                    style={{
                      color: dmsTone('warning').fg, fontWeight: 600, background: dmsTone('warning').bg, padding: '5px 12px',
                      borderRadius: 6, border: '1px solid transparent', cursor: bulkVessel.trim() && !loading ? 'pointer' : 'default',
                      display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 12,
                    }}
                    title={bulkVessel.trim()
                      ? `Click to tag these ${items.filter(i => !i.folder && !i.tags?.vessel).length} missing files with "${bulkVessel.trim()}"`
                      : 'Select or enter a vessel name to tag these missing files'}
                  >
                    <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 12 }} /> {items.filter(i => !i.folder && !i.tags?.vessel).length} file{items.filter(i => !i.folder && !i.tags?.vessel).length > 1 ? 's' : ''} missing vessel tag
                    {bulkVessel.trim() && <span style={{ textDecoration: 'underline', fontWeight: 700, marginLeft: 2 }}>— Tag Now</span>}
                  </button>
                ) : (
                  <span style={{ color: dmsTone('success').fg, fontWeight: 600, background: dmsTone('success').bg, padding: '5px 10px', borderRadius: 6 }}>
                    <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 12 }} /> Ready to auto-tag
                  </span>
                )}
              </span>
            </div>
          )}

          {/* Item list table */}
          <div style={{ background: 'var(--vdms-surface)', border: '1px solid var(--vdms-line)', borderRadius: 14, boxShadow: clay.shadowRaised, overflow: 'hidden' }}>
            {/* Column header */}
            <div style={{
              display: 'grid', gridTemplateColumns: '30px 30px minmax(180px,1fr) minmax(200px,2fr) 140px',
              gap: 12, alignItems: 'center', padding: '10px 14px',
              background: 'var(--vdms-surface-alt)', borderBottom: '1px solid var(--vdms-line)',
              fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', textTransform: 'uppercase', color: 'var(--vdms-text-muted)',
            }}>
              <input
                type="checkbox"
                title={items.length > 0 && items.every(i => selected.has(i.id)) ? 'Deselect all items' : 'Select all items (including sub-folders)'}
                aria-label="Select all items"
                checked={items.length > 0 && items.every(i => selected.has(i.id))}
                onChange={e => {
                  if (e.target.checked) {
                    setSelected(new Set(items.map(i => i.id)));
                  } else {
                    setSelected(new Set());
                  }
                }}
              />
              <span />
              <span style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                Name
                {selected.size > 0 && (() => {
                  const selItems = items.filter(i => selected.has(i.id));
                  const selFolders = selItems.filter(i => i.folder).length;
                  const selFiles = selItems.filter(i => !i.folder).length;
                  return (
                    <span style={{ fontSize: 11, background: dmsTone('accent').bg, color: dmsTone('accent').fg, padding: '1px 8px', borderRadius: 10, fontWeight: 600 }}>
                      {selFolders > 0 && selFiles > 0
                        ? `${selFolders} folder${selFolders > 1 ? 's' : ''}, ${selFiles} file${selFiles > 1 ? 's' : ''} selected`
                        : selFolders > 0
                          ? `${selFolders} folder${selFolders > 1 ? 's' : ''} selected`
                          : `${selFiles} file${selFiles > 1 ? 's' : ''} selected`}
                    </span>
                  );
                })()}
              </span>
              <span>Tags</span>
              <span style={{ textAlign: 'right' }}>Actions</span>
            </div>

            {items.map(item => (
              <div
                key={item.id}
                style={{
                  display: 'grid', gridTemplateColumns: '30px 30px minmax(180px,1fr) minmax(200px,2fr) 140px',
                  gap: 12, alignItems: 'center', padding: '12px 14px', borderBottom: '1px solid var(--vdms-border-soft)',
                  background: editing === item.id ? dmsTone('accent').bg : undefined,
                }}
              >
                {/* Checkbox */}
                <input
                  type="checkbox"
                  checked={selected.has(item.id)}
                  onChange={() => setSelected(previous => {
                    const next = new Set(previous);
                    next.has(item.id) ? next.delete(item.id) : next.add(item.id);
                    return next;
                  })}
                />

                {/* In-app drilldown arrow for folder / preview arrow for file */}
                <button
                  type="button"
                  onClick={() => {
                    if (item.folder) {
                      void loadFolder(context, item.id, item.name, [...crumbs, { id: item.id, name: item.name }]);
                    } else {
                      openSiteFile(item);
                    }
                  }}
                  title={item.folder ? `Open folder ${item.name}` : `Preview ${item.name}`}
                  style={{
                    width: 26, height: 26, padding: 0, border: '1px solid transparent', background: dmsTone('accent').bg,
                    color: dmsTone('accent').fg, borderRadius: 6, cursor: 'pointer', fontSize: 13, fontWeight: 700,
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                  }}
                >
                  {item.folder ? '→' : '↗'}
                </button>

                {/* File / folder name + file count badge for folders */}
                <button
                  type="button"
                  onClick={() => {
                    if (item.folder) {
                      void loadFolder(context, item.id, item.name, [...crumbs, { id: item.id, name: item.name }]);
                    } else {
                      openSiteFile(item);
                    }
                  }}
                  style={{ textAlign: 'left', border: 0, background: 'transparent', color: item.folder ? clay.accent : 'var(--vdms-text)', cursor: 'pointer', fontWeight: 600, fontSize: 13, display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}
                >
                  <span>{item.folder ? <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 13 }} /> : <Icon iconName="Page" aria-hidden="true" style={{ fontSize: 13 }} />} {item.name}</span>
                  {item.folder && (() => {
                    const fc = item.folder_counts;
                    if (fc) {
                      const hasSubfolders = fc.direct_subfolders > 0;
                      const totalFiles = fc.total_files;
                      if (hasSubfolders) {
                        return (
                          <span
                            style={{
                              display: 'inline-flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              background: totalFiles > 0 ? clay.accent : 'var(--vdms-text-faint)',
                              color: DMS_ON_ACCENT,
                              borderRadius: 20,
                              padding: '2px 10px',
                              fontSize: 11,
                              fontWeight: 700,
                              lineHeight: '16px',
                              whiteSpace: 'nowrap',
                            }}
                            title={`${fc.direct_subfolders} direct subfolder${fc.direct_subfolders === 1 ? '' : 's'}, ${totalFiles} total file${totalFiles === 1 ? '' : 's'} across all subfolders`}
                          >
                            <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 11 }} /> {fc.direct_subfolders} {fc.direct_subfolders === 1 ? 'subfolder' : 'subfolders'} · {totalFiles} {totalFiles === 1 ? 'file' : 'files'}
                          </span>
                        );
                      } else {
                        return (
                          <span
                            style={{
                              display: 'inline-flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              background: totalFiles > 0 ? clay.accent : 'var(--vdms-text-faint)',
                              color: DMS_ON_ACCENT,
                              borderRadius: 20,
                              padding: '1px 8px',
                              fontSize: 10,
                              fontWeight: 700,
                              lineHeight: '16px',
                              whiteSpace: 'nowrap',
                            }}
                            title={`${totalFiles} ${totalFiles === 1 ? 'file' : 'files'} inside ${item.name}`}
                          >
                            {totalFiles} {totalFiles === 1 ? 'file' : 'files'}
                          </span>
                        );
                      }
                    }

                    // Fallback while async counting is in progress
                    const fallbackCount = getFolderFileCount(item);
                    return (
                      <span
                        style={{
                          display: 'inline-flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          background: fallbackCount > 0 ? clay.accent : 'var(--vdms-text-faint)',
                          color: DMS_ON_ACCENT,
                          borderRadius: 20,
                          padding: '1px 8px',
                          fontSize: 10,
                          fontWeight: 700,
                          lineHeight: '16px',
                          whiteSpace: 'nowrap',
                        }}
                        title={countsLoading ? 'Updating recursive totals in background...' : `${fallbackCount} direct items`}
                      >
                        {fallbackCount} {fallbackCount === 1 ? 'file' : 'files'}
                      </span>
                    );
                  })()}
                </button>

                {/* Tags column — shows editable inputs or badges */}
                {editing === item.id ? (
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                    {tagNames.map(([key, label]) => (
                      key === 'vessel' ? (
                        <span key={key}>
                          <input
                            aria-label={label}
                            value={tagDraft.vessel}
                            onChange={event => setTagDraft(previous => ({ ...previous, vessel: event.target.value }))}
                            placeholder={label}
                            list="vessel-options-list"
                            style={{ ...dmsControlStyle(), width: 115, padding: 5 }}
                          />
                          <datalist id="vessel-options-list">
                            {detectedVessel && <option value={detectedVessel} />}
                            {vesselOptions.map(v => <option key={v} value={v} />)}
                          </datalist>
                        </span>
                      ) : (
                        <input
                          key={key}
                          aria-label={label}
                          value={tagDraft[key]}
                          onChange={event => setTagDraft(previous => ({ ...previous, [key]: event.target.value }))}
                          placeholder={label}
                          style={{ ...dmsControlStyle(), width: 105, padding: 5 }}
                        />
                      )
                    ))}
                  </div>
                ) : (
                  renderTags(item)
                )}

                {/* Action buttons (same pattern as Documents module) */}
                <div style={{ display: 'flex', gap: 6, alignItems: 'center', justifyContent: 'flex-end' }}>
                  {editing === item.id ? (
                    <>
                      <button
                        type="button"
                        onClick={() => void saveTags(item)}
                        style={{ ...dmsRowBtn('accent'), padding: '5px 9px' }}
                      >
                        Save
                      </button>
                      <button
                        type="button"
                        onClick={() => setEditing(null)}
                        style={{ ...dmsRowBtn('plain'), padding: '5px 8px' }}
                      >
                        Cancel
                      </button>
                    </>
                  ) : (
                    <>
                      {item.folder ? (
                        <button
                          type="button"
                          onClick={() => void promptAutoTagFromPath([item.id])}
                          title={`Auto-tag all files inside ${item.name} from folder path`}
                          style={{ ...dmsRowBtn('accent'), padding: '5px 9px' }}
                        >
                          <Icon iconName="LightningBolt" aria-hidden="true" style={{ fontSize: 12 }} /> Auto-Tag Files
                        </button>
                      ) : (
                        <button
                          type="button"
                          onClick={() => {
                            setEditing(item.id);
                            const existing = item.tags || {};
                            setTagDraft({
                              ...emptyTags,
                              ...existing,
                              vessel: existing.vessel || detectedVessel || bulkVessel || '',
                            });
                          }}
                          style={{ ...dmsRowBtn('plain'), padding: '5px 8px' }}
                        >
                          <Icon iconName="Edit" aria-hidden="true" style={{ fontSize: 12 }} /> Tags
                        </button>
                      )}
                      {/* Open in SharePoint arrow — exact same button as Documents module */}
                      <button
                        type="button"
                        onClick={() => openSharePointLocation(item)}
                        title={item.folder ? `Open ${item.name} folder in SharePoint` : `Open ${item.name} in SharePoint`}
                        aria-label={`Open ${item.name} in SharePoint`}
                        style={{
                          width: 28, height: 27, padding: 0, borderRadius: 6, border: '1px solid transparent',
                          background: dmsTone('accent').bg, color: dmsTone('accent').fg, cursor: 'pointer', fontSize: 16,
                          fontWeight: 700, lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                        }}
                      >
                        ↗
                      </button>
                    </>
                  )}
                </div>
              </div>
            ))}

            {!items.length && !loading && (
              <div style={{ padding: 24, color: 'var(--vdms-text-muted)', textAlign: 'center' }}>This folder is empty.</div>
            )}
            {loading && (
              <div style={{ padding: 24, color: 'var(--vdms-text-muted)', textAlign: 'center' }}>Loading items...</div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}