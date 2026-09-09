import * as React from 'react';
import type VesselEmail from '../VesselEmail';

type Site = { id: string; display_name: string; web_url?: string; description?: string; thumbnail?: string };
type Drive = { id: string; name: string; web_url?: string };
type FolderCounts = {
  direct_subfolders: number;
  direct_files: number;
  total_subfolders: number;
  total_files: number;
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
  tags?: Record<string, string>;
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



const emptyTags: Tags = { department: '', vessel: '', group: '', category: '' };
const tagNames: Array<[keyof Tags, string]> = [
  ['department', 'Department'],
  ['vessel', 'Vessel'],
  ['group', 'Group'],
  ['category', 'Category'],
];

export function SitesPage({ host }: { host: VesselEmail }): React.ReactElement {
  const [sites, setSites] = React.useState<Site[]>([]);
  const [query, setQuery] = React.useState('');
  const [expanded, setExpanded] = React.useState<string | null>(null);
  const [drives, setDrives] = React.useState<Record<string, Drive[]>>({});
  const [context, setContext] = React.useState<Context | null>(null);
  const [items, setItems] = React.useState<Item[]>([]);
  const [crumbs, setCrumbs] = React.useState<Array<{ id: string; name: string }>>([]);
  const [loading, setLoading] = React.useState(false);
  const [message, setMessage] = React.useState('');
  const [selected, setSelected] = React.useState<Set<string>>(new Set());
  const [editing, setEditing] = React.useState<string | null>(null);
  const [tagDraft, setTagDraft] = React.useState<Tags>(emptyTags);
  const [recursive, setRecursive] = React.useState(false);
  const [scanResults, setScanResults] = React.useState<ScanResult[]>([]);
  const [fieldChoices, setFieldChoices] = React.useState<Record<string, Record<keyof Tags, string>>>({});
  const [vesselOptions, setVesselOptions] = React.useState<string[]>([]);
  const [bulkVessel, setBulkVessel] = React.useState('');
  const [detectedVessel, setDetectedVessel] = React.useState('');
  const [detectedTags, setDetectedTags] = React.useState<Tags>(emptyTags);
  const [taggingModal, setTaggingModal] = React.useState<TaggingModal>(null);
  const [summaryCounts, setSummaryCounts] = React.useState<SummaryCounts | null>(null);
  const [countsLoading, setCountsLoading] = React.useState(false);

  const api = host._base();
  const headers = host._headers();

  /** Async backfill for recursive counts so cold-cache folder views never block initial rendering */
  const fetchSubfolderCounts = React.useCallback(async (ctx: Context, folderId: string): Promise<void> => {
    setCountsLoading(true);
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(ctx.site.id)}/drives/${encodeURIComponent(ctx.drive.id)}/folders/${encodeURIComponent(folderId)}/subfolder-counts`, { headers });
      if (response.ok) {
        const data = await response.json();
        const countsMap: Record<string, FolderCounts> = data.counts || {};
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
      // Non-fatal: UI gracefully falls back to direct properties
    } finally {
      setCountsLoading(false);
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
    setLoading(true);
    try {
      const response = await fetch(`${api}/api/sites`, { headers });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Could not load sites');
      setSites(data.sites || []);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Could not load sites');
    } finally {
      setLoading(false);
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

  /** Navigate into a folder and reset scan state */
  const loadFolder = async (nextContext: Context, folderId: string, name: string, nextCrumbs: Array<{ id: string; name: string }>): Promise<void> => {
    setLoading(true); setMessage(''); setScanResults([]); setFieldChoices({});
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(nextContext.site.id)}/drives/${encodeURIComponent(nextContext.drive.id)}/folders/${encodeURIComponent(folderId)}/children`, { headers });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Could not load folder');
      setContext(nextContext);
      const childItems: Item[] = data.items || [];
      setItems(childItems);
      setSummaryCounts(data.summary_counts || null);
      setCrumbs(nextCrumbs.length ? nextCrumbs : [{ id: 'root', name }]);
      setSelected(new Set());
      const detV = data.detected_vessel || '';
      setDetectedVessel(detV);
      if (detV) {
        setBulkVessel(detV);
      }
      setDetectedTags(data.detected_tags || emptyTags);

      // Trigger non-blocking async count resolution if any folder lacks cached counts
      const hasUncached = childItems.some(i => i.folder && !i.folder_counts);
      if (hasUncached) {
        void fetchSubfolderCounts(nextContext, folderId);
      }
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Could not load folder');
    } finally {
      setLoading(false);
    }
  };

  /** Reload folder items without clearing scanResults so the review panel stays visible after scan */
  const refreshFolder = async (ctx: Context, folderId: string, currentCrumbs: Array<{ id: string; name: string }>): Promise<void> => {
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(ctx.site.id)}/drives/${encodeURIComponent(ctx.drive.id)}/folders/${encodeURIComponent(folderId)}/children`, { headers });
      const data = await response.json();
      if (response.ok) {
        const childItems: Item[] = data.items || [];
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

        const hasUncached = childItems.some(i => i.folder && !i.folder_counts);
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
      setEditing(null);
      setMessage(`Tags saved for ${item.name}`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Could not save tags');
    }
  };

  const scan = async (): Promise<void> => {
    if (!context || selected.size === 0) return;
    setLoading(true);
    setScanResults([]);
    setFieldChoices({});
    setMessage('Scanning selected items with OCR and AI classification...');
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/scan-tags`, {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify({ item_ids: Array.from(selected), recursive }),
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
          if (ocrVal && ocrConf >= ocrThreshold) {
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
      setScanResults(results);

      if (crumbs.length > 0) {
        const crumb = crumbs[crumbs.length - 1];
        await refreshFolder(context, crumb.id, crumbs);
      }
      setMessage(`Scanned ${data.scanned || results.length} item(s). Review suggestions and confirm below.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Scan failed');
    } finally {
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
      for (let i = 0; i < reviewable.length; i++) {
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
    }
  };

  /** Update vessel on selected files/folders directly in SharePoint WITHOUT scanning */
  const applyBulkVessel = async (targetItemIds?: string[], specificVessel?: string): Promise<void> => {
    if (!context) return;
    const vesselToApply = (specificVessel !== undefined ? specificVessel : bulkVessel).trim();
    if (!vesselToApply) {
      setMessage('Please select or enter a vessel name.');
      return;
    }

    let idsToUpdate: string[] = [];
    if (targetItemIds && targetItemIds.length > 0) {
      idsToUpdate = targetItemIds;
    } else if (selected.size > 0) {
      idsToUpdate = Array.from(selected);
    } else {
      idsToUpdate = items.map(i => i.id);
    }

    if (idsToUpdate.length === 0) {
      setMessage('No items available for vessel update.');
      return;
    }

    // Safety check for large operations
    const estimatedFiles = summaryCounts?.total_files ?? items.length;
    const hasFolders = idsToUpdate.some(id => items.find(it => it.id === id)?.folder);
    if (hasFolders || recursive) {
      if (estimatedFiles > 200) {
        const ok = window.confirm(`Found ${estimatedFiles} files across sub-folders. For system safety, this run will process the first 200 files (the remaining ${estimatedFiles - 200} can be run in a subsequent batch).\n\nDo you want to proceed?`);
        if (!ok) return;
      } else if (estimatedFiles > 50) {
        const ok = window.confirm(`This will apply the vessel tag "${vesselToApply}" to ${estimatedFiles} files across all sub-folders.\n\nDo you want to proceed?`);
        if (!ok) return;
      }
    }

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
      title: `Updating Vessel Tag: "${vesselToApply}"`,
      feed: initialFeed,
      total: Math.max(estimatedFiles, idsToUpdate.length),
      finished: false,
      summary: '',
    });

    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/bulk-update-tags`, {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify({ item_ids: idsToUpdate, vessel: vesselToApply, recursive: true }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Could not update vessel');

      const results = data.results || [];
      const okCount = results.filter((r: { ok: boolean }) => r.ok).length;
      const failCount = results.filter((r: { ok: boolean }) => !r.ok).length;

      const feed: TagFeedEntry[] = results.map((r: { filename?: string; parent_path?: string; item_id: string; ok: boolean; error?: string; tags?: Record<string, string> }) => {
        const folderName = r.parent_path ? r.parent_path.split('/').pop() || '' : '';
        const tagLabel = r.tags?.vessel ? ` → ${r.tags.vessel}` : ` → ${vesselToApply}`;
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
          name: 'No files found to update in the selected location.',
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
        ? `✓ ${okCount} file(s) updated. ✗ ${failCount} file(s) failed.`
        : `✓ All ${okCount} file(s) successfully updated with vessel "${vesselToApply}"!`;
      if (data.truncated) {
        summary += ` (Capped at 200 files for safety; run again for remaining files.)`;
      }
      setTaggingModal({
        title: `Updating Vessel Tag: "${vesselToApply}"`,
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
      const errMsg = err instanceof Error ? err.message : 'Could not update vessel';
      setTaggingModal(prev => prev ? { ...prev, finished: true, summary: `✗ Error: ${errMsg}` } : prev);
      setMessage(errMsg);
    } finally {
      setLoading(false);
    }
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

    // Safety confirmation for large operations
    const estimatedFiles = summaryCounts?.total_files ?? items.length;
    const hasFolders = idsToUpdate.some(id => items.find(it => it.id === id)?.folder);
    if (hasFolders || recursive) {
      if (estimatedFiles > 200) {
        const ok = window.confirm(`Found ${estimatedFiles} files across sub-folders. For system safety, this run will process the first 200 files (the remaining ${estimatedFiles - 200} can be run in a subsequent batch).\n\nDo you want to proceed?`);
        if (!ok) return;
      } else if (estimatedFiles > 50) {
        const ok = window.confirm(`This will apply folder path tags to ${estimatedFiles} files across all sub-folders.\n\nDo you want to proceed?`);
        if (!ok) return;
      }
    }

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
      title: 'Auto-Tagging from Folder Path',
      feed: initialFeed,
      total: Math.max(estimatedFiles, idsToUpdate.length),
      finished: false,
      summary: '',
    });

    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/bulk-update-tags`, {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify({ item_ids: idsToUpdate, auto_from_path: true, recursive: true }),
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
        summary += ` (Capped at 200 files for safety; run again for remaining files.)`;
      }
      setTaggingModal({
        title: 'Auto-Tagging from Folder Path',
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

  /** Simultaneous Auto-Tag & Review across sub-folders */
  const autoTagAndReview = async (targetItemIds?: string[]): Promise<void> => {
    if (!context) return;
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

    const estimatedFiles = summaryCounts?.total_files ?? items.length;
    if (estimatedFiles > 200) {
      const ok = window.confirm(`Found ${estimatedFiles} files across sub-folders. For system safety, this run will process the first 200 files (remaining files can be run in a subsequent batch).\n\nDo you want to proceed?`);
      if (!ok) return;
    } else if (estimatedFiles > 50) {
      const ok = window.confirm(`This will scan and auto-tag ${estimatedFiles} files across all sub-folders for review.\n\nDo you want to proceed?`);
      if (!ok) return;
    }

    setLoading(true);
    setScanResults([]);
    setFieldChoices({});
    setMessage('Auto-tagging and scanning selected items with OCR & folder taxonomy...');
    try {
      const response = await fetch(`${api}/api/sites/${encodeURIComponent(context.site.id)}/drives/${encodeURIComponent(context.drive.id)}/scan-tags`, {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify({ item_ids: idsToScan, recursive: true }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Scan failed');
      const results: ScanResult[] = data.results || [];

      // Automatically pre-select best tags according to precedence rules
      const initialChoices: Record<string, Record<keyof Tags, string>> = {};
      results.forEach(res => {
        initialChoices[res.item_id] = { ...emptyTags };
        (['department', 'vessel', 'group', 'category'] as Array<keyof Tags>).forEach(field => {
          const ocrVal = res.ocr_suggestion?.[field]?.value;
          const ocrConf = res.ocr_suggestion?.[field]?.confidence || 0;
          const pathVal = res.path_suggestion?.[field]?.value;
          const isVesselOrGroup = field === 'vessel' || field === 'group';
          const ocrThreshold = isVesselOrGroup ? 0.85 : 0.60;
          if (ocrVal && ocrConf >= ocrThreshold) {
            initialChoices[res.item_id][field] = 'ocr';
          } else if (pathVal) {
            initialChoices[res.item_id][field] = 'path';
          } else if (ocrVal && ocrConf >= 0.40) {
            initialChoices[res.item_id][field] = 'ocr';
          } else if (res.proposed_tags?.[field]) {
            initialChoices[res.item_id][field] = 'manual';
          } else {
            initialChoices[res.item_id][field] = 'skip';
          }
        });
      });
      setFieldChoices(initialChoices);
      setScanResults(results);

      if (crumbs.length > 0) {
        const crumb = crumbs[crumbs.length - 1];
        await refreshFolder(context, crumb.id, crumbs);
      }
      const countMsg = data.truncated
        ? `⚡ Scanned & auto-tagged ${results.length} item(s) (capped at 200). Review suggestions below and confirm!`
        : `⚡ Scanned & auto-tagged ${results.length} item(s) across sub-folders. Review suggestions below and confirm!`;
      setMessage(countMsg);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Scan and review failed');
    } finally {
      setLoading(false);
    }
  };

  const filteredSites = sites.filter(site => `${site.display_name} ${site.description || ''}`.toLowerCase().includes(query.toLowerCase()));

  const renderTags = (item: Item): React.ReactElement => (
    <span style={{ display: 'flex', flexWrap: 'wrap', gap: 4, alignItems: 'center' }}>
      {tagNames.filter(([key]) => item.tags?.[key]).map(([key, label]) => (
        <span key={key} style={{ background: '#e0f2fe', color: '#075985', borderRadius: 12, padding: '3px 8px', fontSize: 11, fontWeight: 500 }}>
          {label}: {item.tags?.[key]}
        </span>
      ))}
      {!item.folder && !item.tags?.vessel && (
        <button
          type="button"
          onClick={() => void applyBulkVessel([item.id], detectedVessel || bulkVessel)}
          title={`Click to set vessel "${detectedVessel || bulkVessel || 'vessel'}" without scanning`}
          style={{
            background: '#fef3c7', color: '#92400e', border: '1px dashed #f59e0b',
            borderRadius: 12, padding: '2px 8px', fontSize: 11, fontWeight: 600,
            cursor: (detectedVessel || bulkVessel) ? 'pointer' : 'default',
            display: 'inline-flex', alignItems: 'center', gap: 3,
          }}
        >
          + Vessel{detectedVessel ? `: ${detectedVessel}` : (bulkVessel ? `: ${bulkVessel}` : '')}
        </button>
      )}
      {!tagNames.some(([key]) => item.tags?.[key]) && item.folder && (() => {
        const fileCount = item.folder_counts ? item.folder_counts.total_files : getFolderFileCount(item);
        return (
          <span style={{ color: '#64748b', fontSize: 11, fontStyle: 'italic', display: 'inline-flex', alignItems: 'center', gap: 4 }}>
            📁 Sub-folder ({fileCount} file{fileCount === 1 ? '' : 's'})
          </span>
        );
      })()}
    </span>
  );

  const reviewPanel = scanResults.length > 0 ? (
    <div style={{ marginTop: 16, marginBottom: 20, background: '#fff7ed', border: '1px solid #fed7aa', padding: 16, borderRadius: 8, boxShadow: '0 2px 8px rgba(0,0,0,0.04)' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
        <strong style={{ fontSize: 15, color: '#9a3412' }}>🔍 Review OCR & AI Tag Suggestions ({scanResults.length} item{scanResults.length > 1 ? 's' : ''})</strong>
        <button
          onClick={() => setScanResults([])}
          style={{ border: 0, background: 'transparent', cursor: 'pointer', fontSize: 16, color: '#9a3412' }}
          title="Close review panel"
        >✕</button>
      </div>

      {/* Quick vessel apply bar */}
      <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 12, padding: '8px 12px', background: '#fef9c3', border: '1px solid #fde68a', borderRadius: 6 }}>
        <span style={{ fontSize: 12, fontWeight: 600, color: '#854d0e', whiteSpace: 'nowrap' }}>⚓ Apply vessel to all:</span>
        <select
          value={bulkVessel}
          onChange={e => setBulkVessel(e.target.value)}
          style={{ flex: 1, maxWidth: 260, padding: '5px 8px', borderRadius: 4, border: '1px solid #fbbf24', fontSize: 12, background: '#fff' }}
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
          style={{
            padding: '5px 14px', background: bulkVessel ? '#b45309' : '#d1d5db', color: '#fff',
            border: 0, borderRadius: 4, cursor: bulkVessel ? 'pointer' : 'not-allowed',
            fontSize: 12, fontWeight: 600, whiteSpace: 'nowrap',
          }}
        >
          Apply to all
        </button>
      </div>

      {scanResults.map(result => (
        <div key={result.item_id} style={{ padding: '14px 0', borderBottom: '1px solid #fed7aa' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              <span style={{ fontWeight: 700, fontSize: 13, color: '#1e293b' }}>
                📄 {result.filename || result.item_id}
              </span>
              {result.subfolder_name && (
                <span style={{ background: '#f1f5f9', border: '1px solid #e2e8f0', color: '#475569', fontSize: 11, padding: '2px 8px', borderRadius: 12, fontWeight: 600 }}>
                  📁 {result.subfolder_name}
                </span>
              )}
            </div>
            <span>
              {result.status === 'confirmed'
                ? <span style={{ color: '#15803d', fontWeight: 600 }}>✓ Confirmed and saved to SharePoint</span>
                : result.error
                  ? <span style={{ color: '#b91c1c', fontWeight: 600 }}>✗ {result.error}</span>
                  : <span style={{ color: '#c2410c', fontWeight: 600 }}>⚠ Ready for review & confirmation</span>
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
                return (
                  <div key={field} style={{ background: '#fff', border: '1px solid #fed7aa', borderRadius: 6, padding: '8px 10px' }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: '#475569', marginBottom: 4 }}>
                      {fieldLabel}
                    </div>
                    <select
                      value={choice}
                      onChange={event => setFieldChoices(previous => ({
                        ...previous,
                        [result.item_id]: { ...emptyTags, ...(previous[result.item_id] || {}), [field]: event.target.value },
                      }))}
                      style={{ width: '100%', padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 4, fontSize: 12, background: '#fff' }}
                    >
                      <option value="skip">Skip / Keep unchanged</option>
                      <option value="ocr" disabled={!ocr?.value}>
                        OCR: {ocr?.value || 'None'}{ocr ? ` (${Math.round(ocr.confidence * 100)}%)` : ''}
                      </option>
                      <option value="path" disabled={!path?.value}>
                        Path: {path?.value || 'None'}
                      </option>
                      <option value="manual">Enter manually</option>
                    </select>
                    {choice === 'manual' && (
                      isVessel && vesselOptions.length > 0 ? (
                        <select
                          value={result.proposed_tags?.vessel || ''}
                          onChange={event => updateProposedTag(result.item_id, 'vessel', event.target.value)}
                          style={{ width: '100%', boxSizing: 'border-box', marginTop: 6, padding: 6, border: '1px solid #cbd5e1', borderRadius: 4, fontSize: 12, background: '#fff' }}
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
                          style={{ width: '100%', boxSizing: 'border-box', marginTop: 6, padding: 6, border: '1px solid #cbd5e1', borderRadius: 4, fontSize: 12 }}
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
          style={{
            padding: '10px 16px', background: '#0f766e', color: '#fff', border: 0, borderRadius: 7,
            cursor: loading || !scanResults.some(result => result.status === 'needs_selection') ? 'not-allowed' : 'pointer',
            fontWeight: 600, fontSize: 13, display: 'inline-flex', alignItems: 'center', gap: 6,
          }}
        >
          ⚡ Auto-Select All High Confidence
        </button>

        <button
          disabled={loading || !scanResults.some(result => result.status === 'needs_selection')}
          onClick={() => void confirmTags()}
          style={{
            padding: '10px 18px', background: '#0369a1', color: '#fff', border: 0, borderRadius: 7,
            cursor: loading || !scanResults.some(result => result.status === 'needs_selection') ? 'not-allowed' : 'pointer',
            fontWeight: 600, fontSize: 13, opacity: scanResults.some(result => result.status === 'needs_selection') ? 1 : 0.6,
          }}
        >
          ✓ Confirm & Apply Selected Tags
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
          background: '#fff', borderRadius: 16, padding: '28px 32px', width: '100%', maxWidth: 540,
          boxShadow: '0 20px 60px rgba(0,0,0,0.25)', display: 'flex', flexDirection: 'column', gap: 16,
        }}>
          {/* Header */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: '#eff6ff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>⚓</div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 16, color: '#0f172a' }}>{taggingModal.title}</div>
              <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>
                {taggingModal.finished ? `Completed (${doneCount} file${doneCount === 1 ? '' : 's'} updated)` : `Processing ${doneCount} of ${taggingModal.total} file(s)…`}
              </div>
            </div>
          </div>

          {/* Progress bar */}
          <div style={{ height: 6, background: '#e2e8f0', borderRadius: 99, overflow: 'hidden' }}>
            <div style={{
              height: '100%', borderRadius: 99,
              background: taggingModal.finished && taggingModal.feed.some(f => f.status === 'failed') ? '#ef4444' : '#0284c7',
              width: `${pct}%`, transition: 'width 0.3s ease',
            }} />
          </div>

          {/* File feed list */}
          <div style={{
            maxHeight: 320, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 4,
            background: '#f8fafc', borderRadius: 10, padding: '10px 12px',
            border: '1px solid #e2e8f0',
          }}>
            {taggingModal.feed.map((entry, idx) => (
              <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, padding: '3px 0' }}>
                <span style={{ fontSize: 15, flexShrink: 0 }}>
                  {entry.status === 'pending' ? '⏳' : entry.status === 'ok' ? '✅' : '❌'}
                </span>
                <span style={{
                  flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                  color: entry.status === 'pending' ? '#64748b' : entry.status === 'ok' ? '#15803d' : '#dc2626',
                  fontWeight: entry.status !== 'pending' ? 600 : 400,
                }}>
                  {entry.name}
                </span>
                {entry.status === 'failed' && entry.error && (
                  <span style={{ fontSize: 11, color: '#dc2626', whiteSpace: 'nowrap', maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis' }}
                    title={entry.error}>
                    {entry.error}
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
                background: taggingModal.feed.every(f => f.status !== 'failed') ? '#dcfce7' : '#fff7ed',
                color: taggingModal.feed.every(f => f.status !== 'failed') ? '#15803d' : '#92400e',
                fontWeight: 700, fontSize: 14, textAlign: 'center',
              }}>
                {taggingModal.summary}
              </div>
              <button
                type="button"
                onClick={() => setTaggingModal(null)}
                style={{
                  padding: '10px 20px', background: '#0284c7', color: '#fff', border: 0,
                  borderRadius: 8, fontWeight: 700, fontSize: 14, cursor: 'pointer', alignSelf: 'flex-end',
                }}
              >
                ✓ Close
              </button>
            </>
          )}
        </div>
      </div>
    );
  };

  return (
    <div style={{ padding: '28px 32px', background: '#f7fafc', minHeight: '100%' }}>
      {renderTaggingModal()}
      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 16, alignItems: 'center', marginBottom: 24, flexWrap: 'wrap' }}>
        <div>
          <h1 style={{ margin: 0, color: '#123044', fontSize: 28 }}>Sites</h1>
          <p style={{ color: '#64748b', margin: '6px 0 0' }}>Browse libraries and tag documents across your tenant.</p>
        </div>
        <input
          aria-label="Search sites"
          value={query}
          onChange={event => setQuery(event.target.value)}
          placeholder="Search sites"
          style={{ width: 260, padding: '11px 14px', border: '1px solid #cbd5e1', borderRadius: 8 }}
        />
      </div>

      {message && (
        <div style={{ padding: 12, marginBottom: 16, background: '#fff7ed', color: '#9a3412', borderRadius: 8, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <span>{message}</span>
          <button onClick={() => setMessage('')} style={{ border: 0, background: 'transparent', cursor: 'pointer', fontSize: 16, color: '#9a3412' }}>✕</button>
        </div>
      )}

      {reviewPanel}

      {/* Sites grid */}
      {!context && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(260px,1fr))', gap: 16 }}>
          {loading && <div style={{ color: '#64748b' }}>Loading sites...</div>}
          {filteredSites.map(site => (
            <div key={site.id} style={{ background: '#fff', border: '1px solid #dbe5ec', borderRadius: 10, padding: 18, boxShadow: '0 2px 8px #1230440d' }}>
              <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                <div style={{ width: 42, height: 42, background: '#e0f2fe', color: '#0284c7', display: 'grid', placeItems: 'center', borderRadius: 8, fontWeight: 700 }}>SP</div>
                <div>
                  <h2 style={{ margin: 0, fontSize: 16, color: '#123044' }}>{site.display_name}</h2>
                  <small style={{ color: '#64748b' }}>{site.web_url || site.id}</small>
                </div>
              </div>
              <p style={{ color: '#64748b', fontSize: 13, minHeight: 34 }}>{site.description || 'SharePoint site'}</p>
              <button onClick={() => void loadDrives(site)} style={{ padding: '9px 12px', border: '1px solid #0284c7', color: '#0369a1', background: '#f0f9ff', borderRadius: 7, cursor: 'pointer' }}>
                {expanded === site.id ? 'Hide libraries' : 'Show libraries'}
              </button>
              {expanded === site.id && (
                <div style={{ marginTop: 12, display: 'grid', gap: 6 }}>
                  {(drives[site.id] || []).map(drive => (
                    <button key={drive.id} onClick={() => chooseDrive(site, drive)} style={{ textAlign: 'left', padding: 10, border: '1px solid #e2e8f0', background: '#fff', borderRadius: 6, cursor: 'pointer' }}>
                      📚 {drive.name}
                    </button>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      )}

      {/* Folder / file browser */}
      {context && (
        <div>
          <button
            onClick={() => { setContext(null); setScanResults([]); setMessage(''); }}
            style={{ border: 0, background: 'transparent', color: '#0369a1', cursor: 'pointer', padding: 0, marginBottom: 16 }}
          >
            ← All sites
          </button>

          <h2 style={{ color: '#123044', margin: '0 0 10px' }}>
            {context.site.display_name} / {context.drive.name}
          </h2>

          {/* Breadcrumb + action toolbar */}
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap', marginBottom: 16 }}>
            {crumbs.map((crumb, index) => (
              <React.Fragment key={crumb.id}>
                <button
                  onClick={() => void loadFolder(context, crumb.id, crumb.name, crumbs.slice(0, index + 1))}
                  style={{ border: 0, background: 'transparent', color: '#0369a1', cursor: 'pointer', fontWeight: index === crumbs.length - 1 ? 700 : 400 }}
                >
                  {crumb.name}
                </button>
                {index < crumbs.length - 1 && <span style={{ color: '#94a3b8' }}>/</span>}
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
                style={{
                  padding: '9px 13px', background: '#eff6ff', color: '#1d4ed8', border: '1px solid #bfdbfe', borderRadius: 7,
                  cursor: 'pointer', fontWeight: 600, fontSize: 13, display: 'inline-flex', alignItems: 'center', gap: 6,
                }}
              >
                ↗ Open in SharePoint
              </button>
              <label style={{ fontSize: 13, display: 'flex', gap: 4, alignItems: 'center', cursor: 'pointer' }}>
                <input type="checkbox" checked={recursive} onChange={event => setRecursive(event.target.checked)} />
                Recursive
              </label>
              <button
                disabled={loading}
                onClick={() => void autoTagAndReview()}
                style={{
                  padding: '9px 13px', background: '#0284c7', color: '#fff', border: 0, borderRadius: 7,
                  cursor: loading ? 'not-allowed' : 'pointer',
                  opacity: loading ? 0.6 : 1,
                  fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: 6,
                }}
                title="Automatically tag from folder path & AI, and open interactive review at the same time"
              >
                ⚡🔍 Auto-Tag & Review{selected.size > 0 ? ` (${selected.size})` : ''}
              </button>
              <button
                disabled={selected.size === 0 || loading}
                onClick={() => void scan()}
                style={{
                  padding: '9px 13px', background: '#0f766e', color: '#fff', border: 0, borderRadius: 7,
                  cursor: selected.size === 0 || loading ? 'not-allowed' : 'pointer',
                  opacity: selected.size === 0 || loading ? 0.6 : 1,
                  fontWeight: 600,
                }}
              >
                {loading ? 'Scanning...' : `🔍 Scan & Review${selected.size > 0 ? ` (${selected.size})` : ''}`}
              </button>
              <button
                disabled={loading || !scanResults.some(result => result.status === 'needs_selection')}
                onClick={() => void confirmTags()}
                style={{
                  padding: '9px 13px', background: '#0369a1', color: '#fff', border: 0, borderRadius: 7,
                  cursor: scanResults.some(result => result.status === 'needs_selection') ? 'pointer' : 'not-allowed',
                  opacity: scanResults.some(result => result.status === 'needs_selection') ? 1 : 0.5,
                  fontWeight: 600,
                }}
              >
                ✓ Confirm Tags
              </button>
            </div>
          </div>

          {/* Folder and file count summary (matching Documents module with recursive accuracy) */}
          <div style={{ fontSize: 13, color: '#475569', marginBottom: 14, display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
            <span style={{ fontWeight: 700, color: '#0f172a' }}>
              📁 {summaryCounts ? summaryCounts.direct_folders : items.filter(i => i.folder).length} {(summaryCounts ? summaryCounts.direct_folders : items.filter(i => i.folder).length) === 1 ? 'folder' : 'folders'}
            </span>
            <span>•</span>
            <span style={{ fontWeight: 700, color: '#0f172a', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              📄 {summaryCounts ? `${summaryCounts.total_files} total files` : `${items.filter(i => !i.folder).length} files`}
              {summaryCounts && summaryCounts.direct_files === 0 && summaryCounts.direct_folders > 0 && (
                <span style={{ fontWeight: 400, color: '#64748b' }}>(0 at this level)</span>
              )}
              {summaryCounts && summaryCounts.direct_files > 0 && summaryCounts.direct_folders > 0 && (
                <span style={{ fontWeight: 400, color: '#64748b' }}>({summaryCounts.direct_files} at this level)</span>
              )}
              {countsLoading && (
                <span style={{ fontSize: 11, color: '#0284c7', fontWeight: 500 }}>
                  (calculating subfolders...)
                </span>
              )}
            </span>
          </div>

          {/* Quick Vessel & Tags Action Bar (Without Scanning) */}
          {items.length > 0 && (
            <div style={{
              display: 'flex', flexWrap: 'wrap', gap: 10, alignItems: 'center',
              padding: '10px 14px', background: '#f8fafc', border: '1px solid #cbd5e1',
              borderRadius: 8, marginBottom: 14,
            }}>
              <span style={{ fontSize: 13, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 6 }}>
                ⚓ Vessel:
              </span>

              <select
                value={bulkVessel}
                onChange={e => setBulkVessel(e.target.value)}
                style={{
                  padding: '6px 10px', fontSize: 13, borderRadius: 6,
                  border: '1px solid #94a3b8', background: '#fff', minWidth: 180, fontWeight: 500,
                }}
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
                style={{
                  padding: '6px 10px', fontSize: 13, borderRadius: 6,
                  border: '1px solid #94a3b8', width: 170,
                }}
              />

              <button
                type="button"
                disabled={!bulkVessel.trim() || loading}
                onClick={() => void applyBulkVessel()}
                style={{
                  padding: '7px 14px', background: '#0284c7', color: '#fff', border: 0,
                  borderRadius: 6, fontWeight: 600, fontSize: 13,
                  cursor: (!bulkVessel.trim() || loading) ? 'not-allowed' : 'pointer',
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                  opacity: (!bulkVessel.trim() || loading) ? 0.5 : 1,
                }}
                title="Save VesselName directly to SharePoint for selected items and sub-folders"
              >
                ⚓ Update Vessel {selected.size > 0 ? `(${selected.size} selected)` : `(all ${summaryCounts ? summaryCounts.total_files : items.length} files)`}
              </button>

              <button
                type="button"
                disabled={loading}
                onClick={() => void autoTagFromPath()}
                style={{
                  padding: '7px 12px', background: '#f0fdf4', color: '#15803d', border: '1px solid #bbf7d0',
                  borderRadius: 6, fontWeight: 600, fontSize: 12, cursor: 'pointer',
                  display: 'inline-flex', alignItems: 'center', gap: 5,
                }}
                title="Automatically tag all files from folder path and taxonomy across selected items and sub-folders"
              >
                ⚡ Auto-Tag from Folder Path {detectedVessel ? `(${detectedVessel})` : ''}
              </button>

              <span style={{ marginLeft: 'auto', fontSize: 12 }}>
                {items.filter(i => !i.folder && !i.tags?.vessel).length > 0 ? (
                  <span style={{ color: '#d97706', fontWeight: 600, background: '#fef3c7', padding: '4px 9px', borderRadius: 6 }}>
                    ⚠️ {items.filter(i => !i.folder && !i.tags?.vessel).length} file{items.filter(i => !i.folder && !i.tags?.vessel).length > 1 ? 's' : ''} missing vessel tag
                  </span>
                ) : (
                  <span style={{ color: '#16a34a', fontWeight: 600, background: '#dcfce7', padding: '4px 9px', borderRadius: 6 }}>
                    ✓ Ready to auto-tag
                  </span>
                )}
              </span>
            </div>
          )}

          {/* Item list table */}
          <div style={{ background: '#fff', border: '1px solid #dbe5ec', borderRadius: 10, overflow: 'hidden' }}>
            {/* Column header */}
            <div style={{
              display: 'grid', gridTemplateColumns: '30px 30px minmax(180px,1fr) minmax(200px,2fr) 140px',
              gap: 12, alignItems: 'center', padding: '10px 16px',
              background: '#f1f5f9', borderBottom: '1px solid #e2e8f0',
              fontSize: 12, fontWeight: 700, color: '#475569',
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
                    <span style={{ fontSize: 11, background: '#e0f2fe', color: '#0369a1', padding: '1px 8px', borderRadius: 10, fontWeight: 600 }}>
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
                  gap: 12, alignItems: 'center', padding: '12px 16px', borderBottom: '1px solid #eef2f6',
                  background: editing === item.id ? '#f0f9ff' : undefined,
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
                    width: 26, height: 26, padding: 0, border: '1px solid #bfdbfe', background: '#eff6ff',
                    color: '#1d4ed8', borderRadius: 6, cursor: 'pointer', fontSize: 13, fontWeight: 700,
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
                  style={{ textAlign: 'left', border: 0, background: 'transparent', color: item.folder ? '#0369a1' : '#123044', cursor: 'pointer', fontWeight: 600, fontSize: 13, display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}
                >
                  <span>{item.folder ? '📁' : '📄'} {item.name}</span>
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
                              background: totalFiles > 0 ? '#0284c7' : '#94a3b8',
                              color: '#fff',
                              borderRadius: 20,
                              padding: '2px 10px',
                              fontSize: 11,
                              fontWeight: 700,
                              lineHeight: '16px',
                              whiteSpace: 'nowrap',
                            }}
                            title={`${fc.direct_subfolders} direct subfolder${fc.direct_subfolders === 1 ? '' : 's'}, ${totalFiles} total file${totalFiles === 1 ? '' : 's'} across all subfolders`}
                          >
                            📁 {fc.direct_subfolders} {fc.direct_subfolders === 1 ? 'subfolder' : 'subfolders'} · {totalFiles} {totalFiles === 1 ? 'file' : 'files'}
                          </span>
                        );
                      } else {
                        return (
                          <span
                            style={{
                              display: 'inline-flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              background: totalFiles > 0 ? '#0284c7' : '#94a3b8',
                              color: '#fff',
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
                          background: countsLoading ? '#e2e8f0' : (fallbackCount > 0 ? '#0284c7' : '#94a3b8'),
                          color: countsLoading ? '#64748b' : '#fff',
                          borderRadius: 20,
                          padding: '1px 8px',
                          fontSize: 10,
                          fontWeight: 700,
                          lineHeight: '16px',
                          whiteSpace: 'nowrap',
                        }}
                        title={countsLoading ? 'Calculating recursive totals...' : `${fallbackCount} direct items`}
                      >
                        {countsLoading ? '⏳ ...' : `${fallbackCount} ${fallbackCount === 1 ? 'file' : 'files'}`}
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
                            style={{ width: 115, padding: 5, border: '1px solid #cbd5e1', borderRadius: 4, fontSize: 12 }}
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
                          style={{ width: 105, padding: 5, border: '1px solid #cbd5e1', borderRadius: 4, fontSize: 12 }}
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
                        style={{ border: 0, background: '#0369a1', color: '#fff', borderRadius: 6, padding: '5px 9px', cursor: 'pointer', fontSize: 12, fontWeight: 600, whiteSpace: 'nowrap' }}
                      >
                        Save
                      </button>
                      <button
                        type="button"
                        onClick={() => setEditing(null)}
                        style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '5px 8px', cursor: 'pointer', fontSize: 12 }}
                      >
                        Cancel
                      </button>
                    </>
                  ) : (
                    <>
                      {item.folder ? (
                        <button
                          type="button"
                          onClick={() => void autoTagFromPath([item.id])}
                          title={`Auto-tag all files inside ${item.name} from folder path`}
                          style={{
                            border: '1px solid #7dd3fc',
                            background: '#f0f9ff',
                            color: '#0284c7',
                            borderRadius: 6,
                            padding: '5px 9px',
                            cursor: 'pointer',
                            fontSize: 12,
                            fontWeight: 600,
                            whiteSpace: 'nowrap',
                          }}
                        >
                          ⚡ Auto-Tag Files
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
                          style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '5px 8px', cursor: 'pointer', fontSize: 12, whiteSpace: 'nowrap' }}
                        >
                          ✏ Tags
                        </button>
                      )}
                      {/* Open in SharePoint arrow — exact same button as Documents module */}
                      <button
                        type="button"
                        onClick={() => openSharePointLocation(item)}
                        title={item.folder ? `Open ${item.name} folder in SharePoint` : `Open ${item.name} in SharePoint`}
                        aria-label={`Open ${item.name} in SharePoint`}
                        style={{
                          width: 28, height: 27, padding: 0, borderRadius: 6, border: '1px solid #bfdbfe',
                          background: '#eff6ff', color: '#1d4ed8', cursor: 'pointer', fontSize: 16,
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
              <div style={{ padding: 24, color: '#64748b', textAlign: 'center' }}>This folder is empty.</div>
            )}
            {loading && (
              <div style={{ padding: 24, color: '#64748b', textAlign: 'center' }}>Loading items...</div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
