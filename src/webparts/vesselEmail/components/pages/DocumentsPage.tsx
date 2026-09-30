import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import {
  badge, GROUP_COLORS, DATASOURCE_TAGS_MAP, VESSEL_TYPES, cleanName, suggestTagFromFilename,
  INITIAL_MOCK_DOCUMENTS, INITIAL_MOCK_TEMPLATES, INITIAL_MOCK_APPROVALS,
  INITIAL_MOCK_NOTIFICATIONS, INITIAL_MOCK_USERS, resolveDetectedVesselForFile,
} from '../constants';
import type {
  FlatRow, GroupedRow, VesselRecord,
} from '../types/rows';
import type { BentoEmailLog } from '../types/bento';
import type { AppView, ModalMode } from '../types/view';
import type {
  FormState, DocPreviewItem, DeletedNode, DocumentItem, TemplateItem,
  ApprovalItem, NotificationItem, UserItem,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';
import {
  folderNamesByMainFolder, subfolderNamesByFolder, getAllKnownNestedTemplateSubfolders,
  getCommonShipsFlatRows, getKaizenFlatRows, getVesselTemplateFlatRows,
} from '../vesselFolderTemplate';
import { extractFilesFromDataTransfer } from '../BulkUploadModal';
import { CopilotSearchPanel } from '../copilot/CopilotSearchPanel';
import { FolderTreeSelect, FolderTreeNode } from './FolderTreeSelect';
import { DebouncedSearchInput } from './DebouncedSearchInput';

import { clay } from '../clayTheme';
// ── List view table styles ────────────────────────────────────────────────
// Every cell is top-aligned with the same padding, and each file line in the
// File name / Size / Date columns has the same fixed height, so a file's size
// and upload time always sit on the same line as its name.
const LIST_TH: React.CSSProperties = {
  position: 'sticky', top: 0, zIndex: 1,
  padding: '11px 12px', background: 'var(--vdms-surface-alt)',
  borderBottom: '1px solid var(--vdms-border)',
  color: 'var(--vdms-text-muted)', fontSize: 10.5, fontWeight: 700,
  letterSpacing: '0.06em', textTransform: 'uppercase', textAlign: 'left', whiteSpace: 'nowrap',
};
const LIST_TD: React.CSSProperties = {
  padding: '10px 12px', verticalAlign: 'top', fontSize: 12, lineHeight: '18px', overflow: 'hidden',
};
const LIST_FILE_LINE: React.CSSProperties = {
  display: 'flex', alignItems: 'center', height: 26, minWidth: 0, whiteSpace: 'nowrap',
};
const LIST_ELLIPSIS: React.CSSProperties = {
  display: 'block', minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
};
// Shows the value in full (wraps instead of truncating with "…") — used for
// tag-derived columns (Vessel, Domain, Main folder, Group, Category, Sub
// category) where a cut-off label like "Communicati…" is unreadable/ambiguous.
const LIST_WRAP: React.CSSProperties = {
  display: 'block', minWidth: 0, whiteSpace: 'normal', wordBreak: 'break-word', overflowWrap: 'anywhere',
};
const LIST_PILL: React.CSSProperties = {
  display: 'inline-flex', alignItems: 'center', gap: 5, maxWidth: '100%',
  borderRadius: 999, padding: '2px 9px', fontSize: 11, fontWeight: 600, lineHeight: '18px', whiteSpace: 'nowrap',
};
const LIST_X_DOT: React.CSSProperties = {
  display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
  width: 14, height: 14, borderRadius: '50%', background: '#c2410c', color: '#fff',
  fontSize: 9, fontWeight: 800, lineHeight: 1, flexShrink: 0,
};
const LIST_ACTION_BTN: React.CSSProperties = {
  height: 26, padding: '0 8px', borderRadius: 6, border: '1px solid var(--vdms-border)',
  fontSize: 11, fontWeight: 600, whiteSpace: 'nowrap', boxSizing: 'border-box',
  display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 4,
};

/** Folder view Group/Category search: one backend recursive walk per
 * `${siteId}::${driveId}::${folderId}`, kept for GROUP_CAT_TREE_TTL_MS so
 * re-renders and filter changes don't refetch it. */
interface GroupCatTreeEntry {
  status: 'loading' | 'done' | 'error';
  items: any[];
  truncated: boolean;
  at: number;
  /** First attempt of this search (kept across polls of a still-building index). */
  since?: number;
  /** Waiting for the backend's folder/file index to finish its first build: re-ask at `retryAt`. */
  idle?: boolean;
  retryAt?: number;
  /** Consecutive failed lookups (drives the retry pause). */
  failures?: number;
}
/** Pause before retrying a failed lookup: 4 s, 8 s, 16 s, then 30 s. */
const groupCatRetryDelayMs = (failures: number): number => Math.min(30000, 4000 * Math.pow(2, Math.max(0, failures - 1)));
const GROUP_CAT_TREE_TTL_MS = 5 * 60 * 1000;
/** How long to keep polling the backend's index before falling back to the slow Graph walk. */
const GROUP_CAT_INDEX_WAIT_MS = 150 * 1000;
/** Words a folder name must contain for the backend to pre-select files of a Group (superset of the exact client rule). */
function groupCatPathWords(group: string): string[] {
  const g = (group || '').trim().toLowerCase();
  if (!g || g === 'all') return [];
  if (g === 'drawings') return ['drawing', 'drawings', 'dwg', 'dwgs'];
  if (g === 'manuals') return ['manual', 'manuals'];
  if (g === 'to be classified') return ['classified'];
  return ['drawing', 'drawings', 'dwg', 'dwgs', 'manual', 'manuals'];
}
const groupCatTreeCache: Map<string, GroupCatTreeEntry> = new Map();

/** ✕ marker for a List view cell whose value isn't present in the path. */
function ListMissingMark(props: { title: string }): React.ReactElement {
  return (
    <span
      title={props.title}
      aria-label={props.title}
      style={{
        display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
        width: 20, height: 20, borderRadius: 6,
        background: '#fef2f2', color: '#dc2626', border: '1px solid #fecaca',
        fontSize: 11, fontWeight: 800, lineHeight: 1,
      }}
    >
      ✕
    </span>
  );
}

export function renderDocumentsPage(host: VesselEmail): React.ReactElement {

    const {
      textFilter, vesselFilter, catFilter, attachmentFilter, docViewMode, showAllVesselsInFolderView,
      vessels, allVesselNamesForSearch, rows, docListPage, docListSort, docGroupFilter, docCategoryFilter, docGroupLevelFilter, docLeafCategoryFilter, docSubCategoryFilter,
      docSubfolderOtherFilter,
      docCompareMode, compareBoxSubPaths,
      docScopeType,
      documentVesselCount, documentVesselsLoadingMore, documentFilesLoading, vesselLoadingName,
      docUploadRowKey, docUploadBusy, docUploadMsg, documentsList,
      folderPathStack, uploadedFilesByFolder, docMainFolder,
      folderNavHistory, folderNavIndex, listViewSelectedFiles, folderViewSelectedFiles,
      windowWidth,
    } = host.state;

    const vesselColumns = windowWidth <= 767 ? 1 : windowWidth <= 1024 ? 2 : 4;

    // Shared tokenized search for the Documents module's search bar. Splits
    // the query on whitespace so a combined query — e.g. a vessel name plus
    // a (partial) file name typed together — still matches even though the
    // two pieces live in different fields on the same row/item: every
    // whitespace-separated token just has to appear *somewhere* across the
    // given fields (any field, any file format/extension), and different
    // tokens are free to match different fields. A single-clause query
    // behaves exactly like a plain "contains" search as before.
    // Strips combining diacritical marks after Unicode NFD decomposition
    // (e.g. "é" -> "e" + U+0301, then the mark is dropped), so an accented
    // vessel/file name matches an unaccented search term and vice versa.
    // Falls back to the plain string if normalize() isn't available.
    const foldDiacritics = (s: string): string => {
      try {
        return s.normalize('NFD').replace(/[̀-ͯ]/g, '');
      } catch {
        return s;
      }
    };

    // +, comma or semicolon separates OR clauses — mirrors the backend's
    // /api/dashboard/documents matcher (_dashboard_doc_matches_text in
    // main.py), so "Bow Fighter + Bow Fraternity" finds files under EITHER
    // vessel in one search instead of finding nothing (no single row ever
    // has both vessel names on it, so a plain AND-of-all-words search always
    // came back empty for a multi-vessel query). A clause that names a known
    // vessel only needs ONE such clause to match the row; any other
    // (non-vessel) clause still narrows the result further, e.g.
    // "Bow Fighter + Bow Fraternity safety" only returns Safety files under
    // either vessel. A query with no +/,/; at all behaves exactly as before.
    // `vessels` is the Vessels-module table — scoped to whichever site
    // vesselSiteFilter currently points at, and deliberately NOT refetched
    // when switching the "SharePoint site" dropdown here in Documents (see
    // _switchDocumentSite's comment: reloading it on every site switch would
    // leak other-site vessel data into the Vessels table). That means a
    // clause naming a real vessel under the site currently being browsed can
    // still be missing from `vessels` if it lives under a different site
    // than vesselSiteFilter — allVesselNamesForSearch (tenant-wide, fetched
    // once at mount, never used for anything but this) fills that gap so
    // e.g. "Bow Fighter + plan" is still recognized as a vessel clause even
    // when the Vessels table hasn't loaded Bow Fighter's site.
    const knownVesselNamesLower = Array.from(new Set([
      ...(vessels || []).map(v => (v.name || '').trim().toLowerCase()),
      ...(allVesselNamesForSearch || []).map(n => n.trim().toLowerCase()),
    ])).filter(Boolean);

    const matchesSearchTokens = (query: string, ...fields: Array<string | undefined | null>): boolean => {
      const trimmedQuery = foldDiacritics((query || '').trim().toLowerCase());
      if (!trimmedQuery) return true;
      const haystack = foldDiacritics(fields.filter(Boolean).join(' \u0001 ').toLowerCase());
      const andOfWords = (clause: string): boolean => {
        const tokens = clause.split(/\s+/).filter(Boolean);
        return tokens.length > 0 && tokens.every(token => haystack.includes(token));
      };
      const rawClauses = trimmedQuery.split(/[+,;]+/).map(c => c.trim()).filter(Boolean);
      if (rawClauses.length <= 1) {
        const tokens = trimmedQuery.split(/\s+/).filter(Boolean);
        return tokens.every(token => haystack.includes(token));
      }
      const vesselClauses = rawClauses.filter(c =>
        knownVesselNamesLower.some(name => name === c || name.includes(c) || c.includes(name))
      );
      if (vesselClauses.length > 0) {
        if (!vesselClauses.some(andOfWords)) return false;
        const otherClauses = rawClauses.filter(c => !vesselClauses.includes(c));
        return otherClauses.every(andOfWords);
      }
      // No clause recognized as a known vessel name — e.g. "Bow Fighter +
      // Bow Fraternity" when the vessel list hasn't loaded yet, or "Bow
      // Fighter + monthly report.pdf" mixing a vessel with a filename.
      // +/,/; still means OR here: requiring every clause to match (AND)
      // meant a query naming two different things never matched any single
      // row, which is exactly the "no results" bug this was meant to fix.
      // Each clause is still an AND of its own words.
      return rawClauses.some(andOfWords);
    };

    const getFirstClassSite = (keyHint: string, fallbackIdx: number) => {
      const sites = host.state.documentSites || [];
      const match = sites.find(s =>
        s.site_key === keyHint ||
        (s.sp_site_name && s.sp_site_name.toLowerCase().includes(keyHint.toLowerCase())) ||
        (s.default_library_name && s.default_library_name.toLowerCase().includes(keyHint.toLowerCase()))
      );
      if (match) return match;
      if (sites.length > fallbackIdx) return sites[fallbackIdx];
      return sites[0] || null;
    };

    const triggerFolderRefresh = (
      targetStack: Array<{ id: string; name: string }>,
      targetScope: 'vessels' | 'common' | 'kaizen' | 'sites' | 'shared_docs' | 'documents' = host.state.docScopeType,
      targetVessel: string = host.state.vesselFilter,
      rowGroupKey?: string,
    ): void => {
      if (!targetStack || targetStack.length === 0) return;
      const leafNode = targetStack[targetStack.length - 1];
      const leafId = leafNode?.id || '';

      if (targetScope === 'sites') {
        const siteNode = targetStack[1];
        const rawSiteId = (siteNode?.id || '').replace(/^site:/, '');
        const matchedSite = (host.state.documentSites || []).find(s =>
          s.site_id === rawSiteId || s.site_key === rawSiteId || s.sp_site_name === siteNode?.name
        );
        const effectiveSiteId = matchedSite?.site_id || rawSiteId;
        const driveNode = targetStack[2];
        const rawDriveId = (driveNode?.id || '').replace(/^drive:/, '') || matchedSite?.drive_id || '';

        if (effectiveSiteId && rawDriveId) {
          let folderRef = targetStack.length <= 3 ? 'root' : leafId;
          if (/^sf_/i.test(folderRef)) {
            const segments = targetStack.slice(3).map(n => n.name).filter(Boolean);
            folderRef = segments.join('/');
          }
          setTimeout(() => void host._refreshSiteFolder(effectiveSiteId, rawDriveId, folderRef).catch(() => undefined), 0);
        }

        // NOTE: this used to also detect whether the folder just opened
        // shares a name with a registered vessel and, if so, fire off
        // _loadVesselRowsFromApi + _loadFilesForVessel (a full recursive
        // Graph walk of that vessel's entire folder tree via
        // _mergeLiveSharePointFiles) in the background — purely as a
        // "warm the Vessels module's cache in case the user switches
        // there later" prefetch. It ran on every folder navigation in
        // Sites scope, competing for the same origin's connection pool
        // with the _refreshSiteFolder call right above it for the folder
        // the user is actually waiting on — i.e. opening a vessel-named
        // folder in Sites could genuinely slow down loading that folder's
        // own contents, and left documentFilesLoading/vesselLoadingName
        // set (see the toolbar banner above) for a vessel the user may
        // have already navigated away from by the time it resolved.
        // Removed: the Vessels module already loads a vessel's rows
        // on demand (componentDidUpdate's vesselFilter handler) exactly
        // when the user actually opens it there, so this prefetch traded
        // a real, guaranteed slowdown for a speculative later saving.
      } else if (targetScope === 'shared_docs' || targetScope === 'documents') {
        const firstClass = targetScope === 'shared_docs' ? getFirstClassSite('nksdocman', 0) : getFirstClassSite('dev', 1);
        if (firstClass?.site_id && firstClass?.drive_id) {
          const folderRef = targetStack.length <= 1 ? 'root' : leafId;
          setTimeout(() => void host._refreshSiteFolder(firstClass.site_id, firstClass.drive_id, folderRef).catch(() => undefined), 0);
        }
      } else if (targetScope === 'vessels') {
        const hasSyntheticId = !leafId || /^(sf_|category_|common|kaizen_root|dept_|sites_root|site:|drive:)/.test(leafId) || /^\d+$/.test(leafId);
        if (!hasSyntheticId && leafId) {
          setTimeout(() => void host._refreshFolderFiles(leafId, rowGroupKey || '', true, true).catch(() => undefined), 0);
        } else if (targetVessel && targetVessel !== 'all') {
          host._filesLoadedForVessels.delete(targetVessel);
          setTimeout(() => void host._loadFilesForVessel(targetVessel).catch(() => undefined), 0);
        }
      } else if (targetScope === 'common') {
        const hasSyntheticId = !leafId || /^(sf_|category_|common|kaizen_root|dept_|sites_root|site:|drive:)/.test(leafId);
        if (!hasSyntheticId && leafId) {
          setTimeout(() => void host._refreshFolderFiles(leafId, rowGroupKey || '', true, true).catch(() => undefined), 0);
        }
      } else if (targetScope === 'kaizen') {
        const hasSyntheticId = !leafId || /^(sf_|category_|common|kaizen_root|dept_|sites_root|site:|drive:)/.test(leafId);
        if (!hasSyntheticId && leafId) {
          setTimeout(() => void host._refreshFolderFiles(leafId, rowGroupKey || '', true, true).catch(() => undefined), 0);
        }
      }
    };

    // Filled in once deriveLiveNavFilterState is defined further down, so
    // Back/Forward (defined here) can re-sync the filter bar too.
    const liveNavSync: {
      derive?: (stack: { id: string; name: string }[]) => {
        docCategoryFilter: string; docSubfolderOtherFilter: string; vesselFilter: string; docListPage: number; docGroupLevelFilter: string;
      };
    } = {};

    const canGoBack = folderNavIndex > 0;
    const canGoForward = folderNavIndex < folderNavHistory.length - 1;

    const goBack = (): void => {
      if (!canGoBack) return;
      const prev = folderNavHistory[folderNavIndex - 1];
      host.setState({ folderNavIndex: folderNavIndex - 1, folderPathStack: prev.folderPathStack, docMainFolder: prev.docMainFolder });
      if (prev.folderPathStack[0]?.id === 'sites_root' && liveNavSync.derive) host.setState(liveNavSync.derive(prev.folderPathStack));
      triggerFolderRefresh(prev.folderPathStack);
    };

    const goForward = (): void => {
      if (!canGoForward) return;
      const next = folderNavHistory[folderNavIndex + 1];
      host.setState({ folderNavIndex: folderNavIndex + 1, folderPathStack: next.folderPathStack, docMainFolder: next.docMainFolder });
      if (next.folderPathStack[0]?.id === 'sites_root' && liveNavSync.derive) host.setState(liveNavSync.derive(next.folderPathStack));
      triggerFolderRefresh(next.folderPathStack);
    };

    const PAGE_ROWS = 20;

    // Main folder definitions
    type MainFolderKey =
      | 'Technical & Crewing'
      | 'Commercial & Chartering'
      | 'Insurance'
      | 'Kaizen - Knowledge Bank'
      | 'Knowledge Bank'
      | 'Shared Documents'
      | 'Documents'
      | 'SharePoint Sites';

    const VESSEL_MAIN_FOLDERS: Array<{ key: MainFolderKey; icon: string; emoji: string; color: string; bg: string }> = [];
    const MAIN_FOLDERS: Array<{ key: MainFolderKey; icon: string; emoji: string; color: string; bg: string }> = [];
    const activeLiveSite = (host.state.documentSites || []).find(site => site.site_key === host.state.activeDocumentSite);
    const isNksDocMan = /nksdocman/i.test(`${activeLiveSite?.site_key || ''} ${activeLiveSite?.sp_site_name || ''}`);

    // Resolve the site/drive actually being displayed right now. When browsing
    // via the "SharePoint Sites" breadcrumb tree (docScopeType 'sites', the
    // same folderPathStack[1]/[2] resolution the Level>=2/3 folder-view
    // renderer below uses), that's a different site than whatever
    // activeDocumentSite happens to be — that variable only tracks the
    // top "SharePoint site" quick-switcher, which this navigation mode
    // doesn't use. Falling back to activeLiveSite keeps the quick-switcher
    // scopes (Documents / Vessels / etc.) working as before.
    const isSitesScopeNav = docScopeType === 'sites' && folderPathStack.length >= 2 &&
      (folderPathStack[0]?.id === 'sites_root' || folderPathStack[0]?.name === 'SharePoint Sites' || folderPathStack[0]?.name === 'Sites Documents');
    const sitesScopeSiteNode = isSitesScopeNav ? folderPathStack[1] : null;
    const sitesScopeRawSiteId = (sitesScopeSiteNode?.id || '').replace(/^site:/, '');
    const sitesScopeMatchedSite = sitesScopeSiteNode
      ? (host.state.documentSites || []).find(s =>
          s.site_id === sitesScopeRawSiteId || s.site_key === sitesScopeRawSiteId || s.sp_site_name === sitesScopeSiteNode.name
        )
      : null;
    const sitesScopeSiteId = sitesScopeSiteNode ? (sitesScopeMatchedSite?.site_id || sitesScopeRawSiteId) : '';
    const sitesScopeDriveNode = isSitesScopeNav && folderPathStack.length >= 3 ? folderPathStack[2] : null;
    const sitesScopeDriveId = sitesScopeDriveNode
      ? (sitesScopeDriveNode.id || '').replace(/^drive:/, '')
      : (sitesScopeMatchedSite?.drive_id || '');

    const effectiveLiveSiteId = sitesScopeSiteId || activeLiveSite?.site_id || '';
    const effectiveLiveDriveId = sitesScopeDriveId || activeLiveSite?.drive_id || '';

    const liveLibraryResolved = !!(effectiveLiveSiteId && effectiveLiveDriveId) &&
      (docScopeType === 'sites' || docScopeType === 'shared_docs' || docScopeType === 'documents');
    // The library's folders 3 levels deep, from the backend's drive folder
    // index (one call, no per-folder listing). Used to find vessel folders
    // and as a second source for the top-level folders below.
    const shallowRootTree = liveLibraryResolved
      ? host._getOrLoadSiteFolderTree(effectiveLiveSiteId, effectiveLiveDriveId, 'root', 3)
      : null;
    let directRootItems: any[] = effectiveLiveSiteId && effectiveLiveDriveId
      ? host._getOrLoadSiteFolderChildren(effectiveLiveSiteId, effectiveLiveDriveId, 'root').items
      : [];
    if (liveLibraryResolved) {
      // The root listing can fail or still be loading (Graph throttling, a
      // timed-out request) while the folder index already knows the top-level
      // folders. Without any source the Main folder dropdown collapsed to a
      // made-up "Technical & Crewing" and a vessel could not be opened until
      // the user pressed Retry. Order: live listing → folder index → last
      // good list for this library.
      const rootMemo: Map<string, any[]> = ((host as any)._rootFolderItemsMemo ||= new Map<string, any[]>());
      const rootMemoKey = `${effectiveLiveSiteId}::${effectiveLiveDriveId}`;
      if (!directRootItems.some(i => !!i?.folder)) {
        const top = (shallowRootTree?.folders || []).filter(f => f.depth === 1);
        if (top.length > 0) directRootItems = top.map(f => ({ id: f.id, name: f.name, folder: {} }));
      }
      if (directRootItems.some(i => !!i?.folder)) rootMemo.set(rootMemoKey, directRootItems);
      else if (rootMemo.has(rootMemoKey)) directRootItems = rootMemo.get(rootMemoKey)!;
    }
    const liveRootSource = effectiveLiveSiteId && effectiveLiveDriveId
      ? directRootItems
      : (host.state.documentLiveFolders || []).filter(folder => folder.is_folder !== false && folder.depth === 0);
    const liveRootFolders = liveRootSource
      // `folder.folder` (the Graph driveItem folder facet) is the only
      // reliable positive folder signal across both sources this list can
      // come from: it's present on a raw Graph children response only for
      // real folders, and it's synthesized from documentLiveFolders'
      // is_folder flag the same way (see _getOrLoadSiteFolderChildren).
      // `is_folder !== false` used to be OR'd in here as a fallback, but
      // that's `undefined !== false` → true for every raw Graph *file* too
      // (that shape has no is_folder field at all), which is what let files
      // leak into this "folders only" list.
      .filter(folder => !!folder.folder &&
        !['shared documents', 'documents'].includes((folder.name || '').trim().toLowerCase()))
      .map(folder => ({
        key: folder.name,
        icon: '📁',
        emoji: '📁',
        color: '#0f766e',
        bg: '#ccfbf1',
        liveFolderId: folder.id,
      }));
    // The "SharePoint Sites" tile used to be pinned here as a synthetic
    // level-0 card, but it duplicated the "SharePoint site" dropdown above
    // the grid: clicking it just re-opened a picker for the same sites the
    // dropdown already switches between, which read as redundant/confusing.
    // Site switching now happens only through that dropdown, so this list is
    // just the real top-level folders of the selected site's library.
    const rootFolderCards = [
      ...liveRootFolders,
    ];

    // ── Live Main folder / Sub-folder classification ──────────────────────────
    // "Main folder" options must be the real, current top-level folders of the
    // selected site's Documents library (not a hard-coded list), and any real
    // folder that sits alongside vessel folders but isn't itself a vessel
    // (e.g. "Report", "ACRA CHARGE", "Share with Mr Akase") must still be
    // reachable as its own filter option instead of disappearing. Vessel
    // identity is authoritative from the SharePoint Term Store (merged with
    // DB vessels server-side), not path-guessing alone.
    const liveTermStoreVesselNames: string[] = effectiveLiveSiteId
      ? host._getOrLoadTermStoreVessels(effectiveLiveSiteId)
      : [];
    const liveTermStoreVesselSet = new Set(
      liveTermStoreVesselNames.concat((vessels || []).map(v => v.name)).map(n => (n || '').trim().toLowerCase()).filter(Boolean)
    );
    const KNOWN_MAIN_FOLDER_MAP = folderNamesByMainFolder();
    const KNOWN_MAIN_FOLDER_NAME_SET = new Set(Object.keys(KNOWN_MAIN_FOLDER_MAP).map(n => n.trim().toLowerCase()));
    const KNOWN_MAIN_FOLDER_COMMON_MAP = folderNamesByMainFolder(true);
    const KNOWN_CATEGORY_NAME_SET = new Set(
      Object.values(KNOWN_MAIN_FOLDER_MAP).concat(Object.values(KNOWN_MAIN_FOLDER_COMMON_MAP))
        .reduce((all, names) => all.concat(names), [] as string[])
        .map(n => n.trim().toLowerCase())
    );
    // "Main folder" = every real top-level folder in the site's Documents
    // library, live — known department (Technical & Crewing, ...), a vessel
    // folder sitting flat at root (e.g. "mvtest2209"), or anything else
    // (e.g. "Report", "ACRA CHARGE") — all of it, not just the known set.
    // A top-level folder that is also a recognised vessel is cross-listed in
    // the vessel filter too (see siteVesselNames below), it isn't removed
    // from here.
    const liveMainFolderNames = Array.from(new Set(
      liveRootFolders.map(f => f.key).filter(Boolean)
    ));

    // Hierarchy navigation state
    const stackLevel = folderPathStack.length;
    const atKaizenRoot = stackLevel >= 1 && (folderPathStack[0]?.id === 'kaizen_root' || folderPathStack[0]?.name === 'Kaizen - Knowledge Bank');
    const atSharedDocsRoot = stackLevel >= 1 && (folderPathStack[0]?.id === 'lib:shared_documents' || folderPathStack[0]?.name === 'Shared Documents');
    const atDocsRoot = stackLevel >= 1 && (folderPathStack[0]?.id === 'lib:documents' || folderPathStack[0]?.name === 'Documents');
    const atSitesRoot = stackLevel >= 1 && !atSharedDocsRoot && !atDocsRoot && (
      folderPathStack[0]?.id === 'sites_root' ||
      folderPathStack[0]?.name === 'SharePoint Sites' ||
      folderPathStack[0]?.name === 'Sites Documents'
    );
    const atMainDepartment = stackLevel >= 1 && !atKaizenRoot && !atSharedDocsRoot && !atDocsRoot && !atSitesRoot && MAIN_FOLDERS.some(mf => mf.key === folderPathStack[0]?.name);
    const atDepartmentVesselList = atMainDepartment && stackLevel === 1;
    const showSelectedVesselCategories = atDepartmentVesselList && host.state.docScopeType === 'vessels' && vesselFilter !== 'all';
    const atCommonShips = atMainDepartment && stackLevel >= 2 && (
      folderPathStack[1]?.id === 'common' ||
      folderPathStack[1]?.name === 'Common for all ships' ||
      folderPathStack[1]?.name === 'Common for all vessels' ||
      folderPathStack[1]?.name === 'Common (Not Ship Specific)' ||
      folderPathStack[1]?.name === 'Common Agreements (Not Ship Specific)'
    );
    const vesselNodeInStack = (atMainDepartment && stackLevel >= 2 && !atCommonShips) ? folderPathStack[1] : null;
    const vesselStackIdx = vesselNodeInStack
      ? folderPathStack.findIndex(node => node.id === vesselNodeInStack.id)
      : -1;
    let sitesVesselName: string | null = null;
    if (atSitesRoot && stackLevel > 3) {
      const segs = folderPathStack.slice(3).map(n => n.name);
      const matched = (vessels || []).find(v => segs.some(s => s.toLowerCase() === v.name.toLowerCase()));
      sitesVesselName = matched ? matched.name : (segs[1] || segs[0] || null);
    }
    const currentVesselNameFromStack = atKaizenRoot
      ? 'Kaizen - Knowledge Bank'
      : (atSharedDocsRoot
        ? 'Shared Documents'
        : (atDocsRoot
          ? 'Documents'
          : (atSitesRoot ? (sitesVesselName || 'SharePoint Sites') : (atCommonShips ? 'Common for all vessels' : (vesselNodeInStack?.name || null)))));

    // Auto-repair: if user is on a department breadcrumb but a vessel filter is active,
    // push the missing vessel node so breadcrumb and folder context stay aligned.
    if (docViewMode === 'folder' && showSelectedVesselCategories && docMainFolder && atDepartmentVesselList) {
      const selectedVesselName = (vesselFilter || '').trim();
      if (selectedVesselName) {
        const match = (vessels || []).find(v => (v.name || '').trim().toLowerCase() === selectedVesselName.toLowerCase());
        const vesselId = String(match?.id || selectedVesselName);
        setTimeout(() => {
          host._pushFolderNav([
            { id: docMainFolder, name: docMainFolder },
            { id: vesselId, name: selectedVesselName },
          ], docMainFolder);
        }, 0);
      }
    }

    if (atSitesRoot && (!host.state.documentSites || host.state.documentSites.length === 0)) {
      setTimeout(() => {
        void host._loadDocumentSites().catch(() => undefined);
      }, 0);
    }
    if (docViewMode === 'folder' && atSitesRoot && stackLevel >= 3) {
      const siteDriveNode = folderPathStack[2];
      const activeSiteNode = folderPathStack[1];
      const currentSiteDoc = (host.state.documentSites || []).find(s =>
        s.site_key === activeSiteNode?.id ||
        s.site_id === activeSiteNode?.id ||
        (s.sp_site_name && activeSiteNode?.name && s.sp_site_name.toLowerCase() === activeSiteNode.name.toLowerCase())
      ) || host.state.documentSites[0];
      const activeDriveId = (siteDriveNode?.id || '').replace(/^drive:/, '') || currentSiteDoc?.drive_id || '';
      const activeCurrentFolderNode = folderPathStack[folderPathStack.length - 1];
      const activeFolderId = stackLevel === 3 ? 'root' : (activeCurrentFolderNode?.id || 'root');
      const activeSiteId = currentSiteDoc?.site_id || '';

      if (activeSiteId && activeDriveId && activeFolderId) {
        // Load only the folder currently being viewed. Prefetching every child
        // folder here causes a render/forceUpdate request fan-out.
        host._getOrLoadSiteFolderChildren(activeSiteId, activeDriveId, activeFolderId);
      }
    }

    if (currentVesselNameFromStack && currentVesselNameFromStack !== 'SharePoint Sites' && !host._filesLoadedForVessels.has(currentVesselNameFromStack)) {
      host._filesLoadedForVessels.add(currentVesselNameFromStack);
      setTimeout(() => {
        if (atKaizenRoot) {
          void host._mergeLiveSharePointFiles(['Kaizen - Knowledge Bank']).catch(() => undefined);
        } else if (atCommonShips) {
          void host._mergeLiveSharePointFiles(['Common for all vessels']).catch(() => undefined);
        } else {
          void host._loadFilesForVessel(currentVesselNameFromStack!).catch(() => undefined);
        }
      }, 0);
    }

    const currentFolderNode = folderPathStack.length > 0 ? folderPathStack[folderPathStack.length - 1] : null;
    const currentFolderName = currentFolderNode ? currentFolderNode.name : null;
    const mainFolderPage = atMainDepartment && stackLevel === 1;

    const DEFAULT_VESSEL_MAINS = folderNamesByMainFolder();
    const SUBFOLDERS_MAP = subfolderNamesByFolder();
    const COMMON_DEFAULT_MAINS = folderNamesByMainFolder(true);

    // Build full breadcrumb path from navigation stack
    const fullStackBreadcrumb = folderPathStack
      .map(n => n.name)
      .filter(n => n !== 'Vessels' && n !== 'Specific Vessels')
      .join(' > ');

    const vesselName = currentVesselNameFromStack || (folderPathStack.length > 0 ? folderPathStack[0].name : null);
    const folderName = currentFolderNode ? currentFolderNode.name : null;

    // Resolve real SPO folder ID for the current folder node from live map and rows
    const resolvedCurrentFolderId: string | null = (() => {
      if (!currentFolderNode) return null;
      if (!/^(sf_|category_|common|kaizen_root|dept_)/.test(currentFolderNode.id)) return currentFolderNode.id;

      // 1. Try full stack breadcrumb (e.g. "Technical & Crewing > newvessel37 > Drawings and Manuals > Drawing > Basic")
      if (fullStackBreadcrumb) {
        const liveId = host._getLiveSharePointFolderId(fullStackBreadcrumb);
        if (liveId) return liveId;
      }

      // 2. Try vessel-first and dept-first breadcrumb permutations
      if (vesselName && docMainFolder && folderName) {
        const afterVessel = folderPathStack.slice(2).map(n => n.name);
        const alt1 = [vesselName, docMainFolder, ...afterVessel].join(' > ');
        const liveId1 = host._getLiveSharePointFolderId(alt1);
        if (liveId1) return liveId1;

        const alt2 = [docMainFolder, vesselName, ...afterVessel].join(' > ');
        const liveId2 = host._getLiveSharePointFolderId(alt2);
        if (liveId2) return liveId2;
      }

      // 3. Scan spoFolderMap directly for a matching node
      if (vesselName && folderName) {
        const vClean = vesselName.toLowerCase().replace(/^(mv|m\/v|mt)\s+/i, '').replace(/[^\w\d]/g, '');
        const fClean = folderName.toLowerCase().replace(/[^\w\d]/g, '');
        for (const [id, node] of Array.from(host.state.spoFolderMap.entries())) {
          if (!node || !node.isFolder || node.deleted) continue;
          const sPath = (node.serverRelativePath || '').toLowerCase().replace(/[^\w\d]/g, '');
          const sName = (node.name || '').toLowerCase().replace(/[^\w\d]/g, '');
          if (sPath.includes(vClean) && (sName === fClean || sPath.endsWith(fClean))) {
            return id;
          }
        }
      }

      // 4. Match row
      const parentInStack = folderPathStack.length >= 2 ? folderPathStack[folderPathStack.length - 2] : null;
      const matchRow = host.state.rows.find(r =>
        r.vesselName === vesselName &&
        r.uploadFolderId &&
        !r.uploadFolderId.includes('/') &&
        (r.subCategory === folderName || r.category === folderName) &&
        (docMainFolder ? r.group === docMainFolder : true) &&
        (parentInStack && parentInStack.name !== vesselName
          ? (r.category === parentInStack.name || r.subCategory === parentInStack.name)
          : true)
      );
      if (matchRow?.uploadFolderId && !/^\d+$/.test(matchRow.uploadFolderId)) return matchRow.uploadFolderId;
      return null;
    })();

    const mainsDefaultsSource = atCommonShips ? COMMON_DEFAULT_MAINS : (atKaizenRoot ? folderNamesByMainFolder() : DEFAULT_VESSEL_MAINS);
    const scopedKey = docMainFolder && currentFolderName ? `${docMainFolder} > ${currentFolderName}` : null;
    const isAtCategoryLevel =
      (vesselNodeInStack && stackLevel === 2);

    let subfolderNames: string[] = [];
    if (scopedKey && SUBFOLDERS_MAP[scopedKey]) {
      subfolderNames = SUBFOLDERS_MAP[scopedKey];
    } else if (docMainFolder !== 'Insurance' && currentFolderName && SUBFOLDERS_MAP[currentFolderName]) {
      subfolderNames = SUBFOLDERS_MAP[currentFolderName];
    } else if (docMainFolder && mainsDefaultsSource[docMainFolder]) {
      if (isAtCategoryLevel || showSelectedVesselCategories) {
        subfolderNames = mainsDefaultsSource[docMainFolder];
      }
    }
    const isDisplayableFolderName = (name: string): boolean => {
      const trimmed = (name || '').trim();
      if (!trimmed) return false;
      if (trimmed.includes('||') || trimmed.includes('>') || trimmed.includes('/')) return false;
      if (/^01[A-Za-z0-9]{15,}$/.test(trimmed) || /^[A-Fa-f0-9]{24,}$/.test(trimmed)) return false;

      const lower = trimmed.toLowerCase();
      const forbidden = new Set([
        'vessels', 'specific vessels', 'documents', 'shared documents', 'root',
        'common for all ships', 'common for all vessels', 'common',
        'common (not ship specific)', 'common agreements (not ship specific)',
        'kaizen', 'kaizen - knowledge bank', 'knowledge bank',
        'technical & crewing', 'commercial & chartering', 'insurance',
        'empty list', 'null', 'undefined', 'nan'
      ]);
      if (forbidden.has(lower)) return false;
      return true;
    };

    subfolderNames = subfolderNames.filter(isDisplayableFolderName);
    const templateSubfolderSet = new Set(subfolderNames.map(name => name.trim().toLowerCase()));
    const allNestedTemplateSet = getAllKnownNestedTemplateSubfolders(docMainFolder || undefined);

    // Include dynamically uploaded subfolders so folder uploads are visible as
    // navigable folders in Folder view even when they are not in the template.
    const dynamicSubfolderSet = new Set<string>();

    const tryAddDynamicSubfolder = (childName: string): void => {
      const trimmed = (childName || '').trim();
      if (!trimmed) return;

      // 1. Never add groupKeys or strings containing '||'
      if (trimmed.includes('||')) return;

      // 2. Never add SharePoint Graph Item IDs (e.g. 014Z... or alphanumeric hashes)
      if (/^01[A-Za-z0-9]{15,}$/.test(trimmed) || /^[A-Fa-f0-9]{24,}$/.test(trimmed)) return;

      const lower = trimmed.toLowerCase();

      // 3. Never add main folder names or container names
      if (MAIN_FOLDERS.some(mf => mf.key.toLowerCase() === lower)) return;
      if (!isDisplayableFolderName(trimmed)) return;

      // 4. If we are inside a vessel (stackLevel >= 2) or at vessel level, never add any vessel's name as a subfolder
      if (vessels.some(v => (v.name || '').trim().toLowerCase() === lower)) return;

      // 5. Skip if already in the template categories at this level (case-insensitive deduplication)
      if (templateSubfolderSet.has(lower)) return;

      // 6. If at the top category level of a vessel/department, do NOT hoist known nested sub-sub-folders
      if (isAtCategoryLevel && allNestedTemplateSet.has(lower)) return;

      // 7. Skip if already added in dynamicSubfolderSet
      for (const existing of Array.from(dynamicSubfolderSet)) {
        if (existing.toLowerCase() === lower) return;
      }
      dynamicSubfolderSet.add(trimmed);
    };

    // Helper to get normalized folder segments (stripped of container root, vessel, and dept prefix ordering)
    const getFolderTailSegments = (pathStr: string, currentVessel?: string | null, mainDept?: string | null): string[] => {
      if (!pathStr) return [];
      // If pathStr is a groupKey (contains '||'), extract only the breadcrumb/path portion
      let cleanPath = pathStr;
      if (pathStr.includes('||')) {
        const parts = pathStr.split('||');
        const breadcrumbPart = parts.find(p => p.includes('>') || p.includes('/'));
        if (breadcrumbPart) {
          cleanPath = breadcrumbPart;
        } else {
          return [];
        }
      }

      // Ignore Graph IDs / raw IDs
      if (/^01[A-Za-z0-9]{15,}$/.test(cleanPath.trim())) return [];

      const parts = cleanPath.split(/[>/]/).map(s => s.trim()).filter(Boolean);
      const vName = (currentVessel || '').toLowerCase();
      const dName = (mainDept || '').toLowerCase();
      const knownContainers = new Set([
        'vessels', 'specific vessels', 'documents', 'shared documents',
        'common for all ships', 'common for all vessels', 'common',
        'common (not ship specific)', 'common agreements (not ship specific)',
        'kaizen - knowledge bank', 'kaizen', 'knowledge bank',
        'technical & crewing', 'commercial & chartering', 'insurance', 'root'
      ]);

      const filtered = parts.filter(p => {
        const low = p.toLowerCase();
        if (knownContainers.has(low)) return false;
        if (vName && low === vName) return false;
        if (dName && low === dName) return false;
        // Never allow any known vessel name as a tail segment
        if (vessels.some(v => (v.name || '').trim().toLowerCase() === low)) return false;
        // Never allow Graph item IDs
        if (/^01[A-Za-z0-9]{15,}$/.test(p)) return false;
        return true;
      });
      return filtered;
    };

    const currentTail = getFolderTailSegments(fullStackBreadcrumb, vesselName, docMainFolder);

    // Only inspect subfolder tails when we are inside a vessel/department (stackLevel >= 2)
    if (stackLevel >= 2 || currentTail.length > 0) {
      (rows || []).forEach(r => {
        if (r.uploadFolderId && host._appDeletedItemIds.has(r.uploadFolderId)) return;
        if (vesselName && r.vesselName && r.vesselName.toLowerCase() !== vesselName.toLowerCase()) return;
        if (docMainFolder && r.group && r.group.toLowerCase() !== docMainFolder.toLowerCase()) return;

        const rowTail = getFolderTailSegments(r.subFolderPath, r.vesselName, r.group);
        if (rowTail.length > currentTail.length) {
          const startsWithCurrent = currentTail.every((seg, idx) => seg.toLowerCase() === rowTail[idx]?.toLowerCase());
          if (startsWithCurrent) {
            const nextChild = rowTail[currentTail.length];
            if (nextChild) tryAddDynamicSubfolder(nextChild);
          }
        }
      });

      Object.keys(uploadedFilesByFolder || {}).forEach(key => {
        if (host._appDeletedItemIds.has(key)) return;
        if (/^01[A-Za-z0-9]{15,}$/.test(key.trim())) return;

        // If key is a groupKey or path, ensure it does not belong to a different vessel or department
        if (key.includes('||')) {
          const parts = key.split('||');
          const kVessel = (parts[0] || '').trim().toLowerCase();
          const kGroup = (parts[1] || '').trim().toLowerCase();
          if (vesselName && kVessel && kVessel !== vesselName.toLowerCase()) return;
          if (docMainFolder && kGroup && kGroup !== docMainFolder.toLowerCase()) return;
        } else {
          const otherVessel = vessels.find(v => v.name && key.toLowerCase().includes(v.name.toLowerCase()));
          if (otherVessel && vesselName && otherVessel.name.toLowerCase() !== vesselName.toLowerCase()) return;
        }

        const keyTail = getFolderTailSegments(key, vesselName, docMainFolder);
        if (keyTail.length > currentTail.length) {
          const startsWithCurrent = currentTail.every((seg, idx) => seg.toLowerCase() === keyTail[idx]?.toLowerCase());
          if (startsWithCurrent) {
            const nextChild = keyTail[currentTail.length];
            if (nextChild) tryAddDynamicSubfolder(nextChild);
          }
        }
      });
    }

    if (resolvedCurrentFolderId) {
      if (host.state.spoFolderMap.has(resolvedCurrentFolderId)) {
        const liveNode = host.state.spoFolderMap.get(resolvedCurrentFolderId);
        (liveNode?.children || []).forEach(c => {
          if (c?.isFolder && c.name && (!c.id || !host._appDeletedItemIds.has(c.id))) tryAddDynamicSubfolder(c.name);
        });
      }
      for (const [, node] of Array.from(host.state.spoFolderMap.entries())) {
        if (!node || node.deleted || (node.id && host._appDeletedItemIds.has(node.id))) continue;
        if (node.isFolder && node.parentId === resolvedCurrentFolderId && node.name) {
          tryAddDynamicSubfolder(node.name);
        }
      }
    }

    if (dynamicSubfolderSet.size > 0) {
      const merged = new Set<string>(subfolderNames);
      Array.from(dynamicSubfolderSet).forEach(name => merged.add(name));
      subfolderNames = Array.from(merged).filter(isDisplayableFolderName);
    }

    const uploadedOnlySubfolderNames = Array.from(dynamicSubfolderSet)
      .filter(name => !templateSubfolderSet.has((name || '').trim().toLowerCase()));

    const hasDynamicSubfolders = dynamicSubfolderSet.size > 0;
    const showDebugKeys = typeof window !== 'undefined' && /^(localhost|127\.0\.0\.1)$/i.test(window.location.hostname);
    const keyVariants = (raw: string): string[] => {
      const src = (raw || '').trim().toLowerCase();
      if (!src) return [];
      const toBread = (v: string) => v.replace(/\//g, ' > ').replace(/\s*>\s*/g, ' > ').replace(/\s+/g, ' ').trim();
      const toSlash = (v: string) => v.replace(/\s*>\s*/g, '/').replace(/\/+/g, '/').trim();
      const bread = toBread(src);
      const slash = toSlash(src);
      const parts = bread.split(' > ').map(s => s.trim()).filter(Boolean);
      const swappedBread = parts.length >= 2 ? [parts[1], parts[0], ...parts.slice(2)].join(' > ') : '';
      const swappedSlash = swappedBread ? toSlash(swappedBread) : '';
      return Array.from(new Set([
        src,
        bread,
        slash,
        swappedBread,
        swappedSlash,
      ].filter(Boolean)));
    };
    const getRowDebugKey = (row: GroupedRow): string => {
      const liveId = host._getLiveSharePointFolderId(row.subFolderPath) || 'unresolved';
      return `${row.groupKey} | live:${liveId}`;
    };
    const getMatchedUploadKeyForRow = (row: GroupedRow): string => {
      const normBreadcrumb = (p: string) => (p || '').replace(/\s*>\s*/g, ' > ').trim().toLowerCase();
      const rowTail = getFolderTailSegments(row.subFolderPath, row.vesselName, row.group).map(s => s.toLowerCase());
      const canonicalRowKey = [
        (row.vesselName || '').trim().toLowerCase(),
        (row.group || '').trim().toLowerCase(),
        ...rowTail,
      ].filter(Boolean).join('||');
      const normSub = (row.subFolderPath || '').trim().toLowerCase();
      const rowPathParts = normBreadcrumb(normSub).split(' > ').map(s => s.trim()).filter(Boolean);
      const swappedPathLower = (rowPathParts.length >= 2)
        ? [rowPathParts[1], rowPathParts[0], ...rowPathParts.slice(2)].join(' > ')
        : '';
      const sharePointPathLower = host._sharePointFolderPath(row.subFolderPath, '').trim().toLowerCase();
      const sharePointPathBreadLower = sharePointPathLower ? sharePointPathLower.split('/').filter(Boolean).join(' > ') : '';
      const liveId = host._getLiveSharePointFolderId(row.subFolderPath) || '';

      const candidateKeys = [
        row.groupKey,
        row.subFolderPath,
        sharePointPathLower,
        canonicalRowKey,
        liveId,
        (row.uploadFolderId && !/^\d+$/.test(row.uploadFolderId)) ? row.uploadFolderId : '',
      ].filter(Boolean) as string[];

      const expandedCandidates = Array.from(new Set(candidateKeys.flatMap(keyVariants)));
      const matched = expandedCandidates.find(k => Array.isArray(uploadedFilesByFolder[k]) && (uploadedFilesByFolder[k] || []).length > 0);
      if (matched) return matched;
      return `none | live:${liveId || 'unresolved'}`;
    };

    const isLeafFolder = docViewMode === 'folder' && stackLevel > 0 && !atDepartmentVesselList && subfolderNames.length === 0;

    type DisplayFile = { name: string; size: string; date: string; pending?: boolean; id?: string; uploadedAt?: number };
    const allCurrentFolderFilesRaw: DisplayFile[] = [];

    const addCurrentFolderFile = (file: DisplayFile): void => {
      if (!file?.name) return;
      const existing = allCurrentFolderFilesRaw.find(f => f.name.toLowerCase() === file.name.toLowerCase());
      if (!existing) {
        allCurrentFolderFilesRaw.push(file);
        return;
      }
      if ((!existing.size || existing.size === '—') && file.size && file.size !== '—') existing.size = file.size;
      if ((!existing.date || existing.date === 'Today') && file.date) existing.date = file.date;
      if (!existing.uploadedAt && file.uploadedAt) existing.uploadedAt = file.uploadedAt;
      if ((!existing.id || existing.id === existing.name) && file.id) existing.id = file.id;
    };

    if (docViewMode === 'folder' && currentFolderNode) {
      const currentTail = getFolderTailSegments(fullStackBreadcrumb, vesselName, docMainFolder);
      const currentTailKey = currentTail.map(s => s.toLowerCase()).join(' > ');
      const currentNameNorm = (currentFolderName || '').trim().toLowerCase();
      const mainFolderNamesLower = ['technical & crewing', 'commercial & chartering', 'insurance', 'kaizen - knowledge bank'];
      const belongsToCurrentBranch = (rawKey: string): boolean => {
        const key = (rawKey || '').trim();
        if (!key) return false;
        const keyLower = key.toLowerCase();

        // Group key shape: vessel||group||category||subCategory||breadcrumb
        if (key.includes('||')) {
          const parts = key.split('||');
          const keyVessel = (parts[0] || '').trim().toLowerCase();
          const keyGroup = (parts[1] || '').trim().toLowerCase();
          if (vesselName && keyVessel && keyVessel !== vesselName.trim().toLowerCase()) return false;
          if (docMainFolder && keyGroup && keyGroup !== docMainFolder.trim().toLowerCase()) return false;
          return true;
        }

        // If a textual key explicitly names a main folder, enforce exact same folder.
        if (docMainFolder) {
          const referencedMain = mainFolderNamesLower.find(name => keyLower.includes(name));
          if (referencedMain && referencedMain !== docMainFolder.trim().toLowerCase()) return false;
        }

        // If a textual key includes any known vessel name, enforce current vessel.
        const referencedVessel = vessels.find(v => keyLower.includes((v.name || '').trim().toLowerCase()));
        if (vesselName && referencedVessel && referencedVessel.name.trim().toLowerCase() !== vesselName.trim().toLowerCase()) return false;

        return true;
      };

      (rows || []).forEach(r => {
        if (r.uploadFolderId && host._appDeletedItemIds.has(r.uploadFolderId)) return;
        if (vesselName && r.vesselName && r.vesselName.toLowerCase() !== vesselName.toLowerCase()) return;
        if (docMainFolder && r.group && r.group.toLowerCase() !== docMainFolder.toLowerCase()) return;

        const rowTail = getFolderTailSegments(r.subFolderPath, r.vesselName, r.group);
        const rowTailKey = rowTail.map(s => s.toLowerCase()).join(' > ');

        const isExactMatch = currentTailKey ? rowTailKey === currentTailKey : (
          (r.subCategory || '').toLowerCase() === currentNameNorm ||
          (r.category || '').toLowerCase() === currentNameNorm
        );

        if (isExactMatch && r.fileName) {
          addCurrentFolderFile({
            id: r.fileId || r.fileName,
            name: r.fileName,
            size: r.fileSize || '—',
            date: r.fileUploadedAt ? new Date(r.fileUploadedAt).toLocaleString() : 'Today',
            uploadedAt: r.fileUploadedAt,
          });
        }
      });

      // Gather files from uploadedFilesByFolder across all matching keys
      Object.entries(uploadedFilesByFolder || {}).forEach(([key, list]) => {
        if (!Array.isArray(list) || list.length === 0) return;
        if (host._appDeletedItemIds.has(key)) return;

        const keyLooksScoped = key.includes('||') || key.includes('>') || key.includes('/');
        const keyTail = getFolderTailSegments(key, vesselName, docMainFolder);
        const keyTailKey = keyTail.map(s => s.toLowerCase()).join(' > ');

        const safeCurrentNodeId = (currentFolderNode?.id && !/^\d+$/.test(currentFolderNode.id)) ? currentFolderNode.id : null;
        const isKeyMatch = (
          (keyLooksScoped && currentTailKey && keyTailKey === currentTailKey) ||
          (resolvedCurrentFolderId && key.toLowerCase() === resolvedCurrentFolderId.toLowerCase()) ||
          (safeCurrentNodeId && key.toLowerCase() === safeCurrentNodeId.toLowerCase())
        );

        if (isKeyMatch && belongsToCurrentBranch(key)) {
          list.forEach(item => {
            if (!item?.name) return;
            addCurrentFolderFile({
              id: (item as any).id || item.name,
              name: item.name,
              size: item.size || '—',
              date: item.date || 'Today',
              pending: !!item.pending,
              uploadedAt: (item as any).uploadedAt,
            });
          });
        }
      });

      // Also scan live spoFolderMap for files inside resolvedCurrentFolderId or matching path
      const folderMapCandidates = new Set<string>();
      if (resolvedCurrentFolderId && !/^\d+$/.test(resolvedCurrentFolderId)) folderMapCandidates.add(resolvedCurrentFolderId);
      if (currentFolderNode?.id && !/^\d+$/.test(currentFolderNode.id)) folderMapCandidates.add(currentFolderNode.id);

      for (const fId of Array.from(folderMapCandidates)) {
        if (host.state.spoFolderMap.has(fId)) {
          const liveNode = host.state.spoFolderMap.get(fId);
          (liveNode?.children || []).forEach(c => {
            if (!c.isFolder && c.name && (!c.id || !host._appDeletedItemIds.has(c.id))) {
              addCurrentFolderFile({
                id: c.id,
                name: c.name,
                size: typeof c.size === 'number' ? `${(c.size / 1024).toFixed(1)} KB` : '—',
                date: (c as any).createdDateTime
                  ? new Date((c as any).createdDateTime).toLocaleString()
                  : ((c as any).lastModifiedDateTime ? new Date((c as any).lastModifiedDateTime).toLocaleString() : 'Today'),
                pending: false,
                uploadedAt: (c as any).createdDateTime
                  ? Date.parse((c as any).createdDateTime)
                  : ((c as any).lastModifiedDateTime ? Date.parse((c as any).lastModifiedDateTime) : undefined),
              });
            }
          });
        }
      }
    } else if (docViewMode === 'folder' && stackLevel === 0) {
      // Gather files uploaded at root / recent batch uploads
      Object.entries(uploadedFilesByFolder || {}).forEach(([key, list]) => {
        if (!Array.isArray(list) || list.length === 0) return;
        if (host._appDeletedItemIds.has(key)) return;
        list.forEach(item => {
          if (!item?.name) return;
          addCurrentFolderFile({
            id: (item as any).id || item.name,
            name: item.name,
            size: item.size || '—',
            date: item.date || 'Today',
            pending: !!item.pending,
            uploadedAt: (item as any).uploadedAt,
          });
        });
      });

      // Also gather files from rows
      (rows || []).forEach(r => {
        if (r.uploadFolderId && host._appDeletedItemIds.has(r.uploadFolderId)) return;
        if (r.fileName) {
          addCurrentFolderFile({
            id: r.fileId || r.fileName,
            name: r.fileName,
            size: r.fileSize || '—',
            date: r.fileUploadedAt ? new Date(r.fileUploadedAt).toLocaleString() : 'Today',
            uploadedAt: r.fileUploadedAt,
          });
        }
      });
    }

    const allCurrentFolderFiles = allCurrentFolderFilesRaw
      .slice()
      .sort((a, b) => (b.uploadedAt || 0) - (a.uploadedAt || 0));

    const visibleCurrentFolderFiles = textFilter.trim()
      ? allCurrentFolderFiles.filter(file => matchesSearchTokens(textFilter, file.name))
      : allCurrentFolderFiles;

    const commonTemplateRows = getCommonShipsFlatRows();
    const kaizenTemplateRows = getKaizenFlatRows();

    // 1. Sanitize specific vessel rows from API: remove any Kaizen rows from individual vessels
    const sanitizedVesselRows = (rows || []).filter(r => {
      const isKaizen = (r.group || '').toLowerCase().includes('kaizen') || (r.subFolderPath || '').toLowerCase().includes('kaizen');
      const isCommon = r.vesselName === 'Common for all vessels' || (r.subFolderPath || '').toLowerCase().includes('common for all ships');
      return !isKaizen && !isCommon;
    });

    const activeVesselName = (docScopeType === 'vessels' && vesselFilter !== 'all') ? vesselFilter : null;
    const vesselHasRows = activeVesselName
      ? sanitizedVesselRows.some(r => r.vesselName.trim().toLowerCase() === activeVesselName.trim().toLowerCase())
      : true;
    const effectiveVesselRows = (activeVesselName && !vesselHasRows)
      ? [...sanitizedVesselRows, ...getVesselTemplateFlatRows(activeVesselName)]
      : sanitizedVesselRows;

    // 2. Collect dynamic uploaded files/rows for Common and Kaizen
    const dynamicCommonRows = (rows || []).filter(r =>
      r.vesselName === 'Common for all vessels' || (r.subFolderPath || '').toLowerCase().includes('common for all ships')
    );
    const dynamicKaizenRows = (rows || []).filter(r =>
      ((r.group || '').toLowerCase().includes('kaizen') || (r.subFolderPath || '').toLowerCase().includes('kaizen'))
    ).map(r => ({
      ...r,
      vesselName: 'Kaizen - Knowledge Bank',
      group: 'Kaizen - Knowledge Bank',
      subFolderPath: r.subFolderPath.replace(/^[^>]+>\s*kaizen\s*-\s*knowledge\s*bank/i, 'Kaizen - Knowledge Bank'),
    }));

    const mergedCommonRows = [...commonTemplateRows];
    dynamicCommonRows.forEach(dr => {
      const normSub = (dr.subFolderPath || '').toLowerCase();
      const idx = mergedCommonRows.findIndex(cr => (cr.subFolderPath || '').toLowerCase() === normSub);
      if (idx !== -1) {
        if (dr.fileName) mergedCommonRows[idx] = { ...mergedCommonRows[idx], ...dr };
      } else {
        mergedCommonRows.push(dr);
      }
    });

    const mergedKaizenRows = [...kaizenTemplateRows];
    dynamicKaizenRows.forEach(dk => {
      const normSub = (dk.subFolderPath || '').toLowerCase();
      const idx = mergedKaizenRows.findIndex(kr => (kr.subFolderPath || '').toLowerCase() === normSub);
      if (idx !== -1) {
        if (dk.fileName) mergedKaizenRows[idx] = { ...mergedKaizenRows[idx], ...dk };
      } else {
        mergedKaizenRows.push(dk);
      }
    });

    // ── Intelligent Metadata Parser with Per-Vessel, Department & Structural Fallbacks ──
    const allFleetVesselNames = [
      'Snow Flower', 'Snow Flake', 'Senegal Express', 'Peissy', 'Potiniere',
      'Norse New Haven', 'Norse Ijmuiden', 'Belle Lune', 'Bow Fighter'
    ];

    // Vessel alias tables + keyword sets used by parseSharePointRowMetadata.
    // These depend only on render-time state, so they are built once per
    // render instead of once per parse call (the parser runs several times per
    // row on every render, which made List view and the filters slow).
    const buildVesselAliasTables = () => {
      const allVesselNames = new Set<string>();
      const vesselAliasMap = new Map<string, string>();

      allFleetVesselNames.forEach(name => {
        allVesselNames.add(name);
        vesselAliasMap.set(name.toLowerCase(), name);
      });

      (vessels || []).forEach(v => {
        if (v.name) {
          allVesselNames.add(v.name);
          vesselAliasMap.set(v.name.toLowerCase(), v.name);
        }
      });

      // SharePoint Term Store vessel terms for the active site — the
      // authoritative vessel list per the "Vessel Name" managed metadata
      // column (see GET /api/sites/{site_id}/term-store-vessels). Folders
      // named after a vessel that only exists in the term store (not yet in
      // the app's own vessels table) are still recognised as vessels here
      // instead of falling through to "Not Listed".
      liveTermStoreVesselNames.forEach(name => {
        if (name && !vesselAliasMap.has(name.toLowerCase())) {
          allVesselNames.add(name);
          vesselAliasMap.set(name.toLowerCase(), name);
        }
      });

      const dynamicVesselAliases = host.state.documentVesselAliases || {};
      Object.entries(dynamicVesselAliases).forEach(([canonical, aliasList]) => {
        allVesselNames.add(canonical);
        vesselAliasMap.set(canonical.toLowerCase(), canonical);
        if (Array.isArray(aliasList)) {
          aliasList.forEach(a => {
            if (a) vesselAliasMap.set(a.trim().toLowerCase(), canonical);
          });
        }
      });

      const matchVesselStringCache = new Map<string, string | null>();
      const matchVesselString = (str: string): string | null => {
        if (!str) return null;
        if (matchVesselStringCache.has(str)) return matchVesselStringCache.get(str)!;
        const result = matchVesselStringUncached(str);
        matchVesselStringCache.set(str, result);
        return result;
      };
      const matchVesselStringUncached = (str: string): string | null => {
        const clean = str.trim().toLowerCase().replace(/^(mv|m\/v|mt)\s+/i, '').trim();
        if (vesselAliasMap.has(clean)) return vesselAliasMap.get(clean)!;
        if (vesselAliasMap.has(str.trim().toLowerCase())) return vesselAliasMap.get(str.trim().toLowerCase())!;
        for (const [alias, canonical] of Array.from(vesselAliasMap.entries())) {
          if (alias.length >= 4 && clean.includes(alias)) {
            return canonical;
          }
        }
        return null;
      };

      const nonVesselKeywords = new Set([
        'sharepoint sites', 'sites documents', 'shared documents', 'documents',
        'nksdocman', 'communication site', 'root', 'technical', 'crewing',
        'technical & crewing', 'commercial & chartering', 'insurance',
        'kaizen', 'kaizen - knowledge bank', 'common for all vessels', 'common for all ships',
        'site library', 'general documents'
      ]);
      // Every configured site's own key/display name is a breadcrumb
      // container segment (like "Communication Site" / "NKSDocMan" above),
      // never a vessel or a real folder — e.g. "NissenKaiunExternal".
      (host.state.documentSites || []).forEach(s => {
        if (s.site_key) nonVesselKeywords.add(s.site_key.trim().toLowerCase());
        if (s.sp_site_name) nonVesselKeywords.add(s.sp_site_name.trim().toLowerCase());
      });
      return { allVesselNames, vesselAliasMap, matchVesselString, nonVesselKeywords };
    };
    let vesselAliasTablesCache: ReturnType<typeof buildVesselAliasTables> | null = null;
    const getVesselAliasTables = (): ReturnType<typeof buildVesselAliasTables> => {
      if (!vesselAliasTablesCache) vesselAliasTablesCache = buildVesselAliasTables();
      return vesselAliasTablesCache;
    };

    const parseSharePointRowMetadataUncached = (
      rowPath: string,
      fileName?: string | null,
      rawVessel?: string | null,
      rawGroup?: string | null,
      rawCategory?: string | null,
      rawSubCategory?: string | null
    ): {
      vessel: string;
      mainFolder: string;
      group: string;
      category: string;
      documentSection: string;
      subCategory: string;
    } => {
      const parts = (rowPath || '')
        .replace(/\//g, ' > ')
        .split('>')
        .map(s => s.trim())
        .filter(Boolean);

      const { allVesselNames, matchVesselString, nonVesselKeywords } = getVesselAliasTables();


      let detectedVessel: string | null = null;
      if (rawVessel && !nonVesselKeywords.has(rawVessel.trim().toLowerCase())) {
        detectedVessel = matchVesselString(rawVessel) || (allVesselNames.has(rawVessel) ? rawVessel : null);
      }
      if (!detectedVessel) {
        for (const part of parts) {
          if (nonVesselKeywords.has(part.toLowerCase())) continue;
          const matched = matchVesselString(part);
          if (matched) {
            detectedVessel = matched;
            break;
          }
        }
      }
      if (!detectedVessel && fileName) {
        detectedVessel = matchVesselString(fileName);
      }

      const vesselLabel = detectedVessel || 'Not Listed';

      const deptAliases: Record<string, string[]> = host.state.documentDepartmentAliases || {
        'Technical & Crewing': ['technical and crewing new', 'technical and crewing  new', 'technical & crewing', 'technical', 'crewing'],
        'Commercial & Chartering': ['commercial and chartering', 'commercial & chartering', 'commercial', 'chartering'],
        'Insurance': ['insurance', 'claims'],
        'Kaizen - Knowledge Bank': ['kaizen', 'knowledge bank', 'kaizen - knowledge bank'],
      };

      let detectedMainFolder: string | null = null;
      const matchDepartment = (str: string): string | null => {
        if (!str) return null;
        const low = str.trim().toLowerCase();
        for (const [canonicalDept, aliasList] of Object.entries(deptAliases)) {
          if (low === canonicalDept.toLowerCase()) return canonicalDept;
          if (Array.isArray(aliasList)) {
            for (const alias of aliasList) {
              if (low === alias || low.includes(alias)) return canonicalDept;
            }
          }
        }
        return null;
      };

      for (const part of parts) {
        const matchedDept = matchDepartment(part);
        if (matchedDept) {
          detectedMainFolder = matchedDept;
          break;
        }
      }
      if (!detectedMainFolder && rawGroup) {
        detectedMainFolder = matchDepartment(rawGroup);
      }
      // A file that isn't under one of the known departments still sits
      // under *some* real top-level folder — e.g. "Report", "ACRA CHARGE",
      // or a vessel folder that sits flat at the site root (e.g.
      // "mvtest2209", where the vessel folder itself is the only top-level
      // container). Use the first non-generic path segment as the Main
      // folder label instead of leaving it blank, so the row/column always
      // shows the real folder it's actually in; prefer a non-vessel segment
      // (a genuine "other" folder) but fall back to the vessel's own
      // top-level folder when that's the only real segment there is.
      const isFallbackSegment = (p: string): boolean => {
        const low = p.trim().toLowerCase();
        if (!low || nonVesselKeywords.has(low)) return false;
        if (fileName && p.trim().toLowerCase() === fileName.trim().toLowerCase()) return false;
        if (/\.(pdf|dwg|dxf|xlsx|xls|docx|doc|txt|msg|eml|png|jpg|jpeg|zip)$/i.test(p.trim())) return false;
        return true;
      };
      // Index of the document-library root marker ("Documents"/"Shared
      // Documents"/...) within the path, reused below so both the Main
      // folder fallback and the sub-level split are structural (positional)
      // rather than dependent on enumerating every possible container name
      // — a site display name like "NissenKaiunExternal" sits *before* this
      // marker and is therefore never a candidate, regardless of whether
      // it's registered in any keyword set.
      const libraryRootIdx = parts.findIndex(p => /^(documents|shared documents|sites documents|site library|general documents)$/i.test(p.trim()));

      if (!detectedMainFolder) {
        // Scan only the segments that sit AFTER the document library root
        // ("... > Documents > ACRA CHARGE", "... > Shared Documents >
        // Report"). Falls back to scanning the whole path only if no
        // library-root marker is present.
        const candidateParts = libraryRootIdx >= 0 ? parts.slice(libraryRootIdx + 1) : parts;
        const targetVesselLow = (detectedVessel || '').trim().toLowerCase();
        const nonVesselSegment = candidateParts.find(p => isFallbackSegment(p) && p.trim().toLowerCase() !== targetVesselLow);
        detectedMainFolder = nonVesselSegment || candidateParts.find(isFallbackSegment) || null;
      }

      const mainFolderLabel = detectedMainFolder || '';

      const knownContainers = new Set([
        'sharepoint sites', 'sites documents', 'shared documents', 'documents',
        'nksdocman', 'communication site', 'site library', 'general documents',
        (detectedVessel || '').toLowerCase(),
        (vesselLabel || '').toLowerCase(),
        (detectedMainFolder || '').toLowerCase(),
        'technical and crewing new', 'technical and crewing  new', 'technical', 'crewing',
        'technical & crewing', 'commercial and chartering', 'commercial & chartering',
        'insurance', 'kaizen', 'kaizen - knowledge bank'
      ]);
      // Every configured site's own key/display name is a breadcrumb
      // container segment, never a real sub-folder/category — same reason
      // it's excluded from nonVesselKeywords above.
      (host.state.documentSites || []).forEach(s => {
        if (s.site_key) knownContainers.add(s.site_key.trim().toLowerCase());
        if (s.sp_site_name) knownContainers.add(s.sp_site_name.trim().toLowerCase());
      });

      const isFileString = (s: string): boolean => {
        if (!s) return false;
        if (fileName && s.trim().toLowerCase() === fileName.trim().toLowerCase()) return true;
        return /\.(pdf|dwg|dxf|xlsx|xls|docx|doc|txt|msg|eml|png|jpg|jpeg|zip)$/i.test(s.trim());
      };

      let vesselIdx = -1;
      const targetVesselLower = (detectedVessel || '').toLowerCase();
      if (detectedVessel) {
        vesselIdx = parts.findIndex(p => matchVesselString(p) === detectedVessel || p.toLowerCase() === targetVesselLower);
      }

      // Where the resolved Main folder segment itself sits in the path —
      // everything strictly after it (and after the vessel, if any) is the
      // real sub-folder/category chain. Searching from just after the
      // library root (when present) guards against a same-named segment
      // appearing earlier, e.g. inside the site name itself.
      const mainFolderLow = mainFolderLabel.trim().toLowerCase();
      const mainFolderIdx = mainFolderLow
        ? parts.findIndex((p, i) => i > libraryRootIdx && p.trim().toLowerCase() === mainFolderLow)
        : -1;

      let subLevels: string[];
      if (vesselIdx >= 0) {
        subLevels = parts.slice(vesselIdx + 1).filter(p => !isFileString(p) && (!targetVesselLower || p.toLowerCase() !== targetVesselLower));
      } else if (mainFolderIdx >= 0) {
        subLevels = parts.slice(mainFolderIdx + 1).filter(p => !isFileString(p) && !allVesselNames.has(p));
      } else {
        subLevels = parts.filter(p => !knownContainers.has(p.toLowerCase()) && !allVesselNames.has(p) && !isFileString(p));
      }

      const section = subLevels[0] || (rawCategory && !isFileString(rawCategory) && !knownContainers.has(rawCategory.toLowerCase()) ? rawCategory : '') || '';
      const group = subLevels[1] || (rawGroup && rawGroup !== section && !isFileString(rawGroup) && !knownContainers.has(rawGroup.toLowerCase()) ? rawGroup : '') || '';
      const category = subLevels[2] || (rawCategory && rawCategory !== section && rawCategory !== group && !isFileString(rawCategory) && !knownContainers.has(rawCategory.toLowerCase()) ? rawCategory : '') || '';
      const subCategory = subLevels[3] || (rawSubCategory && rawSubCategory !== category && rawSubCategory !== group && !isFileString(rawSubCategory) ? rawSubCategory : '') || '';

      return {
        vessel: vesselLabel,
        mainFolder: mainFolderLabel,
        documentSection: section || '',
        group: group || '',
        category: category || '',
        subCategory: subCategory || '',
      };
    };

    const parseSharePointRowMetadataCache = new Map<string, ReturnType<typeof parseSharePointRowMetadataUncached>>();
    const parseSharePointRowMetadata = (
      rowPath: string,
      fileName?: string | null,
      rawVessel?: string | null,
      rawGroup?: string | null,
      rawCategory?: string | null,
      rawSubCategory?: string | null
    ): ReturnType<typeof parseSharePointRowMetadataUncached> => {
      const cacheKey = [rowPath, fileName, rawVessel, rawGroup, rawCategory, rawSubCategory].map(v => v || '').join('\u0001');
      const cached = parseSharePointRowMetadataCache.get(cacheKey);
      if (cached) return cached;
      const result = parseSharePointRowMetadataUncached(rowPath, fileName, rawVessel, rawGroup, rawCategory, rawSubCategory);
      parseSharePointRowMetadataCache.set(cacheKey, result);
      return result;
    };

    const getSharePointSiteFlatRows = (): FlatRow[] => {
      const generatedRows: FlatRow[] = [];
      const seenRowKeys = new Set<string>();

      const addFlatRow = (
        srPrefix: string,
        path: string,
        fileName?: string | null,
        fileId?: string | null,
        fileSize?: string,
        uploadedAt?: number,
        folderId?: string,
        tags?: {
          vessel?: string;
          domain?: string;
          department?: string;
          mainFolder?: string;
          group?: string;
          category?: string;
          subCategory?: string;
          sub_category?: string;
          documentSection?: string;
          document_section?: string;
        }
      ) => {
        const cleanPath = path.trim();
        const dedupeKey = `${cleanPath.toLowerCase()}||${(fileName || '').toLowerCase()}`;
        if (seenRowKeys.has(dedupeKey)) return;
        seenRowKeys.add(dedupeKey);

        const labels = parseSharePointRowMetadata(cleanPath, fileName);
        const tagVessel = tags?.vessel && tags.vessel !== 'To Be Classified' && tags.vessel !== 'Unknown' && tags.vessel !== 'Not Listed' ? tags.vessel : null;
        const rowDomain = tags?.domain || '';
        const tagDept = tags?.department || tags?.mainFolder || null;
        const tagSection = tags?.documentSection || tags?.document_section || null;
        const tagGroup = tags?.group || null;
        const tagCat = tags?.category || null;
        const tagSubCat = tags?.subCategory || tags?.sub_category || null;

        const rowVessel = tagVessel || labels.vessel;
        const rowMainFolder = tagDept || (labels.mainFolder !== 'Main folder not assigned' ? labels.mainFolder : (tagGroup || 'Technical & Crewing'));
        const rowSection = tagSection || (labels.documentSection !== 'Document section not assigned' ? labels.documentSection : (tagGroup || 'General'));
        const rowGroup = tagGroup || (labels.group !== 'Group not assigned' ? labels.group : rowSection);
        const rowCat = tagCat || (labels.category !== 'Category not assigned' ? labels.category : (labels.group !== 'Group not assigned' ? labels.group : labels.documentSection));
        const rowSubCat = tagSubCat || (labels.subCategory !== 'Sub-category not assigned' ? labels.subCategory : labels.category);

        generatedRows.push({
          srNo: `${srPrefix}-${generatedRows.length + 1}`,
          vesselName: rowVessel,
          domain: rowDomain,
          group: rowMainFolder,
          category: rowCat,
          subCategory: rowSubCat,
          subFolderPath: cleanPath,
          fileName: fileName || null,
          fileId: fileId || null,
          fileSize: fileSize,
          fileUploadedAt: uploadedAt,
          canUpload: true,
          groupKey: `spo_${folderId || dedupeKey.replace(/[^a-zA-Z0-9]/g, '_')}`,
          uploadFolderId: folderId || '',
          monthDriven: false,
          ...(host.state.activeDocumentSite ? { siteKey: host.state.activeDocumentSite } : {}),
        } as FlatRow);
      };

      // Resolve the active site — always track which site the user has selected
      const activeSiteKey = host.state.activeDocumentSite;
      const baseSite = (activeSiteKey
        ? host.state.documentSites.find(s => s.site_key === activeSiteKey || s.site_id === activeSiteKey)
        : null) || host.state.documentSites[0];
      const baseSiteName = baseSite?.sp_site_name || baseSite?.site_key || 'NKSDocMan';
      const baseLibName = baseSite?.default_library_name || 'Shared Documents';
      // IDs we want to restrict cache entries to (site-scoped filtering)
      const activeSiteId = baseSite?.site_id || '';
      const activeDriveId = baseSite?.drive_id || '';

      // 1. Process Cached items from live SharePoint API calls (includes full SharePoint Online tags)
      //    Only include entries that belong to the currently-selected site to prevent cross-site mixing.
      for (const [cacheKey, cacheEntry] of Array.from(host._siteFolderItemsCache.entries())) {
        if (!cacheEntry?.items) continue;
        // Cache keys are formatted as "siteId:driveId:folderId" — skip entries from other sites
        if (activeSiteId && !cacheKey.startsWith(activeSiteId)) continue;
        if (activeDriveId && !cacheKey.includes(activeDriveId)) continue;
        const parentPath = (cacheEntry as any).parentPath || '';
        cacheEntry.items.forEach(item => {
          if (!item?.name) return;
          const isFolder = item.folder || (!item.file && !item.name.includes('.'));
          const fileName = isFolder ? null : item.name;
          const fileId = isFolder ? null : (item.id || item.name);
          const fileSize = typeof item.size === 'number'
            ? (item.size > 1024 * 1024 ? `${(item.size / (1024 * 1024)).toFixed(1)} MB` : `${(item.size / 1024).toFixed(1)} KB`)
            : undefined;
          const uploadedAt = item.lastModifiedDateTime ? Date.parse(item.lastModifiedDateTime) : (item.createdDateTime ? Date.parse(item.createdDateTime) : undefined);
          const pRef = item.parentReference?.path ? item.parentReference.path.split('root:', 2)[1] : '';
          const pPath = (item.path || pRef || parentPath || '').replace(/^\//, '');
          const rowPath = pPath
            ? (pPath.startsWith('SharePoint Sites') ? pPath.replace(/\//g, ' > ') : `SharePoint Sites > ${baseSiteName} > ${baseLibName} > ${pPath.replace(/\//g, ' > ')}`)
            : `SharePoint Sites > ${baseSiteName} > ${baseLibName}`;
          const folderId = isFolder ? (item.id || '') : (item.parentReference?.id || '');
          addFlatRow('CACHE', rowPath, fileName, fileId, fileSize, uploadedAt, folderId, item.tags);
        });
      }

      // 2. Process Live Folders & Files from recursive scan (with SharePoint Online tags)
      //    documentLiveFolders are already scoped to the activeDocumentSite by _loadDocumentLiveTree.
      (host.state.documentLiveFolders || []).forEach(lf => {
        if (!lf.name && !lf.path) return;
        const isFolder = lf.is_folder !== false;
        const fileName = isFolder ? null : lf.name;
        const fileId = isFolder ? null : lf.id;
        const fileSize = typeof lf.size === 'number'
          ? (lf.size > 1024 * 1024 ? `${(lf.size / (1024 * 1024)).toFixed(1)} MB` : `${(lf.size / 1024).toFixed(1)} KB`)
          : undefined;
        const uploadedAt = lf.created_date_time ? Date.parse(lf.created_date_time) : (lf.last_modified_date_time ? Date.parse(lf.last_modified_date_time) : undefined);
        const folderOnlyPath = lf.path ? lf.path.replace(/\/[^/]+$/, '') : '';
        const rowPath = folderOnlyPath
          ? (folderOnlyPath.startsWith('SharePoint Sites') ? folderOnlyPath.replace(/\//g, ' > ') : `SharePoint Sites > ${baseSiteName} > ${baseLibName} > ${folderOnlyPath.replace(/\//g, ' > ')}`)
          : `SharePoint Sites > ${baseSiteName} > ${baseLibName}`;
        addFlatRow('LIVE', rowPath, fileName, fileId, fileSize, uploadedAt, lf.id, (lf as any).tags);
      });

      // 3. Known files in configured sites — ONLY included if this site matches the active site
      if (host.state.documentSites.length > 1) {
        const s2 = host.state.documentSites[1];
        if (baseSite && (baseSite.site_key === s2.site_key || baseSite.site_id === s2.site_id)) {
          const s2Name = s2.sp_site_name || s2.site_key || 'Communication Site';
          const s2Lib = s2.default_library_name || 'Documents';
          addFlatRow('SPO', `SharePoint Sites > ${s2Name} > ${s2Lib} > 5450 - Ballast Water managemnet plan.pdf`, '5450 - Ballast Water managemnet plan.pdf', '014ZGIJDMKWGFXCUBFJVHKFLAIBOKU3KI7', '1.2 MB');
          addFlatRow('SPO', `SharePoint Sites > ${s2Name} > ${s2Lib} > Boiler - tech specification sheet with FOC from manual.pdf`, 'Boiler - tech specification sheet with FOC from manual.pdf', '014ZGIJDLCHDQMTTTKERFZONUBR63NNCGD', '840 KB');
        }
      }

      // 4. Template folder paths — only add empty folder paths for NKSDocMan when NO actual files have been found yet
      const hasRealFiles = generatedRows.some(r => !!r.fileName);
      const isNksSite = (baseSiteName || '').toLowerCase().includes('nks') || (activeSiteKey || '').toLowerCase().includes('nks');
      if (!hasRealFiles && isNksSite) {
        allFleetVesselNames.forEach(vName => {
          const vPath = `SharePoint Sites > ${baseSiteName} > ${baseLibName} > Technical and Crewing  New > ${vName}`;
          addFlatRow('SPO', `${vPath} > Drawings and Manuals > Drawings`, null, null);
          addFlatRow('SPO', `${vPath} > Drawings and Manuals > Manuals`, null, null);
          addFlatRow('SPO', `${vPath} > Drawings and Manuals > To Be Classified`, null, null);
          addFlatRow('SPO', `${vPath} > Certificates`, null, null);
          addFlatRow('SPO', `${vPath} > Technical`, null, null);
        });
        addFlatRow('SPO', `SharePoint Sites > ${baseSiteName} > ${baseLibName} > Technical and Crewing  New`);
        addFlatRow('SPO', `SharePoint Sites > ${baseSiteName} > ${baseLibName} > Technical`);
      }

      return generatedRows;
    };

    const siteRows: FlatRow[] = getSharePointSiteFlatRows();
    const sharedDocsRows: FlatRow[] = siteRows.filter(r => (r.subFolderPath || '').includes('Shared Documents') || r.group === 'Shared Documents');
    const documentsLibraryRows: FlatRow[] = siteRows.filter(r => (r.subFolderPath || '').includes('Documents') || r.group === 'Documents');

    const allRows: FlatRow[] = [
      ...effectiveVesselRows,
      ...mergedCommonRows,
      ...mergedKaizenRows,
      ...siteRows,
    ];

    // Filter by active scope (vessels, common, kaizen, shared_docs, documents, sites)
    const scopeRows = allRows.filter(r => {
      const isSharePointRow = (r.subFolderPath || '').startsWith('SharePoint Sites') ||
                              (r.groupKey || '').startsWith('spo_') ||
                              (r.groupKey || '').startsWith('live:') ||
                              r.group === 'SharePoint Sites';
      if (docScopeType === 'vessels') {
        return r.vesselName !== 'Common for all vessels' &&
               r.vesselName !== 'Kaizen - Knowledge Bank' &&
               r.group !== 'SharePoint Sites' &&
               r.group !== 'Shared Documents' &&
               r.group !== 'Documents' &&
               !isSharePointRow &&
               !(r.subFolderPath || '').startsWith('Shared Documents') &&
               !(r.subFolderPath || '').startsWith('Documents');
      } else if (docScopeType === 'common') {
        return r.vesselName === 'Common for all vessels';
      } else if (docScopeType === 'kaizen') {
        return r.vesselName === 'Kaizen - Knowledge Bank';
      } else if (docScopeType === 'shared_docs') {
        return r.group === 'Shared Documents' || (r.subFolderPath || '').startsWith('Shared Documents') || (r.subFolderPath || '').includes('Shared Documents');
      } else if (docScopeType === 'documents') {
        return r.group === 'Documents' || (r.subFolderPath || '').startsWith('Documents') || (r.subFolderPath || '').includes('Documents');
      } else if (docScopeType === 'sites') {
        if (!isSharePointRow &&
            !(r.subFolderPath || '').startsWith('Shared Documents') &&
            !(r.subFolderPath || '').startsWith('Documents') &&
            r.group !== 'Shared Documents' &&
            r.group !== 'Documents') {
          return false;
        }
        // Strict active site matching: when a site is selected, only show rows belonging to THAT site
        const activeSiteKeyNow = host.state.activeDocumentSite;
        if (activeSiteKeyNow) {
          const currentSiteObj = (host.state.documentSites || []).find(s =>
            s.site_key === activeSiteKeyNow || s.site_id === activeSiteKeyNow
          );
          const activeKeyClean = activeSiteKeyNow.toLowerCase();
          const activeNameClean = (currentSiteObj?.sp_site_name || currentSiteObj?.site_key || activeSiteKeyNow).toLowerCase();
          const activeIdClean = (currentSiteObj?.site_id || '').toLowerCase();

          // 1. Check live row key (live:siteKey:folderId)
          if ((r.groupKey || '').startsWith('live:')) {
            const rowSiteKey = (r.groupKey.split(':')[1] || '').toLowerCase();
            return rowSiteKey === activeKeyClean || rowSiteKey === activeIdClean;
          }

          // 2. Check explicit siteKey attached to row
          if ((r as any).siteKey) {
            const rowKey = String((r as any).siteKey).toLowerCase();
            return rowKey === activeKeyClean || rowKey === activeIdClean;
          }

          // 3. Check site name in subFolderPath ("SharePoint Sites > <SiteName> > ...")
          const subPath = (r.subFolderPath || '').trim();
          if (subPath.startsWith('SharePoint Sites')) {
            const segments = subPath.split(/\s*>\s*/);
            if (segments.length >= 2) {
              const rowSiteName = segments[1].toLowerCase();
              return rowSiteName === activeNameClean ||
                     rowSiteName === activeKeyClean ||
                     (activeNameClean && (rowSiteName.includes(activeNameClean) || activeNameClean.includes(rowSiteName))) ||
                     (activeKeyClean && (rowSiteName.includes(activeKeyClean) || activeKeyClean.includes(rowSiteName)));
            }
          }

          return false;
        }
        return true;
      }
      return true;
    });

    // Scope rows for dropdown populating
    const scopeVesselRows = activeVesselName
      ? scopeRows.filter(r => r.vesselName.trim().toLowerCase() === activeVesselName.trim().toLowerCase())
      : scopeRows;
    const wholeSiteSearch = Boolean(
      host.state.activeDocumentSite && textFilter.trim().length > 0 &&
      (docScopeType === 'sites' || docScopeType === 'shared_docs' || docScopeType === 'documents')
    );

    const hierarchyForRow = (row: Pick<FlatRow, 'vesselName' | 'group' | 'category' | 'subCategory' | 'subFolderPath'>): { section: string; group: string; category: string; subCategory: string } => {
      const parts = (row.subFolderPath || '').split('>').map(part => part.trim()).filter(Boolean);
      const knownContainers = new Set([
        'vessels', 'specific vessels', 'documents', 'shared documents',
        'common for all ships', 'common', 'kaizen - knowledge bank', 'kaizen',
        'sharepoint sites', 'sites documents', 'nksdocman', 'site library', 'general documents',
        (row.vesselName || '').trim().toLowerCase(),
        (row.group || '').trim().toLowerCase(),
        'technical & crewing', 'commercial & chartering', 'insurance', 'knowledge bank'
      ]);
      // A "SharePoint Sites > <Site display name> > <Library> > …" row (the
      // flattened multi-site list) carries the site's own name as a
      // breadcrumb segment, never a real Group — same as
      // parseSharePointRowMetadataUncached's knownContainers above. Without
      // this, a file with no Group tag falls back to showing the site name
      // itself (e.g. "Communication Site") as its Group.
      (host.state.documentSites || []).forEach(s => {
        if (s.site_key) knownContainers.add(s.site_key.trim().toLowerCase());
        if (s.sp_site_name) knownContainers.add(s.sp_site_name.trim().toLowerCase());
      });

      const subLevels = parts.filter(p => !knownContainers.has(p.toLowerCase()));

      const section = subLevels[0] || (row.category && !knownContainers.has(row.category.toLowerCase()) ? row.category : '') || '';
      const group = subLevels[1] || (row.category && row.category !== section && !knownContainers.has(row.category.toLowerCase()) ? row.category : '') || '';
      const category = subLevels[2] || (row.subCategory && row.subCategory !== section && row.subCategory !== group ? row.subCategory : '') || '';
      const subCategory = subLevels[3] || '';

      return { section, group, category, subCategory };
    };

    type ListViewLabels = {
      vessel: string;
      mainFolder: string;
      documentSection: string;
      group: string;
      category: string;
      subCategory: string;
    };
    // Called several times per row per render (dropdown options, filtering,
    // grouping); memoize per row object for this render.
    const listViewLabelsCache = new WeakMap<object, ListViewLabels>();
    const getListViewLabels = (row: Pick<FlatRow, 'vesselName' | 'group' | 'category' | 'subCategory' | 'subFolderPath'> & { fileName?: string | null }): ListViewLabels => {
      const cachedLabels = listViewLabelsCache.get(row);
      if (cachedLabels) return cachedLabels;
      const computed = getListViewLabelsUncached(row);
      listViewLabelsCache.set(row, computed);
      return computed;
    };
    const getListViewLabelsUncached = (row: Pick<FlatRow, 'vesselName' | 'group' | 'category' | 'subCategory' | 'subFolderPath'> & { fileName?: string | null }): ListViewLabels => {
      const parsed = parseSharePointRowMetadata(
        row.subFolderPath || '',
        row.fileName,
        row.vesselName,
        row.group,
        row.category,
        row.subCategory
      );
      // When parsing leaves fields unassigned, fall back to direct row field values
      // via hierarchyForRow so filter dropdowns are not blank.
      const needsFallback =
        !parsed.documentSection ||
        !parsed.group ||
        !parsed.category ||
        !parsed.subCategory;
      if (needsFallback) {
        const hier = hierarchyForRow(row);
        return {
          vessel: parsed.vessel,
          mainFolder: parsed.mainFolder || '',
          documentSection: parsed.documentSection || hier.section || '',
          group: parsed.group || hier.group || '',
          category: parsed.category || hier.category || '',
          subCategory: parsed.subCategory || hier.subCategory || '',
        };
      }
      return parsed;
    };

    const allGroups = Array.from(new Set(
      scopeVesselRows.map(r => {
        const mf = getListViewLabels(r).mainFolder;
        if (mf && mf !== 'Main folder not assigned' && mf !== 'SharePoint Sites') return mf;
        return (r.group && r.group !== 'SharePoint Sites' && r.group !== 'Main folder not assigned') ? r.group : '';
      }).filter(Boolean)
    )).sort();
    if (allGroups.length === 0) {
      allGroups.push('Commercial & Chartering', 'Insurance', 'Kaizen - Knowledge Bank', 'Technical & Crewing');
    }
    const filteredCategoryRows = scopeVesselRows.filter(r => {
      if (docGroupFilter === 'all') return true;
      const mf = getListViewLabels(r).mainFolder;
      return r.group.trim().toLowerCase() === docGroupFilter.trim().toLowerCase() ||
             mf.trim().toLowerCase() === docGroupFilter.trim().toLowerCase();
    });

    const rowMetadataList = filteredCategoryRows.map(row => ({
      row,
      meta: getListViewLabels(row),
    }));

    // "Main folder" = every real top-level folder of the selected site's
    // Documents library, live (see liveMainFolderNames above) — never a
    // fixed list. liveMainFolderNames only covers a single resolvable site
    // (docScopeType 'documents'/'vessels'/'sites'-breadcrumb); the flattened
    // "Sites" list view spans every connected site at once, so there's no
    // one site/drive to fetch a root listing for and it stays empty there.
    // Union in whatever main-folder value each in-scope row already
    // resolved to (same fallback logic as parseSharePointRowMetadata above)
    // so the dropdown reflects real folders in every scope, not just the
    // single-site ones.
    const rowScannedMainFolderNames = Array.from(new Set(
      scopeVesselRows
        .map(r => getListViewLabels(r).mainFolder)
        .filter(mf => mf && mf !== 'SharePoint Sites' && mf !== 'Main folder not assigned')
    ));
    // getListViewLabels()/parseSharePointRowMetadata() canonicalize a row's
    // folder into the shared department name (e.g. "Technical & Crewing"),
    // not the site's literal top-level folder name. That's fine as a
    // fallback for the flattened multi-site "Sites" list view, which has no
    // single root to list live — but for a resolvable single site (this
    // scope, once liveMainFolderNames is populated from the real root
    // listing), unioning both sources added a phantom option: NKSDocMan's
    // real top-level folders are "Technical" and "Technical and Crewing
    // New", plus a canonical "Technical & Crewing" that matches neither
    // literal folder, so the dropdown showed 3 entries for 2 real folders.
    // Once live data is available, trust it exclusively.
    // For a resolved live library never fall back to the row-derived
    // canonical names ("Technical & Crewing"): while the live list is still
    // loading the dropdown shows just "All main folders" instead of a
    // folder that does not exist on the site.
    const mainFolderOptionsLive = liveMainFolderNames.length > 0
      ? liveMainFolderNames
      : (liveLibraryResolved ? [] : Array.from(new Set([...liveMainFolderNames, ...rowScannedMainFolderNames])));
    // While both sources are still loading (first render after a site
    // switch, or before any row has been scanned) fall back to the known
    // department set so the dropdown isn't empty for a moment.
    const mainFolderOptions = mainFolderOptionsLive.length > 0
      ? mainFolderOptionsLive.sort()
      : (liveLibraryResolved ? [] : Object.keys(KNOWN_MAIN_FOLDER_MAP));

    // "Sub-folder" = a real folder found *nested* one level under the
    // *currently selected/navigated* main folder — not any main folder in
    // scope. When a main folder is selected (via the "All main folders"
    // dropdown, docCategoryFilter — the actual UI control despite its name
    // — or by drilling into one in Folder view) this list is scoped
    // strictly to that one main folder, so it can never show another
    // folder's sub-folders (or the site name, or a bare file) alongside
    // them.
    //
    // A live top-level folder click (e.g. "Share Folder with Tsuneishi
    // China") always pushes docMainFolder as the literal marker
    // 'SharePoint Sites' (see the rootFolderCards click handler), not the
    // folder's own name, so that alone can't identify the folder — use the
    // breadcrumb's own last segment (currentFolderName) for that specific
    // navigation shape instead. docMainFolder itself is deliberately NOT
    // used as a fallback: it isn't reliably cleared when navigating back up
    // to the Documents root, so trusting it let a stale value from a
    // previous folder keep scoping this list after the user had backed all
    // the way out.
    const isTopLevelLiveFolderView = docViewMode === 'folder' && stackLevel === 4 &&
      folderPathStack[0]?.id === 'sites_root' &&
      (folderPathStack[1]?.id || '').startsWith('site:') &&
      (folderPathStack[2]?.id || '').startsWith('drive:');
    // Same "SharePoint Sites" live navigation shape as isTopLevelLiveFolderView,
    // but *deeper* than the main-folder tile view (already inside a vessel /
    // category / sub-folder). folderPathStack[3] is always the main folder in
    // this shape regardless of how many more levels have been drilled into
    // since, so it's still resolvable here — without this, scopedMainFolder-
    // ForSubfolders stayed null the moment the user navigated past the main
    // folder tiles (docCategoryFilter is only set by explicitly picking from
    // the dropdown, which normal folder-click navigation never does), which
    // in turn disabled the live recursive walk below and the "always keep the
    // current folder selectable" fallback — the Sub-folder dropdown then only
    // had whatever Source 3 (DB row scanning) happened to know about, so any
    // folder nested deeper than that (or not yet reflected in a DB row) never
    // appeared, however deep the real SharePoint nesting actually went.
    const isWithinLiveFolderView = docViewMode === 'folder' && stackLevel > 4 &&
      folderPathStack[0]?.id === 'sites_root' &&
      (folderPathStack[1]?.id || '').startsWith('site:') &&
      (folderPathStack[2]?.id || '').startsWith('drive:');
    const scopedMainFolderForSubfolders = docCategoryFilter !== 'all'
      ? docCategoryFilter
      : (isTopLevelLiveFolderView && currentFolderName
        ? currentFolderName
        : (isWithinLiveFolderView ? (folderPathStack[3]?.name || null) : null));
    // vesselFilter names one exact vessel (see the Vessel select above). When
    // it's set, the Sub-folder dropdown below must be scoped to that
    // vessel's own branch — otherwise every vessel folder nested anywhere
    // under the selected main folder gets walked and offered together, so
    // picking e.g. "Elephanta" left every other vessel's sub-folders in the
    // list too. Normalized (not exact-string) so minor punctuation/spacing
    // differences between the DB vessel name and the live folder name still
    // match.
    const normalizeVesselKey = (s: string): string => s.trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
    const activeVesselFilter = vesselFilter && vesselFilter !== 'all' && vesselFilter.trim().toLowerCase() !== 'not listed'
      ? normalizeVesselKey(vesselFilter)
      : null;
    // No confirmed main-folder scope → no sub-folder options at all, rather
    // than falling back to a global scan across every folder in view.
    // Row-derived sub-folders come ONLY from real folder segments that sit
    // *below* the selected main folder in a row's own path. The old version
    // pulled the row's group/category/documentSection labels instead, and
    // those labels fall back to things like the site name — which is how
    // "Communication Site" showed up as a sub-folder of a main folder that
    // contains no such folder.
    const allSiteNameSet = new Set(
      (host.state.documentSites || [])
        .reduce((acc, s) => acc.concat([s.sp_site_name, s.site_key, s.default_library_name]), [] as Array<string | undefined>)
        .concat(['shared documents', 'documents'])
        .filter(Boolean)
        .map(n => String(n).trim().toLowerCase())
    );
    const subfolderOptionsFromRows: string[] = [];
    // Parallel to subfolderOptionsFromRows, but keeps the REAL nesting chain
    // (every ancestor segment, in order) for each name instead of flattening
    // it away. subfolderOptionsFromRows itself has always thrown that shape
    // out — it exists only to answer "is this name a sub-folder at all?" —
    // which is fine for the plain-<select> era this fed, but the tree below
    // (rowOnlyOptions) used to reuse those bare names as ROOT nodes for any
    // folder the live recursive walk hadn't (yet, or ever, e.g. once its
    // 300-entry/15-per-level budget is spent) discovered. A folder several
    // levels deep — say Drawings and Manuals > To Be Classified > Basic >
    // "07 Flow sensor" — landed at the TOP of the tree, as if it were a
    // sibling of "Drawings and Manuals" itself, purely because this scan
    // knew its name but not where it actually lives. Recording the chain
    // here lets that fallback nest it correctly instead of guessing "root".
    const subfolderChainByName = new Map<string, string[]>();
    if (scopedMainFolderForSubfolders) {
      const mainLow = scopedMainFolderForSubfolders.trim().toLowerCase();
      filteredCategoryRows.forEach(r => {
        const segs = (r.subFolderPath || '').split('>').map(s => s.trim()).filter(Boolean);
        // "SharePoint Sites > <site> > <library> > <main folder> > …" —
        // start looking after the site/library prefix so a main folder that
        // happens to share a name with the site can't match there.
        const searchFrom = (segs[0] || '').toLowerCase() === 'sharepoint sites' ? 3 : 0;
        let mainIdx = -1;
        for (let i = searchFrom; i < segs.length; i++) {
          if (segs[i].toLowerCase() === mainLow) { mainIdx = i; break; }
        }
        if (mainIdx < 0) return;
        // With a vessel selected, only take segments *below* that vessel's
        // own segment (wherever it sits under the main folder) — a row not
        // under the selected vessel at all is skipped entirely, rather than
        // contributing its segments starting from the main folder.
        let startIdx = mainIdx + 1;
        if (activeVesselFilter) {
          let vesselIdx = -1;
          for (let i = startIdx; i < segs.length; i++) {
            if (normalizeVesselKey(segs[i]) === activeVesselFilter) { vesselIdx = i; break; }
          }
          if (vesselIdx < 0) return;
          startIdx = vesselIdx + 1;
        }
        const fileLow = (r.fileName || '').trim().toLowerCase();
        const tailSegs = segs.slice(startIdx).filter(seg => {
          const low = seg.toLowerCase();
          if (fileLow && low === fileLow) return false;
          if (/\.[a-z0-9]{2,5}$/i.test(seg)) return false;
          return true;
        });
        tailSegs.forEach((seg, i) => {
          subfolderOptionsFromRows.push(seg);
          const key = seg.trim().toLowerCase();
          const chain = tailSegs.slice(0, i + 1);
          // Prefer the longest/deepest chain seen for a given name (a name
          // reused at a shallower spot elsewhere shouldn't win and truncate
          // a real nested chain already recorded for it).
          const existing = subfolderChainByName.get(key);
          if (!existing || chain.length > existing.length) subfolderChainByName.set(key, chain);
        });
      });
    }
    const subfolderOptionsFromRowsFiltered = subfolderOptionsFromRows.filter(v => {
      const low = v.toLowerCase();
      if (low === 'sharepoint sites' || low === 'main folder not assigned') return false;
      if (allSiteNameSet.has(low)) return false;
      if (KNOWN_MAIN_FOLDER_NAME_SET.has(low) || KNOWN_CATEGORY_NAME_SET.has(low)) return false;
      if (mainFolderOptions.some(n => n.trim().toLowerCase() === low)) return false;
      return true;
    });
    // Row-scanning alone misses a real sub-folder whose files/rows haven't
    // been loaded into state yet — which is the normal state for a main
    // folder the user has only *selected from the dropdown*, not actually
    // opened (opening is what triggers the file/row fetch). Fetch that
    // folder's own children straight from SharePoint, the same live source
    // the tile grid itself uses, and union the two so the list is correct
    // whether or not anything has been browsed into yet.
    //
    // folderPathStack[3]'s id is only trustworthy when its NAME still
    // matches scopedMainFolderForSubfolders. A vessel that sits under more
    // than one main folder (the "is also under" chips / the Main folder
    // dropdown) sets docCategoryFilter to the newly picked folder
    // synchronously, but navigateToLiveMainFolder's own walk to that
    // folder's live id — resolving a not-yet-cached vessel path, or simply
    // the render that lands between the two setState calls — can leave
    // folderPathStack pointing at the *previous* main folder for a beat, or
    // (if navigateToLiveMainFolder's own rootFolderCards lookup ever misses,
    // e.g. a name/casing mismatch) indefinitely. Blindly trusting
    // folderPathStack[3]?.id in that window resolved the Sub-folder walk
    // against the OLD main folder's id, so the dropdown kept listing that
    // folder's sub-folders — appearing "stuck" on whichever folder was
    // selected first — even though the Main folder control had already
    // switched. Falling back to the rootFolderCards lookup (the same one
    // navigateToLiveMainFolder itself uses) whenever the names disagree
    // keeps the Sub-folder list in step with whichever main folder is
    // actually selected right now.
    const breadcrumbMainFolderNode = (isTopLevelLiveFolderView || isWithinLiveFolderView) ? folderPathStack[3] : null;
    const breadcrumbMainFolderMatches = !!breadcrumbMainFolderNode && !!scopedMainFolderForSubfolders &&
      breadcrumbMainFolderNode.name.trim().toLowerCase() === scopedMainFolderForSubfolders.trim().toLowerCase();
    const scopedMainFolderLiveFolderId = breadcrumbMainFolderMatches
      ? (breadcrumbMainFolderNode!.id || null)
      : (scopedMainFolderForSubfolders
        ? (rootFolderCards.find(c => c.key.trim().toLowerCase() === scopedMainFolderForSubfolders.trim().toLowerCase())?.liveFolderId || null)
        : (breadcrumbMainFolderNode?.id || null));
    // The first three breadcrumb entries ("SharePoint Sites" > site > drive)
    // that any live top-level-folder navigation is built on top of. Reused
    // both when a folder tile is clicked directly and when a Main
    // folder / Sub-folder is picked from the filter dropdowns instead, so
    // both paths land on the exact same navigation shape.
    const siteNavPrefix: { id: string; name: string }[] | null = isSitesScopeNav
      ? [folderPathStack[0], folderPathStack[1], folderPathStack[2]].filter(Boolean) as { id: string; name: string }[]
      : (effectiveLiveSiteId && effectiveLiveDriveId
        ? [
          { id: 'sites_root', name: 'SharePoint Sites' },
          { id: `site:${effectiveLiveSiteId}`, name: activeLiveSite?.sp_site_name || activeLiveSite?.site_key || 'Site' },
          { id: `drive:${effectiveLiveDriveId}`, name: activeLiveSite?.sp_site_name || 'Documents' },
        ]
        : null);
    // A "sub-folder" isn't just the main folder's direct children — a
    // folder like "SS366" can itself contain further folders ("SS366
    // FINISH PLAN"), and the filter needs to surface those too, at
    // whatever depth they're actually nested. Walk down live children
    // recursively (bounded, so one huge tree can't fan out into thousands
    // of API calls) and collect every folder found at any level under the
    // selected main folder, along with its own live folder id and the full
    // chain of ids/names from the main folder down to it — that chain is
    // what lets picking it from the dropdown actually navigate there.
    interface LiveSubfolderEntry {
      name: string;
      id: string;
      depth: number; // 1 = direct child of the main folder, 2+ = nested deeper
      pathIds: string[];
      pathNames: string[];
    }
    const liveSiteOwnNameSet = new Set(
      [activeLiveSite?.sp_site_name, activeLiveSite?.site_key, (host.state.documentSites || []).find(s => s.site_id === effectiveLiveSiteId)?.sp_site_name]
        .filter(Boolean)
        .map(n => (n as string).trim().toLowerCase())
    );
    // Total-entry budget across the whole walk, independent of depth — keeps
    // a wide-and-deep tree from ballooning the dropdown (and the Graph call
    // count) even though folders now nest "indefinitely" per level.
    const LIVE_SUBFOLDER_BUDGET = 300;
    const LIVE_SUBFOLDER_MAX_LOAD_DEPTH = 2;
    const collectLiveFolderEntriesRecursive = (
      folderId: string,
      depthRemaining: number,
      depth: number,
      pathIds: string[],
      pathNames: string[],
      out: LiveSubfolderEntry[],
    ): void => {
      if (depthRemaining <= 0 || !folderId || !effectiveLiveSiteId || !effectiveLiveDriveId || out.length >= LIVE_SUBFOLDER_BUDGET) return;
      // Fan-out guard: this walk runs during render, so every not-yet-cached
      // folder it touches becomes a Graph request. Only load new listings for
      // the first two levels; deeper levels use whatever is already cached
      // (filled in as the user browses). Without this, picking a main folder
      // launched up to ~300 requests and tripped 429 throttling.
      if (depth > LIVE_SUBFOLDER_MAX_LOAD_DEPTH &&
        !host._siteFolderItemsCache.has(`${effectiveLiveSiteId}::${effectiveLiveDriveId}::${folderId}`)) return;
      const children = host._getOrLoadSiteFolderChildren(effectiveLiveSiteId, effectiveLiveDriveId, folderId).items || [];
      // Dedup scoped to just THIS folder's own children listing (in case
      // Graph ever repeats an item across a page boundary) — NOT a Set
      // shared across the whole recursive walk. A shared/global Set here
      // meant a folder name reused anywhere else in the tree (e.g. every
      // vessel having its own "Drawings"/"Manuals"/"To Be Classified" —
      // completely normal in this DMS) got silently dropped the instant
      // that name was first seen under a different, often earlier-visited
      // vessel/branch — see the vessel-scoped walk below, which exists
      // specifically to work around this for the *selected* vessel; this
      // fixes it at the source so every branch of the tree collects its own
      // same-named folders correctly, not just the selected vessel's.
      const seenAtThisLevel = new Set<string>();
      children
        // `folder.folder` (the Graph folder facet) is the only reliable
        // positive folder signal — see the liveRootFolders note above for
        // why `is_folder !== false` was unsafe here (it let files through).
        .filter(f => f && !!f.folder && f.name)
        .slice(0, 15)
        .forEach(f => {
          if (out.length >= LIVE_SUBFOLDER_BUDGET) return;
          const name = (f.name || '').trim();
          const childPathIds = f.id ? [...pathIds, f.id] : pathIds;
          const childPathNames = [...pathNames, name];
          if (name) {
            const low = name.toLowerCase();
            if (low !== 'sharepoint sites' && low !== 'main folder not assigned' &&
              !KNOWN_MAIN_FOLDER_NAME_SET.has(low) && !KNOWN_CATEGORY_NAME_SET.has(low) &&
              !liveSiteOwnNameSet.has(low) && !seenAtThisLevel.has(low)) {
              seenAtThisLevel.add(low);
              out.push({ name, id: f.id, depth, pathIds: childPathIds, pathNames: childPathNames });
            }
          }
          if (f.id) collectLiveFolderEntriesRecursive(f.id, depthRemaining - 1, depth + 1, childPathIds, childPathNames, out);
        });
    };
    const liveSubfolderEntries: LiveSubfolderEntry[] = [];
    // FAST PATH: the whole folder-only tree of the main folder comes from ONE
    // cached backend call (/folder-tree) — no per-folder /children requests,
    // no 15-per-level slice, no 300-entry budget (which is what hid the
    // nested folders of every vessel past the first few). The old bounded
    // walk below is kept only as a fallback if that call fails.
    const folderTreeRes = (scopedMainFolderLiveFolderId && effectiveLiveSiteId && effectiveLiveDriveId)
      ? host._getOrLoadSiteFolderTree(effectiveLiveSiteId, effectiveLiveDriveId, scopedMainFolderLiveFolderId)
      : null;
    // Even a partial tree is used as-is (the host quietly re-fetches until it is
    // complete) — falling back to per-folder loading on top of a throttled
    // Graph only made things slower and less complete.
    const liveFolderTreeComplete = !!folderTreeRes && folderTreeRes.folders.length > 0;
    if (scopedMainFolderLiveFolderId && folderTreeRes && folderTreeRes.folders.length > 0) {
      const treeByParent = new Map<string, typeof folderTreeRes.folders>();
      folderTreeRes.folders.forEach(r => {
        const arr = treeByParent.get(r.parent_id);
        if (arr) arr.push(r); else treeByParent.set(r.parent_id, [r]);
      });
      const treeRootId = folderTreeRes.folders.find(r => r.depth === 1)?.parent_id;
      const emitTree = (parentId: string, depth: number, pIds: string[], pNames: string[]): void => {
        const seenAtLevel = new Set<string>();
        (treeByParent.get(parentId) || []).forEach(f => {
          const name = (f.name || '').trim();
          const cIds = [...pIds, f.id];
          const cNames = [...pNames, name];
          const low = name.toLowerCase();
          if (name && low !== 'sharepoint sites' && low !== 'main folder not assigned' &&
            !KNOWN_MAIN_FOLDER_NAME_SET.has(low) && !KNOWN_CATEGORY_NAME_SET.has(low) &&
            !liveSiteOwnNameSet.has(low) && !seenAtLevel.has(low)) {
            seenAtLevel.add(low);
            liveSubfolderEntries.push({ name, id: f.id, depth, pathIds: cIds, pathNames: cNames });
          }
          emitTree(f.id, depth + 1, cIds, cNames);
        });
      };
      if (treeRootId) emitTree(treeRootId, 1, [], []);
    } else if (scopedMainFolderLiveFolderId && !!folderTreeRes && folderTreeRes.error && !folderTreeRes.loading && Date.now() >= host._graphThrottledUntil) {
      // Fallback (tree call failed/empty): the old bounded walk. Depth bound
      // of 6 so one pathological tree can't fan out into hundreds of calls.
      collectLiveFolderEntriesRecursive(scopedMainFolderLiveFolderId, 6, 1, [], [], liveSubfolderEntries);
    }
    // The folder the user is currently inside (and every folder between it
    // and the main folder) must always be selectable, even when the bounded
    // walk above cut it off (15 children per level / 300 total). Otherwise
    // clicking into e.g. "ACRA CHARGE > NKS Canopus" sets the Sub-folder
    // dropdown to a value it has no <option> for, and it silently shows
    // "All sub-folders" while the grid is inside NKS Canopus.
    if (
      scopedMainFolderForSubfolders &&
      folderPathStack[0]?.id === 'sites_root' &&
      folderPathStack.length > 4 &&
      (folderPathStack[3]?.name || '').trim().toLowerCase() === scopedMainFolderForSubfolders.trim().toLowerCase()
    ) {
      const existing = new Set(liveSubfolderEntries.map(e => e.name.trim().toLowerCase()));
      const tail = folderPathStack.slice(4);
      tail.forEach((node, i) => {
        const low = (node.name || '').trim().toLowerCase();
        if (!low || existing.has(low)) return;
        existing.add(low);
        const upto = tail.slice(0, i + 1);
        liveSubfolderEntries.push({
          name: node.name.trim(),
          id: node.id,
          depth: i + 1,
          pathIds: upto.map(n => n.id),
          pathNames: upto.map(n => n.name),
        });
      });
    }
    // The folder the user is CURRENTLY BROWSING (rendered as live tiles with
    // real counts in the main content area, e.g. "Drawings"(7)/"Manuals"(15)/
    // "To Be Classified"(0) under "Drawings and Manuals") must have its own
    // children collected too, not just be selectable itself. The main-folder
    // walk above starts at the *main* folder (e.g. "Technical & Crewing")
    // and is depth-first with a 15-children-per-level slice and a 300-entry
    // total budget — with ~24 vessels sitting under most main folders, that
    // budget is very likely to be exhausted (or a whole vessel skipped by
    // the 15-per-level cap) long before the walk ever reaches deep into the
    // specific vessel/category branch actually being viewed. That's exactly
    // why folders visibly rendered as live tiles right here never showed up
    // in the dropdown. Unconditional — unlike the vessel-scoped walk below,
    // this does NOT require the Vessel filter dropdown to be set, since
    // simply clicking through "SharePoint Sites" breadcrumb tiles (the
    // normal navigation flow) never touches vesselFilter at all.
    if (
      !liveFolderTreeComplete &&
      isWithinLiveFolderView &&
      scopedMainFolderForSubfolders &&
      folderPathStack.length > 4
    ) {
      const tail = folderPathStack.slice(4);
      const currentTailNode = tail[tail.length - 1];
      if (currentTailNode?.id) {
        const currentFolderEntries: LiveSubfolderEntry[] = [];
        collectLiveFolderEntriesRecursive(
          currentTailNode.id, 4, tail.length + 1,
          tail.map(n => n.id), tail.map(n => n.name), currentFolderEntries,
        );
        const existingIds = new Set(liveSubfolderEntries.map(e => e.id));
        currentFolderEntries.forEach(e => {
          if (!existingIds.has(e.id)) {
            existingIds.add(e.id);
            liveSubfolderEntries.push(e);
          }
        });
      }
    }
    // When the Folder view is at/under the selected vessel's own folder
    // (e.g. right after picking the vessel from the Vessel dropdown), walk
    // that vessel folder's children directly. The main-folder walk above
    // only takes 15 children per level, so a vessel folder beyond the first
    // 15 under its main folder (e.g. #20 of "Technical"'s 23) never had its
    // sub-folders collected, leaving the Sub-folder dropdown empty.
    if (
      !liveFolderTreeComplete &&
      activeVesselFilter &&
      scopedMainFolderForSubfolders &&
      folderPathStack[0]?.id === 'sites_root' &&
      folderPathStack.length > 4 &&
      (folderPathStack[3]?.name || '').trim().toLowerCase() === scopedMainFolderForSubfolders.trim().toLowerCase()
    ) {
      const tail = folderPathStack.slice(4);
      const vesselIdxInTail = tail.findIndex(n => normalizeVesselKey(n.name || '') === activeVesselFilter);
      if (vesselIdxInTail >= 0 && tail[vesselIdxInTail].id) {
        const upto = tail.slice(0, vesselIdxInTail + 1);
        // collectLiveFolderEntriesRecursive's own dedup is now scoped per
        // folder (see its definition above), so this no longer needs its
        // own separate seen-set to avoid the selected vessel losing a
        // same-named folder to some other vessel the main walk reached
        // first — this walk exists for a different reason: the main walk
        // above only takes the first 15 children per level, so a vessel
        // folder beyond that cutoff (e.g. #20 of "Technical"'s 23) never
        // gets its own sub-folders collected there at all. Walking directly
        // from the selected vessel's own id guarantees its branch is always
        // covered regardless of where it sits in its main folder's listing.
        const vesselEntries: LiveSubfolderEntry[] = [];
        collectLiveFolderEntriesRecursive(
          tail[vesselIdxInTail].id, 5, vesselIdxInTail + 2,
          upto.map(n => n.id), upto.map(n => n.name), vesselEntries,
        );
        // Pushed AFTER the main walk's entries, so when liveSubfolderEntryByName
        // (below) is built, this vessel's own id for a shared name wins over
        // any other vessel's entry the main walk happened to collect first —
        // navigating the dropdown always resolves to *this* vessel's folder.
        liveSubfolderEntries.push(...vesselEntries);
      }
    }
    // The walk above collects every folder under the *main* folder, across
    // every vessel in it — narrow that down to the selected vessel's own
    // branch before it reaches the dropdown. An entry qualifies when the
    // vessel's name appears somewhere in its own path chain (pathNames)
    // *before* its last segment; that last check excludes the vessel's own
    // folder entry itself (already covered by the Vessel filter, not a
    // "sub-folder" of it). Depth is re-based to the vessel so its direct
    // children start at depth 1 again, matching the indentation the options
    // below expect.
    const vesselIndexInPath = (e: LiveSubfolderEntry): number =>
      activeVesselFilter ? e.pathNames.findIndex(n => normalizeVesselKey(n) === activeVesselFilter) : -1;
    const scopedLiveSubfolderEntries = activeVesselFilter
      ? liveSubfolderEntries.filter(e => {
        const idx = vesselIndexInPath(e);
        return idx >= 0 && idx < e.pathNames.length - 1;
      })
      : liveSubfolderEntries;
    // Kept unscoped (every vessel's entries, by name) so navigation lookups
    // — including from a Sub-folder value chosen before the Vessel filter
    // changed — can still resolve a folder id; only the *displayed* options
    // and their depths above are vessel-scoped.
    const liveSubfolderEntryByName = new Map(liveSubfolderEntries.map(e => [e.name.trim().toLowerCase(), e]));
    // Preserve the live entries' own hierarchy (DFS) order so nesting reads
    // top-to-bottom in the dropdown; row-scanned names with no live match
    // (folders discovered only because a file under them was already loaded)
    // have no known depth, so they're appended as top-level after the live
    // ones, alphabetically.
    const rowOnlyOptions = subfolderOptionsFromRowsFiltered
      .filter(name => !liveSubfolderEntryByName.has(name.trim().toLowerCase()))
      .filter((name, idx, arr) => arr.findIndex(n => n.trim().toLowerCase() === name.trim().toLowerCase()) === idx)
      .sort();
    // Nested tree for the Sub-folder filter (FolderTreeSelect). Three
    // sources are merged into one name-keyed trie, because no single source
    // covers every navigation flow / folder state on its own:
    //   1. scopedLiveSubfolderEntries — the "SharePoint Sites" live recursive
    //      walk (see above); each entry already carries its own full path.
    //   2. host.state.spoFolderMap — the live folder graph the classic
    //      Department > Vessel flow already builds and reads for its own
    //      folder-tile counts (see subfolderFolderCountMap below). This is
    //      what makes a genuinely empty container folder (children, but no
    //      files of its own — e.g. "Drawings and Manuals" with 0 files but 3
    //      sub-folders) show up at all: source 3 below can only ever find a
    //      folder that has a *file* somewhere underneath it.
    //   3. Row / uploaded-folder path scanning — a safety net for a folder
    //      that hasn't been walked live yet in this session but does have
    //      files under it, so the tree isn't empty on a cold load.
    const subfolderTree: FolderTreeNode[] = (() => {
      interface MutableFolderNode { name: string; id: string; children: Map<string, MutableFolderNode>; }
      const roots = new Map<string, MutableFolderNode>();
      const ensureChild = (parent: Map<string, MutableFolderNode>, name: string, id?: string): MutableFolderNode => {
        const key = name.trim().toLowerCase();
        let node = parent.get(key);
        if (!node) {
          node = { name: name.trim(), id: id || name.trim(), children: new Map() };
          parent.set(key, node);
        } else if (id && node.id === node.name) {
          // Upgrade a synthetic (name-based) id to a real live one once a
          // later source learns it, so navigation can use whichever source
          // discovered this folder first.
          node.id = id;
        }
        return node;
      };
      const toFolderTreeNodes = (m: Map<string, MutableFolderNode>): FolderTreeNode[] =>
        Array.from(m.values()).map(n => ({ name: n.name, id: n.id, children: toFolderTreeNodes(n.children) }));
      // Hoisted out of the Source 3 block below so the rowOnlyOptions
      // fallback (end of this IIFE) can also use it to nest a name at its
      // real known depth instead of dropping it at the root.
      const insertPath = (segs: string[]): void => {
        let level = roots;
        segs.filter(isDisplayableFolderName).forEach(seg => {
          const node = ensureChild(level, seg);
          level = node.children;
        });
      };

      // Source 1: SharePoint-Sites live walk.
      scopedLiveSubfolderEntries.forEach(e => {
        const idx = vesselIndexInPath(e);
        const relNames = idx >= 0 ? e.pathNames.slice(idx + 1) : e.pathNames;
        const relIds = idx >= 0 ? e.pathIds.slice(idx + 1) : e.pathIds;
        let level = roots;
        for (let i = 0; i < relNames.length; i++) {
          const node = ensureChild(level, relNames[i], relIds[i]);
          level = node.children;
        }
      });

      // Source 2: Department > Vessel flow's own live folder graph.
      // Not used in the Sites flow — source 1 above already covers it, from
      // its own (better-scoped) recursive Graph walk.
      const deptVesselLiveId: string | null = (() => {
        if (atSitesRoot) return null;
        const raw = vesselNodeInStack?.id;
        if (raw && !/^(sf_|category_|common|kaizen_root|dept_|vessels_root|specific_vessels)/.test(raw)) return raw;
        if (vesselName && docMainFolder) {
          return host._getLiveSharePointFolderId([docMainFolder, vesselName].join(' > ')) || null;
        }
        return null;
      })();
      if (deptVesselLiveId) {
        const MAX_DEPT_TREE_NODES = 500;
        let nodeCount = 0;
        const walkSpoFolderMap = (liveId: string, target: Map<string, MutableFolderNode>, depthLeft: number): void => {
          if (depthLeft <= 0 || nodeCount >= MAX_DEPT_TREE_NODES) return;
          const liveNode = host.state.spoFolderMap.get(liveId);
          (liveNode?.children || []).forEach(child => {
            if (!child.isFolder || !child.name) return;
            if (child.id && host._appDeletedItemIds.has(child.id)) return;
            if (nodeCount >= MAX_DEPT_TREE_NODES) return;
            nodeCount++;
            const node = ensureChild(target, child.name, child.id);
            if (child.id) walkSpoFolderMap(child.id, node.children, depthLeft - 1);
          });
        };
        walkSpoFolderMap(deptVesselLiveId, roots, 6);
      }

      // Source 3: row / uploaded-folder path scanning, full depth (not just
      // one level relative to wherever the user currently is), scoped to
      // this vessel + main folder.
      if (vesselName && docMainFolder) {
        (rows || []).forEach(r => {
          if ((r.vesselName || '').trim().toLowerCase() !== vesselName.trim().toLowerCase()) return;
          if ((r.group || '').trim().toLowerCase() !== docMainFolder.trim().toLowerCase()) return;
          insertPath(getFolderTailSegments(r.subFolderPath, vesselName, docMainFolder));
        });
        Object.keys(uploadedFilesByFolder || {}).forEach(key => {
          if (host._appDeletedItemIds.has(key)) return;
          if (key.includes('||')) {
            const parts = key.split('||');
            const kVessel = (parts[0] || '').trim().toLowerCase();
            const kGroup = (parts[1] || '').trim().toLowerCase();
            if (kVessel && kVessel !== vesselName.trim().toLowerCase()) return;
            if (kGroup && kGroup !== docMainFolder.trim().toLowerCase()) return;
          } else {
            const otherVessel = vessels.find(v => v.name && key.toLowerCase().includes(v.name.toLowerCase()));
            if (otherVessel && otherVessel.name.toLowerCase() !== vesselName.trim().toLowerCase()) return;
          }
          insertPath(getFolderTailSegments(key, vesselName, docMainFolder));
        });
      }

      // rowOnlyOptions: names the Sites-flow's own row-scan found with no
      // CONFIRMED LIVE nesting (i.e. the live walk above hasn't — or, once
      // its budget/depth caps are hit, simply can't — reach them this
      // session). That never meant they're actually top-level: a folder
      // like "07 Flow sensor" living at Drawings and Manuals > To Be
      // Classified > Basic > "07 Flow sensor" used to land here and get
      // added as a bare root (ensureChild(roots, name) with no ancestors),
      // which is what made it render as a sibling of "Drawings and Manuals"
      // instead of nested under it. subfolderChainByName (recorded during
      // the same row scan that found the name) carries that real ancestor
      // chain, so re-use it here via insertPath — same as Source 3 — and
      // only fall back to a bare root when no chain was ever recorded for
      // it (e.g. it surfaced from something other than a row path).
      rowOnlyOptions.forEach(name => {
        if (roots.has(name.trim().toLowerCase())) return;
        const chain = subfolderChainByName.get(name.trim().toLowerCase());
        if (chain && chain.length > 0) insertPath(chain);
        else ensureChild(roots, name);
      });

      return toFolderTreeNodes(roots);
    })();
    // Selecting a Main folder / Sub-folder from the dropdown must not just
    // filter the grid — it must navigate the Folder view to that folder's
    // real contents, exactly as clicking its tile/breadcrumb would.
    const navigateToLiveMainFolder = (folderKey: string, carry?: CarrySubfolder | null): void => {
      const card = rootFolderCards.find(c => c.key.trim().toLowerCase() === folderKey.trim().toLowerCase());
      if (card && card.liveFolderId && siteNavPrefix) {
        // A vessel already selected (it has a folder under this main folder
        // too — the same vessel commonly sits under several main folders)
        // should stay selected and land back inside *its own* folder here,
        // not just at this main folder's root. Landing at the root has no
        // vessel segment in the path, so deriveLiveNavFilterState/liveNavSync
        // below would otherwise clear the Vessel dropdown back to "All
        // vessels" — which is what made switching Main folder look like it
        // silently dropped/broke the vessel selection instead of just
        // showing that vessel's documents under the newly picked main folder.
        if (vesselFilter && vesselFilter !== 'all' && vesselFilter.trim().toLowerCase() !== 'not listed') {
          const vesselPaths = siteVesselFolderPaths.get(vesselFilter.trim().toLowerCase()) || [];
          // Take the SHORTEST recorded path under the new main folder (the
          // vessel's own "<main>/<vessel>" folder), same rule as
          // vesselBasePathUnderMain. The recorded list can also hold deeper
          // nested discoveries ("<main>/<vessel>/Drawings and Manuals"), and
          // whichever came first after a refresh used to win via .find(),
          // dropping the user inside a sub-folder instead of the vessel root.
          // Single-segment paths (the vessel folder sitting at the library
          // root) are not a "home" under this main folder, so skip them.
          const vesselPathUnderNewMain = vesselPaths
            .filter(p => p.includes('/') &&
              (p.split('/')[0] || '').trim().toLowerCase() === folderKey.trim().toLowerCase())
            .sort((a, b) => a.length - b.length)[0];
          if (vesselPathUnderNewMain) {
            const nodes = resolveCachedDriveFolderStack(vesselPathUnderNewMain);
            // Carrying a Sub-folder pick (e.g. Drawings > Hull) from the same
            // vessel under another main folder: open it here too.
            if (nodes && carry) {
              openVesselWithCarry(vesselFilter, nodes, carry);
              return;
            }
            if (nodes) {
              const vesselStack = [...siteNavPrefix, ...nodes];
              host._pushFolderNav(vesselStack, 'SharePoint Sites');
              host.setState({
                ...deriveDocFiltersFromStack(vesselStack),
                docSubfolderOtherFilter: 'all',
                vesselFilter,
                docListPage: 0,
              });
              return;
            }
            // resolveCachedDriveFolderStack only reads what's already in
            // _siteFolderItemsCache and gives up (returning null) the
            // moment a level hasn't been fetched yet — which, for a main
            // folder the user hasn't opened at all this session (exactly
            // the "is also under" chip case, since that chip is populated
            // from a full tree scan, not from this per-level cache), is
            // every level below its root. That's what made clicking the
            // chip strand the user on this main folder's bare root with no
            // further attempt to actually reach the vessel's folder there.
            // Land on the root now so the click still does *something*
            // visible, then keep resolving the real path level-by-level
            // (awaiting each Graph round trip instead of bailing on the
            // first one) and finish the jump the moment it comes back.
            const rootStack = [...siteNavPrefix, { id: card.liveFolderId, name: card.key }];
            host._pushFolderNav(rootStack, 'SharePoint Sites');
            if (liveNavSync.derive) host.setState({ vesselFilter: liveNavSync.derive(rootStack).vesselFilter });
            const targetVessel = vesselFilter;
            resolveLiveDriveFolderStackAsync(vesselPathUnderNewMain).then(resolvedNodes => {
              if (!resolvedNodes) {
                // Vessel folder not reachable here: don't leave a carried
                // Sub-folder pinned on the main folder's root.
                if (carry && subfolderStillIs(carry) && host.state.vesselFilter === targetVessel) {
                  (host as any)._subfolderSelectedPath = undefined;
                  host.setState({ docSubfolderOtherFilter: 'all', docListPage: 0 });
                }
                return;
              }
              // Only complete the jump if the user hasn't since navigated
              // away from the root this redirected them to, or changed the
              // vessel selection — otherwise the fetch's result is stale
              // and finishing the jump would yank them somewhere they
              // didn't ask to go.
              const current = host.state.folderPathStack || [];
              const stillAtRedirectedRoot = current.length === rootStack.length &&
                current[current.length - 1]?.id === rootStack[rootStack.length - 1]?.id &&
                host.state.vesselFilter === targetVessel;
              if (!stillAtRedirectedRoot) return;
              if (carry) {
                openVesselWithCarry(targetVessel, resolvedNodes, carry);
                return;
              }
              const vesselStack = [...siteNavPrefix, ...resolvedNodes];
              host._pushFolderNav(vesselStack, 'SharePoint Sites');
              host.setState({
                ...deriveDocFiltersFromStack(vesselStack),
                docSubfolderOtherFilter: 'all',
                vesselFilter: targetVessel,
                docListPage: 0,
              });
            }).catch(() => undefined);
            return;
          }
        }
        const stack = [...siteNavPrefix, { id: card.liveFolderId, name: card.key }];
        host._pushFolderNav(stack, 'SharePoint Sites');
        if (liveNavSync.derive) host.setState({ vesselFilter: liveNavSync.derive(stack).vesselFilter });
        if (carry) {
          // No vessel folder under this main folder to carry the pick into.
          (host as any)._subfolderSelectedPath = undefined;
          host.setState({ docSubfolderOtherFilter: 'all', docListPage: 0 });
        }
      }
    };
    const navigateToLiveSubfolder = (subfolderName: string, namePath?: string[], nodeId?: string): void => {
      // A bare name is ambiguous ("Safety" exists under both Drawings and
      // Manuals) — when the caller knows the full tree path, resolve by that
      // first so the right branch is opened.
      const pathKey = namePath && namePath.length ? namePath.map(n => n.trim().toLowerCase()).join('/') : '';
      const lazyMap = (host as any)._lazySubfolderEntries as Map<string, LiveSubfolderEntry> | undefined;
      // With a vessel selected the dropdown tree is scoped to that vessel, so
      // its paths are RELATIVE to the vessel (no vessel segment) while
      // liveSubfolderEntries now holds every vessel's folders. Resolve inside
      // the selected vessel first — otherwise a shared name/path such as
      // Drawings > Basic matched whichever vessel came first in the list
      // (Belle Lune), silently jumping to the wrong vessel.
      const lowPath = (names: string[]): string => names.map(n => n.trim().toLowerCase()).join('/');
      const inSelectedVessel = activeVesselFilter
        ? liveSubfolderEntries.filter(e => vesselIndexInPath(e) >= 0)
        : [];
      // Folder ids are unique across vessels, so when the dropdown hands us the
      // clicked node's own id, that is the one unambiguous way to find it —
      // names/paths repeat under every vessel (Drawings > Basic ...).
      const byNodeId = nodeId
        ? (liveSubfolderEntries.find(e => e.id === nodeId)
          || (lazyMap ? Array.from(lazyMap.values()).find(e => e.id === nodeId) : undefined)
          // A folder that was filtered out of the entry list itself (e.g. the
          // "Drawings"/"Manuals" ancestors) still appears in every descendant's
          // id chain — rebuild its entry from that chain.
          || (() => {
            for (const e of liveSubfolderEntries) {
              const idx = e.pathIds.indexOf(nodeId);
              if (idx >= 0) {
                return { name: e.pathNames[idx], id: nodeId, depth: idx + 1, pathIds: e.pathIds.slice(0, idx + 1), pathNames: e.pathNames.slice(0, idx + 1) } as LiveSubfolderEntry;
              }
            }
            return undefined;
          })())
        : undefined;
      const entry = byNodeId || (pathKey && (
          liveSubfolderEntries.find(e => lowPath(e.pathNames) === pathKey)
          || inSelectedVessel.find(e => lowPath(e.pathNames.slice(vesselIndexInPath(e) + 1)) === pathKey)
          || lazyMap?.get('path:' + pathKey)))
        || (namePath && namePath.length >= 2 ? (activeVesselFilter ? inSelectedVessel : liveSubfolderEntries).find(e => {
              const en = e.pathNames.map(n => n.trim().toLowerCase());
              const np = namePath.map(n => n.trim().toLowerCase());
              const k = Math.min(en.length, np.length, 2);
              return k >= 2 && en.slice(-k).join('/') === np.slice(-k).join('/');
            }) : undefined)
        || (activeVesselFilter ? inSelectedVessel.find(e => e.name.trim().toLowerCase() === subfolderName.trim().toLowerCase()) : undefined)
        || liveSubfolderEntryByName.get(subfolderName.trim().toLowerCase())
        || lazyMap?.get(subfolderName.trim().toLowerCase());
      if (entry && siteNavPrefix && scopedMainFolderLiveFolderId && scopedMainFolderForSubfolders) {
        const mainFolderNode = { id: scopedMainFolderLiveFolderId, name: scopedMainFolderForSubfolders };
        const extraStack = entry.pathIds.map((id, i) => ({ id, name: entry.pathNames[i] }));
        const stack = [...siteNavPrefix, mainFolderNode, ...extraStack];
        host._pushFolderNav(stack, 'SharePoint Sites');
        if (liveNavSync.derive) host.setState({ vesselFilter: liveNavSync.derive(stack).vesselFilter });
      }
    };
    // Counterpart to navigateToLiveSubfolder for the classic Department >
    // Vessel flow (not "SharePoint Sites"), used when picking a node from
    // the FolderTreeSelect tree that came from subfolderTree's source 2/3
    // (spoFolderMap / row-scanning) rather than the Sites live walk, which
    // has no entry in liveSubfolderEntryByName to navigate from. Takes the
    // node's full name path (root-to-leaf) since the same folder name can
    // exist at more than one depth.
    const navigateToDeptSubfolder = (pathSegs: string[]): void => {
      if (atSitesRoot || !vesselName || !docMainFolder || pathSegs.length === 0) return;
      const mainNode = folderPathStack[0];
      const mainFolderNodeId = (mainNode && mainNode.name.trim().toLowerCase() === docMainFolder.trim().toLowerCase())
        ? mainNode.id
        : docMainFolder;
      const vesselRaw = vesselNodeInStack?.id;
      const vesselNodeId = (vesselRaw && !/^(sf_|category_|common|kaizen_root|dept_|vessels_root|specific_vessels)/.test(vesselRaw))
        ? vesselRaw
        : (host._getLiveSharePointFolderId([docMainFolder, vesselName].join(' > ')) || vesselName);
      const newStack: { id: string; name: string }[] = [
        { id: mainFolderNodeId, name: docMainFolder },
        { id: vesselNodeId, name: vesselName },
      ];
      pathSegs.forEach((seg, i) => {
        const breadcrumb = [docMainFolder, vesselName, ...pathSegs.slice(0, i + 1)].join(' > ');
        newStack.push({ id: host._getLiveSharePointFolderId(breadcrumb) || breadcrumb, name: seg });
      });
      host._pushFolderNav(newStack, docMainFolder);
      const leaf = newStack[newStack.length - 1];
      void host._refreshFolderFiles(leaf.id, [docMainFolder, vesselName, ...pathSegs].join(' > '), true).catch(() => undefined);
    };
    // Jumping via a breadcrumb segment lands the Folder view on a specific
    // point in the tree, but the Main folder / Sub-folder dropdowns only
    // ever get set by their own onChange handlers above — a breadcrumb click
    // bypasses those entirely, so the dropdowns kept showing whatever was
    // selected before the jump even though the grid had moved on. Derive the
    // correct dropdown values from the breadcrumb stack being navigated to,
    // so every breadcrumb click re-syncs them the same way picking that same
    // folder from the dropdown would have.
    const deriveDocFiltersFromStack = (stack: { id: string; name: string }[]): { docCategoryFilter: string; docSubfolderOtherFilter: string } => {
      // Live "SharePoint Sites" shape: [sites_root, site:, drive:, mainFolder, ...subfolders]
      if (stack[0]?.id === 'sites_root') {
        if (stack.length <= 3) return { docCategoryFilter: 'all', docSubfolderOtherFilter: 'all' };
        if (stack.length === 4) return { docCategoryFilter: stack[3].name, docSubfolderOtherFilter: 'all' };
        return { docCategoryFilter: stack[3].name, docSubfolderOtherFilter: stack[stack.length - 1].name };
      }
      // Live "Shared Documents" / "Documents" library shape (renderLibraryBrowser):
      // [libRoot, mainFolder, ...subfolders] — the same kind of live SharePoint
      // folder browsing as the "SharePoint Sites" shape above, just without the
      // site/drive prefix, since this is a single fixed configured library rather
      // than a user-picked one. A vessel folder commonly sits under more than one
      // top-level main folder here too (e.g. "Technical" and "Technical and
      // Crewing New" both containing "Bow Fighter"), so this needs the same
      // per-navigation re-sync: without it, the Main folder / Sub-folder dropdowns
      // kept whatever value they were last set to (from an earlier dropdown pick,
      // possibly under a completely different main folder) even after the user
      // had clicked folder tiles into a different main folder entirely — which is
      // what let the Sub-folder dropdown keep resolving against the wrong
      // (earlier-selected) main folder's live id while the grid itself had
      // already moved to the newly opened one.
      if (stack[0]?.id === 'lib:shared_documents' || stack[0]?.id === 'lib:documents') {
        if (stack.length <= 1) return { docCategoryFilter: 'all', docSubfolderOtherFilter: 'all' };
        if (stack.length === 2) return { docCategoryFilter: stack[1].name, docSubfolderOtherFilter: 'all' };
        return { docCategoryFilter: stack[1].name, docSubfolderOtherFilter: stack[stack.length - 1].name };
      }
      // Any other breadcrumb shape (department / vessel / Kaizen drill-down)
      // doesn't use the Main folder / Sub-folder dropdowns — those only ever
      // apply to the live SharePoint-folder scope — so landing there via
      // breadcrumb should leave them cleared, same as the existing idx===0 /
      // idx===1 handlers below already do explicitly.
      return { docCategoryFilter: 'all', docSubfolderOtherFilter: 'all' };
    };

    const drawingCategories = ['Hull', 'Basic', 'Electrical', 'Machinery', 'Safety'];
    const manualCategories = ['Automation', 'Auxiliary Engine', 'Boiler', 'Bridge Equipments', 'Cargo', 'Deck Machinery', 'Electrical', 'Main Engine', 'Other Manuals', 'Pollution', 'Propulsion', 'Refrigeration', 'Safety', 'Shafting', 'Steering Gear'];
    // groupOptions / categoryOptions / subCategoryOptions are derived from the
    // rows that actually exist (see right after `filtered` below).

    // ── Group / Category for the List view columns and their filters ──
    // Real folder names only *partly* spell these out, e.g.
    //   "SS367-MAERSK FERRATO-Ship drawings, Plans, Manuals" > … > "ELECT"
    // so both are matched on partial words, never on exact folder names:
    //   • Group: a folder whose name contains the word "drawing(s)"/"dwg"
    //     or "manual(s)". The DEEPEST such folder decides, so
    //     "Drawings and Manuals > Drawings > Hull" is Drawings only; a folder
    //     naming both ("Ship drawings, Plans, Manuals") with nothing more
    //     specific below it counts as both.
    //   • Category: any folder in the path whose words cover the category's
    //     words, where a folder word may be an abbreviation (≥4 letters) of
    //     the category word or vice versa: "ELECT" → Electrical,
    //     "MACH" → Machinery, "BRIDGE EQUIP" → Bridge Equipments,
    //     "Aux Engine" → Auxiliary Engine ("AUX" is 3 letters, so only when
    //     the remaining words also match).
    // File names are never matched — only the folders a file sits in.
    const GC_STOPWORDS = new Set(['and', 'of', 'the', 'for', 'a']);
    const gcWords = (value: string): string[] => (value || '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .split(' ')
      .filter(w => w && !GC_STOPWORDS.has(w));
    const gcWordMatches = (categoryWord: string, folderWord: string, allowShort: boolean): boolean => {
      if (categoryWord === folderWord) return true;
      const minLen = allowShort ? 3 : 4;
      if (folderWord.length >= minLen && categoryWord.indexOf(folderWord) === 0) return true; // ELECT → electrical
      if (categoryWord.length >= 4 && folderWord.indexOf(categoryWord) === 0) return true;    // electricals → electrical
      return false;
    };
    const folderMatchesCategory = (folderName: string, category: string): boolean => {
      const catWords = gcWords(category);
      const folderWords = gcWords(folderName);
      if (catWords.length === 0 || folderWords.length === 0) return false;
      // Multi-word categories may use a 3-letter abbreviation ("Aux") since
      // the other word(s) must still match; single words need ≥4 letters.
      const allowShort = catWords.length > 1;
      return catWords.every(cw => folderWords.some(fw => gcWordMatches(cw, fw, allowShort)));
    };
    const TBC_GROUP_LABEL = 'To Be Classified';
    const isToBeClassifiedName = (name: string): boolean => gcWords(name).join(' ') === 'to be classified';
    const folderGroupsOf = (folderName: string): { drawings: boolean; manuals: boolean } => {
      const words = gcWords(folderName);
      return {
        drawings: words.some(w => /^(drawings?|dwgs?)$/.test(w)),
        manuals: words.some(w => /^manuals?$/.test(w)),
      };
    };
    const pathFolderSegments = (subFolderPath: string): string[] => (subFolderPath || '')
      .split(/\s*>\s*|\//)
      .map(seg => seg.trim())
      .filter(seg => !!seg && !/\.[a-z0-9]{2,5}$/i.test(seg));
    const decisiveGroupOf = (segments: string[]): { idx: number; drawings: boolean; manuals: boolean } | null => {
      for (let i = segments.length - 1; i >= 0; i--) {
        const g = folderGroupsOf(segments[i]);
        if (g.drawings || g.manuals) return { idx: i, drawings: g.drawings, manuals: g.manuals };
      }
      return null;
    };
    /** True when the folder path satisfies the Group and/or Category filter. */
    const pathMatchesGroupCategory = (subFolderPath: string, group: string, category: string): boolean => {
      const segments = pathFolderSegments(subFolderPath);
      if (group !== 'all') {
        const want = group.trim().toLowerCase();
        if (want === TBC_GROUP_LABEL.toLowerCase()) {
          // "To Be Classified" is a folder-based group: any file sitting in a
          // "To Be Classified" folder (at any depth) belongs to it.
          if (!segments.some(isToBeClassifiedName)) return false;
        } else {
          // Files filed under a "To Be Classified" folder belong only to the
          // "To Be Classified" group, never to Drawings / Manuals.
          if (segments.some(isToBeClassifiedName)) return false;
          const g = decisiveGroupOf(segments);
          if (!g) return false;
          if (want === 'drawings' && !g.drawings) return false;
          if (want === 'manuals' && !g.manuals) return false;
        }
      }
      if (category !== 'all') {
        if (!segments.some(seg => folderMatchesCategory(seg, category))) return false;
      }
      return true;
    };
    // Sub-category picked from a classified entry ("<Group>::<Category>::<Sub>",
    // e.g. To Be Classified › Yard Drawing › Basic): the Group › Category it
    // was listed under. Ignored when it contradicts an explicitly chosen
    // Group / Category or is no longer the active Sub-category.
    const subCategoryPickCtx = (): { group: string; category: string } | null => {
      if (docSubCategoryFilter === 'all') return null;
      const raw = (host as any)._subCategoryPickKey as string | undefined;
      if (!raw) return null;
      const parts = raw.split('::');
      if (parts.length !== 3 || !parts[0] || !parts[1]) return null;
      const low = (v: string): string => (v || '').trim().toLowerCase();
      if (low(parts[2]) !== low(docSubCategoryFilter)) return null;
      if (docGroupLevelFilter !== 'all' && low(docGroupLevelFilter) !== low(parts[0])) return null;
      if (docLeafCategoryFilter !== 'all' && low(docLeafCategoryFilter) !== low(parts[1])) return null;
      return { group: parts[0], category: parts[1] };
    };
    /**
     * Does the folder path contain the selected Sub-category folder?
     * With a classified pick the folder must sit exactly where the dropdown
     * found it (same rule libTaxonomy is built with): directly under the
     * picked Category, inside the picked Group — so "To Be Classified ›
     * Yard Drawing › Basic" no longer also matches "Drawings › Basic".
     * Without one, any folder of that name matches (previous behaviour).
     */
    const pathMatchesSubCategory = (folderPath: string | string[], sub: string): boolean => {
      const low = (v: string): string => (v || '').trim().toLowerCase();
      const segments = Array.isArray(folderPath)
        ? folderPath.map(x => (x || '').trim()).filter(Boolean)
        : pathFolderSegments(folderPath);
      const want = low(sub);
      const ctx = subCategoryPickCtx();
      if (!ctx || low(ctx.group) === '') return segments.some(seg => low(seg) === want);
      const grp = low(ctx.group);
      if (grp === TBC_GROUP_LABEL.toLowerCase()) {
        for (let i = 0; i + 2 < segments.length; i++) {
          if (isToBeClassifiedName(segments[i]) && low(segments[i + 1]) === low(ctx.category) && low(segments[i + 2]) === want) return true;
        }
        return false;
      }
      if (segments.some(isToBeClassifiedName)) return false;
      const catMatches = (seg: string): boolean => low(seg) === low(ctx.category) || folderMatchesCategory(seg, ctx.category);
      let inGroup = false;
      for (let i = 0; i + 1 < segments.length; i++) {
        const g = folderGroupsOf(segments[i]);
        if ((grp === 'drawings' && g.drawings) || (grp === 'manuals' && g.manuals)) { inGroup = true; continue; }
        if (inGroup && catMatches(segments[i]) && low(segments[i + 1]) === want) return true;
      }
      return false;
    };
    const drawingsManualsCache = new WeakMap<object, { group: string | null; category: string | null }>();
    const drawingsManualsForRow = (row: Pick<FlatRow, 'subFolderPath'>): { group: string | null; category: string | null } => {
      const cached = drawingsManualsCache.get(row);
      if (cached) return cached;
      const segments = pathFolderSegments(row.subFolderPath || '');
      let result: { group: string | null; category: string | null } = { group: null, category: null };
      const g = decisiveGroupOf(segments);
      if (g) {
        const group = g.drawings && g.manuals ? 'Drawings / Manuals' : (g.drawings ? 'Drawings' : 'Manuals');
        const known = g.drawings && g.manuals
          ? drawingCategories.concat(manualCategories)
          : (g.drawings ? drawingCategories : manualCategories);
        // Category: first folder below the group folder that is (an
        // abbreviation of) a known category, shown by its canonical name;
        // otherwise the folder directly under the group folder, as before.
        let category: string | null = null;
        for (let i = g.idx + 1; i < segments.length && !category; i++) {
          const hit = known.find(c => folderMatchesCategory(segments[i], c));
          if (hit) category = hit;
        }
        if (!category && segments[g.idx + 1]) category = segments[g.idx + 1];
        result = { group, category };
      }
      drawingsManualsCache.set(row, result);
      return result;
    };

    // Sub-folder tree <-> "All categories" dropdown sync. These are two
    // different pickers over what is often the same underlying folder (the
    // tree shows live/scanned sub-folders; the flat dropdown shows the
    // fixed Drawings/Manuals category list) — selecting one should update
    // the other whenever they clearly refer to the same folder, in both
    // directions, so the two never show contradictory selections.
    const findCategoryForFolderName = (folderName: string): string | null => {
      if (!folderName || folderName === 'all') return null;
      const norm = folderName.trim().toLowerCase();
      const exact = categoryOptions.find(c => c.trim().toLowerCase() === norm);
      if (exact) return exact;
      return categoryOptions.find(c => folderMatchesCategory(folderName, c)) || null;
    };
    /** Does a tree path sit inside the given Group / Category selection? */
    const pathFitsGroupContext = (path: string[], ctx?: { group?: string; category?: string }): boolean => {
      if (!ctx) return true;
      const group = (ctx.group || 'all').trim().toLowerCase();
      if (group !== 'all') {
        if (group === TBC_GROUP_LABEL.toLowerCase()) {
          if (!path.some(isToBeClassifiedName)) return false;
        } else {
          if (path.some(isToBeClassifiedName)) return false;
          const wantDrawings = group === 'drawings';
          if (!path.some(seg => { const g = folderGroupsOf(seg); return wantDrawings ? g.drawings : g.manuals; })) return false;
        }
      }
      if (ctx.category && ctx.category !== 'all') {
        if (!path.some(seg => folderMatchesCategory(seg, ctx.category as string))) return false;
      }
      return true;
    };
    const findSubfolderNodeForCategory = (category: string, ctx?: { group?: string; category?: string }): { name: string; path: string[] } | null => {
      if (!category || category === 'all') return null;
      const norm = category.trim().toLowerCase();
      const isTbcQuery = isToBeClassifiedName(category);
      let exactHit: { name: string; path: string[] } | null = null;
      let fuzzyHit: { name: string; path: string[] } | null = null;
      const walk = (nodes: FolderTreeNode[], trail: string[]): void => {
        for (const node of nodes) {
          const ownPath = [...trail, node.name];
          const fits = pathFitsGroupContext(ownPath, ctx);
          if (fits && !exactHit && (node.name.trim().toLowerCase() === norm || (isTbcQuery && isToBeClassifiedName(node.name)))) exactHit = { name: node.name, path: ownPath };
          if (fits && !fuzzyHit && folderMatchesCategory(node.name, category)) fuzzyHit = { name: node.name, path: ownPath };
          if (node.children.length) walk(node.children, ownPath);
        }
      };
      walk(subfolderTree, []);
      return exactHit || fuzzyHit;
    };

    // Vessels the folder tree itself names: a top-level main folder or a
    // (recursively discovered) sub-folder whose name matches a recognized
    // vessel — e.g. a vessel folder sitting flat at the site root, or one
    // nested a few levels down — must be selectable from the Vessel filter
    // too, not just vessels that already have DB rows. "Recognized" means
    // matched against the Term Store / DB vessel set already used elsewhere
    // to classify folders (liveTermStoreVesselSet), not a name guess, so an
    // ordinary folder like "Report" or "Insurance" can't get misread as a
    // vessel. Only the sub-folder tree of the *currently scoped* main folder
    // is walked (liveSubfolderEntries, computed above for the Sub-folder
    // dropdown) — walking every main folder's full subtree on every render
    // would be far too expensive — so this widens as the user browses/picks
    // different main folders, in both Folder view and List view alike.
    const vesselDisplayNameByLower = new Map<string, string>();
    liveTermStoreVesselNames.forEach(n => { if (n) vesselDisplayNameByLower.set(n.trim().toLowerCase(), n); });
    (vessels || []).forEach(v => { if (v.name && !vesselDisplayNameByLower.has(v.name.trim().toLowerCase())) vesselDisplayNameByLower.set(v.name.trim().toLowerCase(), v.name); });
    const folderDetectedVesselNames = Array.from(new Set(
      [...liveMainFolderNames, ...liveSubfolderEntries.map(e => e.name)]
        .map(name => vesselDisplayNameByLower.get(name.trim().toLowerCase()))
        .filter((name): name is string => Boolean(name))
    ));

    const distinctVesselsInScope = Array.from(new Set(
      scopeRows.map(r => getListViewLabels(r).vessel).filter(Boolean).concat(folderDetectedVesselNames)
    )).sort((a, b) => {
      if (a === 'Not Listed') return 1;
      if (b === 'Not Listed') return -1;
      return a.localeCompare(b);
    });

    const vesselFilterOptions = (mainFolderPage && docMainFolder
      ? vessels.filter(v => scopeRows.some(r =>
          r.vesselName.trim().toLowerCase() === v.name.trim().toLowerCase() &&
          r.group.trim().toLowerCase() === docMainFolder.trim().toLowerCase()
        ))
      : vessels
    ).concat(
      // Same minimal shape as KNOWN_SPO_VESSELS below — only id/name are read
      // where this list is rendered (the Vessel <select>'s options), but the
      // full VesselRecord shape (imported from ../types/rows, not available
      // to read from here) is matched as closely as that existing pattern
      // does, with `as unknown as VesselRecord` covering any other field
      // that type declares beyond what's known here.
      folderDetectedVesselNames
        .filter(name => !vessels.some(v => v.name.trim().toLowerCase() === name.trim().toLowerCase()))
        .map(name => ({ id: name, name, status: 'Active', is_provisioned: true } as unknown as VesselRecord))
    );
    // ── Site-specific vessel detection (Vessel filter, folder cards, rows) ──
    // A vessel belongs to the site being viewed only when a real folder on
    // that site is named after it. Candidate vessel names are this site's
    // Term Store vessels plus every vessel created from this app. A folder
    // name matches a vessel when, after normalising:
    //   • exactly — "Bow Fighter", "MV Bow Fighter", "bow-fighter"
    //   • with a leading number stripped — "01 Bow Fighter", "123-Bow Fighter"
    //   • partially, on whole words — "Bow Fighter Drawings" (names ≥5 chars;
    //     the longest vessel name wins, so "Maersk EI Banco" beats "Maersk").
    // Substring matching on file names (what parseSharePointRowMetadata does)
    // is NOT used here — that is how vessels with no folder on the site, e.g.
    // "Dutches Emerald", leaked into the dropdown.
    const normVesselText = (value: string): string => (value || '')
      .toLowerCase()
      .replace(/^(m\/v|mv|mt)[\s.]+/, '')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
    const normFolderText = (value: string): string => normVesselText(
      (value || '').replace(/^[\s\d._\-#()[\]]+/, '')
    );
    const vesselCandidateByKey = new Map<string, string>();
    const addVesselCandidate = (name: string | undefined | null): void => {
      const key = normVesselText(name || '');
      if (key && !vesselCandidateByKey.has(key)) vesselCandidateByKey.set(key, (name || '').trim());
    };
    liveTermStoreVesselNames.forEach(addVesselCandidate);
    (vessels || []).forEach(v => addVesselCandidate(v.name));
    const partialVesselKeys = Array.from(vesselCandidateByKey.keys())
      .filter(key => key.length >= 5)
      .sort((x, y) => y.length - x.length);
    // Typo tolerance for folder names that misspell a vessel, e.g. folder
    // "022024 Duchess Emeralad" for vessel "Duchess Emerald". Deliberately
    // narrow so different vessels can't collide ("mvtest3" vs "mvtest4"):
    //   • only multi-word vessel names;
    //   • every word must line up with a folder word, in order;
    //   • at least one word must match exactly;
    //   • a word may differ by 1 edit (2 for words of 9+ letters), and only
    //     when both words are ≥5 letters and contain no digits.
    const withinEditDistance = (a: string, b: string, max: number): boolean => {
      if (a === b) return true;
      if (Math.abs(a.length - b.length) > max) return false;
      let prev: number[] = [];
      for (let j = 0; j <= b.length; j++) prev.push(j);
      for (let i = 1; i <= a.length; i++) {
        const cur: number[] = [i];
        let rowMin = i;
        for (let j = 1; j <= b.length; j++) {
          const cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
          const val = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
          cur.push(val);
          if (val < rowMin) rowMin = val;
        }
        if (rowMin > max) return false;
        prev = cur;
      }
      return prev[b.length] <= max;
    };
    const vesselWordsClose = (folderWord: string, vesselWord: string): boolean => {
      if (folderWord === vesselWord) return true;
      if (folderWord.length < 5 || vesselWord.length < 5) return false;
      if (/\d/.test(folderWord) || /\d/.test(vesselWord)) return false;
      const max = Math.max(folderWord.length, vesselWord.length) >= 9 ? 2 : 1;
      return withinEditDistance(folderWord, vesselWord, max);
    };
    // Only the START of the folder's (already leading-digit/punct-stripped)
    // word sequence is checked against the vessel's words — a sliding window
    // over the WHOLE folder name used to let a folder that merely mentions
    // the vessel somewhere in a longer descriptive title match too, e.g.
    // "Ship's Copy (USB) for BOW FIGHTER at FUKUOKA SY" (a document title
    // filed under an unrelated "Share with Mashin Shokai Singapore" folder)
    // matched vessel "Bow Fighter" purely because those two words happened
    // to appear back-to-back later in the sentence. A real (possibly typo'd
    // or suffixed, e.g. "Duchess Emeralad" or "Bow Fighter - Handover")
    // vessel folder has the vessel's name leading, not buried mid-sentence.
    const fuzzyVesselKeyMatches = (folderKey: string, vesselKey: string): boolean => {
      const vesselWords = vesselKey.split(' ').filter(Boolean);
      const folderWords = folderKey.split(' ').filter(Boolean);
      if (vesselWords.length < 2 || folderWords.length < vesselWords.length) return false;
      let exactCount = 0;
      for (let k = 0; k < vesselWords.length; k++) {
        const fw = folderWords[k];
        const vw = vesselWords[k];
        if (fw === vw) { exactCount++; continue; }
        if (!vesselWordsClose(fw, vw)) return false;
      }
      return exactCount > 0;
    };
    // Kept across renders (reset only when the candidate vessel set changes):
    // the folder lists feeding this now include the library's folder tree,
    // and re-matching every folder name on every render was wasted work.
    const vesselCandidateSig = Array.from(vesselCandidateByKey.entries()).map(([k, v]) => `${k}=${v}`).join('|');
    const persistedVesselMatch = (host as any)._folderVesselMatchMemo as { sig: string; map: Map<string, string | null> } | undefined;
    const folderVesselMatchCache: Map<string, string | null> = persistedVesselMatch && persistedVesselMatch.sig === vesselCandidateSig
      ? persistedVesselMatch.map
      : new Map<string, string | null>();
    (host as any)._folderVesselMatchMemo = { sig: vesselCandidateSig, map: folderVesselMatchCache };
    const matchFolderToVessel = (folderName: string): string | null => {
      const raw = (folderName || '').trim();
      if (!raw) return null;
      if (folderVesselMatchCache.has(raw)) return folderVesselMatchCache.get(raw)!;
      let result: string | null = null;
      const exactKey = normVesselText(raw);
      const strippedKey = normFolderText(raw);
      if (vesselCandidateByKey.has(exactKey)) {
        result = vesselCandidateByKey.get(exactKey)!;
      } else if (strippedKey && vesselCandidateByKey.has(strippedKey)) {
        result = vesselCandidateByKey.get(strippedKey)!;
      } else if (strippedKey) {
        // Require the vessel key to LEAD the (leading-digit/punct-stripped)
        // folder text — not merely appear anywhere inside it. `indexOf`
        // against the whole padded string used to also match a vessel name
        // sitting in the middle or at the end of an unrelated, longer folder
        // title (e.g. "Ship's Copy (USB) for BOW FIGHTER at FUKUOKA SY" ->
        // matched "Bow Fighter" even though that folder isn't the vessel's
        // own folder at all, just a document filed elsewhere that mentions
        // it). A real vessel folder — "Bow Fighter", "022024 Duchess
        // Emeralad", "Bow Fighter - Handover" — has the vessel name leading
        // once the numeric/punctuation prefix is stripped; startsWith keeps
        // that case working while rejecting a mid-sentence mention.
        const partial = partialVesselKeys.find(key => strippedKey === key || strippedKey.startsWith(`${key} `));
        const fuzzy = partial ? undefined : partialVesselKeys.find(key => fuzzyVesselKeyMatches(strippedKey, key));
        const hit = partial || fuzzy;
        result = hit ? vesselCandidateByKey.get(hit)! : null;
      }
      folderVesselMatchCache.set(raw, result);
      return result;
    };
    const vesselNamesEqual = (x: string, y: string): boolean =>
      normVesselText(x) === normVesselText(y);

    // Every folder we already know exists on the viewed site/drive, as
    // drive-relative paths ("Technical & Crewing/Bow Fighter"). Sources, all
    // scoped to effectiveLiveSiteId/effectiveLiveDriveId: the drive root
    // listing, every folder listing already fetched into
    // _siteFolderItemsCache (Folder view browsing and the List view subtree
    // prefetch), and the live tree when it belongs to this site.
    const knownSiteFolderPaths: string[] = [];
    {
      const seenPaths = new Set<string>();
      const addFolderPath = (path: string): void => {
        const clean = (path || '').replace(/\\/g, '/').replace(/^\/+|\/+$/g, '');
        const low = clean.toLowerCase();
        if (!clean || seenPaths.has(low)) return;
        seenPaths.add(low);
        knownSiteFolderPaths.push(clean);
      };
      liveRootFolders.forEach(f => addFolderPath(f.key));
      // Vessel folders sit BELOW the main folders (Technical/Bow Fighter,
      // ACRA CHARGE/NKS/<vessel>), so at the library root none of them had
      // been listed yet and the Vessel filter came up empty until a main
      // folder was opened. (The folder tiles used to list each main folder's
      // children as a side effect; they now use /subfolder-counts instead.)
      //   • each main folder's own listing — one cached call per main folder,
      //     quick, finds vessels one level down;
      //   • the library's folder tree 3 levels deep — one call, served from
      //     the backend's drive folder index, finds vessels nested deeper.
      if (effectiveLiveSiteId && effectiveLiveDriveId &&
        (docScopeType === 'sites' || docScopeType === 'shared_docs' || docScopeType === 'documents')) {
        if (Date.now() >= host._graphThrottledUntil) {
          liveRootFolders.slice(0, 40).forEach(card => {
            if (card.liveFolderId) host._getOrLoadSiteFolderChildren(effectiveLiveSiteId, effectiveLiveDriveId, card.liveFolderId);
          });
        }
        (shallowRootTree?.folders || []).forEach(f => { if (f && f.path) addFolderPath(f.path); });
      }
      if (effectiveLiveSiteId && effectiveLiveDriveId) {
        const cachePrefix = `${effectiveLiveSiteId}::${effectiveLiveDriveId}::`;
        host._siteFolderItemsCache.forEach((entry, key) => {
          if (!key.startsWith(cachePrefix) || !entry || !Array.isArray(entry.items)) return;
          const parentPath = String((entry as any).parentPath || '').replace(/^\/+|\/+$/g, '');
          entry.items.forEach((item: any) => {
            if (!item?.name || !item.folder) return;
            addFolderPath(parentPath ? `${parentPath}/${item.name}` : item.name);
          });
        });
      }
      if (!isSitesScopeNav || sitesScopeMatchedSite?.site_key === activeLiveSite?.site_key) {
        (host.state.documentLiveFolders || []).forEach(folder => {
          if (folder.is_folder !== false && folder.path) addFolderPath(folder.path);
        });
      }
    }
    // vessel display name (lower-cased) → drive-relative paths of its folders.
    const siteVesselFolderPaths = new Map<string, string[]>();
    const siteVesselByLower = new Map<string, string>();
    const recordSiteVesselFolder = (vesselName: string, path: string): void => {
      const low = vesselName.trim().toLowerCase();
      if (!siteVesselByLower.has(low)) siteVesselByLower.set(low, vesselName.trim());
      const paths = siteVesselFolderPaths.get(low) || [];
      if (paths.indexOf(path) < 0) paths.push(path);
      siteVesselFolderPaths.set(low, paths);
    };
    // A deeper folder whose OWN name merely *mentions* a vessel (e.g. a
    // "Certificate (BOW FIGHTER)" folder several levels inside that vessel's
    // own tree, under "Drawings and Manuals/To Be Classified/…") is not a
    // second home for the vessel — matchFolderToVessel's partial/fuzzy rules
    // exist to tolerate typo'd vessel FOLDER names ("022024 Duchess
    // Emeralad"), not to classify every folder that happens to reference the
    // vessel in its title. Recording those too made "select vessel" land
    // several levels too deep (straight into To Be Classified > Certificate
    // (BOW FIGHTER) instead of the vessel's own folder) whenever that deep
    // path happened to resolve/get tried before the real one. Skip a match
    // when the vessel's name already appears as an EARLIER segment of the
    // same path — that's proof this is a sub-folder nested under the
    // vessel's real folder, not an alternate location for it.
    const isNestedUnderSameVessel = (ancestorSegments: string[], vesselName: string): boolean =>
      ancestorSegments.some(seg => {
        const segVessel = matchFolderToVessel(seg);
        return !!segVessel && vesselNamesEqual(segVessel, vesselName);
      });
    knownSiteFolderPaths.forEach(path => {
      const segments = path.split('/');
      const vesselName = matchFolderToVessel(segments[segments.length - 1]);
      if (!vesselName) return;
      if (isNestedUnderSameVessel(segments.slice(0, -1), vesselName)) return;
      recordSiteVesselFolder(vesselName, path);
    });
    // Folder names seen only through the Main folder / Sub-folder walk.
    liveSubfolderEntries.forEach(entry => {
      const vesselName = matchFolderToVessel(entry.name);
      if (!vesselName) return;
      if (isNestedUnderSameVessel(entry.pathNames.slice(0, -1), vesselName)) return;
      recordSiteVesselFolder(vesselName, entry.pathNames.length ? entry.pathNames.join('/') : entry.name);
    });

    // Port of backend site_alias_matches (backend/app/config.py:74-104).
    const SITE_TOKENS = ['external', 'nissenkaiunexternal', 'nksdocman', 'docman', 'nks', 'local', 'communication site', 'communication', 'dev', 'root', 'default', 'vessel dms'];
    const ROOT_ALIASES = ['dev', 'communication', 'communication site', 'root', 'default', 'vessel dms'];
    const canonicalSiteToken = (value: string): string => {
      const text = value.trim().toLowerCase();
      for (const token of SITE_TOKENS) {
        if (text.indexOf(token) >= 0) return token;
      }
      return text;
    };
    const siteAliasSet = (value: string): Set<string> => {
      const norm = canonicalSiteToken(value);
      const aliases = new Set<string>([norm]);
      const addAll = (list: string[]): void => list.forEach(item => aliases.add(item));
      if (ROOT_ALIASES.indexOf(norm) >= 0) addAll(ROOT_ALIASES);
      if (norm.indexOf('communication') >= 0 || norm === 'dev' || norm === 'root') addAll(['dev', 'communication', 'communication site', 'root']);
      if (norm.indexOf('docman') >= 0 || norm.indexOf('nks') >= 0 || norm === 'local') addAll(['local', 'nksdocman', 'docman', 'nks']);
      if (norm.indexOf('external') >= 0) addAll(['external', 'nissenkaiunexternal']);
      return aliases;
    };
    const siteAliasesMatch = (left: string, right: string): boolean => {
      if (!left.trim() || !right.trim()) return false;
      if (canonicalSiteToken(left) === canonicalSiteToken(right)) return true;
      const leftAliases = siteAliasSet(left);
      return Array.from(siteAliasSet(right)).some(alias => leftAliases.has(alias));
    };
    const viewedSite = sitesScopeMatchedSite || activeLiveSite;
    const viewedSiteKeys = [viewedSite?.site_key, viewedSite?.sp_site_name]
      .map(value => String(value || '').trim())
      .filter(Boolean);
    const viewedSiteIds = [viewedSite?.site_id, sitesScopeRawSiteId]
      .map(value => String(value || '').trim().toLowerCase())
      .filter(Boolean);
    const isVesselProvisionedToViewedSite = (v: VesselRecord): boolean =>
      [...(v.provisioned_site_ids || []), v.provisioned_site_key]
        .map(id => String(id || '').trim())
        .filter(Boolean)
        .some(id =>
          viewedSiteIds.indexOf(id.toLowerCase()) >= 0 ||
          viewedSiteKeys.some(key => siteAliasesMatch(id, key))
        );

    // App-created vessels provisioned to this site whose folder hasn't been
    // seen yet (e.g. nested below folders nobody has opened): confirm the
    // recorded vessel_folder_path really exists by walking it through the
    // cached folder listings, loading each parent level on demand (cached,
    // one request per level). Only a confirmed folder adds the vessel — a
    // vessel whose metadata says "this site" but has no folder here (stale
    // provisioned_site_ids, deleted folder) is left out.
    const resolveDriveFolderPath = (path: string): boolean | undefined => {
      if (!effectiveLiveSiteId || !effectiveLiveDriveId) return undefined;
      const segments = path.replace(/\\/g, '/').split('/').map(seg => seg.trim()).filter(Boolean);
      if (segments.length === 0) return undefined;
      let parentId = 'root';
      for (const segment of segments) {
        const listing = host._getOrLoadSiteFolderChildren(effectiveLiveSiteId, effectiveLiveDriveId, parentId);
        if (listing.loading) return undefined;
        const hit = (listing.items || []).find((item: any) =>
          item?.folder && String(item.name || '').trim().toLowerCase() === segment.toLowerCase());
        if (!hit) return false;
        parentId = hit.id;
      }
      return true;
    };
    if (docScopeType === 'sites') {
      (vessels || []).forEach(v => {
        if (!v.name || siteVesselByLower.has(v.name.trim().toLowerCase())) return;
        if (!v.vessel_folder_path || !isVesselProvisionedToViewedSite(v)) return;
        if (resolveDriveFolderPath(v.vessel_folder_path) === true) {
          recordSiteVesselFolder(v.name, v.vessel_folder_path.replace(/\\/g, '/').replace(/^\/+|\/+$/g, ''));
        }
      });
    }
    const siteVesselNames = Array.from(siteVesselByLower.values()).sort((x, y) => x.localeCompare(y));
    // True when `pathLike` (drive-relative) is the selected vessel's folder,
    // inside it, or one of its ancestor folders.
    const selectedSiteVesselPaths = vesselFilter !== 'all'
      ? (siteVesselFolderPaths.get(vesselFilter.trim().toLowerCase()) || []).map(p => p.toLowerCase())
      : [];
    const isPathOnSelectedVesselBranch = (pathLike: string): boolean => {
      const low = pathLike.replace(/\\/g, '/').replace(/^\/+|\/+$/g, '').toLowerCase();
      if (!low) return false;
      return selectedSiteVesselPaths.some(vp => vp === low || vp.startsWith(`${low}/`) || low.startsWith(`${vp}/`));
    };
    // siteVesselNames only ever contains vessels whose folder has already
    // been *confirmed* on this site — either a folder already seen in a
    // cache we've populated this session (knownSiteFolderPaths /
    // liveSubfolderEntries, both scoped to wherever the user has actually
    // browsed), or one resolved on demand above via resolveDriveFolderPath,
    // which itself needs every parent level of that vessel's folder already
    // cached. Landing several levels deep via a direct jump (e.g. picking a
    // vessel from the dropdown, or a saved/shared deep link) never loads the
    // sibling folders at the main-folder root, so at a breadcrumb like
    // Technical > Bow Fighter > Drawings and Manuals, siteVesselNames only
    // ever has "Bow Fighter" — every other vessel on the fleet is filtered
    // out of the dropdown for no reason other than its folder not having
    // been fetched yet in *this* browsing session. The Vessel filter should
    // always offer every vessel in the fleet, not just the ones the cache
    // happens to already know about — folder-confirmed names are kept as
    // the display spelling when there's a mismatch, DB vessels fill in the
    // rest. Picking a not-yet-confirmed one still works: navigateToSiteVesselFolder
    // already has a fallback path for a vessel whose live folder hasn't been
    // walked yet.
    const allSiteVesselNamesByLower = new Map<string, string>();
    // Only offer DB vessels that belong to the viewed site. A vessel that is
    // provisioned to a different site (e.g. a Communication-site vessel like
    // Piessy) must not show up while browsing NKSDocMan / NissenKaiunExternal.
    // Vessels with no provisioning metadata at all are still offered, and
    // folder-confirmed names (siteVesselNames, below) are always kept.
    const vesselBelongsToViewedSite = (v: VesselRecord): boolean => {
      if (docScopeType !== 'sites' || viewedSiteIds.length === 0) return true;
      const provisioned = [...(v.provisioned_site_ids || []), v.provisioned_site_key]
        .map(id => String(id || '').trim())
        .filter(Boolean);
      if (provisioned.length === 0) return true;
      return isVesselProvisionedToViewedSite(v);
    };
    (vessels || []).forEach(v => {
      if (v.name && v.name.trim() && vesselBelongsToViewedSite(v)) allSiteVesselNamesByLower.set(v.name.trim().toLowerCase(), v.name.trim());
    });
    siteVesselNames.forEach(name => allSiteVesselNamesByLower.set(name.trim().toLowerCase(), name));
    const siteVesselOptions = Array.from(allSiteVesselNamesByLower.values())
      .sort((a, b) => a.localeCompare(b))
      .map(name => ({ id: name, name }));
    // Shared Documents / Documents libraries: vessels from loaded rows plus
    // every vessel whose folder was found in the library (siteVesselNames) —
    // rows alone are empty until something has been opened, which left the
    // Vessel filter with nothing but "All vessels" on first load.
    const libraryVesselOptions: string[] = (() => {
      const byKey = new Map<string, string>();
      distinctVesselsInScope.concat(siteVesselNames).forEach(n => {
        const key = normVesselText(n || '');
        if (key && !byKey.has(key)) byKey.set(key, n);
      });
      return Array.from(byKey.values()).sort((a, b) => {
        if (a === 'Not Listed') return 1;
        if (b === 'Not Listed') return -1;
        return a.localeCompare(b);
      });
    })();

    // Every main folder the *selected* vessel has a known folder under —
    // shared by the "also under" chip strip, the Compare toggle, and the
    // Compare grid itself, so all three agree on the same set/order.
    const vesselMainFolders = vesselFilter !== 'all' && vesselFilter.trim().toLowerCase() !== 'not listed'
      ? Array.from(new Set(
        (siteVesselFolderPaths.get(vesselFilter.trim().toLowerCase()) || [])
          .map(p => (p.split('/')[0] || '').trim())
          // A single-segment recorded path ("Bow Fighter", no "/") means the
          // vessel folder itself was matched as if it were a top-level main
          // folder — a real main folder is never just the vessel's own
          // name, so that's the vessel's folder sitting at the library
          // root, not a second home under some other main folder. Drop it
          // rather than showing a "Bow Fighter" chip alongside the real
          // main folders.
          .filter(name => Boolean(name) && name.trim().toLowerCase() !== vesselFilter.trim().toLowerCase())
      ))
      : [];
    // The vessel's own base folder path under a given main folder (the
    // longest recorded path starting with "<mainFolder>/" — vessels are
    // occasionally recorded via a deeper nested discovery too, but the
    // shortest "<mainFolder>/<vessel>" form is what Compare mode should
    // treat as that box's root).
    const vesselBasePathUnderMain = (mainFolder: string): string | null => {
      const paths = (siteVesselFolderPaths.get(vesselFilter.trim().toLowerCase()) || [])
        .filter(p => (p.split('/')[0] || '').trim().toLowerCase() === mainFolder.trim().toLowerCase());
      if (paths.length === 0) return null;
      return paths.slice().sort((a, b) => a.length - b.length)[0];
    };

    // ── Keep the filter bar in step with Folder view navigation ──
    // Every way of moving through the live SharePoint tree (clicking a folder
    // tile, a breadcrumb, Back/Forward, or picking from the Main folder /
    // Sub-folder dropdowns) goes through deriveLiveNavFilterState so the
    // dropdowns always describe the folder actually on screen:
    //   • Main folder / Sub-folder ← the breadcrumb (deriveDocFiltersFromStack)
    //   • Vessel ← the deepest folder in the path that is a vessel folder
    //     (typo-tolerant, see matchFolderToVessel); if there is none and the
    //     new location is off the selected vessel's branch, the vessel filter
    //     is cleared instead of leaving a stale "No folder for X here".
    type LiveNavFilterState = {
      docCategoryFilter: string;
      docSubfolderOtherFilter: string;
      vesselFilter: string;
      docListPage: number;
      docGroupLevelFilter: string;
    };
    const deriveLiveNavFilterState = (stack: { id: string; name: string }[]): LiveNavFilterState => {
      const base = deriveDocFiltersFromStack(stack);
      const currentVessel = host.state.vesselFilter || 'all';
      // Navigating into (or under) a "To Be Classified" folder selects that
      // Group; navigating out of it releases an auto-selected one.
      const currentGroupLevel = host.state.docGroupLevelFilter || 'all';
      const navGroupLevel = stack.slice(stack[0]?.id === 'sites_root' ? 3 : 1).some(n => isToBeClassifiedName(n.name))
        ? TBC_GROUP_LABEL
        : (currentGroupLevel === TBC_GROUP_LABEL ? 'all' : currentGroupLevel);
      if (stack[0]?.id !== 'sites_root') {
        return { ...base, vesselFilter: currentVessel, docListPage: 0, docGroupLevelFilter: navGroupLevel };
      }
      const folderNames = stack.slice(3).map(n => n.name);
      let vesselInPath: string | null = null;
      for (const name of folderNames) {
        const matched = matchFolderToVessel(name);
        if (matched) vesselInPath = matched;
      }
      let nextVessel = currentVessel;
      if (vesselInPath) {
        nextVessel = vesselInPath;
      } else if (currentVessel !== 'all' && currentVessel.trim().toLowerCase() !== 'not listed') {
        const path = folderNames.join('/');
        const low = path.toLowerCase();
        if (!low) {
          // folderNames is empty exactly when the navigation landed at the
          // drive/library root itself — e.g. clicking the "Documents"
          // breadcrumb segment. `!low` used to short-circuit onBranch to
          // true here, which kept whichever vessel was selected before
          // pinned in the dropdown (and kept the folder list scoped to just
          // that vessel's folders) even though the root is above every
          // vessel's folder, not "still on" any one of them. Landing at the
          // root must clear the vessel filter instead.
          nextVessel = 'all';
        } else {
          const vesselPaths = (siteVesselFolderPaths.get(currentVessel.trim().toLowerCase()) || []).map(p => p.toLowerCase());
          const onBranch = vesselPaths.some(vp => vp === low || vp.startsWith(`${low}/`) || low.startsWith(`${vp}/`));
          if (!onBranch) nextVessel = 'all';
        }
      }
      return { ...base, vesselFilter: nextVessel, docListPage: 0, docGroupLevelFilter: navGroupLevel };
    };
    liveNavSync.derive = deriveLiveNavFilterState;
    const pushLiveFolderNav = (stack: { id: string; name: string }[]): void => {
      host._pushFolderNav(stack, 'SharePoint Sites');
      host.setState(deriveLiveNavFilterState(stack));
    };
    // Resolve a drive-relative folder path ("ACRA CHARGE/NKS/022024 Duchess
    // Emeralad") to breadcrumb nodes. Goes through _getOrLoadSiteFolderChildren
    // rather than reading _siteFolderItemsCache directly: a level the user
    // has never opened (e.g. the main folder above a vessel picked from the
    // Vessel dropdown at the library root) isn't in the cache yet, but
    // _getOrLoadSiteFolderChildren derives it synchronously from the
    // already-loaded live tree (documentLiveFolders). Returns null only if a
    // level genuinely still needs a Graph round-trip.
    // Resolve a drive-relative path to breadcrumb nodes from the folder-index
    // trees already in the browser (root tree 3 levels deep + each main
    // folder's full tree) — no /children call, so opening a vessel works even
    // while folder listings are slow, throttled or failed.
    const resolveStackFromTrees = (segments: string[]): { id: string; name: string }[] | null => {
      if (!effectiveLiveSiteId || !effectiveLiveDriveId || segments.length === 0) return null;
      const prefix = `${effectiveLiveSiteId}::${effectiveLiveDriveId}::`;
      const rootTree = host._siteFolderTreeCache.get(`${prefix}root::d3`);
      const rootId = rootTree?.folders.find(f => f.depth === 1)?.parent_id;
      if (!rootId) return null;
      const childrenOf = new Map<string, { id: string; name: string }[]>();
      host._siteFolderTreeCache.forEach((t, k) => {
        if (!k.startsWith(prefix) || !t || !t.folders || t.folders.length === 0) return;
        t.folders.forEach(f => {
          const arr = childrenOf.get(f.parent_id);
          if (arr) arr.push(f); else childrenOf.set(f.parent_id, [f]);
        });
      });
      const nodes: { id: string; name: string }[] = [];
      let parentId = rootId;
      for (const segment of segments) {
        const hit = (childrenOf.get(parentId) || []).find(f => (f.name || '').trim().toLowerCase() === segment.toLowerCase());
        if (!hit) return null;
        nodes.push({ id: hit.id, name: hit.name });
        parentId = hit.id;
      }
      return nodes;
    };
    const resolveCachedDriveFolderStack = (path: string): { id: string; name: string }[] | null => {
      if (!effectiveLiveSiteId || !effectiveLiveDriveId) return null;
      const segments = path.replace(/\\/g, '/').split('/').map(seg => seg.trim()).filter(Boolean);
      if (segments.length === 0) return null;
      const fromTrees = resolveStackFromTrees(segments);
      if (fromTrees) return fromTrees;
      const nodes: { id: string; name: string }[] = [];
      let parentId = 'root';
      for (const segment of segments) {
        const entry = host._getOrLoadSiteFolderChildren(effectiveLiveSiteId, effectiveLiveDriveId, parentId);
        if (!entry || entry.loading) return null;
        const hit = (entry.items || []).find((item: any) =>
          item?.folder && String(item.name || '').trim().toLowerCase() === segment.toLowerCase());
        if (!hit) return null;
        nodes.push({ id: hit.id, name: hit.name });
        parentId = hit.id;
      }
      return nodes;
    };
    // Counterpart to resolveCachedDriveFolderStack for when a level genuinely
    // isn't cached yet (that function's "returns null only if a level
    // genuinely still needs a Graph round-trip" case) — instead of giving up
    // there, this awaits _loadAndCacheSiteFolderChildren (the same
    // _siteFolderItemsCache, but Promise-based) one level at a time so a
    // main folder that has never been opened this session still resolves,
    // just a beat later. Used by navigateToLiveMainFolder's own fallback so
    // jumping to a vessel's "is also under" main folder for the first time
    // in a session lands inside that vessel's folder there instead of on
    // the main folder's bare root.
    const resolveLiveDriveFolderStackAsync = async (path: string): Promise<{ id: string; name: string }[] | null> => {
      if (!effectiveLiveSiteId || !effectiveLiveDriveId) return null;
      const segments = path.replace(/\\/g, '/').split('/').map(seg => seg.trim()).filter(Boolean);
      if (segments.length === 0) return null;
      const fromTrees = resolveStackFromTrees(segments);
      if (fromTrees) return fromTrees;
      const nodes: { id: string; name: string }[] = [];
      let parentId = 'root';
      for (const segment of segments) {
        const items = await host._loadAndCacheSiteFolderChildren(effectiveLiveSiteId, effectiveLiveDriveId, parentId);
        const hit = (items || []).find((item: any) =>
          item?.folder && String(item.name || '').trim().toLowerCase() === segment.toLowerCase());
        if (!hit) return null;
        nodes.push({ id: hit.id, name: hit.name });
        parentId = hit.id;
      }
      return nodes;
    };

    // Proactively confirm whether the selected vessel also has a folder
    // under each of this site's OTHER known main folders — not just the
    // one(s) already discovered from browsing. Without this, a vessel's
    // second home (the "also under" chips, the Compare button) only ever
    // shows up once the user happens to have browsed into that other main
    // folder at least once this session, which is exactly what made the
    // chips/Compare button silently disappear after landing on Bow Fighter
    // straight from the Vessel dropdown, before "Technical" had ever been
    // opened this session. Each check is just resolveCachedDriveFolderStack
    // for "<mainFolder>/<vessel>": free once cached, and when it isn't yet,
    // its own _getOrLoadSiteFolderChildren side effect primes the cache in
    // the background so the next render picks the folder up naturally
    // through the normal knownSiteFolderPaths walk above (self-healing,
    // same pattern used everywhere else in this file). mainFolderOptions is
    // a small, fixed set of departments, so this stays cheap — at most one
    // extra Graph call per not-yet-checked main folder, once.
    if (vesselFilter !== 'all' && vesselFilter.trim().toLowerCase() !== 'not listed' &&
      (docScopeType === 'sites' || docScopeType === 'shared_docs' || docScopeType === 'documents')) {
      const knownVesselMainsLower = new Set(vesselMainFolders.map(m => m.trim().toLowerCase()));
      mainFolderOptions.forEach(m => {
        if (knownVesselMainsLower.has(m.trim().toLowerCase())) return;
        resolveCachedDriveFolderStack(`${m}/${vesselFilter.trim()}`);
      });
    }

    // ── Compare mode: one browsable box per main folder the selected vessel
    // is under, side by side, instead of Folder view's single location.
    // Each box drills through its own vessel-folder subtree independently
    // (compareBoxSubPaths, keyed by main folder), reusing
    // resolveCachedDriveFolderStack/_getOrLoadSiteFolderChildren the same
    // way the rest of Folder view does — so a box that needs a Graph round
    // trip just shows "Loading…" and resolves itself on the next render,
    // same self-healing behaviour as everywhere else in this file.
    const compareBoxStyle: React.CSSProperties = {
      border: '1px solid var(--vdms-border)', borderRadius: 10, background: 'var(--vdms-surface)',
      display: 'flex', flexDirection: 'column', minWidth: 0,
    };
    const renderCompareBox = (mainFolder: string): React.ReactElement => {
      const basePath = vesselBasePathUnderMain(mainFolder);
      const subPath = compareBoxSubPaths[mainFolder] || [];
      const setSubPath = (next: string[]): void => {
        host.setState({ compareBoxSubPaths: { ...compareBoxSubPaths, [mainFolder]: next } });
      };
      const crumb = [mainFolder, vesselFilter, ...subPath];
      const header = (
        <div style={{ padding: '10px 12px', borderBottom: '1px solid var(--vdms-border-soft)', fontSize: 11, color: 'var(--vdms-text-muted)', display: 'flex', flexWrap: 'wrap', gap: 2 }}>
          {crumb.map((seg, i) => {
            const isLast = i === crumb.length - 1;
            const targetSubPath = i <= 1 ? [] : subPath.slice(0, i - 1);
            return (
              <span key={i}>
                {i > 0 && <span style={{ opacity: 0.5 }}> / </span>}
                <span
                  onClick={isLast ? undefined : () => setSubPath(targetSubPath)}
                  style={{
                    cursor: isLast ? 'default' : 'pointer',
                    textDecoration: isLast ? 'none' : 'underline',
                    fontWeight: isLast ? 700 : 500,
                    color: isLast ? 'var(--vdms-text)' : 'var(--vdms-text-muted)',
                  }}
                >
                  {seg}
                </span>
              </span>
            );
          })}
        </div>
      );
      if (!basePath || !siteNavPrefix || !effectiveLiveSiteId || !effectiveLiveDriveId) {
        return (
          <div key={mainFolder} style={compareBoxStyle}>
            {header}
            <div style={{ padding: 16, fontSize: 12, color: 'var(--vdms-text-muted)' }}>Loading…</div>
          </div>
        );
      }
      const fullPath = subPath.length > 0 ? `${basePath}/${subPath.join('/')}` : basePath;
      const nodes = resolveCachedDriveFolderStack(fullPath);
      if (!nodes || nodes.length === 0) {
        return (
          <div key={mainFolder} style={compareBoxStyle}>
            {header}
            <div style={{ padding: 16, fontSize: 12, color: 'var(--vdms-text-muted)' }}>Loading…</div>
          </div>
        );
      }
      const leafId = nodes[nodes.length - 1].id;
      const listing = host._getOrLoadSiteFolderChildren(effectiveLiveSiteId, effectiveLiveDriveId, leafId);
      const items = listing.items || [];
      const boxFolders = items.filter((it: any) => !!it.folder);
      const boxFiles = items.filter((it: any) => !it.folder);
      return (
        <div key={mainFolder} style={compareBoxStyle}>
          {header}
          <div style={{ maxHeight: 340, overflowY: 'auto', padding: 6 }}>
            {listing.loading && <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', padding: 8 }}>Loading…</div>}
            {!listing.loading && boxFolders.length === 0 && boxFiles.length === 0 && (
              <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', padding: 8 }}>Empty</div>
            )}
            {boxFolders.map((f: any) => (
              <div
                key={f.id}
                onClick={() => setSubPath([...subPath, f.name])}
                style={{ padding: '6px 8px', cursor: 'pointer', borderRadius: 6, fontSize: 12, display: 'flex', alignItems: 'center', gap: 6 }}
                onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = 'var(--vdms-surface-alt)'; }}
                onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = 'transparent'; }}
              >
                📁 {f.name}
              </div>
            ))}
            {boxFiles.map((file: any) => {
              const fileUrl = file.web_url || file.webUrl || file.download_url || '';
              return (
                <div
                  key={file.id || file.name}
                  onClick={() => {
                    if (fileUrl) window.open(fileUrl, '_blank');
                    else void host._openDocumentFile(file.id, file.name, crumb.join(' > '));
                  }}
                  style={{ padding: '6px 8px', cursor: 'pointer', borderRadius: 6, fontSize: 12, color: '#0284c7' }}
                  onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = 'var(--vdms-surface-alt)'; }}
                  onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = 'transparent'; }}
                >
                  📄 {file.name}
                </div>
              );
            })}
          </div>
        </div>
      );
    };
    const renderVesselCompareGrid = (): React.ReactElement => (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: 'var(--vdms-text)' }}>
            Comparing {vesselFilter} across {vesselMainFolders.length} main folders
          </div>
          <button
            type="button"
            onClick={() => host.setState({ docCompareMode: false })}
            style={{
              padding: '5px 12px', borderRadius: 8, border: '1px solid var(--vdms-border)',
              background: 'var(--vdms-surface)', color: 'var(--vdms-text-muted)', fontSize: 12, fontWeight: 600, cursor: 'pointer',
            }}
          >
            ✕ Exit compare
          </button>
        </div>
        {/* One box per main folder; wraps to a new row on narrow widths, and
            each box scrolls its own contents independently once a folder
            has more than ~10 items, rather than growing the whole page. */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 12 }}>
          {vesselMainFolders.map(m => renderCompareBox(m))}
        </div>
      </div>
    );
    // Picking a vessel while the Folder view is somewhere that has nothing
    // for it: jump to that vessel's folder (preferring one inside the folder
    // currently open) instead of showing an empty "No folder for X here".
    // The landed stack is [sites_root, site, drive, <main folder>, …, <vessel>],
    // so deriveDocFiltersFromStack auto-selects the Main folder dropdown and
    // puts the vessel folder in the Sub-folder dropdown.
    //
    // Works from every live-library scope, not just "SharePoint Sites":
    //   • sites_root stack            → [sites_root, site, drive, …]  (3 prefix nodes)
    //   • Shared Documents / Documents → [lib:…, …]                    (1 prefix node)
    // siteNavPrefix already synthesises the sites_root prefix for the latter,
    // the same way navigateToLiveMainFolder / navigateToLiveSubfolder do.
    //
    // ── Carrying the selected Sub-folder to another vessel / main folder ──
    // Changing the Vessel filter, or jumping to the same vessel under another
    // main folder (the "<vessel> is also under:" chips, or the Main folder
    // dropdown with a vessel selected), used to reset the Sub-folder dropdown
    // to "All sub-folders" (and the chips also cleared Group / Category), so
    // the user had to re-pick "Hull" every time. The Sub-folder value really
    // means "this folder inside the vessel", so it is re-resolved *inside the
    // vessel folder being opened* by its path relative to the vessel (e.g.
    // Drawings > Hull): from the main folder's folder tree, the per-level
    // folder cache, and finally the live listing. Only when that vessel folder
    // is confirmed not to have it does the selection fall back to "All
    // sub-folders". Group / Category are never touched here.
    type CarrySubfolder = { name: string; rel: string[] };
    type CarryNode = { id: string; name: string };
    type CarryResult = { nodes: CarryNode[]; rel: string[] };
    const carryLow = (names: string[]): string => names.map(n => n.trim().toLowerCase()).join('/');
    // The selected sub-folder's path relative to the vessel it sits in.
    // Folder view: the breadcrumb (it also follows tile clicks); otherwise the
    // branch remembered from the dropdown pick.
    const buildSubfolderCarry = (): CarrySubfolder | null => {
      const prevSub = (docSubfolderOtherFilter || 'all').trim();
      if (!prevSub || prevSub.toLowerCase() === 'all') return null;
      const prevVesselLabel = activeVesselFilter ? (vesselFilter || '') : '';
      const isPrevVesselSeg = (n: string): boolean => {
        if (activeVesselFilter) {
          if (normalizeVesselKey(n) === activeVesselFilter) return true;
          const m = matchFolderToVessel(n);
          return !!m && vesselNamesEqual(m, prevVesselLabel);
        }
        return !!matchFolderToVessel(n);
      };
      const lastIsPrevSub = (names: string[]): boolean =>
        names.length > 0 && names[names.length - 1].trim().toLowerCase() === prevSub.toLowerCase();
      let rel: string[] = [];
      const prefixLen = folderPathStack[0]?.id === 'sites_root'
        ? 3
        : ((folderPathStack[0]?.id === 'lib:shared_documents' || folderPathStack[0]?.id === 'lib:documents') ? 1 : -1);
      if (docViewMode === 'folder' && prefixLen >= 0 && folderPathStack.length > prefixLen) {
        const names = folderPathStack.slice(prefixLen).map(n => n.name);
        const vi = names.findIndex(isPrevVesselSeg);
        if (vi >= 0 && lastIsPrevSub(names.slice(vi + 1))) rel = names.slice(vi + 1);
      }
      if (rel.length === 0) {
        const sp = (host as any)._subfolderSelectedPath as string[] | undefined;
        if (sp && sp.length && lastIsPrevSub(sp)) {
          const vi = sp.findIndex(isPrevVesselSeg);
          rel = vi >= 0 ? sp.slice(vi + 1) : sp.slice();
        }
      }
      if (rel.length === 0) rel = [prevSub];
      return { name: rel[rel.length - 1], rel };
    };
    // The vessel's own folder paths, those under `mainFolder` (default: the
    // one being browsed) first, the shortest ("<main>/<vessel>") first.
    const carryBasePaths = (vesselName: string, onlyMain?: string): string[] => {
      const curMain = (onlyMain || scopedMainFolderForSubfolders || '').trim().toLowerCase();
      const all = siteVesselFolderPaths.get(vesselName.trim().toLowerCase()) || [];
      const byLen = (a: string, b: string): number => a.length - b.length;
      const underMain = curMain
        ? all.filter(p => p.includes('/') && (p.split('/')[0] || '').trim().toLowerCase() === curMain).sort(byLen)
        : [];
      if (onlyMain) return underMain;
      return [...underMain, ...all.filter(p => underMain.indexOf(p) < 0).sort(byLen)];
    };
    // Exact relative path first; with `exactOnly` false also the same last
    // two folder names (vessels may nest the group folders differently),
    // then the shallowest folder with the same name.
    const pickCarry = (cands: CarryResult[], carry: CarrySubfolder, exactOnly: boolean): CarryResult | null => {
      if (cands.length === 0) return null;
      const want = carry.rel.map(n => n.trim().toLowerCase());
      const byDepth = (a: CarryResult, b: CarryResult): number => a.rel.length - b.rel.length;
      const exact = cands.filter(c => carryLow(c.rel) === want.join('/')).sort(byDepth)[0];
      if (exact || exactOnly) return exact || null;
      return cands.filter(c => {
          const r = c.rel.map(n => n.trim().toLowerCase());
          const k = Math.min(r.length, want.length, 2);
          return k >= 2 && r.slice(-k).join('/') === want.slice(-k).join('/');
        }).sort(byDepth)[0]
        || cands.filter(c => c.rel[c.rel.length - 1].trim().toLowerCase() === want[want.length - 1]).sort(byDepth)[0]
        || null;
    };
    // Every folder below the vessel folder (last of vesselNodes, a
    // drive-root-relative stack) that is already known: the main folder's
    // folder tree plus the Sub-folder dropdown's own entries.
    const carryCandidates = (vesselNodes: CarryNode[]): CarryResult[] => {
      if (vesselNodes.length < 2) return [];
      const main = vesselNodes[0];
      const vesselId = vesselNodes[vesselNodes.length - 1].id;
      const out: CarryResult[] = [];
      const seen = new Set<string>();
      const addChain = (chain: CarryNode[]): void => {
        const vi = chain.findIndex(n => n.id === vesselId);
        if (vi < 0) return;
        for (let j = vi + 1; j < chain.length; j++) {
          if (!chain[j].id || seen.has(chain[j].id)) continue;
          seen.add(chain[j].id);
          out.push({ nodes: [main, ...chain.slice(0, j + 1)], rel: chain.slice(vi + 1, j + 1).map(n => n.name) });
        }
      };
      if (effectiveLiveSiteId && effectiveLiveDriveId) {
        const tree = host._getOrLoadSiteFolderTree(effectiveLiveSiteId, effectiveLiveDriveId, main.id);
        if (tree && tree.folders.length > 0) {
          const byId = new Map(tree.folders.map(f => [f.id, f]));
          const parentIds = new Set(tree.folders.map(f => f.parent_id));
          const treeRootId = tree.folders.find(f => f.depth === 1)?.parent_id || main.id;
          // Leaves only: each chain already covers all of its ancestors.
          tree.folders.forEach(f => {
            if (parentIds.has(f.id)) return;
            const chain: CarryNode[] = [];
            let cur: typeof f | undefined = f;
            let guard = 0;
            while (cur && guard++ < 64) {
              chain.unshift({ id: cur.id, name: cur.name });
              if (cur.parent_id === treeRootId) break;
              cur = byId.get(cur.parent_id);
            }
            if (cur && cur.parent_id === treeRootId) addChain(chain);
          });
        }
      }
      if (scopedMainFolderLiveFolderId === main.id) {
        liveSubfolderEntries.forEach(e => addChain(e.pathIds.map((id, i) => ({ id, name: e.pathNames[i] }))));
      }
      return out;
    };
    const carrySyncFor = (vesselNodes: CarryNode[], carry: CarrySubfolder): CarryResult | null => {
      const cands = carryCandidates(vesselNodes);
      const exact = pickCarry(cands, carry, true);
      if (exact) return exact;
      const cached = resolveCachedDriveFolderStack([...vesselNodes.map(n => n.name), ...carry.rel].join('/'));
      if (cached) return { nodes: cached, rel: carry.rel };
      return pickCarry(cands, carry, false);
    };
    // Live, level by level. 'missing' only when the vessel folder's listing
    // really came back without the sub-folder; 'unknown' when a level could
    // not be read (throttled / failed) — the selection is then kept as is.
    const carryAsyncFor = async (vesselNodes: CarryNode[], carry: CarrySubfolder): Promise<CarryResult | 'missing' | 'unknown'> => {
      if (!effectiveLiveSiteId || !effectiveLiveDriveId || vesselNodes.length === 0) return 'unknown';
      const siteId = effectiveLiveSiteId;
      const driveId = effectiveLiveDriveId;
      const nodes = vesselNodes.slice();
      let parentId = nodes[nodes.length - 1].id;
      for (const seg of carry.rel) {
        const items = await host._loadAndCacheSiteFolderChildren(siteId, driveId, parentId);
        const hit = (items || []).find((item: any) =>
          item?.folder && String(item.name || '').trim().toLowerCase() === seg.trim().toLowerCase());
        if (!hit) {
          const level = host._siteFolderItemsCache.get(`${siteId}::${driveId}::${parentId}`) as any;
          if (!level || level.error || level.loading) return 'unknown';
          // Not at the exact path — the vessel may nest it differently. Give
          // the main folder's tree a moment to arrive, then match loosely.
          for (let i = 0; i < 20; i++) {
            const t = host._getOrLoadSiteFolderTree(siteId, driveId, vesselNodes[0].id);
            if (!t || !t.loading) break;
            await new Promise(r => window.setTimeout(r, 1500));
          }
          return pickCarry(carryCandidates(vesselNodes), carry, false) || 'missing';
        }
        nodes.push({ id: hit.id, name: hit.name });
        parentId = hit.id;
      }
      return { nodes, rel: carry.rel };
    };
    const isSameVesselNow = (vesselName: string): boolean =>
      normalizeVesselKey(String(host.state.vesselFilter || '')) === normalizeVesselKey(vesselName);
    const subfolderStillIs = (carry: CarrySubfolder): boolean =>
      (host.state.docSubfolderOtherFilter || '').trim().toLowerCase() === carry.name.trim().toLowerCase();
    // Opens the vessel folder (drive-root-relative `vesselNodes`) and, when a
    // sub-folder is being carried, that sub-folder inside it instead.
    const openVesselWithCarry = (vesselName: string, vesselNodes: CarryNode[], carry?: CarrySubfolder | null): void => {
      if (!siteNavPrefix) return;
      const prefix = siteNavPrefix;
      const carried = carry ? carrySyncFor(vesselNodes, carry) : null;
      const stack = [...prefix, ...(carried ? carried.nodes : vesselNodes)];
      host._pushFolderNav(stack, 'SharePoint Sites');
      if (carried) {
        (host as any)._subfolderSelectedPath = carried.rel;
        host.setState({
          ...deriveDocFiltersFromStack(stack),
          docSubfolderOtherFilter: carried.rel[carried.rel.length - 1] || (carry ? carry.name : 'all'),
          vesselFilter: vesselName,
          docListPage: 0,
        });
        return;
      }
      // Main folder ← stack[3]. Without a sub-folder to carry the Sub-folder
      // dropdown stays on "All sub-folders": with a vessel selected it lists
      // that vessel's own sub-folders (the vessel folder itself is not one of
      // its options).
      host.setState({
        ...deriveDocFiltersFromStack(stack),
        docSubfolderOtherFilter: carry ? carry.name : 'all',
        vesselFilter: vesselName,
        docListPage: 0,
      });
      if (!carry) return;
      // Not resolvable from what is loaded yet: keep the pick and finish the
      // jump into it once the live listing answers (unless the user has moved
      // on in the meantime).
      (host as any)._subfolderSelectedPath = carry.rel;
      void carryAsyncFor(vesselNodes, carry).then(res => {
        if (!isSameVesselNow(vesselName) || !subfolderStillIs(carry)) return;
        const cur = host.state.folderPathStack || [];
        const stillHere = cur.length === stack.length && cur[cur.length - 1]?.id === stack[stack.length - 1]?.id;
        if (!stillHere || res === 'unknown') return;
        if (res === 'missing') {
          (host as any)._subfolderSelectedPath = undefined;
          host.setState({ docSubfolderOtherFilter: 'all', docListPage: 0 });
          return;
        }
        const deep = [...prefix, ...res.nodes];
        host._pushFolderNav(deep, 'SharePoint Sites');
        (host as any)._subfolderSelectedPath = res.rel;
        host.setState({
          ...deriveDocFiltersFromStack(deep),
          docSubfolderOtherFilter: res.rel[res.rel.length - 1] || carry.name,
          vesselFilter: vesselName,
          docListPage: 0,
        });
      }).catch(() => undefined);
    };
    // List view has no breadcrumb to move: keep the Sub-folder pick (its
    // branch is what the row filter uses) and only drop it once none of the
    // vessel's folders (under `onlyMain`, when given) has that folder.
    const finishCarryInList = (vesselName: string, carry: CarrySubfolder, onlyMain?: string): void => {
      const bases = carryBasePaths(vesselName, onlyMain);
      for (const base of bases) {
        const vesselNodes = resolveCachedDriveFolderStack(base);
        const found = vesselNodes ? carrySyncFor(vesselNodes, carry) : null;
        if (found) { (host as any)._subfolderSelectedPath = found.rel; return; }
      }
      (host as any)._subfolderSelectedPath = carry.rel;
      if (bases.length === 0) return;
      void (async (): Promise<CarryResult | 'missing' | 'unknown'> => {
        let sawMissing = false;
        for (const base of bases) {
          const vesselNodes = await resolveLiveDriveFolderStackAsync(base);
          if (!vesselNodes) continue;
          const res = await carryAsyncFor(vesselNodes, carry);
          if (res === 'missing') { sawMissing = true; continue; }
          if (res !== 'unknown') return res;
        }
        return sawMissing ? 'missing' : 'unknown';
      })().then(res => {
        if (!isSameVesselNow(vesselName) || !subfolderStillIs(carry) || res === 'unknown') return;
        if (res === 'missing') {
          (host as any)._subfolderSelectedPath = undefined;
          host.setState({ docSubfolderOtherFilter: 'all', docListPage: 0 });
          return;
        }
        (host as any)._subfolderSelectedPath = res.rel;
        host.setState({ docSubfolderOtherFilter: res.rel[res.rel.length - 1] || carry.name, docListPage: 0 });
      }).catch(() => undefined);
    };
    // Same vessel, other main folder: the "<vessel> is also under:" chips and
    // the Main folder dropdown (with a vessel selected). Keeps Group /
    // Category and carries the Sub-folder pick into the vessel's folder there.
    const switchVesselMainFolder = (mainFolder: string): void => {
      (host as any)._mainFolderImplicit = undefined;
      const carry = buildSubfolderCarry();
      host.setState({
        docCategoryFilter: mainFolder,
        catFilter: mainFolder,
        docSubfolderOtherFilter: carry ? carry.name : 'all',
        docListPage: 0,
      });
      if (carry) (host as any)._subfolderSelectedPath = carry.rel;
      navigateToLiveMainFolder(mainFolder, carry);
    };
    const navigateToSiteVesselFolder = (vesselName: string, carry?: CarrySubfolder | null): void => {
      if (!siteNavPrefix) return;
      const rootId = folderPathStack[0]?.id;
      const livePrefixLength = rootId === 'sites_root'
        ? 3
        : ((rootId === 'lib:shared_documents' || rootId === 'lib:documents') ? 1 : -1);
      const vesselPaths = siteVesselFolderPaths.get(vesselName.trim().toLowerCase()) || [];
      const candidates = vesselPaths.slice();
      if (candidates.length === 0) {
        // The live folder walk hasn't recorded this vessel's folder (the
        // option came from the flattened rows instead). Fall back to the
        // row's own main folder: "<Main folder>/<Vessel>" if that folder
        // exists, otherwise at least land on (and select) the main folder.
        const row = scopeRows.find(r => vesselNamesEqual(getListViewLabels(r).vessel || '', vesselName));
        const mainFolder = row ? (getListViewLabels(row).mainFolder || '').trim() : '';
        const mainCard = mainFolder
          ? liveRootFolders.find(c => c.key.trim().toLowerCase() === mainFolder.toLowerCase())
          : undefined;
        if (!mainCard) return;
        candidates.push(`${mainCard.key}/${vesselName.trim()}`, mainCard.key);
      }
      // List View (or Folder view before any live root has been resolved
      // into the breadcrumb stack, e.g. right after a site switch) has no
      // breadcrumb to push a folder-nav stack onto — _pushFolderNav only
      // means something in Folder view. But the Main folder / Sub-folder
      // dropdowns must still auto-select the vessel's own main folder here,
      // same as Folder view does below, otherwise picking a vessel from
      // List view leaves "All main folders" selected and the Sub-folder
      // dropdown empty (subfolderTree is scoped off docCategoryFilter,
      // see scopedMainFolderForSubfolders above).
      if (docViewMode !== 'folder' || livePrefixLength < 0) {
        // A vessel commonly has a folder under several main folders (e.g.
        // "Technical & Crewing/<vessel>" AND "Commercial & Chartering/<vessel>"
        // AND "Insurance/<vessel>" all exist for the same vessel). List view's
        // whole purpose is to show the vessel's documents flattened across
        // the site, so the Main folder dropdown must only be pinned to one
        // specific main folder when the vessel actually has just one home —
        // pinning it to whichever main folder happened to be first in
        // `candidates` used to silently drop every row that lived under the
        // vessel's *other* main folders (isLiveMainFolderValue path-segment
        // check further down ANDs docCategoryFilter with the vessel filter).
        // With it left on "All main folders", the vessel filter alone (which
        // already matches by folder segment regardless of which main folder
        // it's under) is enough to pull in all of them.
        const distinctMainFolders = Array.from(new Set(
          candidates.map(p => (p.split('/')[0] || '').trim().toLowerCase()).filter(Boolean)
        ));
        const mainFolderName = distinctMainFolders.length === 1
          ? (candidates[0] || '').split('/')[0]?.trim()
          : '';
        host.setState({
          docCategoryFilter: mainFolderName || 'all',
          docSubfolderOtherFilter: carry ? carry.name : 'all',
          vesselFilter: vesselName,
          docListPage: 0,
        });
        if (carry) finishCarryInList(vesselName, carry);
        return;
      }
      // Main folder is on "All main folders" (the user never picked one):
      // do NOT open the vessel's first main folder (usually "Technical") for
      // them. Stay at the library root instead — Folder view there lists
      // only the main folders that lead to this vessel (vesselScopedChildFolders),
      // so every main folder is shown and the user chooses where to go (or
      // uses the "<vessel> is also under" chips / Compare).
      const implicitAllMain = (host as any)._mainFolderImplicit as { vessel: string } | undefined;
      if (implicitAllMain && implicitAllMain.vessel === normalizeVesselKey(vesselName)) {
        const atLibraryRoot = folderPathStack.length <= livePrefixLength;
        if (!atLibraryRoot) host._pushFolderNav([...siteNavPrefix], 'SharePoint Sites');
        host.setState({
          ...deriveDocFiltersFromStack(siteNavPrefix),
          docCategoryFilter: 'all',
          docSubfolderOtherFilter: 'all',
          vesselFilter: vesselName,
          docListPage: 0,
        });
        return;
      }
      const here = folderPathStack.slice(livePrefixLength).map(n => n.name).join('/').toLowerCase();
      // Already inside the vessel's folder (or deeper)? Stay put. Being at the
      // library root or at an ancestor (e.g. the main folder) is NOT enough —
      // the user wants the vessel's own sub-folder opened. `!here` (and the
      // ancestor check) used to count as "on branch", which made picking a
      // vessel at the root a silent no-op — the same bug
      // deriveLiveNavFilterState had, see its comment above.
      const insideVesselFolder = !!here && vesselPaths.some(p => {
        const vp = p.toLowerCase();
        return vp === here || here.startsWith(`${vp}/`);
      });
      if (insideVesselFolder) return;
      const preferred = candidates.filter(p => here && p.toLowerCase().startsWith(`${here}/`));
      // Coming from another vessel's folder (…/Technical/<old vessel>/Drawings/Hull)
      // nothing starts with `here`, so without this the new vessel could land
      // under a different main folder than the one being browsed.
      const hereMain = here.split('/')[0] || '';
      const sameMain = hereMain
        ? candidates.filter(p => preferred.indexOf(p) < 0 && (p.split('/')[0] || '').trim().toLowerCase() === hereMain)
        : [];
      const ordered = [...preferred, ...sameMain, ...candidates];
      let nodes: { id: string; name: string }[] | null = null;
      for (const candidate of ordered) {
        nodes = resolveCachedDriveFolderStack(candidate);
        if (nodes) break;
      }
      // A previously selected sub-folder (e.g. Drawings > Hull) is opened
      // inside this vessel instead of stopping at the vessel's root.
      if (nodes) { openVesselWithCarry(vesselName, nodes, carry); return; }
      // The vessel's folder isn't in the per-level cache yet (a level that
      // needs a Graph round-trip): resolve it level by level instead of
      // silently leaving the breadcrumb on the previous vessel.
      void (async () => {
        for (const candidate of ordered) {
          const resolved = await resolveLiveDriveFolderStackAsync(candidate);
          if (resolved) return resolved;
        }
        return null;
      })().then(resolved => {
        if (!resolved || !isSameVesselNow(vesselName)) return;
        openVesselWithCarry(vesselName, resolved, carry);
      }).catch(() => undefined);
    };

    // Shared by the Vessel filter <select> below and the search box's
    // results dropdown (selecting a vessel suggestion there must behave
    // exactly like picking it from the dropdown) — factored out so there's
    // only one place that knows how to apply a vessel selection.
    const liveStackPrefixLen = folderPathStack[0]?.id === 'sites_root'
      ? 3
      : ((folderPathStack[0]?.id === 'lib:shared_documents' || folderPathStack[0]?.id === 'lib:documents') ? 1 : -1);
    const breadcrumbMainFolderName: string | null = liveStackPrefixLen >= 0 && folderPathStack.length > liveStackPrefixLen
      ? (folderPathStack[liveStackPrefixLen]?.name || null)
      : null;
    // The Main folder dropdown mirrors the folder on screen in Folder view.
    // Picking a vessel while it reads "All main folders" has to open one of
    // the vessel's main folders (a vessel is often under several), which used
    // to flip the dropdown to that folder: a choice the user never made.
    // `_mainFolderImplicit` remembers that the main folder was picked for the
    // user, so the dropdown keeps reading "All main folders" (the "<vessel> is
    // also under" chips show where you actually are). It ends when the vessel
    // is cleared or changed away, or when a main folder is picked explicitly
    // (dropdown / chip).
    const implicitMainRec = (host as any)._mainFolderImplicit as { vessel: string } | undefined;
    const implicitMainAll = !!implicitMainRec && !!activeVesselFilter && implicitMainRec.vessel === activeVesselFilter;
    if (implicitMainRec && !implicitMainAll) (host as any)._mainFolderImplicit = undefined;
    const mainFolderSelectValue: string = implicitMainAll
      ? 'all'
      : (docCategoryFilter !== 'all'
        ? docCategoryFilter
        : ((breadcrumbMainFolderName && mainFolderOptions.find(o => o.trim().toLowerCase() === breadcrumbMainFolderName.trim().toLowerCase())) || 'all'));
    const applyVesselFilterSelection = (val: string): void => {
      if (docViewMode === 'folder' && mainFolderPage && docMainFolder) {
        if (val && val !== 'all') {
          const vesselMatch = vessels.find(v => v.name.trim().toLowerCase() === val.trim().toLowerCase());
          const vesselId = vesselMatch?.id || val;
          host._pushFolderNav(
            [{ id: docMainFolder, name: docMainFolder }, { id: String(vesselId), name: val }],
            docMainFolder
          );
        } else {
          host._pushFolderNav([{ id: docMainFolder, name: docMainFolder }], docMainFolder);
        }
      }
      const isLiveLibraryScope = docScopeType === 'sites' || docScopeType === 'shared_docs' || docScopeType === 'documents';
      const clearingVessel = !val || val === 'all';
      const wasAllMainFolders = mainFolderSelectValue === 'all';
      if (clearingVessel) (host as any)._mainFolderImplicit = undefined;
      // "All vessels" in a live SharePoint library: the breadcrumb is sitting
      // inside (or below) the previously selected vessel's folder, and only
      // clearing the filter state left it there. Walk the stack back to the
      // folder that contains the vessel folder (normally the main folder), so
      // the breadcrumb, grid and Main folder / Sub-folder dropdowns all agree.
      if (isLiveLibraryScope && docViewMode === 'folder' && (!val || val === 'all')) {
        const rootId = folderPathStack[0]?.id;
        const prefixLen = rootId === 'sites_root'
          ? 3
          : ((rootId === 'lib:shared_documents' || rootId === 'lib:documents') ? 1 : -1);
        const prevVessel = (host.state.vesselFilter || 'all').trim();
        if (prefixLen >= 0 && prevVessel && prevVessel.toLowerCase() !== 'all' && prevVessel.toLowerCase() !== 'not listed') {
          let vesselIdx = -1;
          for (let i = prefixLen; i < folderPathStack.length; i++) {
            const matched = matchFolderToVessel(folderPathStack[i].name);
            if (matched && vesselNamesEqual(matched, prevVessel)) { vesselIdx = i; break; }
          }
          // "All main folders" was what the user had (the vessel jump picked a
          // main folder for them): clearing the vessel goes back to the
          // library root, not to that auto-picked main folder.
          if (vesselIdx >= 0 || implicitMainAll) {
            const trimmed = folderPathStack.slice(0, implicitMainAll ? prefixLen : vesselIdx);
            host._pushFolderNav(trimmed, rootId === 'lib:documents' ? 'Documents' : (rootId === 'lib:shared_documents' ? 'Shared Documents' : 'SharePoint Sites'));
            host.setState({ ...deriveDocFiltersFromStack(trimmed), vesselFilter: 'all', docListPage: 0 });
            return;
          }
        }
      }
      // The Sub-folder pick (e.g. Drawings > Hull) belongs to the folder, not
      // the vessel: carried over into the new vessel (see buildSubfolderCarry).
      const subfolderCarry = (isLiveLibraryScope && val && val !== 'all' && val !== 'Not Listed')
        ? buildSubfolderCarry()
        : null;
      // Remember an "All main folders" selection across the vessel jump (see
      // implicitMainAll above); anything else is the user's own main folder.
      (host as any)._mainFolderImplicit = (isLiveLibraryScope && !clearingVessel && val !== 'Not Listed' && wasAllMainFolders)
        ? { vessel: normalizeVesselKey(val) }
        : undefined;
      if (clearingVessel && implicitMainAll) {
        // List view: the vessel jump may have pinned a main folder; put
        // "All main folders" back with the vessel cleared.
        host.setState({ vesselFilter: 'all', docCategoryFilter: 'all', docSubfolderOtherFilter: 'all', docListPage: 0 });
      } else {
        host.setState({ vesselFilter: val || 'all', docListPage: 0 });
      }
      if (isLiveLibraryScope && val && val !== 'all' && val !== 'Not Listed') {
        navigateToSiteVesselFolder(val, subfolderCarry);
      }
      if (val && val !== 'all' && val !== 'Not Listed' && !isLiveLibraryScope) {
        void host._loadVesselRowsFromApi(val).catch(() => undefined);
        void host._loadFilesForVessel(val).catch(() => undefined);
      }
    };

    // ── mainFolderGroupMap: which groups belong to which main folder (for list view filtering) ──
    const mainFolderGroupMap = folderNamesByMainFolder(docScopeType === 'common');

    // Vessels are returned newest first. Start with four and extend in pages
    // of eight when the user chooses "More vessels".
    const visibleVesselNames = new Set(vessels.slice(0, documentVesselCount).map(v => v.name));

    const scopedVesselInFolderView =
      docViewMode === 'folder' && docScopeType === 'vessels' && currentVesselNameFromStack
        ? currentVesselNameFromStack.trim().toLowerCase()
        : null;

    // skipLevels lets the Group / Category / Sub-category dropdowns compute
    // their options from the rows that pass every OTHER filter:
    //   0 = apply all, 1 = ignore Sub-category, 2 = also ignore Category,
    //   3 = also ignore Group.
    const passesDocFilters = (r: FlatRow, skipLevels: number): boolean => {
      const labels = getListViewLabels(r);

      // A text query in SharePoint Sites scope is site-wide. Do not let the
      // folder currently open in the breadcrumb or its derived filters hide
      // matching files returned by the backend site search.
      //
      // IMPORTANT: this used to skip the Vessel filter entirely whenever a
      // text search was active — `wholeSiteSearch` only checks
      // activeDocumentSite/textFilter/docScopeType, never vesselFilter — so
      // typing a keyword and then picking a vessel from the Vessel filter
      // still showed every vessel's matches for that keyword instead of just
      // the selected one ("only that vessel's files need to filter, searched
      // content should not follow other vessels"). The Vessel filter must
      // still narrow the results here, exactly as it does below for the
      // no-search case (same folder-path-segment matching as that branch).
      if (wholeSiteSearch) {
        const hasVesselFilter = vesselFilter !== 'all' && vesselFilter.trim().toLowerCase() !== 'not listed';
        if (hasVesselFilter) {
          if ((r.groupKey || '').endsWith(':search')) {
            const rowVessel = labels.vessel || r.vesselName || '';
            if (!rowVessel || !vesselNamesEqual(rowVessel, vesselFilter)) return false;
          } else {
            const rowSegments = (r.subFolderPath || '').split(/\s*>\s*|\//).map(seg => seg.trim()).filter(Boolean);
            const libraryIdx = rowSegments.findIndex(seg => /^(documents|shared documents|sites documents|site library|general documents)$/i.test(seg));
            const folderSegments = libraryIdx >= 0 ? rowSegments.slice(libraryIdx + 1) : rowSegments;
            const inSelectedVessel = folderSegments.some(seg => {
              const matched = matchFolderToVessel(seg);
              return !!matched && vesselNamesEqual(matched, vesselFilter);
            });
            if (!inSelectedVessel) return false;
          }
        }
        if ((r.groupKey || '').endsWith(':search')) return true;
        return matchesSearchTokens(
          textFilter,
          labels.vessel,
          labels.mainFolder,
          labels.documentSection,
          labels.group,
          labels.category,
          labels.subCategory,
          r.vesselName,
          r.group,
          r.category,
          r.subCategory,
          r.subFolderPath,
          r.fileName,
        );
      }

      // 1. Vessel filter
      if (vesselFilter !== 'all' && docScopeType === 'sites' && vesselFilter.trim().toLowerCase() !== 'not listed') {
        // SharePoint Sites scope: a row belongs to the selected vessel only
        // when its folder path runs through a folder named after that vessel
        // (same matching as the dropdown) — not the fuzzy row label, which
        // let e.g. "mvtest3" also pick up "mvtest3421212" rows.
        const rowSegments = (r.subFolderPath || '').split(/\s*>\s*|\//).map(seg => seg.trim()).filter(Boolean);
        const libraryIdx = rowSegments.findIndex(seg => /^(documents|shared documents|sites documents|site library|general documents)$/i.test(seg));
        const folderSegments = libraryIdx >= 0 ? rowSegments.slice(libraryIdx + 1) : rowSegments;
        const inSelectedVessel = folderSegments.some(seg => {
          const matched = matchFolderToVessel(seg);
          return !!matched && vesselNamesEqual(matched, vesselFilter);
        });
        if (!inSelectedVessel) return false;
      } else if (vesselFilter !== 'all') {
        const normVesselFilter = vesselFilter.trim().toLowerCase();
        const normRowVessel = labels.vessel.trim().toLowerCase();
        if (normRowVessel !== normVesselFilter) {
          if (normVesselFilter === 'not listed') {
            if (normRowVessel !== 'not listed' && normRowVessel !== 'vessel name not listed') return false;
          } else {
            return false;
          }
        }
      } else if (docScopeType === 'vessels') {
        // When user has drilled into a vessel in Folder view, always scope to
        // that vessel even if the top dropdown still says "all".
        if (scopedVesselInFolderView) {
          if ((labels.vessel || r.vesselName || '').trim().toLowerCase() !== scopedVesselInFolderView) return false;
        } else if (!activeVesselName) {
          if (!visibleVesselNames.has(r.vesselName) && !visibleVesselNames.has(labels.vessel)) return false;
        } else if (labels.vessel.trim().toLowerCase() !== activeVesselName.trim().toLowerCase() && r.vesselName.trim().toLowerCase() !== activeVesselName.trim().toLowerCase()) {
          return false;
        }
      }

      // 1.5. Sub-folder filter — real folders that aren't a known main
      // folder/category and aren't a vessel (e.g. "Report", "ACRA CHARGE").
      // These rows commonly have no vessel at all, so this check stands on
      // its own rather than depending on the vessel filter above.
      // Matched primarily as an exact path SEGMENT of subFolderPath (split
      // on '>', not a raw substring) — a sub-folder can be nested at any
      // depth (e.g. "SS366 FINISH PLAN"), and a substring check both
      // under- and over-matches (misses an exact-but-not-adjacent segment,
      // or false-positives on a folder name that's a substring of another).
      // The label-based checks stay as a fallback for rows whose path
      // string doesn't cleanly carry the segment (legacy vessel/department
      // rows where "sub-folder" maps onto group/category instead).
      // (skipLevels 3 = computing Group / Category / Sub-category dropdown
      // options: those lists must keep offering every value in the folder
      // scope — with the current pick highlighted — instead of shrinking to
      // whatever the sub-folder pick leaves.)
      if (docSubfolderOtherFilter !== 'all' && skipLevels < 3) {
        const normFilter = docSubfolderOtherFilter.trim().toLowerCase();
        const pathSegments = (r.subFolderPath || '').split('>').map(s => s.trim().toLowerCase());
        // When the exact branch is known (tree pick), require that branch —
        // the parent folder must precede the name in the row's path.
        const selPath = ((host as any)._subfolderSelectedPath as string[] | undefined) || [];
        const branchKey = selPath.length >= 2 && selPath[selPath.length - 1].trim().toLowerCase() === normFilter
          ? selPath.slice(-2).map(x => x.trim().toLowerCase()) : null;
        const branchMatch = (): boolean => {
          if (!branchKey) return false;
          const at = pathSegments.indexOf(branchKey[1]);
          return at > 0 && pathSegments.slice(0, at).indexOf(branchKey[0]) !== -1;
        };
        const subfolderMatch = branchKey ? branchMatch() : (pathSegments.indexOf(normFilter) !== -1 ||
          (labels.mainFolder || '').trim().toLowerCase() === normFilter ||
          (labels.group || '').trim().toLowerCase() === normFilter ||
          (labels.category || '').trim().toLowerCase() === normFilter ||
          (labels.documentSection || '').trim().toLowerCase() === normFilter ||
          (r.group || '').trim().toLowerCase() === normFilter);
        if (!subfolderMatch) return false;
      }

      // 2. Main folder / group filter
      const effectiveGroupFilter = docGroupFilter !== 'all' ? docGroupFilter : (docViewMode === 'folder' && docMainFolder ? docMainFolder : null);
      if (effectiveGroupFilter && effectiveGroupFilter !== 'all') {
        const normGroup = (r.group || '').trim().toLowerCase();
        const normMainFolder = (labels.mainFolder || '').trim().toLowerCase();
        const normFilter = effectiveGroupFilter.trim().toLowerCase();
        const siteScopeMatch = docScopeType === 'sites' && normFilter === 'sharepoint sites' && (
          normGroup === 'sharepoint sites' || normGroup === 'shared documents' || normGroup === 'documents' ||
          (r.subFolderPath || '').toLowerCase().startsWith('sharepoint sites >') ||
          (r.subFolderPath || '').toLowerCase().startsWith('shared documents >') ||
          (r.subFolderPath || '').toLowerCase().startsWith('documents >')
        );
        const groupMatch = siteScopeMatch ||
          normMainFolder === normFilter ||
          normGroup === normFilter ||
          normGroup.includes(normFilter.split(' ')[0]) ||
          normMainFolder.includes(normFilter.split(' ')[0]) ||
          (r.subFolderPath || '').toLowerCase().includes(normFilter.split(' ')[0]);
        if (!groupMatch) return false;
      }

      // 3. Main folder filter (live SharePoint-tree scope) / Category /
      // Document section filter (vessel & department scope).
      // docCategoryFilter is overloaded between two unrelated concepts: for
      // vessel/department rows it's always been a "document section" filter
      // (Drawings/Manuals-style, kept below); for the live Main folder /
      // Sub-folder dropdowns built for the SharePoint-sites folder tree it
      // holds a real top-level folder NAME instead, which has nothing to do
      // with documentSection — comparing it there almost never matched,
      // which silently filtered every file out of List View whenever a live
      // Main folder was selected (worse combined with a Sub-folder
      // selection, since both then had to pass). When the current value is
      // one of the live folder names the Main folder dropdown itself
      // offers, match it as an exact path segment instead.
      const isLiveMainFolderValue = docCategoryFilter !== 'all' &&
        mainFolderOptions.some(n => n.trim().toLowerCase() === docCategoryFilter.trim().toLowerCase());
      if (isLiveMainFolderValue) {
        const normFilter = docCategoryFilter.trim().toLowerCase();
        const pathSegments = (r.subFolderPath || '').split('>').map(s => s.trim().toLowerCase());
        if (pathSegments.indexOf(normFilter) === -1) return false;
      } else {
        const effectiveSectionFilter = docCategoryFilter !== 'all' ? docCategoryFilter : (catFilter !== 'all' ? catFilter : 'all');
        if (effectiveSectionFilter !== 'all') {
          const normFilter = effectiveSectionFilter.trim().toLowerCase();
          const sectionMatch = (labels.documentSection || '').trim().toLowerCase() === normFilter ||
            (r.category || '').trim().toLowerCase() === normFilter ||
            (r.subFolderPath || '').toLowerCase().includes(`> ${normFilter}`) ||
            (r.subFolderPath || '').toLowerCase().includes(`${normFilter} >`);
          if (!sectionMatch) return false;
        }
      }
      // Group / Category filters use the same path-based rule as the
      // columns (drawingsManualsForRow), so a filter always matches what the
      // table shows.
      const gcGroup = skipLevels < 3 ? docGroupLevelFilter : 'all';
      const gcCategory = skipLevels < 2 ? docLeafCategoryFilter : 'all';
      if (gcGroup !== 'all' || gcCategory !== 'all') {
        if (!pathMatchesGroupCategory(r.subFolderPath || '', gcGroup, gcCategory)) return false;
      }
      if (skipLevels < 1 && docSubCategoryFilter !== 'all' && subCategoryPickCtx()) {
        // Classified pick: the Group › Category it was listed under counts too.
        if (!pathMatchesSubCategory(r.subFolderPath || '', docSubCategoryFilter)) return false;
      } else if (skipLevels < 1 && docSubCategoryFilter !== 'all') {
        const normFilter = docSubCategoryFilter.trim().toLowerCase();
        const subMatch = (labels.subCategory || '').trim().toLowerCase() === normFilter ||
          (r.subCategory || '').trim().toLowerCase() === normFilter ||
          (r.subFolderPath || '').toLowerCase().includes(`> ${normFilter}`);
        if (!subMatch) return false;
      }

      // 4. Text search filter — tokenized (see matchesSearchTokens above) so
      // a combined query such as "<vessel name> <partial file name>" still
      // matches: every whitespace-separated token in the box just needs to
      // be found somewhere across the row's vessel, folder hierarchy, path,
      // category and file name (any file format/extension), and each token
      // may match a different field. This also fixes plain partial file-name
      // searches where the typed words aren't contiguous in the file name
      // (e.g. spaces vs. underscores/hyphens: "annual report" now matches
      // "Annual_Report_2024.pdf").
      if (textFilter) {
        return matchesSearchTokens(
          textFilter,
          labels.vessel,
          labels.mainFolder,
          labels.documentSection,
          labels.group,
          labels.category,
          labels.subCategory,
          r.vesselName,
          r.group,
          r.category,
          r.subCategory,
          r.subFolderPath,
          r.fileName,
        );
      }
      return true;
    };
    let filtered = scopeRows.filter(r => passesDocFilters(r, 0));

    // ── Group / Category / Sub-category dropdown options ──
    // Built only from files that really are classified, so a value is
    // offered only when picking it would show at least one file; otherwise
    // the dropdown is disabled. Each level is scoped by the levels above it
    // (and by the vessel / main folder / sub-folder / attachment filters).
    const gcBaseRows: FlatRow[] = scopeRows.filter(r => passesDocFilters(r, 3));
    const gcUniquePaths = (skip: number): string[] => {
      const seen = new Set<string>();
      gcBaseRows.forEach(r => {
        const path = r.subFolderPath || '';
        if (skip <= 2 && docGroupLevelFilter !== 'all' && !pathMatchesGroupCategory(path, docGroupLevelFilter, 'all')) return;
        if (skip <= 1 && docLeafCategoryFilter !== 'all' && !pathMatchesGroupCategory(path, 'all', docLeafCategoryFilter)) return;
        seen.add(path);
      });
      return Array.from(seen);
    };
    // Every folder path known for this vessel/main folder — the nested tree,
    // the flat live-scan entries and anything loaded on demand.
    const knownFolderPaths: string[][] = (() => {
      const out: string[][] = [];
      const walk = (nodes: FolderTreeNode[], trail: string[]): void => {
        for (const node of nodes) {
          const own = [...trail, node.name];
          out.push(own);
          if (node.children.length) walk(node.children, own);
        }
      };
      walk(subfolderTree, []);
      liveSubfolderEntries.forEach(e => { if (e.pathNames && e.pathNames.length) out.push(e.pathNames); });
      const lazy = (host as any)._lazySubfolderEntries as Map<string, LiveSubfolderEntry> | undefined;
      lazy?.forEach((e, k) => { if (k.indexOf('path:') === 0 && e.pathNames && e.pathNames.length) out.push(e.pathNames); });
      return out;
    })();
    // Library-wide (All vessels + All main folders): nothing is scanned into
    // the tree yet, so category / sub-category choices come from the cached
    // recursive library walk (folder paths relative to the library root).
    const libraryWideScope = vesselFilter === 'all' && docCategoryFilter === 'all';
    // A vessel is picked: Sub-category options come from THAT vessel's folders
    // only (same Group > Category > Sub-category rule as the library-wide
    // taxonomy, applied to the part of each path below the vessel folder).
    // Before this the list was built from already-loaded file rows alone, so a
    // vessel opened at a folder-only level (e.g. "Drawings and Manuals") had no
    // sub-categories and the dropdown was disabled.
    const vesselTaxonomyKey = (!libraryWideScope && !!vesselFilter && vesselFilter !== 'all' && vesselFilter.trim().toLowerCase() !== 'not listed')
      ? vesselFilter.trim().toLowerCase() : '';
    const vesselTaxNames = new Set<string>();
    if (vesselTaxonomyKey) {
      vesselTaxNames.add(vesselTaxonomyKey);
      (siteVesselFolderPaths.get(vesselTaxonomyKey) || []).forEach(vp => {
        const last = (vp.split('/').pop() || '').trim().toLowerCase();
        if (last) vesselTaxNames.add(last);
      });
    }
    const vesselTaxScopeKey = Array.from(vesselTaxNames).sort().join('|');
    const taxonomyScoped = libraryWideScope || !!vesselTaxonomyKey;
    const libWalkPaths: string[][] = (() => {
      const out: string[][] = [];
      if (!libraryWideScope) return out;
      groupCatTreeCache.forEach((entry, key) => {
        if (entry.status !== 'done' || !/::root(::w:[^:]*)?$/.test(key)) return;
        entry.items.forEach((item: any) => {
          if (!item) return;
          const parts = String(item.path || item.name || '').split('/').filter(Boolean);
          const folderParts = item.is_folder ? parts : parts.slice(0, -1);
          if (folderParts.length) out.push(folderParts);
        });
      });
      return out;
    })();
    // Library-wide taxonomy: every Group > Category > Sub-category folder in
    // the whole library, read from ONE folder-only tree of the library root
    // (backend drive index). Without it nothing is known until a Group /
    // Category is picked (the recursive file walk only runs then), so under
    // All vessels / All main folders / All groups the Category list had no
    // "To Be Classified" entries and the Sub-category list was empty.
    // Memoised on the tree's folder array — it only changes when the tree is
    // re-fetched, not on every render.
    interface LibTaxonomyIndex {
      tbcCategories: string[];
      subs: Map<string, Map<string, string>>; // `${group}::${category}` (lower) -> lower -> display
      subMeta: Map<string, { group: string; category: string }>; // same key -> display group / category
      groups: string[]; // groups found in scope (display labels)
      cats: Map<string, string[]>; // group label (lower) -> categories found in scope
    }
    const libFolderTreeRes = (taxonomyScoped && liveLibraryResolved)
      ? host._getOrLoadSiteFolderTree(effectiveLiveSiteId, effectiveLiveDriveId, 'root', undefined, 20000)
      : null;
    const libTaxonomy: LibTaxonomyIndex = (() => {
      const empty: LibTaxonomyIndex = { tbcCategories: [], subs: new Map(), subMeta: new Map(), groups: [], cats: new Map() };
      const src = libFolderTreeRes?.folders;
      if (!src || src.length === 0) return empty;
      const memo = (host as any)._libTaxonomyMemo as { src: any; scope: string; index: LibTaxonomyIndex } | undefined;
      if (memo && memo.src === src && memo.scope === vesselTaxScopeKey) return memo.index;
      const tbc = new Map<string, string>();
      const subs = new Map<string, Map<string, string>>();
      const subMeta = new Map<string, { group: string; category: string }>();
      const groupsSeen = new Set<string>();
      const catsSeen = new Map<string, Map<string, string>>();
      const noteCat = (group: string, category: string): void => {
        const gk = group.toLowerCase();
        const m = catsSeen.get(gk) || new Map<string, string>();
        if (!m.has(category.trim().toLowerCase())) m.set(category.trim().toLowerCase(), category.trim());
        catsSeen.set(gk, m);
        groupsSeen.add(group);
      };
      const addSub = (group: string, category: string, sub: string): void => {
        const nm = (sub || '').trim();
        if (!nm || /\.[a-z0-9]{2,5}$/i.test(nm)) return;
        const key = `${group.toLowerCase()}::${category.trim().toLowerCase()}`;
        let m = subs.get(key);
        if (!m) { m = new Map(); subs.set(key, m); }
        if (!subMeta.has(key)) subMeta.set(key, { group, category: category.trim() });
        if (!m.has(nm.toLowerCase())) m.set(nm.toLowerCase(), nm);
      };
      const findKnown = (seg: string, list: string[]): string | undefined => {
        const low = seg.trim().toLowerCase();
        return list.find(c => c.toLowerCase() === low) || list.find(c => folderMatchesCategory(seg, c));
      };
      src.forEach(f => {
        let p = String(f.path || f.name || '').split('/').map(x => x.trim()).filter(Boolean);
        if (vesselTaxonomyKey) {
          // Vessel scope: keep only folders inside this vessel's own folder and
          // read the taxonomy from the part below it.
          const vi = p.findIndex(seg => vesselTaxNames.has(seg.toLowerCase()));
          if (vi < 0) return;
          p = p.slice(vi + 1);
        }
        let grp: '' | 'drawings' | 'manuals' | 'both' = '';
        for (let i = 0; i < p.length; i++) {
          const seg = p[i];
          if (isToBeClassifiedName(seg)) {
            groupsSeen.add(TBC_GROUP_LABEL);
            const cat = p[i + 1];
            if (cat) {
              if (!tbc.has(cat.toLowerCase())) tbc.set(cat.toLowerCase(), cat);
              noteCat(TBC_GROUP_LABEL, cat);
              if (p[i + 2]) addSub(TBC_GROUP_LABEL, cat, p[i + 2]);
            }
            return;
          }
          if (grp) {
            // The first known category folder below the group folder; the
            // folder right under it is the sub-category.
            const lists: Array<[string, string[]]> = grp === 'both'
              ? [['Drawings', drawingCategories], ['Manuals', manualCategories]]
              : grp === 'drawings' ? [['Drawings', drawingCategories]] : [['Manuals', manualCategories]];
            let hit = false;
            lists.forEach(([gl, list]) => {
              const cat = findKnown(seg, list);
              if (!cat) return;
              hit = true;
              noteCat(gl, cat);
              if (p[i + 1]) addSub(gl, cat, p[i + 1]);
            });
            if (hit) return;
          }
          const g = folderGroupsOf(seg);
          if (g.drawings !== g.manuals) groupsSeen.add(g.drawings ? 'Drawings' : 'Manuals');
          if (g.drawings || g.manuals) grp = g.drawings && g.manuals ? 'both' : (g.drawings ? 'drawings' : 'manuals');
        }
      });
      const index: LibTaxonomyIndex = {
        tbcCategories: Array.from(tbc.values()).sort((a, b) => a.localeCompare(b)),
        subs,
        subMeta,
        groups: Array.from(groupsSeen),
        cats: new Map(Array.from(catsSeen.entries()).map(([k, m]) => [k, Array.from(m.values())] as [string, string[]])),
      };
      (host as any)._libTaxonomyMemo = { src, scope: vesselTaxScopeKey, index };
      return index;
    })();
    // Library-wide taxonomy not known yet (tree still loading / the backend
    // index still building / a failed request being retried): the
    // Sub-category dropdown says so instead of looking empty.
    const libTaxonomyState: 'ready' | 'loading' | 'retrying' = (() => {
      if (!taxonomyScoped || !liveLibraryResolved) return 'ready';
      if (libTaxonomy.subs.size > 0 || libTaxonomy.tbcCategories.length > 0) return 'ready';
      if (!libFolderTreeRes) return 'loading';
      if (libFolderTreeRes.folders.length > 0) return 'ready';
      if (libFolderTreeRes.error && !libFolderTreeRes.loading) return 'retrying';
      return (libFolderTreeRes.loading || libFolderTreeRes.truncated) ? 'loading' : 'ready';
    })();
    const libSubsFor = (group: string, category: string): string[] => {
      const m = libTaxonomy.subs.get(`${group.toLowerCase()}::${category.trim().toLowerCase()}`);
      return m ? Array.from(m.values()) : [];
    };
    const groupOptions: string[] = (() => {
      let hasDrawings = false;
      let hasManuals = false;
      let hasToBeClassified = false;
      gcUniquePaths(3).forEach(p => {
        const segs = pathFolderSegments(p);
        const g = decisiveGroupOf(segs);
        const inTbc = segs.some(isToBeClassifiedName);
        if (g && g.drawings && !inTbc) hasDrawings = true;
        if (g && g.manuals && !inTbc) hasManuals = true;
        if (inTbc) hasToBeClassified = true;
      });
      // The visible rows shrink to whatever folder is open (e.g. once
      // Drawings > Electrical is picked only Drawings files remain), so also
      // read the groups off the folder tree itself — every group stays
      // selectable, with the current one highlighted.
      const walkGroups = (nodes: FolderTreeNode[], trail: string[]): void => {
        for (const node of nodes) {
          const own = [...trail, node.name];
          const g = folderGroupsOf(node.name);
          const inTbc = own.some(isToBeClassifiedName);
          if (!inTbc && g.drawings && !g.manuals) hasDrawings = true;
          if (!inTbc && g.manuals && !g.drawings) hasManuals = true;
          if (isToBeClassifiedName(node.name)) hasToBeClassified = true;
          if (node.children.length) walkGroups(node.children, own);
        }
      };
      walkGroups(subfolderTree, []);
      // The current selection must always be offered, otherwise the <select>
      // has no matching <option> and silently falls back to "All groups".
      if (docGroupLevelFilter === TBC_GROUP_LABEL) hasToBeClassified = true;
      if (docGroupLevelFilter.trim().toLowerCase() === 'drawings') hasDrawings = true;
      if (docGroupLevelFilter.trim().toLowerCase() === 'manuals') hasManuals = true;
      // "All vessels" + "All main folders" (e.g. the library root, before any
      // folder has been opened): nothing is scanned yet, but every group must
      // still be selectable — picking one searches the whole library.
      const atLibraryWide = vesselFilter === 'all' && docCategoryFilter === 'all';
      if (atLibraryWide) { hasDrawings = true; hasManuals = true; hasToBeClassified = true; }
      // Vessel selected: every group that exists inside THIS vessel's folders.
      if (vesselTaxonomyKey) {
        libTaxonomy.groups.forEach(gl => {
          if (gl === 'Drawings') hasDrawings = true;
          else if (gl === 'Manuals') hasManuals = true;
          else if (gl === TBC_GROUP_LABEL) hasToBeClassified = true;
        });
      }
      return [hasDrawings ? 'Drawings' : '', hasManuals ? 'Manuals' : '', hasToBeClassified ? TBC_GROUP_LABEL : ''].filter(Boolean);
    })();
    // Categories on offer for one Group ('all' = every group, merged). The
    // current category pick is always kept for the group it belongs to.
    const categoriesForGroup = (groupSel: string): string[] => {
      const isCurrentGroup = groupSel === docGroupLevelFilter;
      const normalizedGroup = (groupSel || 'all').trim().toLowerCase();
      // "To Be Classified" has no fixed Drawings/Manuals taxonomy: its
      // categories are simply the folders that were created directly inside
      // it (Basic, Electric Maker, Outfit, ...).
      if (normalizedGroup === TBC_GROUP_LABEL.toLowerCase()) {
        const kids = new Map<string, string>();
        knownFolderPaths.forEach(p => {
          if (p.length >= 2 && isToBeClassifiedName(p[p.length - 2])) {
            const nm = p[p.length - 1].trim();
            if (nm && !kids.has(nm.toLowerCase())) kids.set(nm.toLowerCase(), nm);
          }
        });
        libWalkPaths.forEach(p => {
          for (let i = 0; i + 1 < p.length; i++) {
            if (isToBeClassifiedName(p[i])) {
              const nm = p[i + 1].trim();
              if (nm && !/\.[a-z0-9]{2,5}$/i.test(nm) && !kids.has(nm.toLowerCase())) kids.set(nm.toLowerCase(), nm);
              break;
            }
          }
        });
        libTaxonomy.tbcCategories.forEach(nm => { if (!kids.has(nm.toLowerCase())) kids.set(nm.toLowerCase(), nm); });
        if (isCurrentGroup && docLeafCategoryFilter !== 'all' && !kids.has(docLeafCategoryFilter.trim().toLowerCase())) kids.set(docLeafCategoryFilter.trim().toLowerCase(), docLeafCategoryFilter);
        return Array.from(kids.values()).sort((a, b) => a.localeCompare(b));
      }
      // Vessel selected: every category folder of this vessel for the group
      // ('all' = merged across groups), independent of what is loaded.
      if (vesselTaxonomyKey) {
        const fromTax = (normalizedGroup === 'all' ? ['drawings', 'manuals'] : [normalizedGroup])
          .reduce((acc: string[], gk) => acc.concat(libTaxonomy.cats.get(gk) || []), [])
          .filter((v, i, arr) => arr.findIndex(x => x.toLowerCase() === v.toLowerCase()) === i);
        if (fromTax.length > 0) {
          if (isCurrentGroup && docLeafCategoryFilter !== 'all' && !fromTax.some(c => c.toLowerCase() === docLeafCategoryFilter.trim().toLowerCase())) fromTax.push(docLeafCategoryFilter);
          return fromTax.sort((a, b) => a.localeCompare(b));
        }
      }
      const candidates = normalizedGroup === 'drawings' ? drawingCategories
        : normalizedGroup === 'manuals' ? manualCategories
        : drawingCategories.concat(manualCategories).filter((v, i, arr) => arr.indexOf(v) === i);
      // Library-wide: the fixed taxonomy for the chosen group is always on offer.
      if (libraryWideScope) return candidates.slice();
      const paths = gcUniquePaths(2);
      // Same idea as the groups: also offer every category that exists as a
      // folder under the chosen group in the tree, not only those still
      // present in the (folder-narrowed) visible rows.
      const treeCategories = new Set<string>();
      const walkCats = (nodes: FolderTreeNode[], trail: string[]): void => {
        for (const node of nodes) {
          const own = [...trail, node.name];
          if (pathFitsGroupContext(own, { group: groupSel })) {
            candidates.forEach(c => { if (folderMatchesCategory(node.name, c)) treeCategories.add(c); });
          }
          if (node.children.length) walkCats(node.children, own);
        }
      };
      walkCats(subfolderTree, []);
      return candidates.filter(c =>
        treeCategories.has(c) ||
        (isCurrentGroup && c === docLeafCategoryFilter) ||
        paths.some(p => pathMatchesGroupCategory(p, groupSel, c)));
    };
    const categoryOptions: string[] = categoriesForGroup(docGroupLevelFilter);
    // "All groups": the same categories, classified under the group they
    // belong to (Drawings / Manuals / To Be Classified) — "Electrical" and
    // "Safety" exist in both Drawings and Manuals, and the To Be Classified
    // folders have their own. Picking one also selects its group.
    const categoryOptionGroups: Array<{ group: string; categories: string[] }> = docGroupLevelFilter === 'all'
      ? groupOptions
        .map(g => ({ group: g, categories: categoriesForGroup(g) }))
        .filter(x => x.categories.length > 0)
      : [];
    // A file has a sub-category only when one is really tagged / present in
    // its folder path — not when the value is just its category/group name
    // echoed back (see rowSubCat fallback in addFlatRow) or a "not assigned"
    // placeholder.
    const realSubCategoryOf = (r: FlatRow): string => {
      const labels = getListViewLabels(r);
      const same = new Set([
        labels.category, labels.group, labels.documentSection, labels.mainFolder, labels.vessel,
        r.category, r.group, r.vesselName,
      ].map(v => (v || '').trim().toLowerCase()).filter(Boolean));
      const ownCategory = (r.category || '').trim().toLowerCase();
      for (const cand of [labels.subCategory, r.subCategory]) {
        const v = (cand || '').trim();
        if (!v || /not assigned$/i.test(v) || /\.[a-z0-9]{2,5}$/i.test(v)) continue;
        const low = v.toLowerCase();
        // A file that has a sub-category but NO category: hierarchyForRow
        // copies the sub-category into labels.category as a display fallback,
        // which made it look like an echo and dropped it. It's only an echo
        // when the row's own recorded category is that same value.
        const echoesCategory = ownCategory === low;
        const subOnly = !echoesCategory && low === (labels.category || '').trim().toLowerCase();
        if (echoesCategory || (same.has(low) && !subOnly)) continue;
        return v;
      }
      return '';
    };
    const subCategoryOptions: string[] = (() => {
      const set = new Set<string>();
      const okPaths = new Set(gcUniquePaths(1));
      gcBaseRows.forEach(r => {
        if (!okPaths.has(r.subFolderPath || '')) return;
        const v = realSubCategoryOf(r);
        if (v) set.add(v);
      });
      // Library-wide (or vessel-scoped, via libTaxonomy below): folders sitting
      // directly under the chosen category (or under any category of the
      // chosen group) in the library walk.
      if (taxonomyScoped) {
        const grp = (docGroupLevelFilter || 'all').trim().toLowerCase();
        const cats = docLeafCategoryFilter !== 'all'
          ? [docLeafCategoryFilter]
          : (grp === 'drawings' ? drawingCategories : grp === 'manuals' ? manualCategories : drawingCategories.concat(manualCategories));
        libWalkPaths.forEach(p => {
          const joined = p.join(' > ');
          const inTbc = p.some(isToBeClassifiedName);
          if (grp === TBC_GROUP_LABEL.toLowerCase()) {
            for (let i = 0; i + 2 < p.length; i++) {
              if (!isToBeClassifiedName(p[i])) continue;
              if (docLeafCategoryFilter === 'all' || p[i + 1].trim().toLowerCase() === docLeafCategoryFilter.trim().toLowerCase()) {
                const nm = p[i + 2].trim();
                if (nm && !/\.[a-z0-9]{2,5}$/i.test(nm)) set.add(nm);
              }
              break;
            }
            return;
          }
          if (inTbc && grp !== 'all') return;
          if (!inTbc && grp !== 'all' && !pathMatchesGroupCategory(joined, docGroupLevelFilter, 'all')) return;
          for (let i = 0; i + 1 < p.length; i++) {
            if (cats.some(c => folderMatchesCategory(p[i], c)) && (grp === 'all' || pathMatchesGroupCategory(p.slice(0, i + 1).join(' > '), docGroupLevelFilter, 'all'))) {
              const nm = p[i + 1].trim();
              if (nm && !/\.[a-z0-9]{2,5}$/i.test(nm)) set.add(nm);
              break;
            }
          }
        });
        // ...and every sub-category folder of the library-wide taxonomy
        // (known before any Group / Category is picked).
        libTaxonomy.subs.forEach((m, key) => {
          const sep = key.indexOf('::');
          const kg = key.slice(0, sep);
          const kc = key.slice(sep + 2);
          if (grp !== 'all' && kg !== grp) return;
          if (docLeafCategoryFilter !== 'all' && kc !== docLeafCategoryFilter.trim().toLowerCase()) return;
          m.forEach(nm => set.add(nm));
        });
        if (docSubCategoryFilter !== 'all') set.add(docSubCategoryFilter);
      }
      return Array.from(set)
        .filter((v, i, arr) => arr.findIndex(x => x.toLowerCase() === v.toLowerCase()) === i)
        .sort((a, b) => a.localeCompare(b));
    })();
    // Library-wide with no Category picked: the sub-categories classified
    // under their Group > Category, so the same folder name under two
    // categories stays distinguishable. Picking one changes ONLY the
    // Sub-category filter (Groups / Categories keep the user's choice).
    const subCategoryOptionGroups: Array<{ group: string; category: string; subs: string[] }> = (() => {
      if (!taxonomyScoped || docLeafCategoryFilter !== 'all') return [];
      const out: Array<{ group: string; category: string; subs: string[] }> = [];
      if (vesselTaxonomyKey) {
        // Vessel-specific: every Group > Category that has sub-category folders
        // inside THIS vessel (not limited to categories already loaded).
        const order = ['drawings', 'manuals', TBC_GROUP_LABEL.toLowerCase()];
        const wantGroup = (docGroupLevelFilter || 'all').trim().toLowerCase();
        libTaxonomy.subs.forEach((m, key) => {
          const meta = libTaxonomy.subMeta.get(key);
          if (!meta) return;
          if (wantGroup !== 'all' && meta.group.toLowerCase() !== wantGroup) return;
          const subs = Array.from(m.values()).sort((a, b) => a.localeCompare(b));
          if (subs.length) out.push({ group: meta.group, category: meta.category, subs });
        });
        out.sort((a, b) => {
          const d = order.indexOf(a.group.toLowerCase()) - order.indexOf(b.group.toLowerCase());
          return d !== 0 ? d : a.category.localeCompare(b.category);
        });
        return out;
      }
      const groupsToShow = docGroupLevelFilter === 'all' ? groupOptions : [docGroupLevelFilter];
      groupsToShow.forEach(gl => {
        categoriesForGroup(gl).forEach(c => {
          const subs = libSubsFor(gl, c).sort((a, b) => a.localeCompare(b));
          if (subs.length) out.push({ group: gl, category: c, subs });
        });
      });
      return out;
    })();
    // Sub-categories that exist only on already-loaded rows (no taxonomy
    // folder for them) stay selectable as plain options above the groups.
    const subCategoryUngrouped: string[] = subCategoryOptionGroups.length === 0 ? subCategoryOptions : (() => {
      const grouped = new Set<string>();
      subCategoryOptionGroups.forEach(x => x.subs.forEach(v => grouped.add(v.toLowerCase())));
      return subCategoryOptions.filter(v => !grouped.has(v.toLowerCase()));
    })();

    // De-duplicate the same physical file when the SharePoint Sites scope
    // picked it up from more than one live source. `siteRows` above is built
    // from two independent scans (see getSharePointSiteFlatRows): (1) the
    // on-demand per-folder browse cache (_siteFolderItemsCache, populated
    // when the user actually opens a folder — correctly resolves the file's
    // vessel/path), and (2) the site-wide recursive tree
    // (documentLiveFolders / _loadDocumentLiveTree's own "live:" rows via
    // `rows`), which re-derives vessel/path itself and can mis-resolve a
    // flat (Part C) vessel's file as an unmatched root-level "Not Listed"
    // row when its own vessel-matching runs before `vessels` state is fully
    // populated. Both rows carry the same real fileId, so keep only the
    // best one per fileId: prefer whichever row actually resolved a vessel,
    // then whichever has the deeper/more specific folder path.
    if (docScopeType === 'sites') {
      const bestByFileId = new Map<string, FlatRow>();
      const passthrough: FlatRow[] = [];
      const scoreRow = (row: FlatRow): number => {
        const v = (row.vesselName || '').trim().toLowerCase();
        const hasVessel = v && v !== 'not listed' && v !== 'vessel name not listed' && v !== 'vessel not assigned' ? 1 : 0;
        const depth = (row.subFolderPath || '').split('>').length;
        return hasVessel * 1000 + depth;
      };
      for (const r of filtered) {
        if (!r.fileId || !r.fileName) { passthrough.push(r); continue; }
        const existing = bestByFileId.get(r.fileId);
        if (!existing || scoreRow(r) > scoreRow(existing)) {
          bestByFileId.set(r.fileId, r);
        }
      }
      filtered = [...passthrough, ...Array.from(bestByFileId.values())];
    }

    // ── Per-subfolder file count (for tile badge) ──
    // Count files (rows with fileName) that belong to each direct child subfolder.
    // Uses the already-scoped `filtered` list + live spoFolderMap for accuracy.
    const subfolderFileCountMap = new Map<string, number>();
    const subfolderSeenFileKeys = new Map<string, Set<string>>();

    // Pre-initialise all known subfolder names to 0 so tiles always get a number
    subfolderNames.forEach(sfn => { if (!subfolderFileCountMap.has(sfn)) subfolderFileCountMap.set(sfn, 0); });

    // Tally from filtered rows
    filtered.forEach(r => {
      const rowTail = getFolderTailSegments(r.subFolderPath, r.vesselName, r.group);
      if (rowTail.length > currentTail.length) {
        const startsWithCurrent = currentTail.every((seg, idx) => seg.toLowerCase() === rowTail[idx]?.toLowerCase());
        if (startsWithCurrent) {
          const sfKey = rowTail[currentTail.length];
          // Only count actual file rows to get the "files" number
          if (sfKey && r.fileName) {
            const sfNorm = sfKey.trim().toLowerCase();
            const fileIdentity = (r.fileName || r.fileId || '').trim().toLowerCase();
            const dedupeKey = `${sfNorm}||${fileIdentity}`;
            const seen = subfolderSeenFileKeys.get(sfNorm) || new Set<string>();
            if (!seen.has(dedupeKey)) {
              seen.add(dedupeKey);
              subfolderSeenFileKeys.set(sfNorm, seen);
              subfolderFileCountMap.set(sfKey, (subfolderFileCountMap.get(sfKey) || 0) + 1);
            }
          } else if (sfKey && !subfolderFileCountMap.has(sfKey)) {
            subfolderFileCountMap.set(sfKey, 0);
          }
        }
      }
    });

    // Tally from uploaded/live cache keys as fallback when row files are not
    // yet materialized for a section card.
    Object.entries(uploadedFilesByFolder || {}).forEach(([key, list]) => {
      if (!Array.isArray(list) || list.length === 0) return;
      if (host._appDeletedItemIds.has(key)) return;
      if (/^01[A-Za-z0-9]{15,}$/.test((key || '').trim())) return;

      if (key.includes('||')) {
        const parts = key.split('||');
        const kVessel = (parts[0] || '').trim().toLowerCase();
        const kGroup = (parts[1] || '').trim().toLowerCase();
        if (vesselName && kVessel && kVessel !== vesselName.toLowerCase()) return;
        if (docMainFolder && kGroup && kGroup !== docMainFolder.toLowerCase()) return;
      } else {
        const otherVessel = vessels.find(v => v.name && key.toLowerCase().includes(v.name.toLowerCase()));
        if (otherVessel && vesselName && otherVessel.name.toLowerCase() !== vesselName.toLowerCase()) return;
      }

      const keyTail = getFolderTailSegments(key, vesselName, docMainFolder);
      if (keyTail.length <= currentTail.length) return;
      const startsWithCurrent = currentTail.every((seg, idx) => seg.toLowerCase() === keyTail[idx]?.toLowerCase());
      if (!startsWithCurrent) return;

      const sfKey = keyTail[currentTail.length];
      if (!sfKey) return;
      const sfNorm = sfKey.trim().toLowerCase();
      const seen = subfolderSeenFileKeys.get(sfNorm) || new Set<string>();

      list.forEach((file: any) => {
        const name = String(file?.name || '').trim().toLowerCase();
        if (!name) return;
        const dedupeKey = `${sfNorm}||${name}`;
        if (seen.has(dedupeKey)) return;
        seen.add(dedupeKey);
        subfolderFileCountMap.set(sfKey, (subfolderFileCountMap.get(sfKey) || 0) + 1);
      });

      subfolderSeenFileKeys.set(sfNorm, seen);
      if (!subfolderFileCountMap.has(sfKey)) subfolderFileCountMap.set(sfKey, 0);
    });

    // Sub-folder count map (used alongside subfolderFileCountMap to show separate badges)
    const subfolderFolderCountMap = new Map<string, number>();
    subfolderNames.forEach(sfn => { if (!subfolderFolderCountMap.has(sfn)) subfolderFolderCountMap.set(sfn, 0); });

    // Override with live spoFolderMap counts when available (more accurate)
    if (resolvedCurrentFolderId && host.state.spoFolderMap.has(resolvedCurrentFolderId)) {
      const liveParent = host.state.spoFolderMap.get(resolvedCurrentFolderId);
      (liveParent?.children || []).forEach(child => {
        if (!child.isFolder || !child.name || (child.id && host._appDeletedItemIds.has(child.id))) return;
        const sfNode = child.id ? host.state.spoFolderMap.get(child.id) : null;
        const targetSfKey = subfolderNames.find(s => s.toLowerCase() === child.name.toLowerCase()) || child.name;
        if (sfNode) {
          const liveFiles = (sfNode.children || []).filter(
            fc => !fc.isFolder && (!fc.id || !host._appDeletedItemIds.has(fc.id))
          );
          const liveFolders = (sfNode.children || []).filter(
            fc => fc.isFolder && (!fc.id || !host._appDeletedItemIds.has(fc.id))
          );
          const liveFileNames = new Set(liveFiles.map(fc => (fc.name || '').trim().toLowerCase()));
          const sfNorm = child.name.trim().toLowerCase();

          // Count pending/optimistic local uploads not yet synced to spoFolderMap
          let pendingLocalCount = 0;
          const seen = subfolderSeenFileKeys.get(sfNorm);
          if (seen) {
            seen.forEach(key => {
              const fileName = key.split('||')[1];
              if (fileName && !liveFileNames.has(fileName)) {
                pendingLocalCount++;
              }
            });
          }

          const totalLiveCount = liveFiles.length + pendingLocalCount;
          subfolderFileCountMap.set(targetSfKey, totalLiveCount);
          subfolderFolderCountMap.set(targetSfKey, liveFolders.length);
        } else {
          if (!subfolderFileCountMap.has(targetSfKey)) subfolderFileCountMap.set(targetSfKey, 0);
          if (!subfolderFolderCountMap.has(targetSfKey)) subfolderFolderCountMap.set(targetSfKey, 0);
        }
      });
    }

    // ── Group filtered rows by folder path (category/folder level) ──
    const groupedMap = new Map<string, GroupedRow>();
    for (const r of filtered) {
      const rowTail = getFolderTailSegments(r.subFolderPath, r.vesselName, r.group).map(s => s.toLowerCase());
      const canonicalRowKey = [
        (r.vesselName || '').trim().toLowerCase(),
        (r.group || '').trim().toLowerCase(),
        ...rowTail,
      ].filter(Boolean).join('||');
      const baseDedupeKey = canonicalRowKey || (r.subFolderPath || r.groupKey).trim().toLowerCase();
      const dedupeKey = docScopeType === 'sites' && r.fileName
        ? `${baseDedupeKey}||file:${(r.fileId || r.fileName).trim().toLowerCase()}`
        : baseDedupeKey;
      const normSub = (r.subFolderPath || '').trim().toLowerCase();
      const rowTailKey = rowTail.join(' > ');
      const subLower = (r.subCategory || r.category || '').trim().toLowerCase();
      const vLower = (r.vesselName || '').trim().toLowerCase();
      const gLower = (r.group || '').trim().toLowerCase();
      const normBreadcrumb = (p: string) => (p || '').replace(/\s*>\s*/g, ' > ').trim().toLowerCase();
      const sharePointPathLower = host._sharePointFolderPath(r.subFolderPath, '').trim().toLowerCase();
      const sharePointPathBreadLower = sharePointPathLower ? sharePointPathLower.split('/').filter(Boolean).join(' > ') : '';
      const normSubBreadcrumb = normBreadcrumb(normSub);
      const rowPathParts = normSubBreadcrumb.split(' > ').map(s => s.trim()).filter(Boolean);
      const swappedPathLower = (rowPathParts.length >= 2)
        ? [rowPathParts[1], rowPathParts[0], ...rowPathParts.slice(2)].join(' > ')
        : '';
      const liveId = host._getLiveSharePointFolderId(r.subFolderPath);
      const liveNode = liveId ? host.state.spoFolderMap.get(liveId) : null;
      const memFiles = (liveNode?.children || []).filter(c => !c.isFolder).map(c => ({
        id: c.id,
        name: c.name,
        size: typeof c.size === 'number' ? `${(c.size / 1024).toFixed(1)} KB` : '—',
        date: (c as any).createdDateTime ? new Date((c as any).createdDateTime).toLocaleString() : ((c as any).lastModifiedDateTime ? new Date((c as any).lastModifiedDateTime).toLocaleString() : 'Today'),
        pending: false,
        uploadedAt: (c as any).createdDateTime ? Date.parse((c as any).createdDateTime) : ((c as any).lastModifiedDateTime ? Date.parse((c as any).lastModifiedDateTime) : undefined),
      }));
      const siteCacheFiles: Array<{ id: string; name: string; size?: string; uploadedAt?: number }> = [];
      if (r.uploadFolderId) {
        for (const [cacheKey, cacheEntry] of Array.from(host._siteFolderItemsCache.entries())) {
          if (!cacheKey.endsWith(`::${r.uploadFolderId}`)) continue;
          for (const item of cacheEntry.items || []) {
            const isFile = Boolean(item?.file) || (Boolean(item?.name) && !item.folder && !item.name.includes('/'));
            if (!isFile || !item.name) continue;
            siteCacheFiles.push({
              id: item.id || item.name,
              name: item.name,
              size: typeof item.size === 'number' ? `${(item.size / 1024).toFixed(1)} KB` : undefined,
              uploadedAt: item.lastModifiedDateTime ? Date.parse(item.lastModifiedDateTime) : (item.createdDateTime ? Date.parse(item.createdDateTime) : undefined),
            });
          }
        }
      }

      // Collect files for this row across all possible keys
      const candidateKeys = [
        r.groupKey,
        r.subFolderPath,
        sharePointPathLower,
        dedupeKey,
        liveId,
        (r.uploadFolderId && !/^\d+$/.test(r.uploadFolderId)) ? r.uploadFolderId : '',
      ].filter(Boolean) as string[];

      const expandedCandidateKeys = new Set<string>(candidateKeys.flatMap(keyVariants));

      const matchedUploads: Array<{ name: string; size?: string; date?: string; id?: string; uploadedAt?: number }> = [];
      const seenUploadNames = new Set<string>();

      for (const k of Array.from(expandedCandidateKeys)) {
        const list = uploadedFilesByFolder[k];
        if (Array.isArray(list)) {
          for (const item of list) {
            if (item?.name && !seenUploadNames.has(item.name.toLowerCase())) {
              seenUploadNames.add(item.name.toLowerCase());
              matchedUploads.push(item);
            }
          }
        }
      }

      // Also scan spoFolderMap for this exact row path only (avoid collisions for repeated names like "To be Classified")
      for (const [nodeId, node] of Array.from(host.state.spoFolderMap.entries())) {
        if (!node || node.deleted) continue;
        const nodePath = (node.serverRelativePath || '').toLowerCase();
        const nodeName = (node.name || '').toLowerCase();
        const matchVessel = vLower ? nodePath.includes(vLower) : true;
        const matchGroup = gLower ? nodePath.includes(gLower) : true;
        const matchFolder = subLower ? (nodeName === subLower || nodePath.endsWith(`/${subLower}`) || nodePath.includes(`/${subLower}/`)) : false;
        const matchExactPath = sharePointPathLower
          ? (nodePath === sharePointPathLower || nodePath.endsWith(`/${sharePointPathLower}`))
          : false;

        if (matchVessel && matchGroup && matchFolder && matchExactPath) {
          if (node.isFolder && Array.isArray(node.children)) {
            for (const c of node.children) {
              if (!c.isFolder && c.name && !seenUploadNames.has(c.name.toLowerCase())) {
                seenUploadNames.add(c.name.toLowerCase());
                matchedUploads.push({
                  id: c.id,
                  name: c.name,
                  size: typeof c.size === 'number' ? `${(c.size / 1024).toFixed(1)} KB` : '—',
                  uploadedAt: (c as any).createdDateTime ? Date.parse((c as any).createdDateTime) : ((c as any).lastModifiedDateTime ? Date.parse((c as any).lastModifiedDateTime) : undefined),
                });
              }
            }
          }
          const nodeUploads = uploadedFilesByFolder[nodeId];
          if (Array.isArray(nodeUploads)) {
            for (const item of nodeUploads) {
              if (item?.name && !seenUploadNames.has(item.name.toLowerCase())) {
                seenUploadNames.add(item.name.toLowerCase());
                matchedUploads.push(item);
              }
            }
          }
        }
      }

      // Also scan rows in state, but only exact same folder path/groupKey (no loose subCategory matching)
      for (const stRow of (rows || [])) {
        const stPath = normBreadcrumb(stRow.subFolderPath || '');
        const stTail = getFolderTailSegments(stRow.subFolderPath || '', stRow.vesselName || '', stRow.group || '').map(s => s.toLowerCase()).join(' > ');
        if (
          stRow.fileName &&
          stRow.vesselName?.toLowerCase() === vLower &&
          (stRow.group || '').toLowerCase() === gLower &&
          (
            (stRow.groupKey && stRow.groupKey === r.groupKey) ||
            stPath === normBreadcrumb(normSub) ||
            (rowTailKey && stTail === rowTailKey)
          )
        ) {
          if (!seenUploadNames.has(stRow.fileName.toLowerCase())) {
            seenUploadNames.add(stRow.fileName.toLowerCase());
            matchedUploads.push({
              id: stRow.fileId || stRow.fileName,
              name: stRow.fileName,
              size: stRow.fileSize,
              uploadedAt: stRow.fileUploadedAt,
            });
          }
        }
      }

      const folderUploads = [
        ...matchedUploads,
        ...memFiles,
        ...(r.fileName ? [] : siteCacheFiles),
      ];

      const existing = groupedMap.get(dedupeKey);
      if (!existing) {
        const files: Array<{ id: string; name: string; size?: string; uploadedAt?: number }> = [];
        const addOrUpdateFile = (file: { id: string; name: string; size?: string; uploadedAt?: number }): void => {
          const current = files.find(item => item.name.toLowerCase() === file.name.toLowerCase());
          if (!current) {
            files.push(file);
          } else {
            if ((!current.size || current.size === '—') && file.size && file.size !== '—') current.size = file.size;
            if (!current.uploadedAt && file.uploadedAt) current.uploadedAt = file.uploadedAt;
            if (current.id === current.name && file.id) current.id = file.id;
          }
        };
        if (r.fileName) {
          addOrUpdateFile({ id: r.fileId || r.fileName, name: r.fileName, size: r.fileSize, uploadedAt: r.fileUploadedAt });
        }
        for (const f of folderUploads) {
          if (f?.name) addOrUpdateFile({ id: (f as any).id || f.name, name: f.name, size: (f as any).size, uploadedAt: (f as any).uploadedAt });
        }
        groupedMap.set(dedupeKey, {
          srNo: r.srNo,
          vesselName: r.vesselName,
          domain: r.domain,
          group: r.group,
          category: r.category,
          subCategory: r.subCategory || r.category,
          subFolderPath: r.subFolderPath,
          groupKey: r.groupKey,
          uploadFolderId: r.uploadFolderId,
          monthDriven: r.monthDriven,
          canUpload: r.canUpload,
          files,
        });
      } else {
        if (r.fileName) {
          const current = existing.files.find(f => f.name.toLowerCase() === r.fileName!.toLowerCase());
          if (!current) {
            existing.files.push({ id: r.fileId || r.fileName, name: r.fileName, size: r.fileSize, uploadedAt: r.fileUploadedAt });
          } else {
            if ((!current.size || current.size === '—') && r.fileSize) current.size = r.fileSize;
            if (!current.uploadedAt && r.fileUploadedAt) current.uploadedAt = r.fileUploadedAt;
            if (current.id === current.name && r.fileId) current.id = r.fileId;
          }
        }
        for (const f of folderUploads) {
          const current = f?.name ? existing.files.find(ex => ex.name.toLowerCase() === f.name.toLowerCase()) : undefined;
          if (current) {
            if ((!current.size || current.size === '—') && (f as any).size && (f as any).size !== '—') current.size = (f as any).size;
            if (!current.uploadedAt && (f as any).uploadedAt) current.uploadedAt = (f as any).uploadedAt;
          } else if (f?.name) {
            existing.files.push({ id: (f as any).id || f.name, name: f.name, size: (f as any).size, uploadedAt: (f as any).uploadedAt });
          }
        }
        if (r.uploadFolderId && !r.uploadFolderId.includes('/') && (!existing.uploadFolderId || existing.uploadFolderId.includes('/'))) {
          existing.uploadFolderId = r.uploadFolderId;
        }
      }
    }

    let groupedList = Array.from(groupedMap.values());
    // NOTE: this used to unconditionally drop every row with zero files
    // whenever docScopeType === 'sites', which made sense back when every
    // vessel was auto-provisioned with the full department/category
    // template (dozens of near-always-empty leaf folders per vessel) —
    // without it, List View would have been mostly noise. Since the
    // backend moved to the flat one-root-folder-per-vessel model (no more
    // auto-created subtree; see real_backend.py create_vessel / Part C),
    // a brand-new vessel's root folder IS the row, and it has no files
    // until something is uploaded — so this filter was hiding every
    // freshly created vessel from List View entirely, even though the
    // same folder shows up immediately in Folder View (which never
    // applied this filter). Keep every real folder/file row here so List
    // View matches what Folder View already shows; only actually-attached
    // vs. not-attached filtering is still controlled by attachmentFilter
    // below.
    if (attachmentFilter !== 'all') {
      groupedList = groupedList.filter(row => attachmentFilter === 'attached' ? row.files.length > 0 : row.files.length === 0);
    }
    if (docListSort === 'name_az') {
      groupedList.sort((a, b) => a.category.localeCompare(b.category));
    } else {
      // 'default' and 'newest' both float recently-uploaded rows to the top,
      // ordered by the real upload timestamp. Rows with no timestamped files
      // (nothing uploaded through the app yet) keep their original relative
      // order — this is a stable sort, so nothing else shuffles around.
      groupedList = groupedList
        .map((row, i) => ({
          row, i,
          ts: row.files.reduce((max, f) => Math.max(max, f.uploadedAt || 0), 0),
        }))
        .sort((a, b) => (b.ts - a.ts) || (a.i - b.i))
        .map(x => x.row);
    }

    const totalPages = Math.max(1, Math.ceil(groupedList.length / PAGE_ROWS));
    const safePage = Math.min(docListPage, totalPages - 1);
    const pageGroupedRows = groupedList.slice(safePage * PAGE_ROWS, (safePage + 1) * PAGE_ROWS);

    const MAIN_FOLDER_SET = new Set(MAIN_FOLDERS.map(main => main.key.toLowerCase()));
    const COMMON_FOLDER_NAMES = new Set([
      'common for all ships',
      'common for all vessels',
      'common (not ship specific)',
      'common (for all ships)',
      'common',
    ]);

    const openFolderViewForListRow = (row: GroupedRow | null | undefined): void => {
      if (!row || !row.subFolderPath) {
        host.setState({ docViewMode: 'folder' });
        return;
      }

      const segments = (row.subFolderPath || '').split('>').map(s => s.trim()).filter(Boolean);
      if (segments.length === 0) {
        host.setState({ docViewMode: 'folder' });
        return;
      }

      const lower0 = segments[0].toLowerCase();
      const stack: Array<{ id: string; name: string }> = [];
      let nextMainFolder: MainFolderKey | null = null;
      let nextScope: 'vessels' | 'common' | 'kaizen' | 'sites' | 'shared_docs' | 'documents' = 'vessels';
      let nextVesselFilter = vesselFilter;

      const byLower = (name: string): MainFolderKey | null => {
        const found = MAIN_FOLDERS.find(main => main.key.toLowerCase() === name.toLowerCase());
        return found ? found.key : null;
      };

      if (lower0 === 'kaizen - knowledge bank') {
        nextMainFolder = 'Kaizen - Knowledge Bank';
        nextScope = 'kaizen';
        nextVesselFilter = 'all';
        stack.push({ id: 'kaizen_root', name: 'Kaizen - Knowledge Bank' });

        for (let i = 1; i < segments.length; i++) {
          const name = segments[i];
          const breadcrumb = segments.slice(0, i + 1).join(' > ');
          const liveId = host._getLiveSharePointFolderId(breadcrumb);
          const id = liveId || `sf_${i}_${name.replace(/\s+/g, '_').toLowerCase()}`;
          stack.push({ id, name });
        }
      } else if (lower0 === 'sharepoint sites' || lower0 === 'sites documents') {
        nextMainFolder = 'SharePoint Sites';
        nextScope = 'sites';
        nextVesselFilter = 'all';
        stack.push({ id: 'sites_root', name: 'SharePoint Sites' });

        const siteName = segments[1] || '';
        const libraryName = segments[2] || '';
        const site = (host.state.documentSites || []).find(s =>
          s.sp_site_name?.trim().toLowerCase() === siteName.toLowerCase() ||
          s.site_key?.trim().toLowerCase() === siteName.toLowerCase()
        );
        stack.push({
          id: site?.site_id ? `site:${site.site_id}` : `site_${siteName.replace(/\s+/g, '_').toLowerCase()}`,
          name: site?.sp_site_name || siteName,
        });
        const library = site && (site.default_library_name || '').trim().toLowerCase() === libraryName.toLowerCase()
          ? site.default_library_name
          : libraryName;
        stack.push({
          id: site?.drive_id ? `drive:${site.drive_id}` : `drive_${libraryName.replace(/\s+/g, '_').toLowerCase()}`,
          name: library || libraryName,
        });

        for (let i = 3; i < segments.length; i++) {
          const name = segments[i];
          const isLeaf = i === segments.length - 1;
          const subPathTillNow = segments.slice(0, i + 1).join(' > ');
          const matchingRow = host.state.rows.find(r =>
            r.subFolderPath && r.subFolderPath.trim().toLowerCase() === subPathTillNow.toLowerCase() &&
            r.uploadFolderId && !/^(sf_|category_|common|kaizen_root|dept_|sites_root|site:|drive:)/.test(r.uploadFolderId)
          );
          const driveRelPath = segments.slice(3, i + 1).join('/');
          const id = (isLeaf && row.uploadFolderId && !row.uploadFolderId.startsWith('sf_'))
            ? row.uploadFolderId
            : (matchingRow?.uploadFolderId || driveRelPath);
          stack.push({ id, name });
        }
      } else if (MAIN_FOLDER_SET.has(lower0)) {
        const mainFolder = byLower(segments[0]);
        if (mainFolder) {
          nextMainFolder = mainFolder;
          stack.push({ id: mainFolder, name: mainFolder });

          let startIdx = 1;
          if (segments.length > 1) {
            const second = segments[1];
            if (COMMON_FOLDER_NAMES.has(second.toLowerCase())) {
              nextScope = 'common';
              nextVesselFilter = 'all';
              stack.push({ id: 'common', name: 'Common for all ships' });
              startIdx = 2;
            } else {
              nextScope = 'vessels';
              nextVesselFilter = second;
              const vesselId = vessels.find(v => v.name.trim().toLowerCase() === second.trim().toLowerCase())?.id || second;
              stack.push({ id: String(vesselId), name: second });
              startIdx = 2;
            }
          }

          for (let i = startIdx; i < segments.length; i++) {
            const name = segments[i];
            const breadcrumb = segments.slice(0, i + 1).join(' > ');
            const liveId = host._getLiveSharePointFolderId(breadcrumb);
            const isLeaf = i === segments.length - 1;
            const rowId = row.uploadFolderId && !/^\d+$/.test(row.uploadFolderId) ? row.uploadFolderId : null;
            const id = liveId || (isLeaf ? rowId : null) || `sf_${i}_${name.replace(/\s+/g, '_').toLowerCase()}`;
            stack.push({ id, name });
          }
        }
      }

      if (stack.length > 0) {
        host._pushFolderNav(stack, nextMainFolder);
        host.setState({
          docViewMode: 'folder',
          docScopeType: nextScope,
          vesselFilter: nextVesselFilter || 'all',
        });

        // Immediately refresh live files for the destination folder across all scopes
        triggerFolderRefresh(stack, nextScope, nextVesselFilter, row?.groupKey || '');
      } else {
        host.setState({ docViewMode: 'folder' });
      }
    };

    const openFolderViewFromListContext = (): void => {
      if (docViewMode !== 'list') {
        host.setState({ docViewMode: 'folder' });
        return;
      }

      // Prefer navigating straight to whatever live Main folder / Sub-folder
      // is currently selected (now preserved across the switch — see
      // openListViewFromFolderContext above) over guessing from a selected
      // row: an arbitrary candidate row's own path can point somewhere else
      // entirely, which is exactly the "doesn't preserve the folder" symptom
      // reported when switching back to Folder view.
      if (docCategoryFilter !== 'all') {
        if (docSubfolderOtherFilter !== 'all' && liveSubfolderEntryByName.has(docSubfolderOtherFilter.trim().toLowerCase())) {
          navigateToLiveSubfolder(docSubfolderOtherFilter, (host as any)._subfolderSelectedPath as string[] | undefined);
          host.setState({ docViewMode: 'folder' });
          return;
        }
        const liveMainFolderCard = rootFolderCards.find(c => c.key.trim().toLowerCase() === docCategoryFilter.trim().toLowerCase());
        if (liveMainFolderCard && liveMainFolderCard.liveFolderId) {
          navigateToLiveMainFolder(docCategoryFilter);
          host.setState({ docViewMode: 'folder' });
          return;
        }
      }

      const selectedRow = pageGroupedRows.find(row => row.files.some(file => listViewSelectedFiles.has(file.id)));
      const candidate = selectedRow || pageGroupedRows[0] || groupedList[0] || null;
      openFolderViewForListRow(candidate);
    };

    const openListViewFromFolderContext = (): void => {
      if (docViewMode !== 'folder') {
        host.setState({ docViewMode: 'list' });
        return;
      }

      host._cancelDocumentLiveTree();

      let nextScope: 'vessels' | 'common' | 'kaizen' | 'sites' | 'shared_docs' | 'documents' = host.state.docScopeType || 'vessels';
      let nextVesselFilter = vesselFilter || 'all';
      let nextGroupFilter = docGroupFilter || 'all';
      let hierarchyLevels: string[] = [];

      if (atKaizenRoot) {
        nextScope = 'kaizen';
        nextVesselFilter = 'all';
        nextGroupFilter = 'Kaizen - Knowledge Bank';
        hierarchyLevels = folderPathStack
          .filter(n => n.id !== 'kaizen_root' && n.name !== 'Kaizen - Knowledge Bank')
          .map(n => n.name);
      } else if (atSharedDocsRoot || atDocsRoot || atSitesRoot) {
        nextScope = atSharedDocsRoot ? 'shared_docs' : (atDocsRoot ? 'documents' : 'sites');
        nextGroupFilter = atSharedDocsRoot ? 'Shared Documents' : (atDocsRoot ? 'Documents' : 'SharePoint Sites');
        // Keep the vessel already picked in the dropdown; otherwise take it
        // from a vessel folder in the breadcrumb (same folder→vessel matching
        // as the dropdown, so the value is always one of its options).
        nextVesselFilter = vesselFilter && vesselFilter !== 'all' ? vesselFilter : 'all';
        if (nextVesselFilter === 'all') {
          for (const node of folderPathStack) {
            const matched = matchFolderToVessel(node.name);
            if (matched) {
              nextVesselFilter = matched;
              break;
            }
          }
        }
        hierarchyLevels = [];
      } else if (atCommonShips) {
        nextScope = 'common';
        nextVesselFilter = 'all';
        nextGroupFilter = docMainFolder || nextGroupFilter || 'all';
        hierarchyLevels = folderPathStack.slice(2).map(n => n.name);
      } else if (vesselNodeInStack) {
        nextScope = 'vessels';
        nextVesselFilter = currentVesselNameFromStack || nextVesselFilter || 'all';
        nextGroupFilter = docMainFolder || nextGroupFilter || 'all';
        hierarchyLevels = folderPathStack.slice(2).map(n => n.name);
      } else if (atMainDepartment) {
        nextScope = 'vessels';
        nextGroupFilter = docMainFolder || folderPathStack[0]?.name || nextGroupFilter || 'all';
        nextVesselFilter = 'all';
      }

      const level0 = hierarchyLevels[0] || 'all';
      const level1 = hierarchyLevels[1] || 'all';
      const level2 = hierarchyLevels[2] || 'all';
      const level3 = hierarchyLevels[3] || 'all';

      if (nextScope === 'sites' && folderPathStack.length >= 2) {
        const siteNode = folderPathStack[1];
        const driveNode = folderPathStack[2];
        const site = (host.state.documentSites || []).find(s =>
          s.site_id === (siteNode?.id || '').replace(/^site:/, '') ||
          s.site_key === siteNode?.name ||
          s.sp_site_name === siteNode?.name
        );
        const siteId = site?.site_id || '';
        const driveId = (driveNode?.id || '').replace(/^drive:/, '') || site?.drive_id || '';
        const folderId = folderPathStack.length <= 3
          ? 'root'
          : (folderPathStack[folderPathStack.length - 1]?.id || 'root');
        if (siteId && driveId) {
          // List View must show every file *underneath* the folder the user
          // was browsing, not just the files of folders they happened to
          // click into one at a time in Folder view. Folder view only ever
          // loads one level (the folder currently open), so switching to
          // List View right after opening "Documents" — before manually
          // drilling into each of its sub-folders — used to flatten to
          // nothing, even though Folder view clearly showed sub-folders
          // with content. Walk the subtree recursively (bounded, so a huge
          // library can't hang the browser or spam Graph) instead of the
          // previous one-extra-level prefetch, and let already-cached
          // folders resolve instantly (_loadAndCacheSiteFolderChildren
          // reuses _siteFolderItemsCache, the same cache _getOrLoadSiteFolderChildren
          // reads from).
          void host._prefetchSiteSubtree(siteId, driveId, folderId);
        }
      }

      if (nextScope === 'vessels' && host.state.rows.length === 0) {
        void host._loadData(true).catch(() => undefined);
      }

      // The live SharePoint-sites folder tree (Main folder / Sub-folder
      // dropdowns) is a different navigation concept from the legacy
      // hierarchyLevels computed above (which is empty for this scope
      // anyway — see nextScope === 'sites' branch), so switching to List
      // view from there must keep whatever Main folder / Sub-folder was
      // selected rather than stomping it with level0/level3 — that stomp
      // (always 'all' for this scope) is what made the filters visibly
      // reset to "All main folders" / "All sub-folders" on every Folder ->
      // List switch. Every other scope keeps its existing behavior.
      // hierarchyLevels is always [] for the live SharePoint-sites folder
      // tree (see above), so level0..level3 are always 'all' here — that
      // used to blow away docGroupLevelFilter / docLeafCategoryFilter /
      // docSubCategoryFilter (along with docCategoryFilter /
      // docSubfolderOtherFilter) on every single Folder -> List switch,
      // which is the "filters go blank when I switch views repeatedly"
      // symptom for this scope. docCategoryFilter/docSubfolderOtherFilter
      // were already carved out as an exception (preserved instead of
      // stomped); the other three hierarchy filters need the same
      // treatment for consistency.
      const isLiveSitesFolderNav = folderPathStack[0]?.id === 'sites_root';
      // "All main folders" picked before the vessel: List view must span all
      // of the vessel's main folders, not just the one Folder view opened.
      const nextDocCategoryFilter: string = isLiveSitesFolderNav
        ? ((implicitMainAll && vesselMainFolders.length > 1) ? 'all' : docCategoryFilter)
        : level0;
      const nextDocSubfolderOtherFilter: string = isLiveSitesFolderNav ? docSubfolderOtherFilter : 'all';
      const nextDocGroupLevelFilter: string = isLiveSitesFolderNav ? docGroupLevelFilter : level1;
      const nextDocLeafCategoryFilter: string = isLiveSitesFolderNav ? docLeafCategoryFilter : level2;
      const nextDocSubCategoryFilter: string = isLiveSitesFolderNav ? docSubCategoryFilter : level3;
      host.setState({
        docViewMode: 'list',
        docScopeType: nextScope,
        vesselFilter: nextVesselFilter || 'all',
        docGroupFilter: nextGroupFilter || 'all',
        docCategoryFilter: nextDocCategoryFilter,
        docSubfolderOtherFilter: nextDocSubfolderOtherFilter,
        docGroupLevelFilter: nextDocGroupLevelFilter,
        docLeafCategoryFilter: nextDocLeafCategoryFilter,
        docSubCategoryFilter: nextDocSubCategoryFilter,
        catFilter: 'all',
        textFilter: '',
        docListPage: 0,
      });
    };

    // In List View: trigger on-demand live file refresh for visible page rows if not loaded yet.
    // Skip for sites/shared_docs/documents scopes — those rows are loaded via _getOrLoadSiteFolderChildren
    // which correctly uses each site's own siteId/driveId. _refreshFolderFiles uses this.props.siteId/driveId
    // (the webpart's main site) and would produce 404 Graph errors for cross-site folder IDs.
    if (docViewMode === 'list' && docScopeType !== 'sites' && docScopeType !== 'shared_docs' && docScopeType !== 'documents') {
      pageGroupedRows.forEach(row => {
        if ((row.subFolderPath || '').startsWith('SharePoint Sites') || (row.groupKey || '').startsWith('spo_') || (row.groupKey || '').startsWith('live:')) {
          return;
        }
        const liveId = host._getLiveSharePointFolderId(row.subFolderPath) || row.uploadFolderId;
        if (liveId && !/^(sf_|category_|common|vessels_root|specific_vessels|kaizen_root)/.test(liveId) && !/^\d+$/.test(liveId)) {
          const alreadyLoaded = host._filesLoadedForFolders.has(liveId);
          const inFlight = host._refreshFolderFilesInFlight.has(liveId);
          if (!alreadyLoaded && !inFlight) {
            setTimeout(() => void host._refreshFolderFiles(liveId, row.groupKey, true, true).catch(() => undefined), 0);
          }
        }
      });
    }

    if (activeVesselName && !host._filesLoadedForVessels.has(activeVesselName)) {
      host._filesLoadedForVessels.add(activeVesselName);
      setTimeout(() => void host._loadFilesForVessel(activeVesselName).catch(() => undefined), 0);
    }

    // Navigation folder names that should never appear as vessel cards
    const NAV_FOLDER_NAMES = new Set(['specific vessels', 'common for all ships', 'common (not ship specific)', 'vessels', 'kaizen - knowledge bank', 'knowledge bank', 'common']);
    const KNOWN_SPO_VESSELS: VesselRecord[] = [
      { id: 'spo_snow_flower', name: 'Snow Flower', status: 'Active', is_provisioned: true },
      { id: 'spo_snow_flake', name: 'Snow Flake', status: 'Active', is_provisioned: true },
      { id: 'spo_senegal_express', name: 'Senegal Express', status: 'Active', is_provisioned: true },
      { id: 'spo_peissy', name: 'Peissy', status: 'Active', is_provisioned: true },
      { id: 'spo_potiniere', name: 'Potiniere', status: 'Active', is_provisioned: true },
      { id: 'spo_norse_new_haven', name: 'Norse New Haven', status: 'Active', is_provisioned: true },
      { id: 'spo_norse_ijmuiden', name: 'Norse Ijmuiden', status: 'Active', is_provisioned: true },
      { id: 'spo_belle_lune', name: 'Belle Lune', status: 'Active', is_provisioned: true },
      { id: 'spo_bow_fighter', name: 'Bow Fighter', status: 'Active', is_provisioned: true },
    ];
    // Deduplicate by name and exclude any vessel whose name matches a navigation folder
    const displayVessels = (() => {
      const combined = [...vessels, ...KNOWN_SPO_VESSELS];
      const baseList = combined.filter(
        (v, i, arr) =>
          arr.findIndex(x => x.name.trim().toLowerCase() === v.name.trim().toLowerCase()) === i &&
          !NAV_FOLDER_NAMES.has(v.name.trim().toLowerCase())
      );
      if (vesselFilter !== 'all') {
        const matching = baseList.filter(v => v.name.trim().toLowerCase() === vesselFilter.trim().toLowerCase());
        if (matching.length > 0) return matching;
        return [{ id: vesselFilter, name: vesselFilter }];
      }
      return baseList.slice(0, Math.max(documentVesselCount, baseList.length));
    })();

    const handleDocumentsPageDrop = async (e: React.DragEvent): Promise<void> => {
      e.preventDefault();
      e.stopPropagation();
      if (!e.dataTransfer) return;
      const extracted = await extractFilesFromDataTransfer(e.dataTransfer);
      if (!extracted.length) return;

      const topFolderId = currentFolderNode && !/^(sf_|category_|common|vessels_root|specific_vessels|kaizen_root|sites_root|site:|drive:)/.test(currentFolderNode.id) ? currentFolderNode.id : '';
      let currentVessel = '';
      let subFolderPath = '';
      let fallbackPath = '';
      let targetSiteId: string | undefined = undefined;
      let targetDriveId: string | undefined = undefined;

      if (atSitesRoot) {
        const siteNode = folderPathStack[1];
        const matchedSite = (host.state.documentSites || []).find(s =>
          (siteNode?.id && s.site_id === siteNode.id.replace(/^site:/, '')) ||
          (siteNode?.name && s.sp_site_name?.toLowerCase() === siteNode.name.toLowerCase()) ||
          (siteNode?.name && s.site_key?.toLowerCase() === siteNode.name.toLowerCase())
        );
        targetSiteId = matchedSite?.site_id || (siteNode?.id || '').replace(/^site:/, '') || host.props.siteId;
        const driveNode = folderPathStack[2];
        targetDriveId = (driveNode?.id || '').replace(/^drive:/, '') || matchedSite?.drive_id || host.props.driveId;

        const driveSegments = folderPathStack.slice(3).map(n => n.name);
        subFolderPath = driveSegments.join(' > ');
        fallbackPath = driveSegments.join('/');

        const matchedV = (host.state.vessels || []).find(v =>
          driveSegments.some(seg => seg.toLowerCase() === v.name.toLowerCase())
        );
        currentVessel = matchedV ? matchedV.name : (driveSegments[1] || driveSegments[0] || '');
        const isDriveRoot = stackLevel <= 3;
        const resolvedFolderId = isDriveRoot ? 'root' : (topFolderId || fallbackPath || 'root');

        host._openBulkUpload(extracted, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode, targetSiteId, targetDriveId);
        return;
      } else if (atSharedDocsRoot) {
        targetSiteId = sharedDocsSite?.site_id || host.props.siteId;
        targetDriveId = sharedDocsSite?.drive_id || '';
        const driveSegments = folderPathStack.slice(1).map(n => n.name);
        subFolderPath = driveSegments.join(' > ');
        fallbackPath = driveSegments.join('/');
        const matchedV = (host.state.vessels || []).find(v =>
          driveSegments.some(seg => seg.toLowerCase() === v.name.toLowerCase())
        );
        currentVessel = matchedV ? matchedV.name : (driveSegments[1] || driveSegments[0] || '');
        const isDriveRoot = stackLevel <= 1;
        const resolvedFolderId = isDriveRoot ? 'root' : (topFolderId || fallbackPath || 'root');

        host._openBulkUpload(extracted, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode, targetSiteId, targetDriveId);
        return;
      } else if (atDocsRoot) {
        targetSiteId = docsSite?.site_id || host.props.siteId;
        targetDriveId = docsSite?.drive_id || host.props.driveId;
        const driveSegments = folderPathStack.slice(1).map(n => n.name);
        subFolderPath = driveSegments.join(' > ');
        fallbackPath = driveSegments.join('/');
        const matchedV = (host.state.vessels || []).find(v =>
          driveSegments.some(seg => seg.toLowerCase() === v.name.toLowerCase())
        );
        currentVessel = matchedV ? matchedV.name : (driveSegments[1] || driveSegments[0] || '');
        const isDriveRoot = stackLevel <= 1;
        const resolvedFolderId = isDriveRoot ? 'root' : (topFolderId || fallbackPath || 'root');

        host._openBulkUpload(extracted, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode, targetSiteId, targetDriveId);
        return;
      } else if (atKaizenRoot) {
        currentVessel = 'Kaizen - Knowledge Bank';
        const kaizenFolders = folderPathStack.filter(n => n.id !== 'kaizen_root' && n.name !== 'Kaizen - Knowledge Bank').map(n => n.name);
        subFolderPath = ['Kaizen - Knowledge Bank', ...kaizenFolders].join(' > ');
        fallbackPath = ['Kaizen - Knowledge Bank', ...kaizenFolders].join('/');
      } else if (atCommonShips) {
        currentVessel = 'Common for all vessels';
        const commonFolders = folderPathStack.slice(2).map(n => n.name);
        subFolderPath = [docMainFolder || 'Technical & Crewing', 'Common for all ships', ...commonFolders].join(' > ');
        fallbackPath = [docMainFolder || 'Technical & Crewing', 'Common for all ships', ...commonFolders].join('/');
      } else if (vesselStackIdx !== -1) {
        currentVessel = currentVesselNameFromStack || (vessels.length > 0 ? vessels[0].name : 'Bow Fighter');
        const afterVesselItems = folderPathStack.slice(vesselStackIdx + 1).map(n => n.name);
        subFolderPath = [docMainFolder || 'Technical & Crewing', currentVessel, ...afterVesselItems].join(' > ');
        fallbackPath = [docMainFolder || 'Technical & Crewing', currentVessel, ...afterVesselItems].join('/');
      } else if (docMainFolder) {
        currentVessel = (vesselFilter !== 'all' ? vesselFilter : '') || currentVesselNameFromStack || (vessels.length > 0 ? vessels[0].name : '');
        subFolderPath = currentFolderNode && currentFolderNode.name !== docMainFolder && currentFolderNode.name !== 'Documents'
          ? `${docMainFolder} > ${currentVessel ? `${currentVessel} > ` : ''}${currentFolderNode.name}`
          : `${docMainFolder}${currentVessel ? ` > ${currentVessel}` : ''}`;
        fallbackPath = [docMainFolder, ...(currentVessel ? [currentVessel] : [])].join('/');
      } else {
        const activeDept = 'Technical & Crewing';
        currentVessel = (vesselFilter !== 'all' ? vesselFilter : '') || (vessels.length > 0 ? vessels[0].name : '');
        subFolderPath = currentVessel ? `${activeDept} > ${currentVessel}` : activeDept;
        fallbackPath = [activeDept, ...(currentVessel ? [currentVessel] : [])].join('/');
      }

      const liveId = host._getLiveSharePointFolderId(subFolderPath);
      const matchingRow = host.state.rows.find(r =>
        (currentVessel ? r.vesselName === currentVessel : true) &&
        r.uploadFolderId &&
        !r.uploadFolderId.includes('/') &&
        (docMainFolder ? r.group.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0]) : true)
      );
      const resolvedFolderId = topFolderId || liveId || (matchingRow?.uploadFolderId) || fallbackPath;

      host._openBulkUpload(extracted, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode, targetSiteId, targetDriveId);
    };


    const sharedDocsSite = getFirstClassSite('nksdocman', 0);
    const docsSite = getFirstClassSite('dev', 1);

    // Where the "Archive" toolbar button should root its picker popup when
    // browsing the plain SharePoint folder tree (Sites / Shared Documents /
    // Documents) — those views have no per-file checkbox selection of their
    // own, unlike the vessel-organized List/Folder views below. Mirrors the
    // same site/drive/folder-id resolution each of those views' own renderer
    // already does (see the Level>=3 "sites" branch and renderLibraryBrowser
    // further down) so the popup opens on exactly the folder on screen.
    const resolveArchiveRootContext = (): { siteId: string; driveId: string; folderId: string; folderName: string } | null => {
      if (atSitesRoot) {
        if (stackLevel < 3) return null;
        const siteNode = folderPathStack[1];
        const rawSiteId = (siteNode?.id || '').replace(/^site:/, '');
        const matchedSite = (host.state.documentSites || []).find(s =>
          s.site_id === rawSiteId || s.site_key === rawSiteId || s.sp_site_name === siteNode?.name
        );
        const siteId = matchedSite?.site_id || rawSiteId;
        const driveNode = folderPathStack[2];
        const driveId = (driveNode?.id || '').replace(/^drive:/, '') || matchedSite?.drive_id || '';
        if (!siteId || !driveId) return null;
        const isDriveRoot = stackLevel === 3;
        let folderId = isDriveRoot ? 'root' : (currentFolderNode?.id || 'root');
        if (!isDriveRoot && /^sf_/i.test(folderId)) {
          folderId = folderPathStack.slice(3).map(node => node.name).filter(Boolean).join('/');
        }
        return { siteId, driveId, folderId, folderName: currentFolderName || 'this folder' };
      }
      if (atSharedDocsRoot || atDocsRoot) {
        const site = atSharedDocsRoot ? sharedDocsSite : docsSite;
        if (!site?.site_id || !site?.drive_id) return null;
        const isLibRoot = stackLevel === 1;
        const folderId = isLibRoot ? 'root' : (currentFolderNode?.id || 'root');
        return {
          siteId: site.site_id,
          driveId: site.drive_id,
          folderId,
          folderName: currentFolderName || site.default_library_name || (atSharedDocsRoot ? 'Shared Documents' : 'Documents'),
        };
      }
      return null;
    };

    // Mirrors the site/folder resolution `handleDocumentsPageDrop` above
    // does for the vessel-organized Folder/List views (flat-root template
    // structure, not the plain SharePoint browsing `resolveArchiveRootContext`
    // covers) — used by the "New Folder" toolbar button to reuse the
    // existing Add Folder popup (host._openAddFolderDialog), which needs a
    // real, non-synthetic SharePoint drive-item id to create a subfolder
    // under.
    const resolveVesselFolderContext = (): { folderId: string; folderLabel: string; vesselName: string } | null => {
      const topFolderId = currentFolderNode && !/^(sf_|category_|common|vessels_root|specific_vessels|kaizen_root|sites_root|site:|drive:)/.test(currentFolderNode.id) ? currentFolderNode.id : '';
      let currentVessel = '';
      let subFolderPath = '';
      let fallbackPath = '';

      if (atKaizenRoot) {
        currentVessel = 'Kaizen - Knowledge Bank';
        const kaizenFolders = folderPathStack.filter(n => n.id !== 'kaizen_root' && n.name !== 'Kaizen - Knowledge Bank').map(n => n.name);
        subFolderPath = ['Kaizen - Knowledge Bank', ...kaizenFolders].join(' > ');
        fallbackPath = ['Kaizen - Knowledge Bank', ...kaizenFolders].join('/');
      } else if (atCommonShips) {
        currentVessel = 'Common for all vessels';
        const commonFolders = folderPathStack.slice(2).map(n => n.name);
        subFolderPath = [docMainFolder || 'Technical & Crewing', 'Common for all ships', ...commonFolders].join(' > ');
        fallbackPath = [docMainFolder || 'Technical & Crewing', 'Common for all ships', ...commonFolders].join('/');
      } else if (vesselStackIdx !== -1) {
        currentVessel = currentVesselNameFromStack || (vessels.length > 0 ? vessels[0].name : 'Bow Fighter');
        const afterVesselItems = folderPathStack.slice(vesselStackIdx + 1).map(n => n.name);
        subFolderPath = [docMainFolder || 'Technical & Crewing', currentVessel, ...afterVesselItems].join(' > ');
        fallbackPath = [docMainFolder || 'Technical & Crewing', currentVessel, ...afterVesselItems].join('/');
      } else if (docMainFolder) {
        currentVessel = (vesselFilter !== 'all' ? vesselFilter : '') || currentVesselNameFromStack || (vessels.length > 0 ? vessels[0].name : '');
        subFolderPath = currentFolderNode && currentFolderNode.name !== docMainFolder && currentFolderNode.name !== 'Documents'
          ? `${docMainFolder} > ${currentVessel ? `${currentVessel} > ` : ''}${currentFolderNode.name}`
          : `${docMainFolder}${currentVessel ? ` > ${currentVessel}` : ''}`;
        fallbackPath = [docMainFolder, ...(currentVessel ? [currentVessel] : [])].join('/');
      } else {
        const activeDept = 'Technical & Crewing';
        currentVessel = (vesselFilter !== 'all' ? vesselFilter : '') || (vessels.length > 0 ? vessels[0].name : '');
        subFolderPath = currentVessel ? `${activeDept} > ${currentVessel}` : activeDept;
        fallbackPath = [activeDept, ...(currentVessel ? [currentVessel] : [])].join('/');
      }

      const liveId = host._getLiveSharePointFolderId(subFolderPath);
      const matchingRow = host.state.rows.find(r =>
        (currentVessel ? r.vesselName === currentVessel : true) &&
        r.uploadFolderId &&
        !r.uploadFolderId.includes('/') &&
        (docMainFolder ? r.group.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0]) : true)
      );
      const resolvedFolderId = topFolderId || liveId || matchingRow?.uploadFolderId || '';
      if (!resolvedFolderId || resolvedFolderId.includes('/')) return null;
      return { folderId: resolvedFolderId, folderLabel: subFolderPath, vesselName: currentVessel };
    };

    // Folder-tile counts. `documentLiveFolders` is the site's full recursive
    // tree (GET .../folders/root/recursive), so when a tile's folder is in it
    // we report the TOTAL sub-folders and files anywhere inside that folder.
    // Only when the tree isn't available (not loaded / timed out) do we fall
    // back to the direct children of the folder loaded from Graph.
    const liveTreeIndex = (() => {
      const items: any[] = (host.state.documentLiveFolders || []) as any[];
      const byParent = new Map<string, any[]>();
      const ids = new Set<string>();
      items.forEach(it => {
        if (!it || !it.id) return;
        ids.add(it.id);
        const pid = it.parent_id || '';
        const list = byParent.get(pid);
        if (list) list.push(it); else byParent.set(pid, [it]);
      });
      return { byParent, ids };
    })();
    const isFolderItem = (item: any): boolean => !!(item.folder || (!item.file && item.name && !item.name.includes('.')));
    const getFolderTileCounts = (
      folderId: string | undefined,
      directItems: any[] | null,
    ): { folders: number; files: number } | null => {
      if (folderId && liveTreeIndex.ids.has(folderId)) {
        let folders = 0;
        let files = 0;
        const seen = new Set<string>([folderId]);
        const stack: string[] = [folderId];
        while (stack.length > 0) {
          const id = stack.pop() as string;
          (liveTreeIndex.byParent.get(id) || []).forEach(child => {
            if (seen.has(child.id)) return;
            seen.add(child.id);
            if (child.is_folder === false) {
              files += 1;
            } else {
              folders += 1;
              stack.push(child.id);
            }
          });
        }
        return { folders, files };
      }
      if (!directItems) return null;
      const folders = directItems.filter(isFolderItem).length;
      return { folders, files: directItems.length - folders };
    };

    // Count chips from server-computed FULL-DEPTH counts (/subfolder-counts):
    // total nested folders + total files anywhere under the folder. Tooltip
    // shows the direct-vs-nested breakdown. "~" = still partial/estimated.
    const renderServerCountPill = (
      fc: {
        direct_subfolders: number; total_subfolders?: number;
        direct_files?: number; total_files: number;
        is_partial?: boolean; is_estimated?: boolean;
      },
      name: string,
    ): React.ReactElement => {
      const fmt = (n: number): string => n.toLocaleString();
      const totalFolders = Math.max(fc.total_subfolders ?? 0, fc.direct_subfolders);
      const totalFiles = fc.total_files;
      const approx = !!(fc.is_partial || fc.is_estimated);
      const nestedFolders = totalFolders - fc.direct_subfolders;
      const directFiles = fc.direct_files ?? 0;
      const title = [
        `${name}${approx ? ' (partial count — still loading or throttled)' : ''}`,
        `Folders: ${fmt(totalFolders)} total (${fmt(fc.direct_subfolders)} direct, ${fmt(nestedFolders)} nested)`,
        `Files: ${fmt(totalFiles)} total (${fmt(directFiles)} directly inside, ${fmt(Math.max(totalFiles - directFiles, 0))} in subfolders)`,
      ].join('\n');
      const chip = (on: boolean, bg: string, fg: string): React.CSSProperties => ({
        display: 'inline-flex', alignItems: 'center', gap: 5, borderRadius: 8,
        padding: '2px 9px', fontSize: 11, fontWeight: 700, lineHeight: '16px',
        whiteSpace: 'nowrap', fontVariantNumeric: 'tabular-nums',
        background: on ? bg : 'var(--vdms-border-soft)',
        color: on ? fg : 'var(--vdms-text-faint)',
      });
      return (
        <span title={title} style={{ display: 'inline-flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
          <span style={chip(totalFolders > 0, '#e0f2fe', '#0369a1')}>
            <span aria-hidden="true">📁</span>
            {approx ? '~' : ''}{fmt(totalFolders)}
            <span style={{ fontWeight: 600, opacity: 0.8 }}>{totalFolders === 1 ? 'folder' : 'folders'}</span>
          </span>
          <span style={chip(totalFiles > 0, '#dcfce7', '#15803d')}>
            <span aria-hidden="true">📄</span>
            {approx ? '~' : ''}{fmt(totalFiles)}
            <span style={{ fontWeight: 600, opacity: 0.8 }}>{totalFiles === 1 ? 'file' : 'files'}</span>
          </span>
        </span>
      );
    };

    const renderLibraryBrowser = (
      siteId: string,
      driveId: string,
      libraryTitle: string,
      librarySubtitle: string,
      folderGroupKey: 'Shared Documents' | 'Documents'
    ): React.ReactElement => {
      const isLibRoot = stackLevel === 1;
      const currentNode = folderPathStack[folderPathStack.length - 1];
      const currentFolderId = isLibRoot ? 'root' : (currentNode?.id || 'root');
      const folderData = host._getOrLoadSiteFolderChildren(siteId, driveId, currentFolderId);

      if (folderData.loading && folderData.items.length === 0) {
        return (
          <div style={{ padding: 48, textAlign: 'center', color: 'var(--vdms-text-muted)' }}>
            <div style={{ fontSize: 28, marginBottom: 8, animation: 'spin 1s linear infinite', display: 'inline-block' }}>⏳</div>
            <div style={{ fontSize: 14, fontWeight: 600 }}>Loading {currentNode?.name || libraryTitle}...</div>
          </div>
        );
      }

      const baseChildFolders = folderData.items.filter(item => item.folder || (!item.file && item.name && !item.name.includes('.')));
      const childFiles = folderData.items.filter(item => item.file || (item.name && item.name.includes('.')));

      // NOTE: this used to merge host.state.rows (vessel rows, aggregated across ALL
      // sites since GET /api/vessels has no site scoping) and host.state.uploadedFilesByFolder
      // (keyed by a bare path string with no site/drive affinity) into the folder list
      // shown here. Neither source carries any site/drive identity, so at a library root
      // it injected other sites' vessel names as phantom folders (e.g. vessel folders
      // showing up inside "Communication Site" > Documents). This is a real SharePoint
      // library browser, so it must show exactly what Graph returns for this site/drive.
      const childFolders = baseChildFolders;
      const folderServerCounts = folderData.error ? null : host._getOrLoadFolderCounts(siteId, driveId, currentFolderId);

      return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
            <div>
              <div style={{ fontSize: 13, fontWeight: 700, color: '#0369a1', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                {currentNode?.name || libraryTitle}
              </div>
              <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2 }}>
                {librarySubtitle} · {childFolders.length} folders, {childFiles.length} files{folderServerCounts?.summary ? ` · ${folderServerCounts.summary.total_files} total files` : ''}
              </div>
            </div>
          </div>

          {childFolders.length > 0 && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 800, color: '#0369a1', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                Folders ({childFolders.length})
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 }}>
                {childFolders.map((sf, idx) => {
                  const sfInTree = !!sf.id && liveTreeIndex.ids.has(sf.id);
                  // Server counts (one /subfolder-counts call) already feed every tile pill, so
                  // don't also fire a full /children call per tile just to count.
                  const sfCountsPending = !folderServerCounts && host._folderCountsLoading(siteId, driveId, currentFolderId);
                  const sfChildData = sf.id && !sfInTree && !folderServerCounts && !sfCountsPending ? host._getOrLoadSiteFolderChildren(siteId, driveId, sf.id) : null;
                  const sfLoading = sfChildData ? (sfChildData.loading && sfChildData.items.length === 0) : (sfCountsPending && !sfInTree);
                  const sfCounts = sfLoading ? null : getFolderTileCounts(sf.id, sfChildData ? sfChildData.items : null);
                  const sfFolderCount = sfCounts ? sfCounts.folders : null;
                  const sfFileCount  = sfCounts ? sfCounts.files : null;
                  const sfTotal = sf.folder?.childCount ?? 0;
                  const sfServerCounts = folderServerCounts?.counts[sf.id] || null;
                  return (
                    <div
                      key={sf.id || sf.name + idx}
                      onClick={() => {
                        // Re-sync the Main folder / Sub-folder dropdowns to the
                        // folder actually being opened, the same way a
                        // breadcrumb click or dropdown pick already does (see
                        // deriveDocFiltersFromStack). Without this, clicking
                        // folder tiles here never touched docCategoryFilter at
                        // all, so it kept whatever main folder was selected
                        // earlier (e.g. from the Main folder dropdown, or the
                        // "is also under" chips) even after the grid had moved
                        // into a different main folder's tiles — which is what
                        // left the Sub-folder dropdown scoped to that stale,
                        // no-longer-current main folder's live id.
                        const newStack = [...folderPathStack, { id: sf.id, name: sf.name }];
                        host._pushFolderNav(newStack, folderGroupKey);
                        host.setState({
                          ...deriveDocFiltersFromStack(newStack),
                          docGroupLevelFilter: 'all',
                          docLeafCategoryFilter: 'all',
                          docSubCategoryFilter: 'all',
                          docListPage: 0,
                        });
                      }}
                      style={{
                        background: 'var(--vdms-surface)', borderRadius: 14, border: '1px solid var(--vdms-border)', padding: 18,
                        display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                        boxShadow: '0 1px 3px rgba(0,0,0,0.05)', transition: 'all 0.15s ease',
                      }}
                    >
                      <div style={{
                        width: 44, height: 44, borderRadius: 10, background: '#e0f2fe',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7',
                      }}>
                        📁
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--vdms-text)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                          {sf.name}
                        </div>
                        <div title={sfCounts ? `${sfCounts.folders} sub-folder(s) and ${sfCounts.files} file(s) in total` : undefined} style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 3, display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                          {sfServerCounts ? renderServerCountPill(sfServerCounts, sf.name) : sfLoading ? (
                            <span style={{ color: 'var(--vdms-text-faint)', fontSize: 10 }}>{sfTotal > 0 ? `${sfTotal} items` : '···'}</span>
                          ) : (
                            <>
                              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 3, background: (sfFolderCount ?? 0) > 0 ? '#e0f2fe' : 'var(--vdms-border-soft)', color: (sfFolderCount ?? 0) > 0 ? '#0369a1' : 'var(--vdms-text-faint)', borderRadius: 20, padding: '1px 8px', fontSize: 10, fontWeight: 700, lineHeight: '16px' }}>📁 {sfFolderCount ?? 0}</span>
                              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 3, background: (sfFileCount ?? 0) > 0 ? '#dcfce7' : 'var(--vdms-border-soft)', color: (sfFileCount ?? 0) > 0 ? '#15803d' : 'var(--vdms-text-faint)', borderRadius: 20, padding: '1px 8px', fontSize: 10, fontWeight: 700, lineHeight: '16px' }}>📄 {sfFileCount ?? 0}</span>
                            </>
                          )}
                        </div>
                      </div>
                      <span style={{ color: '#0284c7', fontSize: 16 }}>›</span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {childFiles.length > 0 && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 800, color: '#15803d', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                Files ({childFiles.length})
              </div>
              <div style={{ background: 'var(--vdms-surface)', borderRadius: 12, border: '1px solid var(--vdms-border)', overflow: 'hidden' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: 'var(--vdms-surface-alt)', borderBottom: '1px solid var(--vdms-border)', color: 'var(--vdms-text-muted)', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
                      <th style={{ padding: '10px 16px' }}>FILE NAME</th>
                      <th style={{ padding: '10px 16px' }}>SIZE</th>
                      <th style={{ padding: '10px 16px' }}>DATE MODIFIED</th>
                      <th style={{ padding: '10px 16px', textAlign: 'right' }}>ACTION</th>
                    </tr>
                  </thead>
                  <tbody>
                    {childFiles.map((file, idx) => {
                      const fileSize = typeof file.size === 'number'
                        ? (file.size > 1024 * 1024 ? `${(file.size / (1024 * 1024)).toFixed(1)} MB` : `${(file.size / 1024).toFixed(1)} KB`)
                        : '—';
                      const fileDate = file.lastModifiedDateTime ? new Date(file.lastModifiedDateTime).toLocaleString() : '—';
                      const fileUrl = file.web_url || file.webUrl || file.download_url || '';

                      return (
                        <tr key={file.id || file.name + idx} style={{ borderBottom: '1px solid var(--vdms-border-soft)' }}>
                          <td style={{ padding: '12px 16px', fontWeight: 600, color: 'var(--vdms-text)', display: 'flex', alignItems: 'center', gap: 10 }}>
                            <span style={{ fontSize: 18 }}>📄</span>
                            <span
                              onClick={() => {
                                if (fileUrl) {
                                  window.open(fileUrl, '_blank');
                                } else {
                                  void host._openDocumentFile(file.id, file.name, folderPathStack.map(n => n.name).join(' > '));
                                }
                              }}
                              style={{ cursor: 'pointer', color: '#0284c7', textDecoration: 'underline' }}
                              title={`Click to view/download ${file.name}`}
                            >
                              {file.name}
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px', color: 'var(--vdms-text-muted)' }}>{fileSize}</td>
                          <td style={{ padding: '12px 16px', color: 'var(--vdms-text-muted)' }}>{fileDate}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                            <button
                              type="button"
                              onClick={() => {
                                if (fileUrl) {
                                  window.open(fileUrl, '_blank');
                                } else {
                                  void host._openDocumentFile(file.id, file.name, folderPathStack.map(n => n.name).join(' > '));
                                }
                              }}
                              style={{
                                background: '#0284c7', color: '#fff', border: 'none', borderRadius: 6,
                                padding: '5px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                              }}
                            >
                              Open / Download
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {childFolders.length === 0 && childFiles.length === 0 && (
            <div style={{ background: 'var(--vdms-surface)', borderRadius: 14, border: '1px dashed var(--vdms-border)', padding: 48, textAlign: 'center', color: 'var(--vdms-text-faint)' }}>
              <div style={{ fontSize: 32, marginBottom: 8 }}>📂</div>
              <div style={{ fontWeight: 600, color: 'var(--vdms-text-secondary)', fontSize: 15 }}>This folder is empty</div>
              <div style={{ fontSize: 13, marginTop: 4 }}>No files or subfolders found in this directory.</div>
            </div>
          )}
        </div>
      );
    };

    return (
      <div
        onDragOver={e => e.preventDefault()}
        onDrop={handleDocumentsPageDrop}
        style={{ display: 'flex', flexDirection: 'column', gap: 16 }}
      >
        <style>{`
          @keyframes dmsNavPulse {
            0%, 100% { box-shadow: 0 4px 10px rgba(14, 165, 233, 0.18); }
            50% { box-shadow: 0 6px 18px rgba(168, 85, 247, 0.34); }
          }
          .dms-nav-arrow:not(:disabled) { animation: dmsNavPulse 2.8s ease-in-out infinite; }
          .dms-nav-arrow:not(:disabled):hover { transform: translateY(-2px) scale(1.08); filter: saturate(1.2); }
          .dms-nav-arrow:not(:disabled):active { transform: translateY(0) scale(0.96); }
        `}</style>
        {/* Sticky header: breadcrumb + module header + filter toolbar stay pinned
            while the file table scrolls underneath. */}
        <div style={{ position: 'sticky', top: 0, zIndex: 30, background: clay.bg, display: 'flex', flexDirection: 'column', gap: 16, paddingBottom: 8, boxShadow: '0 6px 8px -6px rgba(0,0,0,0.18)' }}>
        {/* Breadcrumb Navigation Trail */}
        <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
          {/* Back / Forward navigation buttons */}
          <button
            onClick={goBack}
            disabled={!canGoBack}
            aria-label="Go back"
            title="Go back (Left Arrow)"
            className="dms-nav-arrow"
            style={{
              width: 34, height: 34, borderRadius: 10, border: canGoBack ? '1px solid #38bdf8' : '1px solid var(--vdms-border)',
              background: canGoBack ? 'linear-gradient(135deg, #0ea5e9, #2563eb)' : 'var(--vdms-border-soft)',
              color: canGoBack ? '#fff' : 'var(--vdms-border)',
              cursor: canGoBack ? 'pointer' : 'not-allowed',
              fontSize: 20, fontWeight: 800, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              transition: 'transform 160ms ease, filter 160ms ease, box-shadow 160ms ease',
              flexShrink: 0,
            }}
          >←</button>
          <button
            onClick={goForward}
            disabled={!canGoForward}
            aria-label="Go forward"
            title="Go forward (Right Arrow)"
            className="dms-nav-arrow"
            style={{
              width: 34, height: 34, borderRadius: 10, border: canGoForward ? '1px solid #c084fc' : '1px solid var(--vdms-border)',
              background: canGoForward ? 'linear-gradient(135deg, #8b5cf6, #ec4899)' : 'var(--vdms-border-soft)',
              color: canGoForward ? '#fff' : 'var(--vdms-border)',
              cursor: canGoForward ? 'pointer' : 'not-allowed',
              fontSize: 20, fontWeight: 800, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              transition: 'transform 160ms ease, filter 160ms ease, box-shadow 160ms ease',
              flexShrink: 0,
            }}
          >→</button>
        {folderPathStack.map((item, idx) => {
          if (idx === 0 && (item.id === 'sites_root' || item.name === 'SharePoint Sites' || item.name === 'Sites Documents')) {
            return null;
          }
            const isLast = idx === folderPathStack.length - 1;
            return (
              <React.Fragment key={item.id + idx}>
                <span>›</span>
                <span
                  onClick={() => {
                    if (atKaizenRoot && idx === 0) {
                      const newStack = folderPathStack.slice(0, 1);
                      host._pushFolderNav(newStack, 'Kaizen - Knowledge Bank');
                      host.setState(deriveDocFiltersFromStack(newStack));
                      triggerFolderRefresh(newStack, 'kaizen');
                    } else if (idx === 0) {
                      const mainFolder = item.name as MainFolderKey;
                      const newStack = folderPathStack.slice(0, 1);
                      host._pushFolderNav(newStack, mainFolder);
                      // Always re-sync the Main folder / Sub-folder dropdowns to
                      // this breadcrumb jump, not only for the three named
                      // departments below — a live SharePoint top-level folder
                      // name landing here needs the same reset.
                      host.setState(deriveDocFiltersFromStack(newStack));
                      if (mainFolder === 'Technical & Crewing' || mainFolder === 'Commercial & Chartering' || mainFolder === 'Insurance') {
                        host.setState({
                          vesselFilter: 'all',
                          docScopeType: 'vessels',
                          docGroupLevelFilter: 'all',
                          docLeafCategoryFilter: 'all',
                          docSubCategoryFilter: 'all',
                          catFilter: 'all',
                          docListPage: 0,
                        });
                      }
                      // Refresh for sites/shared_docs/documents breadcrumb home nodes
                      const breadcrumbScope = host.state.docScopeType;
                      if (breadcrumbScope === 'sites' || breadcrumbScope === 'shared_docs' || breadcrumbScope === 'documents') {
                        triggerFolderRefresh(newStack, breadcrumbScope);
                      }
                    } else if (idx === 1 && atMainDepartment && !atCommonShips) {
                      const newStack = folderPathStack.slice(0, idx + 1);
                      host._pushFolderNav(newStack, docMainFolder);
                      host.setState({
                        vesselFilter: item.name,
                        docScopeType: 'vessels',
                        ...deriveDocFiltersFromStack(newStack),
                        docGroupLevelFilter: 'all',
                        docLeafCategoryFilter: 'all',
                        docSubCategoryFilter: 'all',
                        catFilter: 'all',
                        docListPage: 0,
                      });
                      triggerFolderRefresh(newStack, 'vessels', item.name);
                    } else {
                      const newStack = folderPathStack.slice(0, idx + 1);
                      host._pushFolderNav(newStack, docMainFolder);
                      // Breadcrumb navigation must re-sync the Main folder /
                      // Sub-folder dropdowns to match the folder now being
                      // viewed — the dropdowns' own onChange handlers already
                      // do this when a folder is picked from them, but a
                      // breadcrumb click bypasses those, so it needs its own
                      // sync here.
                      const crumbFilters = newStack[0]?.id === 'sites_root'
                        ? deriveLiveNavFilterState(newStack)
                        : deriveDocFiltersFromStack(newStack);
                      // Jumping to an ancestor of the selected vessel's folder
                      // (e.g. "Technical" above "Belle Lune"): the vessel is no
                      // longer in the path, so its filter must be released —
                      // deriveLiveNavFilterState deliberately keeps it for
                      // ancestors of its folder, which left the previous vessel
                      // pinned and the results showing only that vessel.
                      const crumbPrefixLen = newStack[0]?.id === 'sites_root' ? 3 : 1;
                      const vesselStillInPath = newStack.slice(crumbPrefixLen).some(n => !!matchFolderToVessel(n.name));
                      const crumbNav = crumbFilters as Partial<LiveNavFilterState>;
                      host.setState({
                        docCategoryFilter: crumbFilters.docCategoryFilter,
                        docSubfolderOtherFilter: crumbFilters.docSubfolderOtherFilter,
                        docGroupLevelFilter: crumbNav.docGroupLevelFilter !== undefined ? crumbNav.docGroupLevelFilter : (host.state.docGroupLevelFilter || 'all'),
                        vesselFilter: !vesselStillInPath ? 'all' : (crumbNav.vesselFilter !== undefined ? crumbNav.vesselFilter : (host.state.vesselFilter || 'all')),
                        docListPage: 0,
                      });
                      triggerFolderRefresh(newStack);
                    }
                  }}
                  style={{ cursor: isLast ? 'default' : 'pointer', color: isLast ? 'var(--vdms-text)' : '#0284c7', fontWeight: isLast ? 600 : 400 }}
                >
                  {item.name}
                </span>
              </React.Fragment>
            );
          })}
        </div>

        {/* Module Header — the Folder/List toggle is pinned as the last,
            non-shrinking child so it sits at the same top-right spot in both
            views; the variable-width action buttons wrap to its left. */}
        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 12 }}>
          <div style={{ flex: '0 1 auto', minWidth: 160 }}>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: 'var(--vdms-text)', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 20 }}>📁</span>
              {currentFolderNode ? currentFolderNode.name : 'Documents'}
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: 'var(--vdms-text-muted)' }}>
              {docViewMode === 'list'
                ? `${filtered.length} rows · flattened list view`
                : stackLevel === 0
                  ? `${MAIN_FOLDERS.length} main folders & libraries`
                  : atSharedDocsRoot
                    ? 'NKSDocMan · Shared Documents'
                    : atDocsRoot
                      ? 'Communication Site · Documents'
                      : atSitesRoot
                        ? 'SharePoint Sites'
                        : atDepartmentVesselList
                          ? `${displayVessels.length} vessels · Common for all ships`
                          : `${subfolderNames.length} sections · ${allCurrentFolderFiles.length} files`}
            </p>
            {docViewMode === 'folder' && currentFolderNode && subfolderNames.length === 0
              && !atSitesRoot && !atSharedDocsRoot && !atDocsRoot
              && !host.state.documentFilesLoading
              && (
              <div style={{ marginTop: 6, fontSize: 12, fontWeight: 600, color: allCurrentFolderFiles.length > 0 ? '#15803d' : 'var(--vdms-text-muted)' }}>
                {allCurrentFolderFiles.length > 0 ? '✅ Attached' : '⚪ Not Attached'}
              </div>
            )}
          </div>

          {(host.state.documentSites || []).length > 1 && (
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: 'var(--vdms-text)', fontWeight: 600 }}>
              SharePoint site
              <select
                aria-label="Select SharePoint site"
                value={host.state.activeDocumentSite || ''}
                onChange={event => {
                  const siteKey = event.target.value || '';
                  void host._switchDocumentSite(siteKey).catch(error => {
                    alert(error?.message || 'Could not switch site.');
                  });
                }}
                style={{ minWidth: 230, padding: '9px 12px', border: '1px solid var(--vdms-text-faint)', borderRadius: 7, background: 'var(--vdms-surface)', color: 'var(--vdms-text)' }}
              >
                {host.state.documentSites.map(site => (
                  <option key={site.site_key} value={site.site_key}>
                    {site.sp_site_name || site.site_key}
                  </option>
                ))}
              </select>
            </label>
          )}

          <div style={{ flex: '1 1 auto', minWidth: 0, display: 'flex', alignItems: 'center', justifyContent: 'flex-end', flexWrap: 'wrap', gap: 10 }}>

            {/* New Folder Button — sits just before Archive. In the plain
                SharePoint folder tree (Sites / Shared Documents / Documents)
                it opens a popup (CreateFolderModal) that creates a folder at
                — or at an edited variant of — the path currently open, via
                Graph directly, with an option to create a vessel instead.
                In the vessel-organized Folder/List views it reuses the
                existing Add Folder popup, which needs a real SharePoint
                folder id (so it's disabled where none can be resolved yet,
                e.g. before a vessel/department has been opened). */}
            {(() => {
              const usesArchivePicker = atSitesRoot || atSharedDocsRoot || atDocsRoot;
              const rootCtx = usesArchivePicker ? resolveArchiveRootContext() : null;
              const vesselCtx = usesArchivePicker ? null : resolveVesselFolderContext();
              const canCreateFolder = usesArchivePicker ? !!rootCtx : !!vesselCtx;
              return (
                <button
                  disabled={!canCreateFolder}
                  onClick={() => {
                    if (usesArchivePicker) {
                      if (!rootCtx) { alert('Open a folder first to create a folder inside it.'); return; }
                      host._openCreateFolderDialog(rootCtx.siteId, rootCtx.driveId, rootCtx.folderId, rootCtx.folderName);
                      return;
                    }
                    if (!vesselCtx) { alert('Open a vessel or department folder first to create a folder inside it.'); return; }
                    host._openAddFolderDialog(vesselCtx);
                  }}
                  style={{
                    background: canCreateFolder ? '#0284c7' : 'var(--vdms-border-soft)',
                    color: canCreateFolder ? '#fff' : 'var(--vdms-text-faint)',
                    border: 'none', borderRadius: 8,
                    padding: '7px 16px', fontSize: 13, fontWeight: 600, cursor: canCreateFolder ? 'pointer' : 'not-allowed',
                    display: 'inline-flex', alignItems: 'center', gap: 6,
                  }}
                  title={canCreateFolder ? 'Create a new folder here' : 'Open a folder first to create a folder inside it'}
                >
                  📁 New Folder
                </button>
              );
            })()}

            {/* Archive Button — in the plain SharePoint folder tree (Sites /
                Shared Documents / Documents) there's no per-file checkbox
                selection to archive, so this opens a popup to pick folders
                and/or files instead. In the vessel-organized List/Folder
                views it still archives whatever's already checked there. */}
            {(() => {
              const usesArchivePicker = atSitesRoot || atSharedDocsRoot || atDocsRoot;
              const selectionCount = docViewMode === 'list' ? listViewSelectedFiles.size : folderViewSelectedFiles.size;
              const isActive = usesArchivePicker || selectionCount > 0;
              return (
                <button
                  disabled={!usesArchivePicker && selectionCount === 0}
                  onClick={() => {
                    if (usesArchivePicker) {
                      const ctx = resolveArchiveRootContext();
                      if (!ctx) {
                        alert('Open a folder first to archive files or folders from it.');
                        return;
                      }
                      host._openArchivePicker(ctx.siteId, ctx.driveId, ctx.folderId, ctx.folderName);
                      return;
                    }
                    const selectedFiles = docViewMode === 'list'
                      ? groupedList.flatMap(row => row.files
                        .filter(file => listViewSelectedFiles.has(file.id))
                        .map(file => ({ id: file.id, name: file.name, folderPath: row.subFolderPath, department: row.group, vesselName: row.vesselName })))
                      : allCurrentFolderFiles
                        .filter(file => folderViewSelectedFiles.has(file.id || file.name))
                        .map(file => ({ id: file.id || file.name, name: file.name, folderPath: currentFolderName || '', department: docMainFolder || '', vesselName: currentVesselNameFromStack || '' }));
                    void host._archiveSelectedDocuments(selectedFiles);
                  }}
                  style={{
                    background: isActive ? '#d97706' : 'var(--vdms-border-soft)',
                    color: isActive ? '#fff' : 'var(--vdms-text-faint)',
                    border: 'none', borderRadius: 8,
                    padding: '7px 16px', fontSize: 13, fontWeight: 600, cursor: isActive ? 'pointer' : 'not-allowed',
                    display: 'inline-flex', alignItems: 'center', gap: 6,
                  }}
                  title={usesArchivePicker ? 'Choose folders or files to archive' : undefined}
                >
                  📦 Archive{(!usesArchivePicker && selectionCount > 0) ? ` (${selectionCount})` : ''}
                </button>
              );
            })()}

            {/* Global Delete Button - List View */}
            {docViewMode === 'list' && (
              <button
                disabled={listViewSelectedFiles.size === 0}
                onClick={() => {
                  if (listViewSelectedFiles.size === 0) return;
                  const filesToDelete = groupedList
                    .flatMap(r => r.files.filter(f => listViewSelectedFiles.has(f.id)).map(f => ({ id: f.id, name: f.name, folderId: r.uploadFolderId, folderPath: r.subFolderPath })))
                    .filter(f => f.id);
                  host._openFileDeleteDialog(filesToDelete);
                }}
                style={{
                  background: listViewSelectedFiles.size > 0 ? '#fff5f5' : 'var(--vdms-surface-alt)',
                  color: listViewSelectedFiles.size > 0 ? '#ef4444' : 'var(--vdms-border)',
                  border: `1px solid ${listViewSelectedFiles.size > 0 ? '#fca5a5' : 'var(--vdms-border)'}`,
                  borderRadius: 8, padding: '7px 16px', fontSize: 13, fontWeight: 600,
                  cursor: listViewSelectedFiles.size > 0 ? 'pointer' : 'not-allowed',
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                }}
                title={listViewSelectedFiles.size > 0 ? `Delete ${listViewSelectedFiles.size} selected file(s)` : 'Select files to delete'}
              >
                🗑 Delete{listViewSelectedFiles.size > 0 ? ` (${listViewSelectedFiles.size})` : ''}
              </button>
            )}

            {/* Global Delete Button - Folder View */}
            {docViewMode === 'folder' && subfolderNames.length === 0 && visibleCurrentFolderFiles.length > 0 && (
              <button
                disabled={folderViewSelectedFiles.size === 0}
                onClick={() => {
                  if (folderViewSelectedFiles.size === 0) return;
                  const filesToDelete = visibleCurrentFolderFiles
                    .filter(file => folderViewSelectedFiles.has((file as any).id || file.name))
                    .map(file => ({
                      id: (file as any).id || file.name,
                      name: file.name,
                      folderId: currentFolderNode?.id || '',
                      folderPath: currentFolderNode?.name || '',
                    }));
                  host.setState({ folderViewSelectedFiles: new Set() });
                  host._openFileDeleteDialog(filesToDelete);
                }}
                style={{
                  background: folderViewSelectedFiles.size > 0 ? '#fff5f5' : 'var(--vdms-surface-alt)',
                  color: folderViewSelectedFiles.size > 0 ? '#ef4444' : 'var(--vdms-border)',
                  border: `1px solid ${folderViewSelectedFiles.size > 0 ? '#fca5a5' : 'var(--vdms-border)'}`,
                  borderRadius: 8, padding: '7px 16px', fontSize: 13, fontWeight: 600,
                  cursor: folderViewSelectedFiles.size > 0 ? 'pointer' : 'not-allowed',
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                }}
                title={folderViewSelectedFiles.size > 0 ? `Delete ${folderViewSelectedFiles.size} selected file(s)` : 'Select files to delete'}
              >
                🗑 Delete{folderViewSelectedFiles.size > 0 ? ` (${folderViewSelectedFiles.size})` : ''}
              </button>
            )}

            {/* Top-Right Action Buttons — Always available in Folder View and List View (Root, Department, Vessel, Category, Subfolder) */}
            {(() => {
              const resolveDocPageUploadTarget = (): {
                currentVessel: string;
                subFolderPath: string;
                fallbackPath: string;
                resolvedFolderId: string;
                targetSiteId?: string;
                targetDriveId?: string;
              } => {
                const topFolderId = currentFolderNode && !/^(sf_|category_|common|vessels_root|specific_vessels|kaizen_root|sites_root|site:|drive:)/.test(currentFolderNode.id) ? currentFolderNode.id : '';
                let currentVessel = '';
                let subFolderPath = '';
                let fallbackPath = '';
                let targetSiteId: string | undefined = undefined;
                let targetDriveId: string | undefined = undefined;

                if (atSitesRoot) {
                  const siteNode = folderPathStack[1];
                  const matchedSite = (host.state.documentSites || []).find(s =>
                    (siteNode?.id && s.site_id === siteNode.id.replace(/^site:/, '')) ||
                    (siteNode?.name && s.sp_site_name?.toLowerCase() === siteNode.name.toLowerCase()) ||
                    (siteNode?.name && s.site_key?.toLowerCase() === siteNode.name.toLowerCase())
                  );
                  targetSiteId = matchedSite?.site_id || (siteNode?.id || '').replace(/^site:/, '') || host.props.siteId;
                  const driveNode = folderPathStack[2];
                  targetDriveId = (driveNode?.id || '').replace(/^drive:/, '') || matchedSite?.drive_id || host.props.driveId;

                  const driveSegments = folderPathStack.slice(3).map(n => n.name);
                  subFolderPath = driveSegments.join(' > ');
                  fallbackPath = driveSegments.join('/');

                  const matchedV = (host.state.vessels || []).find(v =>
                    driveSegments.some(seg => seg.toLowerCase() === v.name.toLowerCase())
                  );
                  currentVessel = matchedV ? matchedV.name : (driveSegments[1] || driveSegments[0] || '');
                  const isDriveRoot = stackLevel <= 3;
                  const resolvedFolderId = isDriveRoot ? 'root' : (topFolderId || fallbackPath || 'root');

                  return { currentVessel, subFolderPath, fallbackPath, resolvedFolderId, targetSiteId, targetDriveId };
                } else if (atSharedDocsRoot) {
                  targetSiteId = sharedDocsSite?.site_id || host.props.siteId;
                  targetDriveId = sharedDocsSite?.drive_id || '';
                  const driveSegments = folderPathStack.slice(1).map(n => n.name);
                  subFolderPath = driveSegments.join(' > ');
                  fallbackPath = driveSegments.join('/');
                  const matchedV = (host.state.vessels || []).find(v =>
                    driveSegments.some(seg => seg.toLowerCase() === v.name.toLowerCase())
                  );
                  currentVessel = matchedV ? matchedV.name : (driveSegments[1] || driveSegments[0] || '');
                  const isDriveRoot = stackLevel <= 1;
                  const resolvedFolderId = isDriveRoot ? 'root' : (topFolderId || fallbackPath || 'root');
                  return { currentVessel, subFolderPath, fallbackPath, resolvedFolderId, targetSiteId, targetDriveId };
                } else if (atDocsRoot) {
                  targetSiteId = docsSite?.site_id || host.props.siteId;
                  targetDriveId = docsSite?.drive_id || host.props.driveId;
                  const driveSegments = folderPathStack.slice(1).map(n => n.name);
                  subFolderPath = driveSegments.join(' > ');
                  fallbackPath = driveSegments.join('/');
                  const matchedV = (host.state.vessels || []).find(v =>
                    driveSegments.some(seg => seg.toLowerCase() === v.name.toLowerCase())
                  );
                  currentVessel = matchedV ? matchedV.name : (driveSegments[1] || driveSegments[0] || '');
                  const isDriveRoot = stackLevel <= 1;
                  const resolvedFolderId = isDriveRoot ? 'root' : (topFolderId || fallbackPath || 'root');
                  return { currentVessel, subFolderPath, fallbackPath, resolvedFolderId, targetSiteId, targetDriveId };
                } else if (atKaizenRoot) {
                  currentVessel = 'Kaizen - Knowledge Bank';
                  const kaizenFolders = folderPathStack.filter(n => n.id !== 'kaizen_root' && n.name !== 'Kaizen - Knowledge Bank').map(n => n.name);
                  subFolderPath = ['Kaizen - Knowledge Bank', ...kaizenFolders].join(' > ');
                  fallbackPath = ['Kaizen - Knowledge Bank', ...kaizenFolders].join('/');
                } else if (atCommonShips) {
                  currentVessel = 'Common for all vessels';
                  const commonFolders = folderPathStack.slice(2).map(n => n.name);
                  subFolderPath = [docMainFolder || 'Technical & Crewing', 'Common for all ships', ...commonFolders].join(' > ');
                  fallbackPath = [docMainFolder || 'Technical & Crewing', 'Common for all ships', ...commonFolders].join('/');
                } else if (vesselStackIdx !== -1) {
                  currentVessel = currentVesselNameFromStack || '';
                  const afterVesselItems = folderPathStack.slice(vesselStackIdx + 1).map(n => n.name);
                  subFolderPath = [docMainFolder || 'Technical & Crewing', currentVessel, ...afterVesselItems].filter(Boolean).join(' > ');
                  fallbackPath = [docMainFolder || 'Technical & Crewing', currentVessel, ...afterVesselItems].filter(Boolean).join('/');
                } else if (docMainFolder) {
                  currentVessel = (vesselFilter !== 'all' ? vesselFilter : '') || currentVesselNameFromStack || '';
                  subFolderPath = currentFolderNode && currentFolderNode.name !== docMainFolder && currentFolderNode.name !== 'Documents'
                    ? `${docMainFolder}${currentVessel ? ` > ${currentVessel}` : ''} > ${currentFolderNode.name}`
                    : `${docMainFolder}${currentVessel ? ` > ${currentVessel}` : ''}`;
                  fallbackPath = [docMainFolder, ...(currentVessel ? [currentVessel] : [])].join('/');
                } else {
                  // At root Documents page (stackLevel === 0)
                  const activeDept = 'Technical & Crewing';
                  currentVessel = (vesselFilter !== 'all' ? vesselFilter : '') || currentVesselNameFromStack || '';
                  subFolderPath = currentVessel ? `${activeDept} > ${currentVessel}` : activeDept;
                  fallbackPath = [activeDept, ...(currentVessel ? [currentVessel] : [])].join('/');
                }

                const liveId = host._getLiveSharePointFolderId(subFolderPath);
                const matchingRow = host.state.rows.find(r =>
                  (currentVessel ? r.vesselName === currentVessel : true) &&
                  r.uploadFolderId &&
                  !r.uploadFolderId.includes('/') &&
                  (docMainFolder ? r.group.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0]) : true)
                );
                const resolvedFolderId = topFolderId || liveId || (matchingRow?.uploadFolderId) || fallbackPath;

                return { currentVessel, subFolderPath, fallbackPath, resolvedFolderId, targetSiteId, targetDriveId };
              };

              return (
                <div style={{ display: 'inline-flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                  {/* 1. Upload Multiple Files */}
                  <label style={{
                    background: '#0284c7', color: '#fff', border: 'none', borderRadius: 8,
                    padding: '7px 14px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                    display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 4px rgba(2,132,199,0.2)',
                    transition: 'opacity 0.15s, transform 0.15s',
                  }}>
                    <input
                      type="file"
                      multiple
                      style={{ display: 'none' }}
                      onChange={e => {
                        const filesList = Array.from(e.target.files || []);
                        if (filesList.length === 0) return;
                        e.target.value = '';
                        const { currentVessel, subFolderPath, resolvedFolderId, targetSiteId, targetDriveId } = resolveDocPageUploadTarget();
                        const bulkFiles = filesList.map(f => ({ file: f, relativePath: f.name }));
                        host._openBulkUpload(bulkFiles, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode, targetSiteId, targetDriveId);
                      }}
                    />
                    <span>⬆</span> Upload Files
                  </label>

                  {/* 2. Upload Entire Folder */}
                  <label style={{
                    background: '#059669', color: '#fff', border: 'none', borderRadius: 8,
                    padding: '7px 14px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                    display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 4px rgba(5,150,105,0.2)',
                    transition: 'opacity 0.15s, transform 0.15s',
                  }}>
                    <input
                      type="file"
                      {...({ webkitdirectory: '', directory: '' } as any)}
                      multiple
                      style={{ display: 'none' }}
                      onChange={e => {
                        const filesList = Array.from(e.target.files || []);
                        if (filesList.length === 0) return;
                        e.target.value = '';
                        const { currentVessel, subFolderPath, resolvedFolderId, targetSiteId, targetDriveId } = resolveDocPageUploadTarget();
                        const bulkFiles = filesList.map(f => ({
                          file: f,
                          relativePath: (f as any).webkitRelativePath || f.name,
                        }));
                        host._openBulkUpload(bulkFiles, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode, targetSiteId, targetDriveId);
                      }}
                    />
                    <span>📁</span> Upload Folder
                  </label>

                  {/* 3. Delete Uploaded Folder Button (if uploaded folders exist) */}
                  {uploadedOnlySubfolderNames.length > 0 && (
                    <button
                      onClick={() => {
                        const { currentVessel, subFolderPath } = resolveDocPageUploadTarget();
                        host._openFolderDeleteDialog({
                          folderId: '',
                          folderName: '',
                          folderPath: subFolderPath,
                          vesselName: currentVessel,
                          mainFolder: docMainFolder || 'Technical & Crewing',
                          subfolderNames: uploadedOnlySubfolderNames,
                        });
                      }}
                      style={{
                        background: '#fff1f2', color: '#e11d48', border: '1px solid #fecdd3', borderRadius: 8,
                        padding: '7px 14px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                        display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 4px rgba(225,29,72,0.1)',
                        transition: 'all 0.15s',
                      }}
                      title="Delete uploaded folders (non-template only)"
                    >
                      <span>🗑</span> Delete Uploaded Folder
                    </button>
                  )}
                </div>
              );
            })()}

          </div>

          {/* Folder view / List view Pill Toggle — fixed position */}
          <div style={{ display: 'inline-flex', flexShrink: 0, background: 'var(--vdms-surface)', border: '1px solid var(--vdms-border)', borderRadius: 8, padding: 3, gap: 2 }}>
            <button
              onClick={openFolderViewFromListContext}
              style={{
                padding: '5px 12px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                background: docViewMode === 'folder' ? 'var(--vdms-toggle-active-bg)' : 'transparent',
                color: docViewMode === 'folder' ? 'var(--vdms-toggle-active-text)' : 'var(--vdms-text-muted)',
              }}
            >
              ⊞ Folder view
            </button>
            <button
              onClick={openListViewFromFolderContext}
              style={{
                padding: '5px 12px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                background: docViewMode === 'list' ? 'var(--vdms-toggle-active-bg)' : 'transparent',
                color: docViewMode === 'list' ? 'var(--vdms-toggle-active-text)' : 'var(--vdms-text-muted)',
              }}
            >
              ☰ List view
            </button>
          </div>
        </div>

        {/* ── Filter Toolbar ── */}
        <div style={{ background: 'var(--vdms-surface)', borderRadius: 10, padding: '10px 12px', border: '1px solid var(--vdms-border)', display: 'flex', flexDirection: 'column', gap: 10 }}>
          {docViewMode === 'list' && (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 8, paddingBottom: 6, borderBottom: '1px solid var(--vdms-border-soft)' }}>
              <span style={{ fontSize: 12, color: 'var(--vdms-text-muted)' }}>
                {docScopeType === 'vessels' && `Showing documents for ${vesselFilter !== 'all' ? vesselFilter : 'individual vessels'}`}
                {docScopeType === 'common' && 'Showing documents shared across all vessels (Common for all ships)'}
                {docScopeType === 'kaizen' && 'Showing global Kaizen - Knowledge Bank documents'}
                {docScopeType === 'shared_docs' && 'Showing NKSDocMan · Shared Documents library'}
                {docScopeType === 'documents' && 'Showing Communication Site · Documents library'}
                {docScopeType === 'sites' && 'Showing connected SharePoint sites & document libraries'}
              </span>
              <button
                type="button"
                onClick={() => host._exportVesselsExcel()}
                disabled={host.state.vesselsExcelExportBusy}
                title="Download a full vessel + folder summary (.xlsx) — independent of the filters above"
                style={{
                  border: '1px solid #bae6fd', background: host.state.vesselsExcelExportBusy ? 'var(--vdms-border-soft)' : '#f0f9ff',
                  color: host.state.vesselsExcelExportBusy ? 'var(--vdms-text-faint)' : '#0369a1', borderRadius: 8,
                  padding: '6px 12px', fontSize: 12, fontWeight: 600,
                  cursor: host.state.vesselsExcelExportBusy ? 'not-allowed' : 'pointer',
                  display: 'inline-flex', alignItems: 'center', gap: 6, whiteSpace: 'nowrap',
                }}
              >
                {host.state.vesselsExcelExportBusy ? '⏳ Exporting…' : '⬇ Export vessels report (.xlsx)'}
              </button>
            </div>
          )}

          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
            <div style={{ position: 'relative', flex: '1 1 180px', minWidth: 160 }}>
              <span style={{ position: 'absolute', left: 9, top: '50%', transform: 'translateY(-50%)', color: 'var(--vdms-text-faint)', fontSize: 12, zIndex: 1 }}>🔍</span>
              {/* Debounced: typing feeds this page's expensive per-render
                  recompute (groupCatActive's recursive folder walk / fleet-wide
                  row scan below), which used to run once per keystroke and made
                  the box freeze for 15-20s while it caught up. DebouncedSearchInput
                  keeps the box itself responsive (local state) and only pushes
                  into host.setState / _scheduleGlobalSearch 300ms after typing
                  stops. See DebouncedSearchInput.tsx's doc comment. */}
              <DebouncedSearchInput
                value={textFilter}
                placeholder="Search vessel, file name (partial ok), folder, group, category, path..."
                title="Matches any format/file type. You can combine terms — e.g. a vessel name plus a partial file name — separated by spaces; each word can match a different field."
                style={{ width: '100%', padding: '6px 26px 6px 28px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, outline: 'none', boxSizing: 'border-box' }}
                onChange={nextValue => {
                  host.setState({ textFilter: nextValue, docListPage: 0, searchDropdownActiveIndex: -1 });
                  // Fleet-wide backend lookup, debounced: finds which
                  // vessels (beyond whatever's already loaded into `rows`)
                  // contain a match and lazily loads them, so this box
                  // actually searches the whole fleet rather than only
                  // vessels the user has already opened this session.
                  host._scheduleGlobalSearch(nextValue);
                }}
              />
            </div>
            {docScopeType !== 'kaizen' && MAIN_FOLDERS.length > 0 && (
              <select
                aria-label="Group filter"
                value={docGroupFilter}
                onChange={e => {
                  const val = e.target.value;
                  const mainFolder = val === 'all' ? null : val as MainFolderKey;
                  if (docViewMode === 'folder' && mainFolder) {
                    const nextStack = folderPathStack.length > 0
                      ? [{ id: mainFolder, name: mainFolder }, ...folderPathStack.slice(1)]
                      : [{ id: mainFolder, name: mainFolder }];
                    host._pushFolderNav(nextStack, mainFolder);
                  }
                  if (docViewMode === 'folder') {
                    host.setState({ docGroupFilter: val, docCategoryFilter: 'all', docGroupLevelFilter: 'all', docLeafCategoryFilter: 'all', docSubCategoryFilter: 'all', catFilter: 'all', docListPage: 0, docMainFolder: mainFolder });
                  } else {
                    host.setState({ docGroupFilter: val, docCategoryFilter: 'all', docGroupLevelFilter: 'all', docLeafCategoryFilter: 'all', docSubCategoryFilter: 'all', catFilter: 'all', docListPage: 0 });
                  }
                }}
                style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface)', outline: 'none', maxWidth: 160 }}
              >
                <option value="all">All main folders</option>
                {allGroups.map(g => (
                  <option key={g} value={g}>{g}</option>
                ))}
              </select>
            )}
            {/* Vessel filter — available in both Folder view and List view,
                not only at the department vessel-tile page or in List view
                (mainFolderPage-only case used to hide it everywhere else in
                Folder view browsing). */}
            {(docScopeType === 'vessels' || docScopeType === 'sites' || docScopeType === 'shared_docs' || docScopeType === 'documents') && (
              <select
                aria-label="Vessel filter"
                value={mainFolderPage ? (vesselFilter === 'all' ? '' : vesselFilter) : vesselFilter}
                onChange={e => applyVesselFilterSelection(e.target.value)}
                style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface)', outline: 'none', maxWidth: 160 }}
              >
                {!mainFolderPage && <option value="all">All vessels</option>}
                {mainFolderPage && <option value="">Select vessel</option>}
                {docScopeType === 'vessels' && mainFolderPage ? (
                  vesselFilterOptions.map(v => (
                    <option key={v.id || v.name} value={v.name}>{v.name}</option>
                  ))
                ) : docScopeType === 'sites' ? (
                  siteVesselOptions.map(v => (
                    <option key={v.id || v.name} value={v.name}>{v.name}</option>
                  ))
                ) : (
                  (docScopeType === 'shared_docs' || docScopeType === 'documents' ? libraryVesselOptions : distinctVesselsInScope).map(vName => (
                    <option key={vName} value={vName}>{vName}</option>
                  ))
                )}
              </select>
            )}
            <select
              aria-label="Main folder filter"
              value={mainFolderSelectValue}
              onChange={e => {
                const val = e.target.value;
                // An explicit pick ends the "All main folders" placeholder.
                (host as any)._mainFolderImplicit = undefined;
                // "All main folders" while inside a live main folder: go back to
                // the library root (site + drive) so every main folder shows
                // again, with the vessel / sub-folder / category filters cleared.
                if (val === 'all' && docViewMode === 'folder' && breadcrumbMainFolderName && liveStackPrefixLen >= 0) {
                  const rootStack = folderPathStack.slice(0, liveStackPrefixLen);
                  host._pushFolderNav(rootStack, liveStackPrefixLen === 3 ? 'SharePoint Sites' : ((folderPathStack[0]?.id === 'lib:documents') ? 'Documents' : 'Shared Documents'));
                  host.setState({
                    docCategoryFilter: 'all',
                    docSubfolderOtherFilter: 'all',
                    vesselFilter: 'all',
                    docGroupLevelFilter: 'all',
                    docLeafCategoryFilter: 'all',
                    docSubCategoryFilter: 'all',
                    catFilter: 'all',
                    docListPage: 0,
                  });
                  return;
                }
                // With a vessel selected, picking another main folder is the
                // same as its "is also under" chip: keep Sub-folder / Group /
                // Category and open the same sub-folder in the vessel there.
                if (val !== 'all' && activeVesselFilter && docSubfolderOtherFilter !== 'all' &&
                  (docScopeType === 'sites' || docScopeType === 'shared_docs' || docScopeType === 'documents') &&
                  carryBasePaths(vesselFilter, val).length > 0) {
                  switchVesselMainFolder(val);
                  return;
                }
                host.setState({ docCategoryFilter: val, docGroupLevelFilter: 'all', docLeafCategoryFilter: 'all', docSubCategoryFilter: 'all', docSubfolderOtherFilter: 'all', catFilter: val, docListPage: 0 });
                if (val !== 'all') {
                  navigateToLiveMainFolder(val);
                } else if (docViewMode === 'folder') {
                  // "All main folders": leave the current main folder and show
                  // every main folder again, i.e. go back to the library root
                  // (site + drive) instead of staying inside the old folder.
                  const rootId = folderPathStack[0]?.id;
                  const rootLen = rootId === 'sites_root' ? 3
                    : ((rootId === 'lib:shared_documents' || rootId === 'lib:documents') ? 1 : -1);
                  if (rootLen >= 0 && folderPathStack.length > rootLen) {
                    const rootStack = folderPathStack.slice(0, rootLen);
                    host._pushFolderNav(rootStack, rootId === 'lib:documents' ? 'Documents' : (rootId === 'lib:shared_documents' ? 'Shared Documents' : 'SharePoint Sites'));
                    const rootFilters = deriveDocFiltersFromStack(rootStack);
                    host.setState({
                      docCategoryFilter: rootFilters.docCategoryFilter,
                      docSubfolderOtherFilter: rootFilters.docSubfolderOtherFilter,
                      vesselFilter: 'all',
                      docListPage: 0,
                    });
                  }
                }
              }}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface)', outline: 'none', maxWidth: 180 }}
            >
              <option value="all">All main folders</option>
              {mainFolderOptions.map(folder => <option key={folder} value={folder}>{folder}</option>)}
            </select>
            {subfolderTree.length > 0 ? (
              <FolderTreeSelect
                tree={subfolderTree}
                value={docSubfolderOtherFilter}
                selectedPath={(() => {
                  const sp = (host as any)._subfolderSelectedPath as string[] | undefined;
                  return sp && sp.length && sp[sp.length - 1].trim().toLowerCase() === docSubfolderOtherFilter.trim().toLowerCase() ? sp : undefined;
                })()}
                onChange={(val, path, nodeId) => {
                  // Keep the "All categories" dropdown in step with whatever
                  // was just picked here: a folder that matches one of the
                  // fixed category names selects that category too, and
                  // clearing the tree back to "All sub-folders" clears the
                  // category dropdown the same way.
                  // Category / Sub-category come from the picked PATH, not just
                  // the folder's own name: Manuals > Other Manuals > Electric_Maker
                  // is category "Other Manuals" + sub-category "Electric_Maker"
                  // (matching the name alone wrongly gave "Electrical").
                  let pathCategory = '';
                  let pathSubCategory = 'all';
                  if (val !== 'all') {
                    const segs = path || [];
                    let gi = -1;
                    for (let i = segs.length - 1; i >= 0; i--) {
                      const fg = folderGroupsOf(segs[i]);
                      if (fg.drawings !== fg.manuals) { gi = i; break; }
                    }
                    if (gi >= 0) {
                      const known = folderGroupsOf(segs[gi]).drawings ? drawingCategories : manualCategories;
                      for (let i = gi + 1; i < segs.length && !pathCategory; i++) {
                        const hit = known.find(c => segs[i].trim().toLowerCase() === c.trim().toLowerCase())
                          || known.find(c => folderMatchesCategory(segs[i], c));
                        if (hit) {
                          pathCategory = hit;
                          if (segs[i + 1]) pathSubCategory = segs[i + 1];
                        }
                      }
                    }
                  }
                  const matchedCategory = val === 'all' ? 'all' : (pathCategory || findCategoryForFolderName(val) || docLeafCategoryFilter);
                  // Same for the Groups dropdown: picking "To Be Classified"
                  // (or anything nested under it) selects that group; picking
                  // a folder outside it releases a previously auto-selected
                  // "To Be Classified" group.
                  const inToBeClassified = val !== 'all' && (path || []).some(isToBeClassifiedName);
                  // Picking a folder under "Drawings" / "Manuals" selects that
                  // group too, so a name that exists in both (e.g. Safety)
                  // resolves to the branch that was actually clicked.
                  const pathGroup = val === 'all' ? '' : ((path || []).map(seg => seg.trim().toLowerCase())
                    .map(seg => (seg === 'drawings' || seg === 'drawing') ? 'Drawings' : ((seg === 'manuals' || seg === 'manual') ? 'Manuals' : ''))
                    .filter(Boolean).pop() || '');
                  const nextGroupLevel = inToBeClassified
                    ? TBC_GROUP_LABEL
                    : (pathGroup || (docGroupLevelFilter === TBC_GROUP_LABEL ? 'all' : docGroupLevelFilter));
                  (host as any)._subCategoryPickKey = undefined;
                  host.setState({
                    docGroupLevelFilter: nextGroupLevel,
                    docSubfolderOtherFilter: val,
                    docLeafCategoryFilter: matchedCategory,
                    docSubCategoryFilter: pathSubCategory,
                    docListPage: 0,
                  });
                  // Remember the exact branch clicked — the filter value is only
                  // the folder's name, and e.g. "Electrical" exists under both
                  // Drawings and Manuals.
                  (host as any)._subfolderSelectedPath = val === 'all' ? undefined : path;
                  if (val === 'all') return;
                  if (atSitesRoot) navigateToLiveSubfolder(val, path, nodeId);
                  else navigateToDeptSubfolder(path);
                }}
                allLabel="All sub-folders"
                canLoadChildren={node => !liveFolderTreeComplete && !!effectiveLiveSiteId && !!effectiveLiveDriveId && !!node.id && node.id !== node.name && !/^(sf_|category_|common|kaizen_root|dept_|vessels_root|specific_vessels|sites_root|site:|drive:|lib:)/.test(node.id)}
                loadChildren={async (node, treePath) => {
                  if (!effectiveLiveSiteId || !effectiveLiveDriveId) return [];
                  const lazyEntries: Map<string, LiveSubfolderEntry> = ((host as any)._lazySubfolderEntries ||= new Map<string, LiveSubfolderEntry>());
                  const lowName = node.name.trim().toLowerCase();
                  const parentPathKey = 'path:' + (treePath || []).map(n => n.trim().toLowerCase()).join('/');
                  const parent = lazyEntries.get(parentPathKey)
                    || liveSubfolderEntries.find(e => e.pathNames.map(n => n.trim().toLowerCase()).join('/') === parentPathKey.slice(5))
                    || liveSubfolderEntryByName.get(lowName) || lazyEntries.get(lowName);
                  const parentIds = parent ? parent.pathIds : [node.id];
                  const parentNames = parent ? parent.pathNames : [node.name];
                  const items = await host._loadAndCacheSiteFolderChildren(effectiveLiveSiteId, effectiveLiveDriveId, node.id);
                  const seen = new Set<string>();
                  const kids: FolderTreeNode[] = [];
                  (items || []).forEach((f: any) => {
                    const name = String(f?.name || '').trim();
                    if (!f || !f.folder || !f.id || !name) return;
                    const low = name.toLowerCase();
                    if (seen.has(low)) return;
                    seen.add(low);
                    const lazyEntry: LiveSubfolderEntry = {
                      name, id: f.id, depth: (parent ? parent.depth : 1) + 1,
                      pathIds: [...parentIds, f.id], pathNames: [...parentNames, name],
                    };
                    // Keyed by full tree path (unambiguous) and, for older
                    // callers, by bare name (last one wins).
                    lazyEntries.set('path:' + [...(treePath || []), name].map(n => n.trim().toLowerCase()).join('/'), lazyEntry);
                    lazyEntries.set(low, lazyEntry);
                    kids.push({ name, id: f.id, children: [] });
                  });
                  return kids;
                }}
                title="Folders that aren't a main department or a vessel (e.g. Report, Share with...)"
              />
            ) : scopedMainFolderForSubfolders ? (
              <select
                value="all"
                disabled
                style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text-faint)', outline: 'none', maxWidth: 180, cursor: 'not-allowed' }}
                title={`"${scopedMainFolderForSubfolders}" has no sub-folders`}
              >
                <option value="all">No folder found</option>
              </select>
            ) : null}
            {/* "Groups" filter — fixed Drawings/Manuals taxonomy (see
                drawingCategories/manualCategories above), bound to
                docGroupLevelFilter. Picking a group scopes the "All
                categories" dropdown right after it (categoryOptions) to
                that group's categories, and resets anything deeper
                (category / sub-category / sub-folder tree) so the three
                stay in sync. */}
            <select
              aria-label="Document type filter"
              value={groupOptions.length === 0 ? 'all' : docGroupLevelFilter}
              disabled={groupOptions.length === 0 && !(vesselTaxonomyKey && libTaxonomyState !== 'ready')}
              title={groupOptions.length === 0 ? 'No classified groups found in these folders' : undefined}
              onChange={e => {
                const val = e.target.value;
                // Mirror into the Sub-folder tree: "To Be Classified" jumps
                // to that folder node when it exists; leaving it clears a
                // sub-folder that was only there because of it.
                if (val === TBC_GROUP_LABEL) {
                  const tbcNode = libraryWideScope ? null : findSubfolderNodeForCategory(TBC_GROUP_LABEL, { group: TBC_GROUP_LABEL });
                  host.setState({
                    docGroupLevelFilter: val,
                    docLeafCategoryFilter: 'all',
                    docSubCategoryFilter: 'all',
                    docSubfolderOtherFilter: tbcNode ? tbcNode.name : docSubfolderOtherFilter,
                    docListPage: 0,
                  });
                  if (tbcNode) {
                    (host as any)._subfolderSelectedPath = tbcNode.path;
                    if (atSitesRoot) navigateToLiveSubfolder(tbcNode.name, tbcNode.path);
                    else navigateToDeptSubfolder(tbcNode.path);
                  }
                  return;
                }
                // Drawings / Manuals: jump the Sub-folder tree to that group's
                // own folder (never one under "To Be Classified"). "All
                // groups": drop a sub-folder that was only there for a group.
                // (Library-wide: opening one vessel's group folder would also
                // change the Vessel / Main folder filters — just filter.)
                const groupNode = (val === 'all' || libraryWideScope) ? null : findSubfolderNodeForCategory(val, { group: val });
                (host as any)._subfolderSelectedPath = groupNode ? groupNode.path : undefined;
                host.setState({
                  docGroupLevelFilter: val,
                  docLeafCategoryFilter: 'all',
                  docSubCategoryFilter: 'all',
                  docSubfolderOtherFilter: groupNode
                    ? groupNode.name
                    : (val === 'all' || docGroupLevelFilter !== 'all' ? 'all' : docSubfolderOtherFilter),
                  docListPage: 0,
                });
                if (groupNode) {
                  if (atSitesRoot) navigateToLiveSubfolder(groupNode.name, groupNode.path);
                  else navigateToDeptSubfolder(groupNode.path);
                }
              }}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface)', outline: 'none', maxWidth: 160, cursor: groupOptions.length === 0 ? 'not-allowed' : 'pointer', opacity: groupOptions.length === 0 ? 0.6 : 1 }}
            >
              <option value="all">{groupOptions.length === 0 ? (vesselTaxonomyKey && libTaxonomyState !== 'ready' ? 'Loading groups…' : 'No groups found') : 'All groups'}</option>
              {groupOptions.map(g => <option key={g} value={g}>{g}</option>)}
            </select>
            {(categoryOptions.length > 0 || categoryOptionGroups.length > 0 || libraryWideScope || (!!vesselTaxonomyKey && libTaxonomyState !== 'ready')) ? (
              <select
                aria-label="Category filter"
                value={(() => {
                  // Classified (All groups) options carry "<Group>::<Category>";
                  // show the exact entry the user clicked (a name such as
                  // "Electrical" is listed under both Drawings and Manuals).
                  if (docGroupLevelFilter !== 'all' || categoryOptionGroups.length === 0 || docLeafCategoryFilter === 'all') return docLeafCategoryFilter;
                  const picked = (host as any)._categoryPickKey as string | undefined;
                  if (picked && picked.slice(picked.indexOf('::') + 2) === docLeafCategoryFilter &&
                    categoryOptionGroups.some(x => `${x.group}::${docLeafCategoryFilter}` === picked && x.categories.indexOf(docLeafCategoryFilter) !== -1)) return picked;
                  const owner = categoryOptionGroups.find(x => x.categories.indexOf(docLeafCategoryFilter) !== -1);
                  return owner ? `${owner.group}::${docLeafCategoryFilter}` : docLeafCategoryFilter;
                })()}
                onChange={e => {
                  let val = e.target.value;
                  // "<Group>::<Category>" (All groups, classified list): only the
                  // Category changes — the Groups filter stays exactly as the
                  // user set it. The group is used just to find the matching
                  // folder in the Sub-folder tree.
                  let groupForPick = docGroupLevelFilter;
                  const sepAt = val.indexOf('::');
                  (host as any)._categoryPickKey = sepAt > 0 ? val : undefined;
                  if (sepAt > 0) {
                    groupForPick = val.slice(0, sepAt);
                    val = val.slice(sepAt + 2);
                  }
                  // Mirror the pick into the Sub-folder tree so the two
                  // filters never disagree: picking "all" here clears the
                  // tree back to "All sub-folders" too, and picking a real
                  // category jumps the tree to whichever sub-folder node
                  // matches it, when one exists.
                  if (val === 'all') {
                    // Fall back to the selected group's folder, if any.
                    const groupNode = (!libraryWideScope && docGroupLevelFilter !== 'all') ? findSubfolderNodeForCategory(docGroupLevelFilter, { group: docGroupLevelFilter }) : null;
                    (host as any)._subfolderSelectedPath = groupNode ? groupNode.path : undefined;
                    host.setState({ docLeafCategoryFilter: 'all', docSubCategoryFilter: 'all', docSubfolderOtherFilter: groupNode ? groupNode.name : 'all', docListPage: 0 });
                    if (groupNode) {
                      if (atSitesRoot) navigateToLiveSubfolder(groupNode.name, groupNode.path);
                      else navigateToDeptSubfolder(groupNode.path);
                    }
                    return;
                  }
                  // Opening the matching folder would re-derive Group / Vessel /
                  // Main folder from its path, so only do it when the Group is
                  // already chosen and the view isn't library-wide.
                  const match = (!libraryWideScope && docGroupLevelFilter !== 'all')
                    ? findSubfolderNodeForCategory(val, { group: groupForPick })
                    : null;
                  (host as any)._subfolderSelectedPath = match ? match.path : (host as any)._subfolderSelectedPath;
                  host.setState({
                    docLeafCategoryFilter: val,
                    docSubCategoryFilter: 'all',
                    docSubfolderOtherFilter: match ? match.name : docSubfolderOtherFilter,
                    docListPage: 0,
                  });
                  if (match) {
                    if (atSitesRoot) navigateToLiveSubfolder(match.name, match.path);
                    else navigateToDeptSubfolder(match.path);
                  }
                }}
                style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface)', outline: 'none', maxWidth: 180 }}
              >
                <option value="all">All categories</option>
                {docGroupLevelFilter === 'all' && categoryOptionGroups.length > 0 ? (
                  <>
                    {docLeafCategoryFilter !== 'all' && !categoryOptionGroups.some(x => x.categories.indexOf(docLeafCategoryFilter) !== -1) && (
                      <option value={docLeafCategoryFilter}>{docLeafCategoryFilter}</option>
                    )}
                    {categoryOptionGroups.map(x => (
                      <optgroup key={x.group} label={x.group}>
                        {x.categories.map(category => <option key={`${x.group}::${category}`} value={`${x.group}::${category}`}>{category}</option>)}
                      </optgroup>
                    ))}
                  </>
                ) : categoryOptions.map(category => <option key={category} value={category}>{category}</option>)}
              </select>
            ) : (
              <select aria-label="Category filter" value="all" disabled style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text-faint)', outline: 'none', maxWidth: 180, cursor: 'not-allowed' }}>
                <option value="all">No categories found</option>
              </select>
            )}
            {(subCategoryOptions.length > 0 || subCategoryOptionGroups.length > 0 || taxonomyScoped) ? (
              <select
                aria-label="Sub-category filter"
                value={(() => {
                  // Classified options carry "<Group>::<Category>::<Sub>"; show
                  // the exact entry the user clicked.
                  if (docSubCategoryFilter === 'all' || subCategoryOptionGroups.length === 0) return docSubCategoryFilter;
                  const keyOf = (x: { group: string; category: string }): string => `${x.group}::${x.category}::${docSubCategoryFilter}`;
                  const picked = (host as any)._subCategoryPickKey as string | undefined;
                  const hits = subCategoryOptionGroups.filter(x => x.subs.indexOf(docSubCategoryFilter) !== -1);
                  if (picked && hits.some(x => keyOf(x) === picked)) return picked;
                  return hits.length ? keyOf(hits[0]) : docSubCategoryFilter;
                })()}
                onChange={e => {
                  const raw = e.target.value;
                  // Only the Sub-category changes — Groups, Categories, Vessel and
                  // Main folder stay exactly as the user set them.
                  // Classified entries carry "<Group>::<Category>::<Sub-category>";
                  // the group/category there are display-only.
                  const parts = raw.split('::');
                  const encoded = parts.length === 3 && !!parts[0] && !!parts[1] && !!parts[2];
                  (host as any)._subCategoryPickKey = encoded ? raw : undefined;
                  const val = encoded ? parts[2] : raw;
                  // Opening the matching folder re-derives Group / Vessel / Main
                  // folder from that folder's path (deriveLiveNavFilterState),
                  // so only keep the Sub-folder tree in step when Group AND
                  // Category are already chosen and the view isn't library-wide.
                  const canSyncTree = !libraryWideScope && docGroupLevelFilter !== 'all' && docLeafCategoryFilter !== 'all';
                  const ctx = { group: docGroupLevelFilter, category: docLeafCategoryFilter };
                  const node = !canSyncTree ? null
                    : (val !== 'all'
                      ? findSubfolderNodeForCategory(val, ctx)
                      : findSubfolderNodeForCategory(docLeafCategoryFilter, { group: docGroupLevelFilter }));
                  if (node) (host as any)._subfolderSelectedPath = node.path;
                  host.setState({
                    docSubCategoryFilter: val,
                    docSubfolderOtherFilter: node ? node.name : docSubfolderOtherFilter,
                    docListPage: 0,
                  });
                  if (node) {
                    if (atSitesRoot) navigateToLiveSubfolder(node.name, node.path);
                    else navigateToDeptSubfolder(node.path);
                  }
                }}
                style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface)', outline: 'none', maxWidth: 180 }}
              >
                <option value="all">All sub-categories</option>
                {libTaxonomyState !== 'ready' && (
                  <option value="__vdms_loading" disabled>
                    {libTaxonomyState === 'retrying' ? 'Couldn\'t load sub-categories — retrying…' : 'Loading sub-categories…'}
                  </option>
                )}
                {subCategoryUngrouped.map(subCategory => <option key={subCategory} value={subCategory}>{subCategory}</option>)}
                {subCategoryOptionGroups.map(x => (
                  <optgroup key={`${x.group}::${x.category}`} label={docGroupLevelFilter === 'all' ? `${x.group} › ${x.category}` : x.category}>
                    {x.subs.map(sub => <option key={`${x.group}::${x.category}::${sub}`} value={`${x.group}::${x.category}::${sub}`}>{sub}</option>)}
                  </optgroup>
                ))}
              </select>
            ) : (
              <select aria-label="Sub-category filter" value="all" disabled style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text-faint)', outline: 'none', maxWidth: 180, cursor: 'not-allowed' }}>
                <option value="all">No sub-categories found</option>
              </select>
            )}
            <select
              aria-label="Attachment status filter"
              value={attachmentFilter}
              onChange={e => host.setState({ attachmentFilter: e.target.value as any, docListPage: 0 })}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface)', outline: 'none', maxWidth: 180 }}
            >
              <option value="all">All attachment status</option>
              <option value="attached">Attachment Available</option>
              <option value="not_attached">Attachment Required</option>
            </select>
            <select
              aria-label="Sort filter"
              value={docListSort}
              onChange={e => host.setState({ docListSort: e.target.value as any, docListPage: 0 })}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid var(--vdms-border)', fontSize: 12, background: 'var(--vdms-surface)', outline: 'none' }}
            >
              <option value="default">Default order</option>
              <option value="name_az">Name A–Z</option>
              <option value="newest">Newest</option>
            </select>

            {/* vesselLoadingName / documentVesselsLoadingMore / documentFilesLoading
                are ALL per-vessel background loads: every setter of
                documentFilesLoading (_loadVesselRowsFromApi,
                _mergeLiveSharePointFiles — walked by _loadFilesForVessel,
                _fetchApprovedFilesForVessel, and the Sites-scope
                "this folder matches a known vessel" prefetch below) sets it
                alongside vesselLoadingName for a specific vessel, and that
                vessel is frequently NOT the one the user is currently
                looking at — e.g. browsing SharePoint Sites can trigger a
                background load for a vessel-named folder, and the flag
                stays on (with that vessel's name) even after navigating
                elsewhere, since nothing resets it on scope/folder change.
                Previously only the vesselLoadingName/documentVesselsLoadingMore
                *branch* was scope-gated while documentFilesLoading was OR'd
                in unconditionally — so this still rendered "Loading
                <stale vessel>..." on Shared Documents whenever a background
                vessel walk happened to be in flight. None of these three
                flags mean anything outside the vessels scope, so the whole
                banner is gated on it. */}
            {(docScopeType === 'vessels' && (documentFilesLoading || vesselLoadingName || documentVesselsLoadingMore)) && (
              <div style={{
                display: 'inline-flex', alignItems: 'center', gap: 6, padding: '4px 10px',
                borderRadius: 8, background: '#eff6ff', border: '1px solid #bfdbfe',
                fontSize: 11, color: '#1d4ed8', fontWeight: 600,
              }}>
                <span style={{ display: 'inline-block', animation: 'spin 1s linear infinite' }}>⏳</span>
                <span>{vesselLoadingName ? `Loading ${vesselLoadingName}...` : (documentVesselsLoadingMore ? 'Loading vessels...' : 'Syncing live files...')}</span>
              </div>
            )}
            {/* The Folder view / List view toggle used to be duplicated here
                (a second, compact "::"/"☰" pair) in addition to the pill
                toggle in the Module Header above. Because this Filter
                Toolbar row's own contents change between view modes (which
                selects are shown/enabled) and wraps on narrow widths, a
                `marginLeft: 'auto'`-positioned toggle here visibly jumped
                position on every switch. There is now a single toggle, in
                the Module Header, which does not move. */}
          </div>
        </div>
        </div>{/* /sticky header */}

        <CopilotSearchPanel host={host} />

        {/* Vessel's other main folders — a vessel commonly has a folder under
            several main folders (Technical & Crewing, Commercial &
            Chartering, Insurance, ...). This is the "which main folders have
            this vessel" side strip: every main folder siteVesselFolderPaths
            recorded a folder for the selected vessel under, as a row of
            chips next to the Main folder dropdown. Clicking one jumps
            straight into that vessel's folder there (navigateToLiveMainFolder
            already prefers the vessel's own sub-folder over the main
            folder's root when a vessel is selected — see its definition
            above), and the chip for wherever you're currently standing is
            highlighted. Only rendered when the vessel actually has more than
            one home; a single-main-folder vessel gets nothing extra here. */}
        {docViewMode === 'folder' && !docCompareMode &&
          (docScopeType === 'sites' || docScopeType === 'shared_docs' || docScopeType === 'documents') &&
          vesselMainFolders.length > 1 && (
            <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 6 }}>
              <span style={{ fontSize: 11, color: 'var(--vdms-text-muted)', fontWeight: 600 }}>
                {vesselFilter} is also under:
              </span>
              {vesselMainFolders.map(m => {
                const active = docCategoryFilter.trim().toLowerCase() === m.trim().toLowerCase();
                return (
                  <button
                    key={m}
                    type="button"
                    onClick={() => {
                      // Mirror the "All main folders" dropdown's own
                      // onChange exactly (below): it sets the filter fields
                      // itself, immediately, rather than leaving them to be
                      // derived later from wherever navigateToLiveMainFolder
                      // manages to land. The dropdown visibly updates the
                      // moment it's used; this chip, calling
                      // navigateToLiveMainFolder alone with no filter update
                      // of its own, showed nothing changing at all whenever
                      // that call's own navigation didn't resolve (a slow or
                      // failed Graph fetch for a main folder never opened
                      // this session) — the click looked like a complete
                      // no-op instead of at least switching the Main folder
                      // filter/highlighted chip right away.
                      //
                      // Same vessel, other main folder: the Sub-folder pick
                      // (e.g. Drawings > Hull) and Group / Category are kept
                      // and the same sub-folder is opened under `m` — they
                      // used to be reset to "All …" here.
                      switchVesselMainFolder(m);
                    }}
                    disabled={active}
                    title={active ? `Currently viewing ${vesselFilter} under ${m}` : `Jump to ${vesselFilter} under ${m}`}
                    style={{
                      padding: '3px 10px', borderRadius: 999, fontSize: 11, fontWeight: 600,
                      border: active ? '1px solid #2563eb' : '1px solid var(--vdms-border)',
                      background: active ? '#dbeafe' : 'var(--vdms-surface)',
                      color: active ? '#1d4ed8' : 'var(--vdms-text-muted)',
                      cursor: active ? 'default' : 'pointer',
                    }}
                  >
                    {m}
                  </button>
                );
              })}
              <button
                type="button"
                onClick={() => host.setState({ docCompareMode: true })}
                title={`Show ${vesselFilter}'s folders from all ${vesselMainFolders.length} main folders side by side`}
                style={{
                  padding: '3px 10px', borderRadius: 999, fontSize: 11, fontWeight: 600,
                  border: '1px dashed #94a3b8', background: 'var(--vdms-surface)', color: 'var(--vdms-text-muted)', cursor: 'pointer',
                  marginLeft: 4,
                }}
              >
                ⊞⊞ Compare all {vesselMainFolders.length}
              </button>
            </div>
          )}

        {/* Success Message Banner */}
        {docUploadMsg && (
          <div style={{ background: '#f0fdf4', border: '1px solid #86efac', borderRadius: 8, padding: '8px 14px', fontSize: 12, color: '#16a34a', display: 'flex', justifyContent: 'space-between' }}>
            <span>✓ {docUploadMsg}</span>
            <button onClick={() => host.setState({ docUploadMsg: null })} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#16a34a', fontWeight: 700 }}>✕</button>
          </div>
        )}

        {docCompareMode && vesselMainFolders.length > 1 ? (
          renderVesselCompareGrid()
        ) : docViewMode === 'folder' ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 24, marginTop: 4 }}>

            {/* Level 0: Main Departments + Kaizen - Knowledge Bank */}
            {stackLevel === 0 ? (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                        {rootFolderCards.filter(item => {
                          if (docGroupFilter !== 'all') {
                            const itemNorm = item.key.toLowerCase();
                            const filterNorm = docGroupFilter.toLowerCase();
                            if (!itemNorm.includes(filterNorm) && !filterNorm.includes(itemNorm)) return false;
                          }
                          if (textFilter) {
                            const q = textFilter.toLowerCase();
                            if (!item.key.toLowerCase().includes(q)) return false;
                          }
                          if (catFilter !== 'all') {
                            const section = catFilter.trim().toLowerCase();
                            const hasSection = scopeRows.some(row =>
                              row.group.trim().toLowerCase() === item.key.trim().toLowerCase() &&
                              (row.category || '').trim().toLowerCase() === section
                            );
                            if (!hasSection) return false;
                          }
                          return true;
                        }).map(item => (
                          <div
                            key={item.key}
                            onClick={() => {
                              if (item.key === 'Kaizen - Knowledge Bank') {
                                host._pushFolderNav([{ id: 'kaizen_root', name: 'Kaizen - Knowledge Bank' }], 'Kaizen - Knowledge Bank');
                                host.setState({ vesselFilter: 'all', docScopeType: 'kaizen', docGroupFilter: 'all' });
                              } else if ((item.key === 'Shared Documents' || item.key === 'Documents') && activeLiveSite?.drive_id) {
                                host._pushFolderNav([
                                  { id: 'sites_root', name: 'SharePoint Sites' },
                                  { id: `site:${activeLiveSite.site_id}`, name: activeLiveSite.sp_site_name || activeLiveSite.site_key },
                                  { id: `drive:${activeLiveSite.drive_id}`, name: item.key },
                                ], 'SharePoint Sites');
                                host.setState({ vesselFilter: 'all', docScopeType: 'sites', docGroupFilter: 'all' });
                              } else if (item.key === 'Shared Documents') {
                                host._pushFolderNav([{ id: 'lib:shared_documents', name: 'Shared Documents' }], 'Shared Documents');
                                host.setState({ vesselFilter: 'all', docScopeType: 'shared_docs', docGroupFilter: 'all' });
                              } else if (item.key === 'Documents') {
                                host._pushFolderNav([{ id: 'lib:documents', name: 'Documents' }], 'Documents');
                                host.setState({ vesselFilter: 'all', docScopeType: 'documents', docGroupFilter: 'all' });
                              } else if (item.key === 'SharePoint Sites') {
                                host._pushFolderNav([{ id: 'sites_root', name: 'SharePoint Sites' }], 'SharePoint Sites');
                                host.setState({ vesselFilter: 'all', docScopeType: 'sites', docGroupFilter: 'all' });
                                if ((host.state.documentSites || []).length === 0) {
                                  void host._loadDocumentSites().catch(() => undefined);
                                }
                              } else if (item.liveFolderId && activeLiveSite) {
                                host._pushFolderNav([
                                  { id: 'sites_root', name: 'SharePoint Sites' },
                                  { id: `site:${activeLiveSite.site_id}`, name: activeLiveSite.sp_site_name || activeLiveSite.site_key },
                                  { id: `drive:${activeLiveSite.drive_id}`, name: activeLiveSite.sp_site_name || 'Documents' },
                                  { id: item.liveFolderId, name: item.key },
                                ], 'SharePoint Sites');
                              } else {
                                host._pushFolderNav([{ id: item.key, name: item.key }], item.key as MainFolderKey);
                                host.setState({ vesselFilter: 'all', docScopeType: 'vessels' });
                              }
                            }}
                            style={{
                              background: 'var(--vdms-surface)', borderRadius: 14, border: '1px solid var(--vdms-border)', padding: 18,
                              display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                              boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                            }}
                          >
                            <div style={{ width: 44, height: 44, borderRadius: 10, background: item.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>{item.emoji}</div>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--vdms-text)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.key}</div>
                              <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2 }}>
                                {item.key === 'Kaizen - Knowledge Bank' ? 'Knowledge base' :
                                 item.key === 'Shared Documents' ? `${activeLiveSite?.sp_site_name || 'Selected site'} · SharePoint Library` :
                                 item.key === 'Documents' ? 'Site Documents Library' :
                                 item.liveFolderId ? 'Live SharePoint folder' :
                                 item.key === 'SharePoint Sites' ? `${(host.state.documentSites || []).length || 4} connected sites` :
                                 `${displayVessels.length} vessels`}
                              </div>
                            </div>
                            <span style={{ color: 'var(--vdms-text-faint)', fontSize: 16 }}>›</span>
                          </div>
                        ))}
                      </div>
                {(textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || docCategoryFilter !== 'all' || docGroupLevelFilter !== 'all' || docLeafCategoryFilter !== 'all' || docSubCategoryFilter !== 'all' || attachmentFilter !== 'all') && (
                  <div style={{ marginTop: 24, background: 'var(--vdms-surface)', border: '1px solid var(--vdms-border)', borderRadius: 12, overflow: 'hidden' }}>
                    <div style={{ padding: '14px 16px', borderBottom: '1px solid var(--vdms-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12 }}>
                      <strong style={{ color: 'var(--vdms-text)', fontSize: 14 }}>Filtered document results</strong>
                      <span style={{ color: 'var(--vdms-text-muted)', fontSize: 12 }}>{groupedList.length} folder{groupedList.length === 1 ? '' : 's'}</span>
                    </div>
                    {groupedList.length === 0 ? (
                      <div style={{ padding: 28, textAlign: 'center', color: 'var(--vdms-text-faint)', fontSize: 13 }}>No documents match the selected filters.</div>
                    ) : (
                      <div style={{ display: 'flex', flexDirection: 'column' }}>
                        {groupedList.slice(0, PAGE_ROWS).map(result => (
                          <button
                            key={result.groupKey}
                            type="button"
                            onClick={() => host._pushFolderNav([{ id: result.groupKey, name: result.subFolderPath.split(' > ').pop() || result.subCategory }], result.group as 'Technical & Crewing' | 'Commercial & Chartering' | 'Insurance' | 'Kaizen - Knowledge Bank' | 'Knowledge Bank')}
                            style={{ border: 0, borderBottom: '1px solid var(--vdms-border-soft)', background: 'var(--vdms-surface)', padding: '12px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 16, textAlign: 'left', cursor: 'pointer' }}
                          >
                            <span style={{ minWidth: 0 }}>
                              <strong style={{ display: 'block', color: 'var(--vdms-text)', fontSize: 13 }}>{result.subFolderPath.split(' > ').pop() || result.subCategory}</strong>
                              <span style={{ display: 'block', color: 'var(--vdms-text-muted)', fontSize: 11, marginTop: 3 }}>{result.vesselName} · {result.group} · {result.category}</span>
                            </span>
                            <span style={{ color: '#0284c7', fontSize: 12, whiteSpace: 'nowrap' }}>{result.files.length} file{result.files.length === 1 ? '' : 's'} ›</span>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>

            ) : atSitesRoot ? (
              (() => {
                // ── Level 1: List of SharePoint Sites ──
                if (stackLevel === 1) {
                  const sitesList = host.state.documentSites || [];

                  if (sitesList.length === 0) {
                    return (
                      <div style={{ background: 'var(--vdms-surface)', borderRadius: 14, border: '1px dashed var(--vdms-border)', padding: 48, textAlign: 'center', color: 'var(--vdms-text-faint)' }}>
                        <div style={{ fontSize: 32, marginBottom: 8 }}>🌐</div>
                        <div style={{ fontWeight: 600, color: 'var(--vdms-text-secondary)', fontSize: 15 }}>No configured document site found</div>
                        <div style={{ fontSize: 13, marginTop: 4 }}>Configure a SharePoint document site before browsing its libraries.</div>
                      </div>
                    );
                  }

                  return (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: '#0369a1', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                          Connected SharePoint Sites ({sitesList.length})
                        </div>
                        <span style={{ fontSize: 12, color: 'var(--vdms-text-muted)' }}>Select a SharePoint site to view document libraries</span>
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 16 }}>
                        {sitesList
                        .filter(site => !textFilter || (site.sp_site_name || site.site_key || '').toLowerCase().includes(textFilter.trim().toLowerCase()))
                        .map((site, sIdx) => {
                          const siteName = site.sp_site_name || site.site_key;
                          const siteId = site.site_id || site.site_key;
                          const isActive = site.site_key === host.state.activeDocumentSite;
                          return (
                            <div
                              key={siteId + sIdx}
                              onClick={() => {
                                const siteNodeId = `site:${siteId}`;
                                host._pushFolderNav([...folderPathStack, { id: siteNodeId, name: siteName }], 'SharePoint Sites');
                              }}
                              style={{
                                background: 'var(--vdms-surface)', borderRadius: 14, border: isActive ? '2px solid #0284c7' : '1px solid var(--vdms-border)',
                                padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                                boxShadow: '0 1px 3px rgba(0,0,0,0.05)', transition: 'all 0.15s ease',
                              }}
                            >
                              <div style={{
                                width: 48, height: 48, borderRadius: 12, background: isActive ? '#e0f2fe' : 'var(--vdms-surface-alt)',
                                display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 24,
                              }}>
                                🌐
                              </div>
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontWeight: 700, fontSize: 15, color: 'var(--vdms-text)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                                  {siteName}
                                </div>
                                <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2, display: 'flex', alignItems: 'center', gap: 6 }}>
                                  <span>SharePoint Site</span>
                                  {isActive && (
                                    <span style={{ background: '#dcfce7', color: '#15803d', fontSize: 10, fontWeight: 700, padding: '1px 6px', borderRadius: 4 }}>
                                      Active
                                    </span>
                                  )}
                                </div>
                              </div>
                              <span style={{ color: '#0284c7', fontSize: 18, fontWeight: 700 }}>›</span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                }

                // ── Level 2: Site Libraries / Drives ──
                const siteNode = folderPathStack[1];
                const rawSiteId = (siteNode?.id || '').replace(/^site:/, '');
                const matchedSite = (host.state.documentSites || []).find(s =>
                  s.site_id === rawSiteId || s.site_key === rawSiteId || s.sp_site_name === siteNode?.name
                );
                const effectiveSiteId = matchedSite?.site_id || rawSiteId;

                if (stackLevel === 2) {
                  const cachedDrives = host._getOrLoadSiteDrives(effectiveSiteId);
                  const drivesList = cachedDrives || (matchedSite?.drive_id ? [{ id: matchedSite.drive_id, name: matchedSite.default_library_name || 'Documents' }] : null);

                  if (!drivesList) {
                    return (
                      <div style={{ padding: 40, textAlign: 'center', color: 'var(--vdms-text-muted)' }}>
                        <div style={{ fontSize: 24, marginBottom: 8, animation: 'spin 1s linear infinite', display: 'inline-block' }}>⏳</div>
                        <div>Loading SharePoint document libraries for {siteNode?.name}...</div>
                      </div>
                    );
                  }

                  // SharePoint auto-provisions "system" libraries (Site
                  // Assets, Style Library, Form Templates, ...) on every
                  // site — a site with more features on can easily return
                  // 4-6 of these alongside the real content library, and
                  // they're never where a user's own files live. The
                  // backend now flags each drive with `is_system` and a
                  // best-effort `item_count` (immediate children of the
                  // library root); hide system libraries by default here so
                  // this grid isn't cluttered with libraries no one uses,
                  // and show an item-count badge on the rest so it's obvious
                  // at a glance which one is actually empty.
                  const systemLibraries = drivesList.filter(d => d.is_system);
                  const showSystem = host._systemLibrariesRevealed.has(effectiveSiteId);
                  const visibleDrives = (showSystem ? drivesList : drivesList.filter(d => !d.is_system))
                    .filter(drive => !textFilter || (drive.name || '').toLowerCase().includes(textFilter.trim().toLowerCase()));

                  return (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: '#0369a1', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                          Libraries in {siteNode?.name} ({visibleDrives.length})
                        </div>
                        <span style={{ fontSize: 12, color: 'var(--vdms-text-muted)' }}>Select a document library to browse folders and files</span>
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 16 }}>
                        {visibleDrives.map((drive, dIdx) => {
                          const hasCount = typeof drive.item_count === 'number';
                          const isEmpty = hasCount && drive.item_count === 0;
                          return (
                          <div
                            key={drive.id + dIdx}
                            onClick={() => {
                              pushLiveFolderNav([...folderPathStack, { id: `drive:${drive.id}`, name: drive.name }]);
                            }}
                            style={{
                              background: 'var(--vdms-surface)', borderRadius: 14, border: '1px solid var(--vdms-border)',
                              padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                              boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                            }}
                          >
                            <div style={{
                              width: 48, height: 48, borderRadius: 12, background: '#fef3c7',
                              display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 24,
                            }}>
                              📚
                            </div>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontWeight: 700, fontSize: 15, color: 'var(--vdms-text)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                                {drive.name}
                              </div>
                              <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2, display: 'flex', alignItems: 'center', gap: 6 }}>
                                <span>Document Library</span>
                                {hasCount && (
                                  <span style={{
                                    fontSize: 10, fontWeight: 700, padding: '1px 6px', borderRadius: 4,
                                    background: isEmpty ? 'var(--vdms-border)' : '#dcfce7',
                                    color: isEmpty ? 'var(--vdms-text-muted)' : '#15803d',
                                  }}>
                                    {isEmpty ? 'empty' : `${drive.item_count} item${drive.item_count === 1 ? '' : 's'}`}
                                  </span>
                                )}
                              </div>
                            </div>
                            <span style={{ color: '#0284c7', fontSize: 18, fontWeight: 700 }}>›</span>
                          </div>
                          );
                        })}
                      </div>

                      {systemLibraries.length > 0 && (
                        <button
                          onClick={() => {
                            if (showSystem) { host._systemLibrariesRevealed.delete(effectiveSiteId); }
                            else { host._systemLibrariesRevealed.add(effectiveSiteId); }
                            host.forceUpdate();
                          }}
                          style={{ alignSelf: 'flex-start', border: 'none', background: 'transparent', color: '#0284c7', cursor: 'pointer', fontSize: 12, padding: '4px 2px' }}
                        >
                          {showSystem
                            ? `Hide ${systemLibraries.length} system librar${systemLibraries.length === 1 ? 'y' : 'ies'}`
                            : `Show ${systemLibraries.length} more (SharePoint system librar${systemLibraries.length === 1 ? 'y' : 'ies'})`}
                        </button>
                      )}
                    </div>
                  );
                }

                // ── Level >= 3: Folders & Files within Drive ──
                const driveNode = folderPathStack[2];
                const rawDriveId = (driveNode?.id || '').replace(/^drive:/, '') || matchedSite?.drive_id || '';
                const currentNode = folderPathStack[folderPathStack.length - 1];
                const isDriveRoot = stackLevel === 3;
                let currentFolderId = isDriveRoot ? 'root' : (currentNode?.id || 'root');

                // Breadcrumbs rebuilt from list rows can contain synthetic sf_* IDs.
                // Resolve those IDs from the active site's live path before asking
                // Graph for the folder children.
                if (!isDriveRoot && /^sf_/i.test(currentFolderId)) {
                  const breadcrumbSegments = folderPathStack.slice(3).map(node => node.name).filter(Boolean);
                  const breadcrumbPath = breadcrumbSegments.join('/').toLowerCase();
                  const libraryRelativePath = /^(shared documents|documents)$/i.test(breadcrumbSegments[0] || '')
                    ? breadcrumbSegments.slice(1).join('/').toLowerCase()
                    : breadcrumbPath;
                  const pathCandidates = [breadcrumbPath, libraryRelativePath].filter(Boolean);
                  const liveMatch = (host.state.documentLiveFolders || []).find(folder => {
                    const livePath = (folder.path || '').replace(/\\/g, '/').replace(/^\/+|\/+$/g, '').toLowerCase();
                    return pathCandidates.some(candidate => livePath === candidate || livePath.endsWith(`/${candidate}`));
                  });
                  currentFolderId = liveMatch?.id || breadcrumbSegments.join('/');
                }

                // Cannot render without a resolved drive — show a loading/error state
                if (!rawDriveId) {
                  return (
                    <div style={{ padding: 40, textAlign: 'center', color: 'var(--vdms-text-muted)' }}>
                      <div style={{ fontSize: 24, marginBottom: 8 }}>⏳</div>
                      <div>Resolving SharePoint document library{siteNode?.name ? ` for ${siteNode.name}` : ''}...</div>
                    </div>
                  );
                }

                const folderData = host._getOrLoadSiteFolderChildren(effectiveSiteId, rawDriveId, currentFolderId);

                if (folderData.loading && folderData.items.length === 0) {
                  return (
                    <div style={{ padding: 40, textAlign: 'center', color: 'var(--vdms-text-muted)' }}>
                      <div style={{ fontSize: 24, marginBottom: 8, animation: 'spin 1s linear infinite', display: 'inline-block' }}>⏳</div>
                      <div>Loading contents of {currentNode?.name || 'folder'}...</div>
                    </div>
                  );
                }

                const baseChildFolders = folderData.items.filter(item => item.folder || (!item.file && item.name && !item.name.includes('.')));
                const childFiles = folderData.items.filter(item => item.file || (item.name && item.name.includes('.')));

                // NOTE: this used to merge host.state.rows (vessel rows, aggregated across
                // ALL sites since GET /api/vessels has no site scoping) and
                // host.state.uploadedFilesByFolder (keyed by a bare path string with no
                // site/drive affinity) into the folder list shown here. Neither source
                // carries any site/drive identity, so at a drive root it injected other
                // sites' vessel names (and, one level deeper, vessel folder-template
                // category names) as phantom folders — exactly the "different folders"
                // / "mvtest45" / "Technical & Crewing" symptom reported for the Sites
                // browser, whose fabricated paths then 404 against the real Graph
                // /children endpoint. This is a real SharePoint site/drive browser, so it
                // must show exactly what Graph returns for the current site/drive.
                // Vessel filter in Folder view: once a vessel is picked, show only
                // that vessel's folder — or the parent folders leading to it —
                // until the user is inside it, then show its contents as usual.
                const currentDrivePath = folderPathStack.slice(3).map(node => node.name).join('/');
                const insideSelectedVessel = vesselFilter === 'all' || folderPathStack.slice(3).some(node => {
                  const matched = matchFolderToVessel(node.name);
                  return !!matched && vesselNamesEqual(matched, vesselFilter);
                });
                const vesselScopedChildFolders = insideSelectedVessel
                  ? baseChildFolders
                  : baseChildFolders.filter(item => {
                    const matched = matchFolderToVessel(item.name);
                    if (matched && vesselNamesEqual(matched, vesselFilter)) return true;
                    return isPathOnSelectedVesselBranch(currentDrivePath ? `${currentDrivePath}/${item.name}` : item.name);
                  });
                // ── Group / Category filter in Folder view ──
                // A Graph folder listing has no group/category of its own, so
                // when either filter is set Folder view shows every file *under*
                // the current folder whose folder path passes
                // pathMatchesGroupCategory (partial-word rule, same as List view).
                //
                // Source of truth: the backend recursive walk of the current
                // folder (GET …/folders/{id}/recursive, names/paths only — no
                // tag lookups), fetched once per folder and cached here for 5
                // minutes. The old client-side walk (≤200 folders, level by
                // level) never reached e.g. Documents > … > SS367-MAERSK FERRATO
                // > ELECT from the library root, so "Electrical" alone found
                // nothing. While the walk is in flight, or if it fails, results
                // from folder listings already cached in the browser are shown.
                // Standing in the folder that IS the selected group (e.g. the
                // "To Be Classified" folder with the To Be Classified group
                // picked, no category) is just normal folder browsing: show
                // its sub-folders, not a flat list of every nested file.
                const openFolderName = (folderPathStack[folderPathStack.length - 1]?.name || '').trim();
                const openFolderIsSelectedGroup = docGroupLevelFilter !== 'all' && docLeafCategoryFilter === 'all' && docSubCategoryFilter === 'all' && !!openFolderName && (
                  docGroupLevelFilter === TBC_GROUP_LABEL
                    ? isToBeClassifiedName(openFolderName)
                    : (() => {
                        const fg = folderGroupsOf(openFolderName);
                        return docGroupLevelFilter.trim().toLowerCase() === 'drawings' ? (fg.drawings && !fg.manuals) : (fg.manuals && !fg.drawings);
                      })()
                );
                // A Sub-category picked on its own (Groups / Categories left at
                // "All") searches too — before, nothing happened until a Group
                // or Category was also set.
                const groupCatOnly = (docGroupLevelFilter !== 'all' || docLeafCategoryFilter !== 'all' || docSubCategoryFilter !== 'all') && !openFolderIsSelectedGroup;
                // Also drive the recursive subtree walk off a plain text search, not just
                // the Group/Category filter. Previously a text search only matched files
                // directly inside the folder currently open (vesselScopedChildFiles below);
                // a file several levels deeper that matched the search term never surfaced
                // unless the user happened to open every intermediate folder first.
                // Text search only launches the heavy recursive walk from 2 characters
                // up — a single letter matches nearly everything and just burns a
                // full-tree Graph walk (the debounce already covers fast typing).
                const groupCatActive = groupCatOnly || (textFilter || '').trim().length >= 2;
                const groupCatMatchRelPath = new Map<string, string[]>();
                let groupCatMatchedFiles: any[] = [];
                let groupCatSearching = false;
                let groupCatFailed = false;
                let groupCatTruncated = false;
                if (groupCatActive && effectiveSiteId && rawDriveId) {
                  // The backend answers this from its drive index (one in-memory
                  // pass) and can pre-select the Group's files by folder-name
                  // words, so a Group pick at the library root no longer walks
                  // the whole library folder by folder through Graph. The
                  // exact Group / Category rule still runs below on what it
                  // returns. The pre-selection is skipped when a folder
                  // already open in the breadcrumb carries the Group word.
                  const preWordsAll = groupCatPathWords(docGroupLevelFilter);
                  const stackHasGroupWord = folderPathStack.some(n => gcWords(n.name).some(w => preWordsAll.indexOf(w) >= 0));
                  // A Sub-category is an exact folder name, so its own words are
                  // a much narrower (still complete) pre-selection than the
                  // Group's: "Machinery Drawings" → files under a folder with
                  // "machinery" in its name, instead of every Drawings file
                  // (which overflowed the 8000-item cap at the library root →
                  // "No … files here / Large folder"). Group words and filler
                  // words are dropped unless nothing else is left.
                  const subPick = docSubCategoryFilter !== 'all' ? docSubCategoryFilter.trim().toLowerCase() : '';
                  const stackInSub = !!subPick && folderPathStack.some(n => (n.name || '').trim().toLowerCase() === subPick);
                  const subWordsAll = subPick && !stackInSub ? subPick.split(/[^a-z0-9]+/).filter(Boolean) : [];
                  const subWordsNarrow = subWordsAll.filter(w => ['drawing', 'drawings', 'dwg', 'dwgs', 'manual', 'manuals', 'and', 'of', 'the', 'for', 'a'].indexOf(w) < 0);
                  const subWords = subWordsNarrow.length > 0 ? subWordsNarrow : subWordsAll;
                  const preWords = subWords.length > 0
                    ? subWords
                    : (preWordsAll.length > 0 && !stackHasGroupWord ? preWordsAll : []);
                  const treeKey = `${effectiveSiteId}::${rawDriveId}::${currentFolderId}${preWords.length > 0 ? `::w:${preWords.join('|')}` : ''}`;
                  let tree = groupCatTreeCache.get(treeKey);
                  let searchSince = tree?.since;
                  let priorFailures = tree?.failures || 0;
                  if (tree && tree.status !== 'loading' && Date.now() - tree.at > GROUP_CAT_TREE_TTL_MS) {
                    groupCatTreeCache.delete(treeKey);
                    tree = undefined;
                    searchSince = undefined;
                    priorFailures = 0;
                  }
                  // A failed lookup (timeout / throttled) is retried after a
                  // short pause instead of being remembered for 5 minutes.
                  if (tree && tree.status === 'error' && Date.now() - tree.at > groupCatRetryDelayMs(tree.failures || 1)) {
                    groupCatTreeCache.delete(treeKey);
                    tree = undefined;
                  }
                  // Backend index still building: ask again when the pause is over.
                  if (tree && tree.idle && (tree.retryAt || 0) <= Date.now()) tree = undefined;
                  if (!tree) {
                    const since = searchSince || Date.now();
                    const loading: GroupCatTreeEntry = { status: 'loading', items: [], truncated: false, at: Date.now(), since, failures: priorFailures };
                    groupCatTreeCache.set(treeKey, loading);
                    tree = loading;
                    const useIndex = Date.now() - since < GROUP_CAT_INDEX_WAIT_MS;
                    const url = `${host._base()}/api/sites/${encodeURIComponent(effectiveSiteId)}/drives/${encodeURIComponent(rawDriveId)}` +
                      `/folders/${encodeURIComponent(currentFolderId)}/recursive?include_tags=false&max_items=8000` +
                      (preWords.length > 0 ? `&path_words=${encodeURIComponent(preWords.join(','))}` : '') +
                      (useIndex ? '' : '&use_index=false');
                    host._fetchJson(url)
                      .then((data: any) => {
                        if (data?.building) {
                          const waitMs = Math.max(3000, Math.min(10000, (Number(data.retry_after) || 0) * 1000));
                          groupCatTreeCache.set(treeKey, {
                            status: 'loading', items: [], truncated: false, at: Date.now(), since, idle: true, retryAt: Date.now() + waitMs, failures: priorFailures,
                          });
                          window.setTimeout(() => host._scheduleForceUpdate(), waitMs);
                          return;
                        }
                        groupCatTreeCache.set(treeKey, {
                          status: 'done',
                          items: Array.isArray(data?.folders) ? data.folders : [],
                          truncated: !!data?.truncated,
                          at: Date.now(),
                        });
                      })
                      .catch(() => {
                        const failures = priorFailures + 1;
                        groupCatTreeCache.set(treeKey, { status: 'error', items: [], truncated: false, at: Date.now(), since, failures });
                        // Nothing else re-renders the page at the end of the
                        // pause, so ask for it: the retry starts on that render.
                        window.setTimeout(() => host._scheduleForceUpdate(), groupCatRetryDelayMs(failures) + 200);
                      })
                      .then(() => host._scheduleForceUpdate());
                  }

                  const baseNames = folderPathStack.slice(3).map(n => n.name);
                  const seenIds = new Set<string>();
                  const considerFile = (file: any, relFolderNames: string[]): void => {
                    const id = String(file.id || '');
                    if (id && seenIds.has(id)) return;
                    const fullNames = [...baseNames, ...relFolderNames];
                    if (!pathMatchesGroupCategory(fullNames.join(' > '), docGroupLevelFilter, docLeafCategoryFilter)) return;
                    if (docSubCategoryFilter !== 'all' && !pathMatchesSubCategory(fullNames, docSubCategoryFilter)) return;
                    if (!insideSelectedVessel) {
                      const inVessel = fullNames.some(seg => {
                        const matched = matchFolderToVessel(seg);
                        return !!matched && vesselNamesEqual(matched, vesselFilter);
                      });
                      if (!inVessel) return;
                    }
                    if (textFilter && !matchesSearchTokens(textFilter, file.name, ...relFolderNames)) return;
                    if (id) seenIds.add(id);
                    groupCatMatchedFiles.push(file);
                    if (id) groupCatMatchRelPath.set(id, relFolderNames);
                  };

                  if (tree.status === 'done') {
                    // Recursive walk: `path` is relative to the current folder
                    // and ends with the item's own name.
                    tree.items.forEach((item: any) => {
                      if (!item || item.is_folder || !item.name) return;
                      const parts = String(item.path || item.name).split('/').filter(Boolean);
                      considerFile({
                        id: item.id,
                        name: item.name,
                        size: item.size,
                        lastModifiedDateTime: item.last_modified_date_time || item.created_date_time,
                        web_url: item.web_url,
                      }, parts.slice(0, -1));
                    });
                    groupCatTruncated = tree.truncated;
                  } else {
                    // A failed lookup is retried automatically, so it still reads
                    // as "searching" — never as "No files here".
                    groupCatSearching = tree.status === 'loading' || tree.status === 'error';
                    groupCatFailed = tree.status === 'error';
                    // Interim / fallback: folder listings already in the browser
                    // cache (folders the user has opened or that were prefetched).
                    const queue: { id: string; names: string[] }[] = [{ id: currentFolderId, names: [] }];
                    let visited = 0;
                    while (queue.length > 0 && visited < 2000) {
                      const node = queue.shift()!;
                      visited++;
                      const entry = host._siteFolderItemsCache.get(`${effectiveSiteId}::${rawDriveId}::${node.id}`);
                      if (!entry || entry.loading) {
                        if (tree.status === 'error') groupCatTruncated = true;
                        continue;
                      }
                      (entry.items || []).forEach((it: any) => {
                        if (!it?.name) return;
                        // Graph facets first; the name-has-a-dot guess only when
                        // neither facet is present ("NK CERT.&Shipyard Cert" is
                        // a folder despite the dot).
                        const isFolderItem = it.folder ? true : (it.file ? false : !String(it.name).includes('.'));
                        if (isFolderItem) {
                          if (it.id) queue.push({ id: it.id, names: [...node.names, it.name] });
                          return;
                        }
                        considerFile(it, node.names);
                      });
                    }
                  }

                  // Merge in results from the backend's full-text document
                  // search (VesselEmail.tsx _triggerGlobalSearch → GET
                  // /api/dashboard/documents, kicked off debounced by every
                  // keystroke in the search box above and stashed in
                  // host.state.rows tagged groupKey ending ":search"). That
                  // backend search scans every document the site has
                  // indexed (name + full folder path + vessel), unlike the
                  // walk just above, which is capped at max_items (currently
                  // 8000 nodes) and 8 folder levels deep — so on a large
                  // tree (e.g. NKSDocMan's Shared Documents root, with every
                  // vessel under it) a real match past that cap was silently
                  // reported as "no files here" even though it existed. This
                  // is what makes combined queries reliable — vessel name
                  // alone, vessel + main folder/category, vessel + file
                  // name, vessel + sub-category + file name — regardless of
                  // where in the tree, or how deep, the match actually
                  // lives.
                  if (textFilter) {
                    (host.state.rows || []).forEach(r => {
                      if (!(r.groupKey || '').endsWith(':search')) return;
                      if (!r.fileId || seenIds.has(r.fileId)) return;
                      if (!insideSelectedVessel) {
                        const rowVessel = getListViewLabels(r).vessel;
                        if (!rowVessel || !vesselNamesEqual(rowVessel, vesselFilter)) return;
                      }
                      if (!matchesSearchTokens(textFilter, r.fileName, r.vesselName, r.subFolderPath)) return;
                      seenIds.add(r.fileId);
                      // subFolderPath is "SharePoint Sites > <site> > <library> > …";
                      // drop that 3-segment prefix so the breadcrumb shown here
                      // matches the "relative to the current folder" shape the
                      // Graph-walk results above use.
                      const relFolderNamesForRow = (r.subFolderPath || '')
                        .split('>')
                        .map(s => s.trim())
                        .filter(Boolean)
                        .slice(3);
                      groupCatMatchedFiles.push({
                        id: r.fileId,
                        name: r.fileName,
                        size: r.fileSize,
                        lastModifiedDateTime: r.fileUploadedAt,
                        web_url: null,
                      });
                      groupCatMatchRelPath.set(r.fileId, relFolderNamesForRow);
                    });
                  }

                  // Group / Category picked with no text: the recursive walk above
                  // is capped (8000 items, 8 levels), so on a big library it can
                  // miss whole branches. Also take every matching file already in
                  // the fleet-wide row set (same source List view uses), limited to
                  // what sits under the folder currently open.
                  if (groupCatOnly) {
                    (scopeRows || []).forEach((r: any) => {
                      if (!r || !r.fileId || seenIds.has(r.fileId)) return;
                      const segsAll = (r.subFolderPath || '').split('>').map((x: string) => x.trim()).filter(Boolean);
                      const rel = segsAll.slice(3);
                      if (baseNames.length > rel.length) return;
                      for (let i = 0; i < baseNames.length; i++) {
                        if (rel[i].toLowerCase() !== baseNames[i].toLowerCase()) return;
                      }
                      const relBelow = rel.slice(baseNames.length);
                      if (!pathMatchesGroupCategory(rel.join(' > '), docGroupLevelFilter, docLeafCategoryFilter)) return;
                      if (docSubCategoryFilter !== 'all' && !pathMatchesSubCategory(rel, docSubCategoryFilter)) return;
                      if (!insideSelectedVessel) {
                        const rowVessel = getListViewLabels(r).vessel;
                        if (!rowVessel || !vesselNamesEqual(rowVessel, vesselFilter)) return;
                      }
                      seenIds.add(r.fileId);
                      groupCatMatchedFiles.push({
                        id: r.fileId,
                        name: r.fileName,
                        size: r.fileSize,
                        lastModifiedDateTime: r.fileUploadedAt,
                        web_url: r.webUrl || r.fileUrl,
                      });
                      groupCatMatchRelPath.set(r.fileId, relBelow);
                    });
                  }
                }

                const folderServerCounts2 = folderData.error ? null : host._getOrLoadFolderCounts(effectiveSiteId, rawDriveId, currentFolderId);
                const childFolders = groupCatActive
                  ? []
                  : (textFilter
                    ? vesselScopedChildFolders.filter(item => matchesSearchTokens(textFilter, item.name))
                    : vesselScopedChildFolders);

                return (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
                    {/* Folders Section */}
                    {childFolders.length > 0 && (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                        <div style={{ fontSize: 12, fontWeight: 800, color: '#0369a1', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                          Folders ({childFolders.length})
                        </div>
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 }}>
                          {childFolders.map((sf, idx) => {
                            const sfInTree2 = !!sf.id && liveTreeIndex.ids.has(sf.id);
                            const sfCountsPending2 = !folderServerCounts2 && host._folderCountsLoading(effectiveSiteId, rawDriveId, currentFolderId);
                            const sfChildData2 = sf.id && !sfInTree2 && !folderServerCounts2 && !sfCountsPending2 ? host._getOrLoadSiteFolderChildren(effectiveSiteId, rawDriveId, sf.id) : null;
                            const sfLoading2 = sfChildData2 ? (sfChildData2.loading && sfChildData2.items.length === 0) : (sfCountsPending2 && !sfInTree2);
                            const sfCounts2 = sfLoading2 ? null : getFolderTileCounts(sf.id, sfChildData2 ? sfChildData2.items : null);
                            const sfFolderCount2 = sfCounts2 ? sfCounts2.folders : null;
                            const sfFileCount2  = sfCounts2 ? sfCounts2.files : null;
                            const sfTotal2 = sf.folder?.childCount ?? 0;
                            const sfServerCounts2 = folderServerCounts2?.counts[sf.id] || null;
                            return (
                              <div
                                key={sf.id || sf.name + idx}
                                onClick={() => {
                                  pushLiveFolderNav([...folderPathStack, { id: sf.id, name: sf.name }]);
                                }}
                                style={{
                                  background: 'var(--vdms-surface)', borderRadius: 14, border: '1px solid var(--vdms-border)', padding: 18,
                                  display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                                  boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                                }}
                              >
                                <div style={{
                                  width: 44, height: 44, borderRadius: 10, background: '#e0f2fe',
                                  display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7',
                                }}>
                                  📁
                                </div>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--vdms-text)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                                    {sf.name}
                                  </div>
                                  <div title={sfCounts2 ? `${sfCounts2.folders} sub-folder(s) and ${sfCounts2.files} file(s) in total` : undefined} style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 3, display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                                    {sfServerCounts2 ? renderServerCountPill(sfServerCounts2, sf.name) : sfLoading2 ? (
                                      <span style={{ color: 'var(--vdms-text-faint)', fontSize: 10 }}>{sfTotal2 > 0 ? `${sfTotal2} items` : '···'}</span>
                                    ) : (
                                      <>
                                        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 3, background: (sfFolderCount2 ?? 0) > 0 ? '#e0f2fe' : 'var(--vdms-border-soft)', color: (sfFolderCount2 ?? 0) > 0 ? '#0369a1' : 'var(--vdms-text-faint)', borderRadius: 20, padding: '1px 8px', fontSize: 10, fontWeight: 700, lineHeight: '16px' }}>📁 {sfFolderCount2 ?? 0}</span>
                                        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 3, background: (sfFileCount2 ?? 0) > 0 ? '#dcfce7' : 'var(--vdms-border-soft)', color: (sfFileCount2 ?? 0) > 0 ? '#15803d' : 'var(--vdms-text-faint)', borderRadius: 20, padding: '1px 8px', fontSize: 10, fontWeight: 700, lineHeight: '16px' }}>📄 {sfFileCount2 ?? 0}</span>
                                      </>
                                    )}
                                  </div>
                                </div>
                                <span style={{ color: 'var(--vdms-text-faint)', fontSize: 16 }}>›</span>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    {/* Files Section */}
                    {(() => {
                      // NOTE: this used to also merge host.state.uploadedFilesByFolder,
                      // matched by a bare path/name string with no site or drive
                      // affinity — the same cross-site contamination risk as the
                      // folders merge above. This is a real SharePoint site/drive
                      // browser, so files shown here must come only from Graph
                      // (childFiles), not from a non-site-scoped local cache.
                      // Loose files outside the selected vessel's folder aren't that
                      // vessel's documents.
                      const vesselScopedChildFiles = insideSelectedVessel ? childFiles : [];
                      const textFilteredChildFiles = textFilter
                        ? vesselScopedChildFiles.filter(item => matchesSearchTokens(textFilter, item.name))
                        : vesselScopedChildFiles;
                      // Group / Category active → the subtree search results
                      // computed above (path-based, same rule as List view);
                      // otherwise this folder's own files.
                      const displayChildFiles = groupCatActive ? groupCatMatchedFiles : textFilteredChildFiles;
                      const subPickCtx = subCategoryPickCtx();
                      const groupCatLabel = [
                        docGroupLevelFilter !== 'all' ? docGroupLevelFilter : (subPickCtx ? subPickCtx.group : null),
                        docLeafCategoryFilter !== 'all' ? docLeafCategoryFilter : (subPickCtx ? subPickCtx.category : null),
                        docSubCategoryFilter !== 'all' ? docSubCategoryFilter : null,
                      ].filter(Boolean).join(' › ');

                      // Folder-view pagination: 10 files per page, unlimited pages.
                      // Page resets to 1 whenever the folder or any filter changes
                      // (tracked via a signature stored next to the page number).
                      const FOLDER_PAGE_SIZE = 10;
                      const fvState = host.state as any;
                      const fvSig = [effectiveSiteId, currentFolderId, textFilter, vesselFilter, docGroupLevelFilter, docLeafCategoryFilter, docCategoryFilter, docListSort, fvState.attachmentFilter].join('|');
                      const fvTotalPages = Math.max(1, Math.ceil(displayChildFiles.length / FOLDER_PAGE_SIZE));
                      const fvPage = Math.min(fvState.docFolderPageSig === fvSig ? (fvState.docFolderPage || 0) : 0, fvTotalPages - 1);
                      const pagedChildFiles = displayChildFiles.slice(fvPage * FOLDER_PAGE_SIZE, (fvPage + 1) * FOLDER_PAGE_SIZE);
                      const goFvPage = (n: number): void => {
                        host.setState({ docFolderPage: Math.max(0, Math.min(fvTotalPages - 1, n)), docFolderPageSig: fvSig } as any);
                      };
                      const fvPageNums: (number | '…')[] = [];
                      {
                        const want = new Set<number>([0, fvTotalPages - 1]);
                        for (let k = fvPage - 2; k <= fvPage + 2; k++) if (k >= 0 && k < fvTotalPages) want.add(k);
                        const sorted = Array.from(want).sort((a, b) => a - b);
                        sorted.forEach((n, i) => {
                          if (i > 0 && n - sorted[i - 1] > 1) fvPageNums.push('…');
                          fvPageNums.push(n);
                        });
                      }
                      const fvBtn = (disabled: boolean, active = false): React.CSSProperties => ({
                        border: active ? 'none' : '1px solid var(--vdms-border)',
                        background: active ? '#0078d4' : 'var(--vdms-surface)',
                        color: active ? '#fff' : 'var(--vdms-text)',
                        borderRadius: 4, padding: '3px 9px', fontSize: 11, fontWeight: active ? 700 : 400,
                        cursor: disabled ? 'not-allowed' : 'pointer', opacity: disabled ? 0.4 : 1,
                      });

                      return (
                        <>
                          {displayChildFiles.length > 0 && (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                              <div style={{ fontSize: 12, fontWeight: 800, color: '#15803d', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                                <span>
                                  {groupCatOnly ? `${groupCatLabel} files in this folder (${displayChildFiles.length})` : `Files (${displayChildFiles.length})`}
                                </span>
                                {groupCatActive && groupCatSearching && (
                                  <span style={{ fontSize: 11, fontWeight: 600, color: '#1d4ed8', textTransform: 'none', letterSpacing: 0 }}>Searching sub-folders…</span>
                                )}
                                {groupCatActive && !groupCatSearching && groupCatTruncated && (
                                  <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--vdms-text-muted)', textTransform: 'none', letterSpacing: 0 }}>Large folder — open a sub-folder to search it fully</span>
                                )}
                              </div>
                              <div style={{ background: 'var(--vdms-surface)', borderRadius: 12, border: '1px solid var(--vdms-border)', overflow: 'hidden' }}>
                                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                                  <thead>
                                    <tr style={{ background: 'var(--vdms-surface-alt)', borderBottom: '1px solid var(--vdms-border)', color: 'var(--vdms-text-muted)', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
                                      <th style={{ padding: '10px 16px' }}>FILE NAME</th>
                                      <th style={{ padding: '10px 16px' }}>VESSEL</th>
                                      <th style={{ padding: '10px 16px' }}>FOLDER PATH</th>
                                      <th style={{ padding: '10px 16px' }}>SIZE</th>
                                      <th style={{ padding: '10px 16px' }}>DATE MODIFIED</th>
                                      <th style={{ padding: '10px 16px', textAlign: 'right' }}>ACTION</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {pagedChildFiles.map((file, idx) => {
                                      const fileSize = typeof file.size === 'number'
                                        ? (file.size > 1024 * 1024 ? `${(file.size / (1024 * 1024)).toFixed(1)} MB` : `${(file.size / 1024).toFixed(1)} KB`)
                                        : (file.size || '—');
                                      const fileDate = (() => {
                                        if (!file.lastModifiedDateTime) return '—';
                                        const parsed = new Date(file.lastModifiedDateTime);
                                        return isNaN(parsed.getTime()) ? String(file.lastModifiedDateTime) : parsed.toLocaleString();
                                      })();
                                      const fileUrl = file.web_url || file.webUrl || file.download_url || '';
                                      // Vessel for this file: prefer a vessel name found in the
                                      // current folder path itself (a live folder/sub-folder that
                                      // IS a recognized vessel), then fall back to the matching DB
                                      // row's own vessel metadata; "Not Listed" (with the same X
                                      // mark used in List view) when neither resolves one.
                                      // Sub-folder path below the current folder (only set
                                      // for Group/Category subtree results).
                                      const relFolderNames: string[] = (file.id && groupCatMatchRelPath.get(file.id)) || [];
                                      const fileFolderNames = [...folderPathStack.map(n => n.name), ...relFolderNames];
                                      const fileBreadcrumb = fileFolderNames.join(' > ');
                                      // matchFolderToVessel (not an exact-name lookup) so a
                                      // misspelt vessel folder such as "022024 Duchess Emeralad"
                                      // still resolves to its vessel.
                                      const folderPathVesselMatch = fileFolderNames
                                        .map(name => matchFolderToVessel(name))
                                        .filter(Boolean)
                                        .pop() || null;
                                      const matchRow = (host.state.rows || []).find(r => r.fileId === file.id);
                                      const rowVessel = matchRow ? getListViewLabels(matchRow).vessel : null;
                                      const effectiveVessel = folderPathVesselMatch ||
                                        (rowVessel && rowVessel !== 'Not Listed' && rowVessel !== 'Vessel name not listed' ? rowVessel : null);

                                      return (
                                        <tr key={file.id || file.name + idx} style={{ borderBottom: '1px solid var(--vdms-border-soft)' }}>
                                          <td style={{ padding: '12px 16px', fontWeight: 600, color: 'var(--vdms-text)', display: 'flex', alignItems: 'center', gap: 10 }}>
                                            <span style={{ fontSize: 18 }}>📄</span>
                                            <span style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
                                              <span
                                                onClick={() => {
                                                  if (fileUrl) {
                                                    window.open(fileUrl, '_blank');
                                                  } else {
                                                    void host._openDocumentFile(file.id, file.name, fileBreadcrumb);
                                                  }
                                                }}
                                                style={{ cursor: 'pointer', color: '#0284c7', textDecoration: 'underline' }}
                                                title={`Click to view/download ${file.name}`}
                                              >
                                                {file.name}
                                              </span>
                                              {relFolderNames.length > 0 && (
                                                <span style={{ fontSize: 11, fontWeight: 400, color: 'var(--vdms-text-muted)', marginTop: 2 }} title={fileBreadcrumb}>
                                                  {relFolderNames.join(' › ')}
                                                </span>
                                              )}
                                            </span>
                                          </td>
                                          <td style={{ padding: '12px 16px' }}>
                                            {effectiveVessel ? (
                                              <span style={{ color: 'var(--vdms-text)', fontWeight: 600 }}>{effectiveVessel}</span>
                                            ) : (
                                              <span
                                                style={{ display: 'inline-flex', alignItems: 'center', gap: 4, background: '#ffedd5', color: '#c2410c', borderRadius: 6, padding: '2px 7px', fontSize: 11, fontWeight: 700 }}
                                                title="No vessel associated with this file"
                                              >
                                                <span aria-hidden="true" style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: 14, height: 14, borderRadius: '50%', background: '#c2410c', color: '#fff', fontSize: 9, fontWeight: 800, lineHeight: 1 }}>✕</span>
                                                Not Listed
                                              </span>
                                            )}
                                          </td>
                                          <td style={{ padding: '12px 16px', color: 'var(--vdms-text-muted)', fontSize: 12 }} title={fileBreadcrumb || undefined}>
                                            {fileBreadcrumb || '—'}
                                          </td>
                                          <td style={{ padding: '12px 16px', color: 'var(--vdms-text-muted)' }}>{fileSize}</td>
                                          <td style={{ padding: '12px 16px', color: 'var(--vdms-text-muted)' }}>{fileDate}</td>
                                          <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                                            <button
                                              type="button"
                                              onClick={() => {
                                                if (fileUrl) {
                                                  window.open(fileUrl, '_blank');
                                                } else {
                                                  void host._openDocumentFile(file.id, file.name, fileBreadcrumb);
                                                }
                                              }}
                                              style={{
                                                background: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: 6,
                                                padding: '4px 10px', fontSize: 12, fontWeight: 600, color: '#0284c7', cursor: 'pointer',
                                              }}
                                            >
                                              Download / Open
                                            </button>
                                          </td>
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>
                                <div style={{ padding: '10px 14px', color: 'var(--vdms-text-muted)', fontSize: 11, display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8, borderTop: '1px solid var(--vdms-border-soft)', background: 'var(--vdms-surface-alt)' }}>
                                  <span>
                                    Showing {fvPage * FOLDER_PAGE_SIZE + 1}–{Math.min((fvPage + 1) * FOLDER_PAGE_SIZE, displayChildFiles.length)} of {displayChildFiles.length} files
                                  </span>
                                  <div style={{ display: 'flex', gap: 4, alignItems: 'center', flexWrap: 'wrap' }}>
                                    <button type="button" onClick={() => goFvPage(fvPage - 1)} disabled={fvPage === 0} style={fvBtn(fvPage === 0)}>← Prev</button>
                                    {fvPageNums.map((n, i) => n === '…' ? (
                                      <span key={'gap' + i} style={{ padding: '0 4px' }}>…</span>
                                    ) : (
                                      <button key={n} type="button" onClick={() => goFvPage(n)} style={fvBtn(false, n === fvPage)}>{n + 1}</button>
                                    ))}
                                    <button type="button" onClick={() => goFvPage(fvPage + 1)} disabled={fvPage >= fvTotalPages - 1} style={fvBtn(fvPage >= fvTotalPages - 1)}>Next →</button>
                                  </div>
                                </div>
                              </div>
                            </div>
                          )}

                          {childFolders.length === 0 && displayChildFiles.length === 0 && (
                            folderData.error ? (
                              <div style={{ background: 'var(--vdms-surface)', borderRadius: 14, border: '1px dashed #fca5a5', padding: 48, textAlign: 'center', color: 'var(--vdms-text-faint)' }}>
                                <div style={{ fontSize: 32, marginBottom: 8 }}>⚠️</div>
                                {(() => {
                                  // Graph quota cooldown: the listing re-loads by itself when it ends.
                                  const waitSec = folderData.throttled ? Math.max(0, Math.ceil(((folderData.retryAt || host._graphThrottledUntil) - Date.now()) / 1000)) : 0;
                                  return folderData.throttled ? (
                                    <>
                                      <div style={{ fontWeight: 600, color: '#b45309', fontSize: 15 }}>SharePoint is busy</div>
                                      <div style={{ fontSize: 13, marginTop: 4 }}>
                                        Microsoft limited requests for a moment. This folder will load again automatically{waitSec > 0 ? ` in about ${waitSec}s` : ''}.
                                      </div>
                                    </>
                                  ) : (
                                    <>
                                      <div style={{ fontWeight: 600, color: '#b91c1c', fontSize: 15 }}>Couldn't load this folder</div>
                                      <div style={{ fontSize: 13, marginTop: 4 }}>The request to SharePoint failed — this may not actually be empty. Try again.</div>
                                    </>
                                  );
                                })()}
                                <button
                                  type="button"
                                  onClick={() => void host._refreshSiteFolder(effectiveSiteId, rawDriveId, currentFolderId)}
                                  style={{ marginTop: 12, border: '1px solid #fca5a5', background: '#fef2f2', color: '#b91c1c', borderRadius: 8, padding: '6px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}
                                >
                                  ↻ Retry
                                </button>
                              </div>
                            ) : (
                              <div style={{ background: 'var(--vdms-surface)', borderRadius: 14, border: '1px dashed var(--vdms-border)', padding: 48, textAlign: 'center', color: 'var(--vdms-text-faint)' }}>
                                <div style={{ fontSize: 32, marginBottom: 8 }}>📂</div>
                                {groupCatActive ? (
                                  groupCatSearching ? (
                                    <>
                                      <div style={{ fontWeight: 600, color: 'var(--vdms-text-secondary)', fontSize: 15 }}>Searching for {groupCatLabel} files…</div>
                                      <div style={{ fontSize: 13, marginTop: 4 }}>
                                        {groupCatFailed
                                          ? 'That took longer than expected. Trying again automatically…'
                                          : `Looking through the sub-folders of ${currentNode?.name || 'this folder'}.`}
                                      </div>
                                    </>
                                  ) : (
                                    <>
                                      <div style={{ fontWeight: 600, color: 'var(--vdms-text-secondary)', fontSize: 15 }}>No {groupCatLabel} files here</div>
                                      <div style={{ fontSize: 13, marginTop: 4 }}>
                                        No folder under {currentNode?.name || 'this folder'} matches
                                        {docGroupLevelFilter !== 'all' ? ` ${docGroupLevelFilter}` : ''}
                                        {docGroupLevelFilter !== 'all' && docLeafCategoryFilter !== 'all' ? ' and' : ''}
                                        {docLeafCategoryFilter !== 'all' ? ` ${docLeafCategoryFilter}` : ''}
                                        {docSubCategoryFilter !== 'all' ? `${docGroupLevelFilter !== 'all' || docLeafCategoryFilter !== 'all' ? ' and' : ''} ${docSubCategoryFilter}` : ''}
                                        {vesselFilter !== 'all' ? ` for ${vesselFilter}` : ''}, or those folders have no files.
                                        {groupCatTruncated ? ' Large folder — open a sub-folder to search it fully.' : ''}
                                      </div>
                                    </>
                                  )
                                ) : insideSelectedVessel ? (
                                  <>
                                    <div style={{ fontWeight: 600, color: 'var(--vdms-text-secondary)', fontSize: 15 }}>This folder is empty</div>
                                    <div style={{ fontSize: 13, marginTop: 4 }}>No files or subfolders found in this directory. You can upload files or folders using the buttons above.</div>
                                  </>
                                ) : (
                                  <>
                                    <div style={{ fontWeight: 600, color: 'var(--vdms-text-secondary)', fontSize: 15 }}>No folder for {vesselFilter} here</div>
                                    <div style={{ fontSize: 13, marginTop: 4 }}>Nothing in this folder belongs to the selected vessel. Choose "All vessels" to see everything.</div>
                                  </>
                                )}
                              </div>
                            )
                          )}
                        </>
                      );
                    })()}
                  </div>
                );
              })()
            ) : atSharedDocsRoot ? (
              sharedDocsSite ? (
                renderLibraryBrowser(
                  sharedDocsSite.site_id,
                  sharedDocsSite.drive_id || '',
                  sharedDocsSite.default_library_name || 'Shared Documents',
                  `${sharedDocsSite.sp_site_name || sharedDocsSite.site_key} · SharePoint Library`,
                  'Shared Documents'
                )
              ) : (
                <div style={{ padding: 40, textAlign: 'center', color: 'var(--vdms-text-muted)' }}>
                  Loading configured document site...
                </div>
              )
            ) : atDocsRoot ? (
              docsSite ? (
                renderLibraryBrowser(
                  docsSite.site_id,
                  docsSite.drive_id || '',
                  docsSite.default_library_name || 'Documents',
                  `${docsSite.sp_site_name || docsSite.site_key} · Document Library`,
                  'Documents'
                )
              ) : (
                <div style={{ padding: 40, textAlign: 'center', color: 'var(--vdms-text-muted)' }}>
                  Loading configured document site...
                </div>
              )
            ) : atDepartmentVesselList && !showSelectedVesselCategories ? (

              /* Level 1: Inside a Main Department — Show Vessels + Common for all ships */
              <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
                {(() => {
                  const commonDisplayName = docMainFolder === 'Insurance'
                    ? 'Common (Not Ship Specific)'
                    : (docMainFolder === 'Commercial & Chartering'
                      ? 'Common Agreements (Not Ship Specific)'
                      : 'Common for all ships');
                  return (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                      <div style={{ fontSize: 12, fontWeight: 800, color: '#92400e', textTransform: 'uppercase', letterSpacing: '0.06em' }}>
                        Common Folder
                      </div>
                      <div style={{ display: 'grid', gridTemplateColumns: `repeat(${vesselColumns}, minmax(0, 1fr))`, gap: 16 }}>
                        <div
                          onClick={() => {
                            host._pushFolderNav([...folderPathStack, { id: 'common', name: commonDisplayName }], docMainFolder);
                            host.setState({ vesselFilter: 'all', docScopeType: 'common' });
                          }}
                          style={{
                            background: 'var(--vdms-surface)', borderRadius: 14, border: '1px solid var(--vdms-border)', padding: 18,
                            display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                            boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                          }}
                        >
                          <div style={{ width: 44, height: 44, borderRadius: 10, background: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>📁</div>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--vdms-text)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{commonDisplayName}</div>
                            <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2 }}>Shared {docMainFolder} documents</div>
                          </div>
                          <span style={{ color: 'var(--vdms-text-faint)', fontSize: 16 }}>›</span>
                        </div>
                      </div>
                    </div>
                  );
                })()}

                <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                  <div style={{ fontSize: 12, fontWeight: 800, color: '#0369a1', textTransform: 'uppercase', letterSpacing: '0.06em' }}>
                    Vessel Folders
                  </div>
                  <div style={{ display: 'grid', gridTemplateColumns: `repeat(${vesselColumns}, minmax(0, 1fr))`, gap: 16 }}>
                    {displayVessels
                    .filter(v => vesselFilter === 'all' || v.name.trim().toLowerCase() === vesselFilter.trim().toLowerCase())
                    .filter(v => catFilter === 'all' || scopeRows.some(row =>
                      row.vesselName.trim().toLowerCase() === v.name.trim().toLowerCase() &&
                      (!docMainFolder || row.group.trim().toLowerCase() === docMainFolder.trim().toLowerCase()) &&
                      row.category.trim().toLowerCase() === catFilter.trim().toLowerCase()
                    ))
                    .filter(v => !textFilter || v.name.toLowerCase().includes(textFilter.trim().toLowerCase()))
                    .map(v => (
                    <div
                      key={v.id}
                      onClick={() => {
                        host._pushFolderNav([...folderPathStack, { id: v.id, name: v.name }], docMainFolder);
                        host.setState({ vesselFilter: v.name, docScopeType: 'vessels' });
                        host._loadFilesForVessel(v.name).catch(() => undefined);
                      }}
                      style={{ background: 'var(--vdms-surface)', borderRadius: 14, border: '1px solid var(--vdms-border)', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' }}
                    >
                      <div style={{ width: 44, height: 44, borderRadius: 10, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7' }}>🚢</div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--vdms-text)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{v.name}</div>
                        <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2 }}>Vessel</div>
                      </div>
                      <span style={{ color: 'var(--vdms-text-faint)', fontSize: 16 }}>›</span>
                    </div>
                    ))}

                    {documentVesselCount < vessels.length && (
                      <div
                        onClick={() => {
                          if (!documentVesselsLoadingMore) {
                            void host._loadMoreDocumentVessels();
                          }
                        }}
                        style={{
                          background: documentVesselsLoadingMore ? 'var(--vdms-surface-alt)' : '#f0f9ff',
                          borderRadius: 14,
                          border: '2px dashed #0284c7',
                          padding: 18,
                          display: 'flex',
                          alignItems: 'center',
                          gap: 14,
                          cursor: documentVesselsLoadingMore ? 'wait' : 'pointer',
                          boxShadow: '0 1px 3px rgba(0,0,0,0.03)',
                          transition: 'all 0.15s ease',
                        }}
                      >
                        <div style={{
                          width: 44, height: 44, borderRadius: 10,
                          background: '#0284c7', color: '#fff',
                          display: 'flex', alignItems: 'center', justifyContent: 'center',
                          fontSize: 20, fontWeight: 700, flexShrink: 0,
                        }}>
                          {documentVesselsLoadingMore ? '⏳' : '+'}
                        </div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 700, fontSize: 14, color: '#0284c7', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                            {documentVesselsLoadingMore ? 'Loading vessels...' : `More vessels (+${Math.min(8, vessels.length - documentVesselCount)})`}
                          </div>
                          <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2 }}>
                            {documentVesselsLoadingMore ? 'Please wait...' : `Load next batch (${documentVesselCount} of ${vessels.length} shown)`}
                          </div>
                        </div>
                        <span style={{ color: '#0284c7', fontSize: 18, fontWeight: 700 }}>›</span>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            ) : subfolderNames.length > 0 ? (
              /* Subfolders Grid */
              <div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 }}>
                  {subfolderNames.filter(sf => {
                    const selectedCategoryIsCurrentFolder = catFilter !== 'all' && currentFolderName &&
                      currentFolderName.trim().toLowerCase() === catFilter.trim().toLowerCase();
                    if (catFilter !== 'all' && !selectedCategoryIsCurrentFolder && sf.trim().toLowerCase() !== catFilter.trim().toLowerCase()) return false;
                    if (textFilter && !sf.toLowerCase().includes(textFilter.trim().toLowerCase())) return false;
                    return true;
                  }).map((sfName, idx) => {
                    const stackNames = folderPathStack.map(n => n.name).filter(n => n !== 'Vessels' && n !== 'Specific Vessels');
                    const fullBreadcrumbWithSf = [...stackNames, sfName].join(' > ');
                    const folderDebugLiveId = host._getLiveSharePointFolderId(fullBreadcrumbWithSf) || 'unresolved';
                    return (
                    <div
                      key={sfName + idx}
                      onClick={() => {
                        const vesselName = currentVesselNameFromStack || (vesselFilter !== 'all' ? vesselFilter : (folderPathStack.length > 0 ? folderPathStack[0].name : null));
                        const candidateSubPath = vesselName ? `${vesselName} > ${docMainFolder || ''} > ${sfName}` : '';
                        const liveFolderId = host._getLiveSharePointFolderId(fullBreadcrumbWithSf) ||
                          (candidateSubPath ? host._getLiveSharePointFolderId(candidateSubPath) : null);
                        const matchingSubRow = vesselName ? host.state.rows.find(r => {
                          if (r.vesselName.toLowerCase() !== vesselName.toLowerCase()) return false;
                          if (docMainFolder && r.group && r.group.toLowerCase() !== docMainFolder.toLowerCase()) return false;
                          const tail = getFolderTailSegments(r.subFolderPath, r.vesselName, r.group);
                          return (tail.length > 0 && tail[tail.length - 1].toLowerCase() === sfName.toLowerCase()) ||
                            (r.subCategory || '').toLowerCase() === sfName.toLowerCase() ||
                            (r.category || '').toLowerCase() === sfName.toLowerCase();
                        }) : null;
                        const realFolderId = liveFolderId || (matchingSubRow?.uploadFolderId && !matchingSubRow.uploadFolderId.includes('/') ? matchingSubRow.uploadFolderId : null) || fullBreadcrumbWithSf;
                        const newStack = [...folderPathStack, { id: realFolderId, name: sfName }];
                        host._pushFolderNav(newStack, docMainFolder);
                        // Immediately refresh files for this folder from live SPO
                        const subGroupKey = matchingSubRow?.groupKey || fullBreadcrumbWithSf;
                        void host._refreshFolderFiles(realFolderId, subGroupKey, true).catch(() => undefined);
                      }}
                      style={{
                        background: 'var(--vdms-surface)', borderRadius: 14, border: '1px solid var(--vdms-border)', padding: 18,
                        display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                        boxShadow: '0 1px 3px rgba(0,0,0,0.05)', transition: 'transform 0.15s, box-shadow 0.15s',
                      }}
                    >
                      <div style={{
                        width: 44, height: 44, borderRadius: 10, background: '#e0f2fe',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7',
                      }}>
                        📁
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--vdms-text)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{sfName}</div>
                        <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 3, display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                          <span>{isAtCategoryLevel ? 'Document Section' : (stackLevel === 3 ? 'Category' : 'Sub-Category')}</span>
                          {(() => {
                            const fc = subfolderFileCountMap.get(sfName) ?? subfolderFileCountMap.get(sfName.trim().toLowerCase()) ?? 0;
                            const fsc = subfolderFolderCountMap.get(sfName) ?? subfolderFolderCountMap.get(sfName.trim().toLowerCase()) ?? 0;
                            return (
                              <>
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 3, background: fsc > 0 ? '#e0f2fe' : 'var(--vdms-border-soft)', color: fsc > 0 ? '#0369a1' : 'var(--vdms-text-faint)', borderRadius: 20, padding: '1px 8px', fontSize: 10, fontWeight: 700, lineHeight: '16px', whiteSpace: 'nowrap' }}>📁 {fsc}</span>
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 3, background: fc > 0 ? '#dcfce7' : 'var(--vdms-border-soft)', color: fc > 0 ? '#15803d' : 'var(--vdms-text-faint)', borderRadius: 20, padding: '1px 8px', fontSize: 10, fontWeight: 700, lineHeight: '16px', whiteSpace: 'nowrap' }}>📄 {fc}</span>
                              </>
                            );
                          })()}
                        </div>
                        {showDebugKeys && (
                          <div style={{ marginTop: 4, fontSize: 10, color: '#7c3aed', fontFamily: 'monospace', wordBreak: 'break-all' }}>
                            key: {fullBreadcrumbWithSf} | live:{folderDebugLiveId}
                          </div>
                        )}
                      </div>
                      <span style={{ color: 'var(--vdms-text-faint)', fontSize: 16 }}>›</span>
                    </div>
                    );
                  })}
                </div>

              </div>
            ) : visibleCurrentFolderFiles.length > 0 ? (
              /* Folder File Items List */
              <div style={{ background: 'var(--vdms-surface)', borderRadius: 12, border: '1px solid var(--vdms-border)', overflow: 'hidden' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: 'var(--vdms-surface-alt)', borderBottom: '1px solid var(--vdms-border)', color: 'var(--vdms-text-muted)', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
                      <th style={{ padding: '10px 16px', width: 40, textAlign: 'center' }}></th>
                      <th style={{ padding: '10px 16px' }}>FILE NAME</th>
                      <th style={{ padding: '10px 16px' }}>SIZE</th>
                      <th style={{ padding: '10px 16px' }}>DATE & TIME UPLOADED</th>
                      <th style={{ padding: '10px 16px', textAlign: 'right' }}>ACTION</th>
                    </tr>
                  </thead>
                  <tbody>
                    {visibleCurrentFolderFiles.map((file, idx) => {
                      const fileId = (file as any).id || file.name;
                      const isSelected = folderViewSelectedFiles.has(fileId);
                      const currentFolderSharePointPath = atKaizenRoot
                        ? ['Kaizen - Knowledge Bank', ...folderPathStack.slice(1).map(n => n.name)].join(' > ')
                        : (atCommonShips
                          ? ['Common for all vessels', ...(docMainFolder ? [docMainFolder] : []), ...folderPathStack.slice(1).map(n => n.name)].join(' > ')
                          : (currentVesselNameFromStack && docMainFolder
                            ? [currentVesselNameFromStack, docMainFolder, ...folderPathStack.slice(2).map(n => n.name)].filter(Boolean).join(' > ')
                            : (currentFolderNode?.name || '')));
                      return (
                        <tr key={file.name + idx} style={{ borderBottom: '1px solid var(--vdms-border-soft)' }}>
                          <td style={{ padding: '12px 16px', textAlign: 'center', width: 40 }}>
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={() => {
                                host.setState(prev => {
                                  const next = new Set(prev.folderViewSelectedFiles);
                                  if (next.has(fileId)) next.delete(fileId); else next.add(fileId);
                                  return { folderViewSelectedFiles: next };
                                });
                              }}
                              style={{ width: 16, height: 16, accentColor: '#ef4444', cursor: 'pointer' }}
                            />
                          </td>
                          <td style={{ padding: '12px 16px', fontWeight: 600, color: 'var(--vdms-text)', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                            <span style={{ fontSize: 18 }}>{(file as any).pending ? '⏳' : '📄'}</span>
                            <span
                              onClick={() => {
                                if ((file as any).pending) {
                                  alert(`File "${file.name}" is pending — it will be available after approval.`);
                                  return;
                                }
                                const currentPath = atKaizenRoot
                                  ? ['Kaizen - Knowledge Bank', ...folderPathStack.slice(1).map(n => n.name)].join(' > ')
                                  : (atCommonShips
                                    ? ['Common for all vessels', ...(docMainFolder ? [docMainFolder] : []), ...folderPathStack.slice(1).map(n => n.name)].join(' > ')
                                    : (currentVesselNameFromStack && docMainFolder
                                      ? [currentVesselNameFromStack, docMainFolder, ...folderPathStack.slice(2).map(n => n.name)].filter(Boolean).join(' > ')
                                      : (currentFolderNode?.name || '')));
                                void host._openDocumentFile(fileId, file.name, currentPath);
                              }}
                              style={{ cursor: 'pointer', color: '#0284c7', textDecoration: 'underline' }}
                              title={`Click to view/download ${file.name}`}
                            >
                              {file.name}
                            </span>
                            {(file as any).pending && <span style={{ fontSize: 11, color: '#d97706', fontWeight: 600, background: '#fef3c7', borderRadius: 4, padding: '1px 6px' }}>Pending Approval</span>}
                            {(() => {
                              const detectedVessel = resolveDetectedVesselForFile(file, host, currentVesselNameFromStack);
                              const isUnidentified = (host?.state?.ocrUnidentifiedFiles || []).some(
                                n => (n || '').trim().toLowerCase() === file.name.trim().toLowerCase()
                              );

                              if (detectedVessel && !isUnidentified) {
                                return (
                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      host._openVesselSuggestions([new File([], file.name)], detectedVessel);
                                    }}
                                    title={`View OCR Vessel Suggestion: ${detectedVessel}`}
                                    style={{
                                      background: 'rgba(59,130,246,0.1)', border: '1px solid rgba(59,130,246,0.3)',
                                      borderRadius: 6, padding: '1px 6px', fontSize: 10, color: '#0284c7',
                                      cursor: 'pointer', fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: 3,
                                    }}
                                  >
                                    <span>✨</span> {detectedVessel}
                                  </button>
                                );
                              } else {
                                return (
                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      host._openVesselSuggestions([new File([], file.name)], currentVesselNameFromStack || '');
                                    }}
                                    title="Vessel name not detected by OCR — click to assign vessel"
                                    style={{
                                      background: '#fff7ed', border: '1px solid #fed7aa',
                                      borderRadius: 6, padding: '1px 6px', fontSize: 10, color: '#c2410c',
                                      cursor: 'pointer', fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: 3,
                                    }}
                                  >
                                    <span>⚠️</span> Vessel: Not detected
                                  </button>
                                );
                              }
                            })()}
                          </td>
                          <td style={{ padding: '12px 16px', color: 'var(--vdms-text-muted)' }}>{file.size}</td>
                          <td style={{ padding: '12px 16px', color: 'var(--vdms-text-muted)' }}>{file.date}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                            <div style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
                              <button
                                onClick={() => {
                                  if ((file as any).pending) {
                                    alert(`File "${file.name}" is pending — it will be available after approval.`);
                                    return;
                                  }
                                  const currentPath = atKaizenRoot
                                    ? ['Kaizen - Knowledge Bank', ...folderPathStack.slice(1).map(n => n.name)].join(' > ')
                                    : (atCommonShips
                                      ? ['Common for all vessels', ...(docMainFolder ? [docMainFolder] : []), ...folderPathStack.slice(1).map(n => n.name)].join(' > ')
                                      : (currentVesselNameFromStack && docMainFolder
                                        ? [currentVesselNameFromStack, docMainFolder, ...folderPathStack.slice(2).map(n => n.name)].filter(Boolean).join(' > ')
                                        : (currentFolderNode?.name || '')));
                                  void host._openDocumentFile(fileId, file.name, currentPath);
                                }}
                                style={{ border: '1px solid var(--vdms-border)', background: 'var(--vdms-surface)', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#0078d4' }}
                                title={`View or download ${file.name}`}
                              >
                                View / Download
                              </button>
                              {!String(fileId).startsWith('file_') && !/^\d+$/.test(String(fileId)) && (
                                <button
                                  onClick={() => void host._archiveDocumentFile(String(fileId), file.name, currentFolderName || '', docMainFolder || '', currentVesselNameFromStack || '')}
                                  style={{ border: '1px solid #c4b5fd', background: '#f5f3ff', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer', color: '#6d28d9' }}
                                  title={`Archive ${file.name}`}
                                >
                                  📦 Archive
                                </button>
                              )}
                              <button
                                onClick={() => host._openFileDeleteDialog([{ id: fileId, name: file.name, folderId: currentFolderNode?.id || '', folderPath: currentFolderNode?.name || '' }])}
                                style={{ border: '1px solid #fca5a5', background: '#fff5f5', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#ef4444' }}
                                title="Delete file"
                              >
                                🗑 Delete
                              </button>
                              <button
                                type="button"
                                onClick={() => void host._openSharePointFolder({ subFolderPath: currentFolderSharePointPath } as any)}
                                title="Open this folder in SharePoint"
                                aria-label={`Open ${currentFolderSharePointPath} in SharePoint`}
                                style={{ width: 28, height: 27, padding: 0, borderRadius: 6, border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8', cursor: 'pointer', fontSize: 16, fontWeight: 700, lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
                              >
                                ↗
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            ) : (
              /* Empty Folder View (Screenshot 2) */
              <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', padding: '60px 20px', width: '100%' }}>
                <div style={{
                  background: 'rgba(240, 249, 255, 0.6)', border: '1px solid #e0f2fe',
                  borderRadius: 24, padding: '48px 40px', maxWidth: 500, width: '100%',
                  textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12,
                }}>
                  <div style={{
                    width: 56, height: 56, borderRadius: 14, background: '#e0f2fe',
                    display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 24, color: '#0284c7',
                    marginBottom: 4,
                  }}>
                    📁
                  </div>
                  <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: 'var(--vdms-text)' }}>This folder is empty</h3>
                  <p style={{ margin: 0, fontSize: 13, color: 'var(--vdms-text-muted)', maxWidth: 320, lineHeight: 1.5 }}>
                    Use the Upload button in the top-right to add a document.
                  </p>
                </div>
              </div>
            )}

            {/* ── Segregated Section: Subfolder Level Unmatched Items Inside Vessel ── */}
            {(() => {
              const currentVessel = currentVesselNameFromStack || (folderPathStack.length > 0 ? folderPathStack[0].name : null);
              const subAnomalies = (host.state.folderAnomalies || []).filter(a =>
                a.anomaly_type === 'subfolder_unmatched' && (!currentVessel || a.vessel_name === currentVessel)
              );
              if (subAnomalies.length === 0) return null;

              return (
                <div style={{ background: '#fff7ed', border: '1px solid #ffedd5', borderRadius: 14, padding: 18, marginTop: 16 }}>
                  <h4 style={{ margin: '0 0 8px', fontSize: 14, fontWeight: 700, color: '#c2410c', display: 'flex', alignItems: 'center', gap: 8 }}>
                    ⚠️ Other / Unclassified Items Inside Vessel Tree ({subAnomalies.length})
                  </h4>
                  <p style={{ margin: '0 0 12px', fontSize: 12, color: '#9a3412' }}>
                    These items were added inside the vessel folder in SharePoint but are not part of the standard template structure.
                  </p>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                    {subAnomalies.map(item => (
                      <div
                        key={item.id}
                        style={{
                          background: 'var(--vdms-surface)', borderRadius: 8, border: '1px solid #fed7aa', padding: '10px 14px',
                          display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                          <span style={{ fontSize: 18 }}>{item.item_type === 'folder' ? '📁' : '📄'}</span>
                          <div>
                            <span style={{ fontWeight: 600, fontSize: 13, color: 'var(--vdms-text)' }}>{item.name}</span>
                            <span style={{ fontSize: 11, color: 'var(--vdms-text-muted)', marginLeft: 8 }}>Path: {item.spo_path}</span>
                          </div>
                        </div>
                        <button
                          onClick={() => host._dismissAnomaly(item.id)}
                          style={{ background: 'var(--vdms-surface)', color: '#c2410c', border: '1px solid #fed7aa', borderRadius: 6, padding: '4px 10px', fontSize: 11, cursor: 'pointer' }}
                        >
                          ✕ Dismiss
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}
          </div>
        ) : (

          /* ── LIST VIEW ── */
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            {/* Table wrapper */}
            <div style={{ background: 'var(--vdms-surface)', borderRadius: 12, border: '1px solid var(--vdms-border)', boxShadow: '0 1px 2px rgba(15,23,42,0.04)', overflow: 'hidden', width: '100%' }}>
             <div style={{ overflowX: 'auto', maxHeight: '70vh', overflowY: 'auto' }}>
              <table style={{ width: '100%', minWidth: 1420, borderCollapse: 'separate', borderSpacing: 0, tableLayout: 'fixed', fontSize: 12 }}>
                <colgroup>
                  <col style={{ width: 48 }} />
                  <col style={{ width: 150 }} />
                  <col style={{ width: 130 }} />
                  <col style={{ width: 160 }} />
                  <col style={{ width: 130 }} />
                  <col style={{ width: 140 }} />
                  <col style={{ width: 140 }} />
                  <col style={{ width: 220 }} />
                  <col />
                  <col style={{ width: 82 }} />
                  <col style={{ width: 150 }} />
                  <col style={{ width: 150 }} />
                </colgroup>
                <thead>
                  <tr>
                    <th style={{ ...LIST_TH, textAlign: 'center' }}>#</th>
                    <th style={LIST_TH}>Vessel</th>
                    <th style={LIST_TH}>Domain</th>
                    <th style={LIST_TH}>Main folder</th>
                    <th style={LIST_TH}>Group</th>
                    <th style={LIST_TH}>Category</th>
                    <th style={LIST_TH}>Sub category</th>
                    <th style={LIST_TH}>Folder path</th>
                    <th style={LIST_TH}>File name</th>
                    <th style={{ ...LIST_TH, textAlign: 'right' }}>Size</th>
                    <th style={LIST_TH}>Uploaded</th>
                    <th style={{ ...LIST_TH, textAlign: 'center' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {pageGroupedRows.length === 0 ? (
                    <tr>
                      <td colSpan={12} style={{ padding: '48px 16px', textAlign: 'center' }}>
                        {(host.state.loading || (Boolean(vesselLoadingName) && filtered.length === 0) || (documentFilesLoading && allRows.length === 0) || (documentVesselsLoadingMore && filtered.length === 0) || (docScopeType === 'sites' && Array.from(host._siteFolderItemsCache.values()).some(entry => entry.loading))) ? (
                          <div style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
                            <div style={{
                              width: 32, height: 32, border: '3px solid #e0f2fe',
                              borderTop: '3px solid #0284c7', borderRadius: '50%',
                              animation: 'spin 0.8s linear infinite',
                            }} />
                            <div style={{ fontWeight: 700, fontSize: 13, color: 'var(--vdms-text)' }}>
                              {vesselLoadingName
                                ? `Loading documents and attachments for ${vesselLoadingName}...`
                                : (documentVesselsLoadingMore ? 'Loading more vessels...' : 'Loading vessel documents from SharePoint...')}
                            </div>
                            <div style={{ fontSize: 11, color: 'var(--vdms-text-muted)', maxWidth: 360 }}>
                              Please wait while folder structures and live files are loaded.
                            </div>
                          </div>
                        ) : (
                          <div style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 8 }}>
                            <div style={{ fontSize: 28 }}>🔍</div>
                            <div style={{ fontWeight: 600, fontSize: 13, color: 'var(--vdms-text-secondary)' }}>No documents found</div>
                            <div style={{ color: 'var(--vdms-text-faint)', fontSize: 12, maxWidth: 400 }}>
                              {textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || docCategoryFilter !== 'all' || docGroupLevelFilter !== 'all' || docLeafCategoryFilter !== 'all' || docSubCategoryFilter !== 'all'
                                ? 'No rows match your current filter criteria.'
                                : 'No documents or folders are available for this section.'}
                            </div>
                            {(textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || docCategoryFilter !== 'all' || docGroupLevelFilter !== 'all' || docLeafCategoryFilter !== 'all' || docSubCategoryFilter !== 'all') && (
                              <button
                                type="button"
                                onClick={() => host.setState({
                                  textFilter: '',
                                  vesselFilter: 'all',
                                  docGroupFilter: 'all',
                                  docCategoryFilter: 'all',
                                  docGroupLevelFilter: 'all',
                                  docLeafCategoryFilter: 'all',
                                  docSubCategoryFilter: 'all',
                                  catFilter: 'all',
                                  docListPage: 0,
                                })}
                                style={{
                                  marginTop: 6,
                                  background: '#eff6ff',
                                  color: '#2563eb',
                                  border: '1px solid #bfdbfe',
                                  borderRadius: 6,
                                  padding: '4px 12px',
                                  fontSize: 12,
                                  fontWeight: 600,
                                  cursor: 'pointer',
                                }}
                              >
                                Clear filters
                              </button>
                            )}
                          </div>
                        )}
                      </td>
                    </tr>
                  ) : pageGroupedRows.map((r, idx) => {
                    const globalIdx = safePage * PAGE_ROWS + idx + 1;
                    const isUploading = docUploadRowKey === r.groupKey && docUploadBusy;
                    const hasFiles = r.files.length > 0;
                    const rowFileIds = r.files.map(f => f.id);
                    const rowSelectedCount = rowFileIds.filter(id => listViewSelectedFiles.has(id)).length;
                    const listViewLabels = getListViewLabels(r);
                    const rowDomain = (r.domain || '').trim();
                    const zebra = idx % 2 === 1;
                    // Folder path shown from the library root down ("Documents ›
                    // type of vessel › Bow Fraternity › Sea-Trial"); the full
                    // path stays in the tooltip.
                    const pathSegments = (r.subFolderPath || '').split(/\s*>\s*/).map(seg => seg.trim()).filter(Boolean);
                    const libraryIdx = pathSegments.findIndex(seg => /^(documents|shared documents|sites documents|site library|general documents)$/i.test(seg));
                    const shortPath = (libraryIdx >= 0 ? pathSegments.slice(libraryIdx) : pathSegments).join(' › ');

                    return (
                      <tr key={`${r.groupKey}-${idx}`}
                        style={{ borderBottom: '1px solid var(--vdms-border-soft)', background: zebra ? 'var(--vdms-surface-alt)' : 'var(--vdms-surface)', transition: 'background 0.1s' }}
                        onMouseEnter={e => (e.currentTarget.style.background = '#f0f9ff')}
                        onMouseLeave={e => (e.currentTarget.style.background = zebra ? 'var(--vdms-surface-alt)' : 'var(--vdms-surface)')}
                      >
                        <td style={{ ...LIST_TD, textAlign: 'center', color: 'var(--vdms-text-faint)', fontSize: 11, fontVariantNumeric: 'tabular-nums' }}>{globalIdx}</td>
                        <td style={{ ...LIST_TD, fontWeight: 700, color: 'var(--vdms-text)' }}>
                          {listViewLabels.vessel === 'Not Listed' || listViewLabels.vessel === 'Vessel name not listed' ? (
                            <span style={{ ...LIST_PILL, background: '#ffedd5', color: '#c2410c' }} title="No vessel associated with this file">
                              <span aria-hidden="true" style={LIST_X_DOT}>✕</span>
                              Not Listed
                            </span>
                          ) : listViewLabels.vessel === 'Common for all vessels' ? (
                            <span style={{ ...LIST_PILL, background: '#fef3c7', color: '#92400e' }}>📁 Common for all vessels</span>
                          ) : listViewLabels.vessel === 'Kaizen - Knowledge Bank' ? (
                            <span style={{ ...LIST_PILL, background: '#ede9fe', color: '#6b21a8' }}>📚 Kaizen - Knowledge Bank</span>
                          ) : (
                            <span style={LIST_WRAP} title={listViewLabels.vessel}>{listViewLabels.vessel}</span>
                          )}
                        </td>
                        <td style={LIST_TD}>
                          {rowDomain ? (
                            <span style={{ ...LIST_PILL, background: '#e0e7ff', color: '#3730a3', whiteSpace: 'normal' }} title={rowDomain}>
                              <span style={LIST_WRAP}>{rowDomain}</span>
                            </span>
                          ) : <ListMissingMark title="No domain tag assigned" />}
                        </td>
                        <td style={LIST_TD}>
                          {listViewLabels.mainFolder ? (
                            <span
                              title={listViewLabels.mainFolder}
                              style={{
                                ...LIST_PILL, maxWidth: '100%', whiteSpace: 'normal',
                                background: listViewLabels.mainFolder === 'Kaizen - Knowledge Bank' ? '#ede9fe' : (listViewLabels.mainFolder === 'Insurance' ? '#fef3c7' : (listViewLabels.mainFolder === 'Commercial & Chartering' ? '#dcfce7' : '#eff6ff')),
                                color: listViewLabels.mainFolder === 'Kaizen - Knowledge Bank' ? '#6b21a8' : (listViewLabels.mainFolder === 'Insurance' ? '#b45309' : (listViewLabels.mainFolder === 'Commercial & Chartering' ? '#15803d' : '#2563eb')),
                              }}
                            >
                              <span style={LIST_WRAP}>{listViewLabels.mainFolder}</span>
                            </span>
                          ) : <ListMissingMark title="No main folder" />}
                        </td>
                        <td style={LIST_TD}>
                          {listViewLabels.group ? (
                            <span style={{ ...LIST_PILL, background: '#f3e8ff', color: '#7e22ce', whiteSpace: 'normal' }}>
                              <span style={LIST_WRAP}>{listViewLabels.group}</span>
                            </span>
                          ) : <ListMissingMark title="No group tag assigned" />}
                        </td>
                        <td style={{ ...LIST_TD, fontWeight: 600, color: 'var(--vdms-text)' }}>
                          {listViewLabels.category
                            ? <span style={LIST_WRAP} title={listViewLabels.category}>{listViewLabels.category}</span>
                            : <ListMissingMark title="No category tag assigned" />}
                        </td>
                        <td style={LIST_TD}>
                          {listViewLabels.subCategory
                            ? <span style={LIST_WRAP} title={listViewLabels.subCategory}>{listViewLabels.subCategory}</span>
                            : <ListMissingMark title="No sub category tag assigned" />}
                        </td>
                        <td style={LIST_TD} title={r.subFolderPath}>
                          <button
                            type="button"
                            onClick={() => openFolderViewForListRow(r)}
                            style={{
                              border: 'none', background: 'transparent', color: '#0369a1', cursor: 'pointer',
                              textAlign: 'left', padding: 0, fontSize: 11.5, lineHeight: '17px', maxWidth: '100%',
                              display: '-webkit-box', WebkitLineClamp: 3, WebkitBoxOrient: 'vertical', overflow: 'hidden',
                              wordBreak: 'break-word',
                            } as React.CSSProperties}
                            title={`Open in Folder view: ${r.subFolderPath}`}
                          >
                            {shortPath || r.subFolderPath}
                          </button>
                          {showDebugKeys && (
                            <div style={{ marginTop: 4, color: '#7c3aed', fontSize: 10, fontFamily: 'monospace', wordBreak: 'break-all' }}>
                              {getRowDebugKey(r)}
                            </div>
                          )}
                        </td>
                        <td style={LIST_TD}>
                          {hasFiles ? (
                            <div style={{ display: 'flex', flexDirection: 'column' }}>
                              {r.files.map(file => (
                                <div key={file.id || file.name} style={{ ...LIST_FILE_LINE, gap: 8 }}>
                                  <input
                                    type="checkbox"
                                    checked={listViewSelectedFiles.has(file.id)}
                                    onChange={() => {
                                      host.setState(prev => {
                                        const next = new Set(prev.listViewSelectedFiles);
                                        if (next.has(file.id)) next.delete(file.id); else next.add(file.id);
                                        return { listViewSelectedFiles: next };
                                      });
                                    }}
                                    style={{ width: 14, height: 14, accentColor: '#0284c7', cursor: 'pointer', flexShrink: 0, margin: 0 }}
                                    aria-label={`Select ${file.name}`}
                                  />
                                  <span aria-hidden="true" style={{ fontSize: 13, flexShrink: 0 }}>📄</span>
                                  <span
                                    onClick={() => {
                                      if (file.id && !file.id.startsWith('file_')) {
                                        void host._openDocumentFile(file.id, file.name);
                                      } else {
                                        alert(`File "${file.name}" is pending — it will be available after approval.`);
                                      }
                                    }}
                                    style={{ ...LIST_ELLIPSIS, flex: '1 1 auto', color: '#0369a1', fontWeight: 600, cursor: 'pointer' }}
                                    title={file.id && /^\d+$/.test(file.id) ? `${file.name} (pending approval - click to preview staged copy)` : file.name}
                                  >
                                    {file.name}{file.id && /^\d+$/.test(file.id) ? ' ⏳' : ''}
                                  </span>
                                  {(() => {
                                    const detectedVessel = resolveDetectedVesselForFile(file, host, r.vesselName);
                                    const isUnidentified = (host?.state?.ocrUnidentifiedFiles || []).some(
                                      n => (n || '').trim().toLowerCase() === file.name.trim().toLowerCase()
                                    );
                                    if (!detectedVessel || isUnidentified) return null;
                                    return (
                                      <button
                                        type="button"
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          host._openVesselSuggestions([new File([], file.name)], detectedVessel);
                                        }}
                                        title={`View OCR Vessel Suggestion: ${detectedVessel}`}
                                        style={{
                                          background: 'rgba(59,130,246,0.1)', border: '1px solid rgba(59,130,246,0.3)',
                                          borderRadius: 6, padding: '1px 6px', fontSize: 10, color: '#0284c7', flexShrink: 0,
                                          cursor: 'pointer', fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: 3,
                                        }}
                                      >
                                        <span>✨</span> {detectedVessel}
                                      </button>
                                    );
                                  })()}
                                </div>
                              ))}
                            </div>
                          ) : (
                            <span style={{ ...LIST_FILE_LINE, color: 'var(--vdms-text-faint)' }}>—</span>
                          )}
                        </td>
                        <td style={{ ...LIST_TD, textAlign: 'right', color: 'var(--vdms-text-muted)', fontSize: 11, fontVariantNumeric: 'tabular-nums' }}>
                          {hasFiles
                            ? r.files.map(file => <div key={file.id || file.name} style={{ ...LIST_FILE_LINE, justifyContent: 'flex-end' }}>{file.size || '—'}</div>)
                            : <span style={LIST_FILE_LINE}>—</span>}
                        </td>
                        <td style={{ ...LIST_TD, color: 'var(--vdms-text-muted)', fontSize: 11, fontVariantNumeric: 'tabular-nums' }}>
                          {hasFiles
                            ? r.files.map(file => (
                              <div key={file.id || file.name} style={LIST_FILE_LINE} title={file.uploadedAt ? new Date(file.uploadedAt).toISOString() : undefined}>
                                {file.uploadedAt ? new Date(file.uploadedAt).toLocaleString() : '—'}
                              </div>
                            ))
                            : <span style={LIST_FILE_LINE}>—</span>}
                        </td>
                        <td style={LIST_TD}>
                          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 6 }}>
                            <label
                              title="Upload files into this folder"
                              style={{
                                ...LIST_ACTION_BTN,
                                background: isUploading ? 'var(--vdms-border-soft)' : 'var(--vdms-surface)',
                                color: isUploading ? 'var(--vdms-text-faint)' : 'var(--vdms-text)',
                                cursor: isUploading ? 'not-allowed' : 'pointer',
                              }}
                            >
                              <input
                                type="file"
                                multiple
                                style={{ display: 'none' }}
                                disabled={isUploading}
                                onChange={e => {
                                  const filesList = Array.from(e.target.files || []);
                                  if (filesList.length === 0) return;
                                  e.target.value = '';
                                  const bulkFiles = filesList.map(f => ({ file: f, relativePath: f.name }));
                                  host._openBulkUpload(bulkFiles, r.uploadFolderId, r.subFolderPath, r.vesselName);
                                }}
                              />
                              {isUploading ? '⏳' : '↑ Upload'}
                            </label>
                            <button
                              type="button"
                              disabled={rowSelectedCount === 0}
                              onClick={() => {
                                if (rowSelectedCount === 0) return;
                                const currentSelected = host.state.listViewSelectedFiles;
                                const filesToDelete = r.files
                                  .filter(f => currentSelected.has(f.id))
                                  .map(f => ({ id: f.id, name: f.name, folderId: r.uploadFolderId, folderPath: r.subFolderPath }));
                                if (filesToDelete.length === 0) return;
                                host.setState({ listViewSelectedFiles: new Set() });
                                host._openFileDeleteDialog(filesToDelete);
                              }}
                              style={{
                                ...LIST_ACTION_BTN,
                                border: `1px solid ${rowSelectedCount > 0 ? '#fca5a5' : 'var(--vdms-border)'}`,
                                background: rowSelectedCount > 0 ? '#fff5f5' : 'var(--vdms-surface-alt)',
                                cursor: rowSelectedCount > 0 ? 'pointer' : 'not-allowed',
                                color: rowSelectedCount > 0 ? '#ef4444' : 'var(--vdms-text-faint)',
                              }}
                              title={rowSelectedCount > 0 ? `Delete ${rowSelectedCount} selected file(s)` : 'Select files to delete'}
                            >
                              🗑 {rowSelectedCount > 0 ? `(${rowSelectedCount})` : 'Delete'}
                            </button>
                            {r.canUpload && r.uploadFolderId && !r.uploadFolderId.includes('/') ? (
                              <button
                                type="button"
                                onClick={() => host._openAddFolderDialog({
                                  folderId: r.uploadFolderId,
                                  folderLabel: r.subFolderPath,
                                  vesselName: r.vesselName,
                                })}
                                style={{ ...LIST_ACTION_BTN, border: '1px solid #bae6fd', background: '#f0f9ff', color: '#0369a1' }}
                                title="Add a subfolder here"
                              >
                                + Folder
                              </button>
                            ) : <span />}
                            <button
                              type="button"
                              onClick={() => { void host._openSharePointFolder(r); }}
                              title="Open the folder containing these files in SharePoint Online"
                              aria-label={`Open ${r.subFolderPath} in SharePoint Online`}
                              style={{ ...LIST_ACTION_BTN, border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8' }}
                            >
                              ↗ Open
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
             </div>

              {/* Pagination footer */}
              <div style={{ padding: '10px 14px', color: 'var(--vdms-text-muted)', fontSize: 11, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'var(--vdms-surface-alt)', borderTop: '1px solid var(--vdms-border)' }}>
                <span>
                  Showing {filtered.length === 0 ? 0 : safePage * PAGE_ROWS + 1}–{Math.min((safePage + 1) * PAGE_ROWS, filtered.length)} of {filtered.length} rows
                </span>
                <div style={{ display: 'flex', gap: 3, alignItems: 'center' }}>
                  <button
                    onClick={() => host.setState({ docListPage: Math.max(0, safePage - 1) })}
                    disabled={safePage === 0}
                    style={{ border: '1px solid var(--vdms-border)', background: 'var(--vdms-surface)', borderRadius: 4, padding: '3px 8px', fontSize: 11, cursor: safePage === 0 ? 'not-allowed' : 'pointer', opacity: safePage === 0 ? 0.4 : 1 }}
                  >‹</button>
                  {Array.from({ length: totalPages }, (_, i) => (
                    <button
                      key={i}
                      onClick={() => host.setState({ docListPage: i })}
                      style={{
                        border: i === safePage ? 'none' : '1px solid var(--vdms-border)',
                        background: i === safePage ? '#0078d4' : 'var(--vdms-surface)',
                        color: i === safePage ? '#fff' : 'var(--vdms-text)',
                        borderRadius: 4, padding: '3px 8px', fontSize: 11,
                        fontWeight: i === safePage ? 700 : 400,
                        cursor: 'pointer',
                        minWidth: 26,
                      }}
                    >{i + 1}</button>
                  ))}
                  <button
                    onClick={() => host.setState({ docListPage: Math.min(totalPages - 1, safePage + 1) })}
                    disabled={safePage >= totalPages - 1}
                    style={{ border: '1px solid var(--vdms-border)', background: 'var(--vdms-surface)', borderRadius: 4, padding: '3px 8px', fontSize: 11, cursor: safePage >= totalPages - 1 ? 'not-allowed' : 'pointer', opacity: safePage >= totalPages - 1 ? 0.4 : 1 }}
                  >›</button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    );



}
