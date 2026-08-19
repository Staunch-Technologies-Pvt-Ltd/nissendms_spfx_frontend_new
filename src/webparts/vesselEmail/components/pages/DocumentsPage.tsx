import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import {
  badge, GROUP_COLORS, DATASOURCE_TAGS_MAP, VESSEL_TYPES, cleanName, suggestTagFromFilename,
  INITIAL_MOCK_DOCUMENTS, INITIAL_MOCK_TEMPLATES, INITIAL_MOCK_APPROVALS,
  INITIAL_MOCK_NOTIFICATIONS, INITIAL_MOCK_USERS,
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
import { renderClassifyDialog } from './VesselsPage';
import {
  folderNamesByMainFolder, subfolderNamesByFolder,
  getCommonShipsFlatRows, getKaizenFlatRows,
} from '../vesselFolderTemplate';

export function renderDocumentsPage(host: VesselEmail): React.ReactElement {

    const {
      textFilter, vesselFilter, catFilter, attachmentFilter, docViewMode, showAllVesselsInFolderView,
      vessels, rows, docListPage, docListSort, docGroupFilter,
      documentVesselCount, documentVesselsLoadingMore,
      docUploadRowKey, docUploadBusy, docUploadMsg, documentsList,
      folderPathStack, uploadedFilesByFolder, docMainFolder,
      folderNavHistory, folderNavIndex, listViewSelectedFiles, folderViewSelectedFiles,
    } = host.state;

    const canGoBack = folderNavIndex > 0;
    const canGoForward = folderNavIndex < folderNavHistory.length - 1;

    const goBack = (): void => {
      if (!canGoBack) return;
      const prev = folderNavHistory[folderNavIndex - 1];
      host.setState({ folderNavIndex: folderNavIndex - 1, folderPathStack: prev.folderPathStack, docMainFolder: prev.docMainFolder });
    };

    const goForward = (): void => {
      if (!canGoForward) return;
      const next = folderNavHistory[folderNavIndex + 1];
      host.setState({ folderNavIndex: folderNavIndex + 1, folderPathStack: next.folderPathStack, docMainFolder: next.docMainFolder });
    };

    const PAGE_ROWS = 10;

    // Main folder definitions
    type MainFolderKey = 'Technical & Crewing' | 'Commercial & Chartering' | 'Insurance' | 'Kaizen - Knowledge Bank' | 'Knowledge Bank';
    const VESSEL_MAIN_FOLDERS: Array<{ key: MainFolderKey; icon: string; emoji: string; color: string; bg: string }> = [
      { key: 'Technical & Crewing', icon: '⚙️', emoji: '⚙️', color: '#dc2626', bg: '#fee2e2' },
      { key: 'Commercial & Chartering', icon: '💼', emoji: '💼', color: '#16a34a', bg: '#dcfce7' },
      { key: 'Insurance', icon: '🛡️', emoji: '🛡️', color: '#d97706', bg: '#fef3c7' },
    ];
    const MAIN_FOLDERS: Array<{ key: MainFolderKey; icon: string; emoji: string; color: string; bg: string }> = [
      ...VESSEL_MAIN_FOLDERS,
      { key: 'Kaizen - Knowledge Bank', icon: '📚', emoji: '📚', color: '#7c3aed', bg: '#ede9fe' },
    ];

    // Hierarchy navigation state
    const stackLevel = folderPathStack.length;
    const atVesselsRoot = stackLevel === 1 && folderPathStack[0]?.id === 'vessels_root';
    const atSpecificVessels = stackLevel === 2 && folderPathStack[1]?.id === 'specific_vessels';
    const atCommonShips = stackLevel === 1 && folderPathStack[0]?.id === 'common';
    // Vessel node is always at index 2 when path is [vessels_root, specific_vessels, vessel]
    // For other paths (e.g. kaizen), fall back to searching by exclusion
    const KNOWN_NAV_IDS = new Set(['vessels_root', 'specific_vessels', 'common', 'kaizen_root']);
    const vesselStackIdx = (() => {
      // Standard path: vessels_root > specific_vessels > {vessel}
      if (stackLevel >= 3 && folderPathStack[0]?.id === 'vessels_root' && folderPathStack[1]?.id === 'specific_vessels') {
        return 2;
      }
      // Fallback: find first node that is not a known nav node and not a main folder name
      return folderPathStack.findIndex(n =>
        !KNOWN_NAV_IDS.has(n.id) && !VESSEL_MAIN_FOLDERS.some(mf => mf.key === n.name)
      );
    })();
    const vesselNodeInStack = vesselStackIdx !== -1 ? folderPathStack[vesselStackIdx] : null;
    const atVesselMainFolderSelect = vesselNodeInStack !== null && !docMainFolder && stackLevel === vesselStackIdx + 1;
    const currentVesselNameFromStack = vesselNodeInStack?.name || null;

    if (currentVesselNameFromStack && !host._filesLoadedForVessels.has(currentVesselNameFromStack)) {
      host._filesLoadedForVessels.add(currentVesselNameFromStack);
      setTimeout(() => host._loadFilesForVessel(currentVesselNameFromStack!).catch(() => undefined), 0);
    }

    const currentFolderNode = folderPathStack.length > 0 ? folderPathStack[folderPathStack.length - 1] : null;
    const currentFolderName = currentFolderNode ? currentFolderNode.name : null;

    const DEFAULT_VESSEL_MAINS = folderNamesByMainFolder();
    const SUBFOLDERS_MAP = subfolderNamesByFolder();
    const COMMON_DEFAULT_MAINS = folderNamesByMainFolder(true);

    // Resolve real SPO folder ID for the current folder node from rows
    const resolvedCurrentFolderId: string | null = (() => {
      if (!currentFolderNode) return null;
      if (!/^(sf_|category_|common|vessels_root|specific_vessels|kaizen_root)/.test(currentFolderNode.id)) return currentFolderNode.id;
      // Try to find a matching row by vessel + folder name chain
      const vesselName = currentVesselNameFromStack || (folderPathStack.length > 0 ? folderPathStack[0].name : null);
      if (!vesselName) return null;
      const folderName = currentFolderNode.name;
      const candidateSubPath = `${vesselName} > ${docMainFolder || ''} > ${folderName}`;
      const liveId = host._getLiveSharePointFolderId(candidateSubPath);
      if (liveId) return liveId;
      // Determine the parent folder name from the navigation stack so we can
      // disambiguate "To be Classified" (or any folder name that appears under
      // multiple parents).  e.g. "To be Classified" under "Month End Reports"
      // vs "To be Classified" directly under "Technical & Crewing".
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
      return matchRow?.uploadFolderId || null;
    })();

    // Determine subfolders for current depth
   // Determine subfolders for current depth
    const mainsDefaultsSource = atCommonShips ? COMMON_DEFAULT_MAINS : DEFAULT_VESSEL_MAINS;
    let subfolderNames: string[] = [];
    if (currentFolderName && SUBFOLDERS_MAP[currentFolderName]) {
      subfolderNames = SUBFOLDERS_MAP[currentFolderName];
    } else if (docMainFolder && mainsDefaultsSource[docMainFolder]) {
      // If we are at the main folder level inside a vessel (or inside Common for all ships)
      if (
        !currentFolderName ||
        currentFolderName === docMainFolder ||
        (vesselNodeInStack && currentFolderNode?.id === vesselNodeInStack.id) ||
        atVesselsRoot ||
        atSpecificVessels ||
        atCommonShips ||
        folderPathStack.length <= 3
      ) {
        subfolderNames = mainsDefaultsSource[docMainFolder];
      }
    }

    type DisplayFile = { name: string; size: string; date: string; pending?: boolean; id?: string; uploadedAt?: number };
    const allCurrentFolderFilesRaw: DisplayFile[] = [];

    // Collect all keys under which files for the current folder may be stored
    const currentFolderLiveId = resolvedCurrentFolderId
      ? host._getLiveSharePointFolderId(
          currentVesselNameFromStack && docMainFolder && currentFolderNode
            ? `${currentVesselNameFromStack} > ${docMainFolder} > ${currentFolderNode.name}`
            : ''
        ) || resolvedCurrentFolderId
      : null;
    // Also find the matching row's groupKey for this folder — scope by main folder
    // and parent in stack so that "To be Classified" under different parents resolves correctly.
    const parentInStack = folderPathStack.length >= 2 ? folderPathStack[folderPathStack.length - 2] : null;
    const currentFolderMatchRow = currentFolderNode && currentVesselNameFromStack
      ? host.state.rows.find(r =>
          r.vesselName === currentVesselNameFromStack &&
          (r.subCategory === currentFolderNode.name || r.category === currentFolderNode.name) &&
          (docMainFolder ? r.group === docMainFolder : true) &&
          (parentInStack && parentInStack.name !== currentVesselNameFromStack
            ? (r.category === parentInStack.name || r.subCategory === parentInStack.name)
            : true)
        )
      : null;

    // Trigger live SPO file refresh when we have a real folder ID
    if (resolvedCurrentFolderId && !/^(sf_|category_|common)/.test(resolvedCurrentFolderId) && !/^\d+$/.test(resolvedCurrentFolderId)) {
      const alreadyLoaded = host._filesLoadedForFolders.has(resolvedCurrentFolderId);
      const inFlight = host._refreshFolderFilesInFlight.has(resolvedCurrentFolderId);
      if (!alreadyLoaded && !inFlight) {
        const triggerGroupKey = currentFolderMatchRow?.groupKey || resolvedCurrentFolderId;
        setTimeout(() => host._refreshFolderFiles(resolvedCurrentFolderId!, triggerGroupKey, true).catch(() => undefined), 0);
      }
    }

    const seenFileNames = new Set<string>();
    const addFiles = (list: DisplayFile[] | undefined): void => {
      (list || []).forEach(f => {
        if (f?.name && !seenFileNames.has(f.name.toLowerCase())) {
          seenFileNames.add(f.name.toLowerCase());
          allCurrentFolderFilesRaw.push(f);
        }
      });
    };
    // Check all possible keys: resolvedCurrentFolderId, live upload ID, groupKey, node ID, display name
    if (resolvedCurrentFolderId) addFiles(uploadedFilesByFolder[resolvedCurrentFolderId]);
    if (currentFolderLiveId && currentFolderLiveId !== resolvedCurrentFolderId) addFiles(uploadedFilesByFolder[currentFolderLiveId]);
    if (currentFolderMatchRow?.groupKey) addFiles(uploadedFilesByFolder[currentFolderMatchRow.groupKey]);
    if (currentFolderMatchRow?.uploadFolderId && currentFolderMatchRow.uploadFolderId !== resolvedCurrentFolderId) addFiles(uploadedFilesByFolder[currentFolderMatchRow.uploadFolderId]);
    if (currentFolderNode && !/^(sf_|category_|common)/.test(currentFolderNode.id) && currentFolderNode.id !== resolvedCurrentFolderId) addFiles(uploadedFilesByFolder[currentFolderNode.id]);
    // Also check by display folder name (used as key by the top-level upload handler)
    if (currentFolderName) addFiles(uploadedFilesByFolder[currentFolderName]);
    // Also check by the full breadcrumb path (lowercase) — stored by the top-level upload handler as normSub
    if (currentVesselNameFromStack && docMainFolder && currentFolderName) {
      const breadcrumbKey = `${currentVesselNameFromStack} > ${docMainFolder} > ${currentFolderName}`.trim().toLowerCase();
      addFiles(uploadedFilesByFolder[breadcrumbKey]);
      // Also check all rows matching this folder for their subFolderPath keys — scoped by parent context
      host.state.rows
        .filter(r => r.vesselName === currentVesselNameFromStack && (r.subCategory === currentFolderName || r.category === currentFolderName) &&
          (docMainFolder ? r.group === docMainFolder : true) &&
          (parentInStack && parentInStack.name !== currentVesselNameFromStack
            ? (r.category === parentInStack.name || r.subCategory === parentInStack.name)
            : true))
        .forEach(r => {
          if (r.subFolderPath) addFiles(uploadedFilesByFolder[r.subFolderPath.trim().toLowerCase()]);
          if (r.groupKey) addFiles(uploadedFilesByFolder[r.groupKey]);
          if (r.uploadFolderId && r.uploadFolderId !== resolvedCurrentFolderId) addFiles(uploadedFilesByFolder[r.uploadFolderId]);
        });
    }
    // Scan all uploadedFilesByFolder entries whose key contains the current folder's Graph ID
    // (handles the case where the upload stored files under the confirmed Graph folder ID)
    if (currentFolderNode) {
      const knownIds = new Set([resolvedCurrentFolderId, currentFolderLiveId, currentFolderMatchRow?.uploadFolderId, currentFolderNode.id].filter(Boolean) as string[]);
      for (const [key, files] of Object.entries(uploadedFilesByFolder)) {
        if (!knownIds.has(key) && key !== currentFolderName) {
          // Only include if the key looks like a Graph drive item ID (not a path or display name)
          // and matches one of the IDs we know about via the spoFolderMap
          if (host.state.spoFolderMap.has(key) && (key === resolvedCurrentFolderId || key === currentFolderLiveId ||
              (currentFolderMatchRow && key === currentFolderMatchRow.uploadFolderId))) {
            addFiles(files);
          }
        }
      }
    }
    // Also include files from backend rows matching this folder
    if (currentFolderNode) {
      const matchIds = new Set<string>();
      if (!/^(sf_|category_|common)/.test(currentFolderNode.id)) matchIds.add(currentFolderNode.id);
      if (resolvedCurrentFolderId) matchIds.add(resolvedCurrentFolderId);
      rows.filter(r => matchIds.has(r.uploadFolderId) && r.fileName && !r.filePending)
        .forEach(r => {
          if (!seenFileNames.has(r.fileName!.toLowerCase())) {
            seenFileNames.add(r.fileName!.toLowerCase());
            allCurrentFolderFilesRaw.push({ name: r.fileName!, size: '—', date: '—', uploadedAt: r.fileUploadedAt });
          }
        });
    }
    // Also include files from rows matched by groupKey (covers post-upload state updates)
    if (currentFolderMatchRow) {
      rows.filter(r => r.groupKey === currentFolderMatchRow.groupKey && r.fileName && !r.filePending)
        .forEach(r => {
          if (!seenFileNames.has(r.fileName!.toLowerCase())) {
            seenFileNames.add(r.fileName!.toLowerCase());
            allCurrentFolderFilesRaw.push({ name: r.fileName!, size: '—', date: '—', id: r.fileId || undefined, uploadedAt: r.fileUploadedAt });
          }
        });
    }
    const fileTimestamp = (file: DisplayFile): number => {
      if (typeof file.uploadedAt === 'number' && Number.isFinite(file.uploadedAt)) return file.uploadedAt;
      const parsed = Date.parse(file.date);
      return Number.isFinite(parsed) ? parsed : 0;
    };
    const allCurrentFolderFiles = [...allCurrentFolderFilesRaw].sort(
      (a, b) => fileTimestamp(b) - fileTimestamp(a) || a.name.localeCompare(b.name)
    );
    const visibleCurrentFolderFiles = attachmentFilter === 'not_attached' ? [] : allCurrentFolderFiles;
    // ── Build source rows with 3 segregated options ──
    const docScopeType = host.state.docScopeType || 'vessels';
    const commonTemplateRows = getCommonShipsFlatRows();
    const kaizenTemplateRows = getKaizenFlatRows();

    // 1. Sanitize specific vessel rows from API: remove any Kaizen rows from individual vessels
    const sanitizedVesselRows = (rows || []).filter(r => {
      const isKaizen = (r.group || '').toLowerCase().includes('kaizen') || (r.subFolderPath || '').toLowerCase().includes('kaizen');
      const isCommon = r.vesselName === 'Common for all vessels' || (r.subFolderPath || '').toLowerCase().includes('common for all ships');
      return !isKaizen && !isCommon;
    });

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

    const allRows: FlatRow[] = [
      ...sanitizedVesselRows,
      ...mergedCommonRows,
      ...mergedKaizenRows,
    ];

    // Filter by active scope (1st option: Specific Vessels, 2nd option: Common for all vessels, 3rd option: Kaizen - Knowledge Bank)
    const scopeRows = allRows.filter(r => {
      if (docScopeType === 'vessels') {
        return r.vesselName !== 'Common for all vessels' && r.vesselName !== 'Kaizen - Knowledge Bank';
      } else if (docScopeType === 'common') {
        return r.vesselName === 'Common for all vessels';
      } else if (docScopeType === 'kaizen') {
        return r.vesselName === 'Kaizen - Knowledge Bank';
      }
      return true;
    });

    // Filtered list rows for List view
    const activeVesselName = (docScopeType === 'vessels' && vesselFilter !== 'all') ? vesselFilter : null;
    const allGroups = Array.from(new Set(scopeRows.map(r => r.group))).sort();
    const allCategories = Array.from(
      new Set(scopeRows.filter(r => docGroupFilter === 'all' || r.group === docGroupFilter).map(r => r.category))
    ).sort();

    // ── mainFolderGroupMap: which groups belong to which main folder (for list view filtering) ──
    const mainFolderGroupMap = folderNamesByMainFolder(docScopeType === 'common');

    // Vessels are returned newest first. Start with four and extend in pages
    // of eight when the user chooses "More vessels".
    const visibleVesselNames = new Set(vessels.slice(0, documentVesselCount).map(v => v.name));

    let filtered = scopeRows.filter(r => {
      if (docScopeType === 'vessels') {
        if (!activeVesselName) {
          if (!visibleVesselNames.has(r.vesselName)) return false;
        } else if (r.vesselName !== activeVesselName) {
          return false;
        }
      }
      // Filter by active main folder
      if (docMainFolder) {
        const allowedCats = mainFolderGroupMap[docMainFolder] || [];
        const mainMatch = r.group === docMainFolder ||
          r.group.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0].toLowerCase()) ||
          allowedCats.includes(r.group) ||
          allowedCats.includes(r.category) ||
          (r.subFolderPath || '').toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0].toLowerCase());
        if (!mainMatch) return false;
      }
      if (docGroupFilter !== 'all' && r.group !== docGroupFilter) return false;
      if (catFilter !== 'all' && r.category !== catFilter && r.group !== catFilter) return false;
      if (textFilter) {
        const q = textFilter.toLowerCase();
        return r.vesselName.toLowerCase().includes(q) || r.group.toLowerCase().includes(q) ||
          r.category.toLowerCase().includes(q) || (r.subCategory || '').toLowerCase().includes(q) ||
          (r.subFolderPath || '').toLowerCase().includes(q) ||
          (r.fileName || '').toLowerCase().includes(q);
      }
      return true;
    });

    // ── Group filtered rows by folder path (category/folder level) ──
    const groupedMap = new Map<string, GroupedRow>();
    for (const r of filtered) {
      const dedupeKey = (r.subFolderPath || r.groupKey).trim().toLowerCase();
      const normSub = (r.subFolderPath || '').trim().toLowerCase();
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

      const liveUploads = liveId ? (uploadedFilesByFolder[liveId] || []) : [];
      const subFolderUploads = normSub ? (uploadedFilesByFolder[normSub] || []) : [];
      const dedupeUploads = dedupeKey ? (uploadedFilesByFolder[dedupeKey] || []) : [];
      const groupUploads = r.groupKey ? (uploadedFilesByFolder[r.groupKey] || []) : [];
      // Do NOT read from r.uploadFolderId — it may be a backend DB ID shared across multiple "To be Classified" rows
      // Only read from live Graph folder IDs and unique per-row keys (groupKey, normSub, dedupeKey)

      const folderUploads = [
        ...groupUploads,
        ...liveUploads,
        ...subFolderUploads,
        ...dedupeUploads,
        ...memFiles,
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

    // Navigation folder names that should never appear as vessel cards
    const NAV_FOLDER_NAMES = new Set(['specific vessels', 'common for all ships', 'vessels', 'kaizen - knowledge bank', 'knowledge bank', 'common']);
    // Deduplicate by name and exclude any vessel whose name matches a navigation folder
    const displayVessels = vessels.slice(0, documentVesselCount).filter(
      (v, i, arr) =>
        arr.findIndex(x => x.name.trim().toLowerCase() === v.name.trim().toLowerCase()) === i &&
        !NAV_FOLDER_NAMES.has(v.name.trim().toLowerCase())
    );

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        {/* Breadcrumb Navigation Trail */}
        <div style={{ fontSize: 12, color: '#64748b', display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
          {/* Back / Forward navigation buttons */}
          <button
            onClick={goBack}
            disabled={!canGoBack}
            title="Go back"
            style={{
              width: 28, height: 28, borderRadius: 6, border: '1px solid #cbd5e1',
              background: canGoBack ? '#fff' : '#f1f5f9',
              color: canGoBack ? '#334155' : '#cbd5e1',
              cursor: canGoBack ? 'pointer' : 'not-allowed',
              fontSize: 14, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              flexShrink: 0,
            }}
          >‹</button>
          <button
            onClick={goForward}
            disabled={!canGoForward}
            title="Go forward"
            style={{
              width: 28, height: 28, borderRadius: 6, border: '1px solid #cbd5e1',
              background: canGoForward ? '#fff' : '#f1f5f9',
              color: canGoForward ? '#334155' : '#cbd5e1',
              cursor: canGoForward ? 'pointer' : 'not-allowed',
              fontSize: 14, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              flexShrink: 0,
            }}
          >›</button>
          <span style={{ cursor: 'pointer', color: '#0284c7' }} onClick={() => host._pushFolderNav([], null)}>Home</span>
          <span>›</span>
          <span
            style={{ cursor: stackLevel === 0 && !docMainFolder ? 'default' : 'pointer', color: stackLevel === 0 && !docMainFolder ? '#0f172a' : '#0284c7', fontWeight: stackLevel === 0 && !docMainFolder ? 600 : 400 }}
            onClick={() => host._pushFolderNav([], null)}
          >
            Documents
          </span>
         {folderPathStack.map((item, idx) => {
            const mainFolderFollows = !!docMainFolder && idx === vesselStackIdx;
            const isLast = idx === folderPathStack.length - 1 && !mainFolderFollows;
            return (
              <React.Fragment key={item.id + idx}>
                <span>›</span>
                <span
                  onClick={() => host._pushFolderNav(folderPathStack.slice(0, idx + 1), idx < vesselStackIdx ? null : docMainFolder)}
                  style={{ cursor: isLast ? 'default' : 'pointer', color: isLast ? '#0f172a' : '#0284c7', fontWeight: isLast ? 600 : 400 }}
                >
                  {item.name}
                </span>
                {mainFolderFollows && (
                  <>
                    <span>›</span>
                    <span
                      onClick={() => host._pushFolderNav(folderPathStack.slice(0, idx + 1), docMainFolder)}
                      style={{
                        cursor: idx === folderPathStack.length - 1 ? 'default' : 'pointer',
                        color: idx === folderPathStack.length - 1 ? '#0f172a' : '#0284c7',
                        fontWeight: idx === folderPathStack.length - 1 ? 600 : 400,
                      }}
                    >
                      {docMainFolder}
                    </span>
                  </>
                )}
              </React.Fragment>
            );
          })}
        </div>

        {/* Module Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 20 }}>📁</span>
              {(currentFolderNode && currentFolderNode !== vesselNodeInStack && currentFolderNode.id !== 'vessels_root' && currentFolderNode.id !== 'specific_vessels' ? currentFolderNode.name : docMainFolder) || currentFolderName || 'Documents'}
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>
              {docViewMode === 'list'
                ? `${filtered.length} rows · flattened list view`
                : stackLevel === 0 && !docMainFolder
                  ? '3 top-level folders'
                  : atVesselsRoot
                    ? '2 folders · Specific Vessels & Common'
                    : atSpecificVessels
                      ? `${vessels.length} vessels`
                      : atVesselMainFolderSelect
                        ? '3 main folders'
                        : `${subfolderNames.length} folders · ${allCurrentFolderFiles.length} files`}
            </p>
            {docViewMode === 'folder' && currentFolderNode && (
              <div style={{ marginTop: 6, fontSize: 12, fontWeight: 600, color: allCurrentFolderFiles.length > 0 ? '#15803d' : '#64748b' }}>
                {allCurrentFolderFiles.length > 0 ? '✅ Attached' : '⚪ Not Attached'}
              </div>
            )}
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            {/* Folder view / List view Pill Toggle */}
            <div style={{ display: 'inline-flex', background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8, padding: 3, gap: 2 }}>
              <button
                onClick={() => host.setState({ docViewMode: 'folder' })}
                style={{
                  padding: '5px 12px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                  background: docViewMode === 'folder' ? '#0f172a' : 'transparent',
                  color: docViewMode === 'folder' ? '#fff' : '#64748b',
                }}
              >
                ⊞ Folder view
              </button>
              <button
                onClick={() => host.setState({ docViewMode: 'list', docMainFolder: null, folderPathStack: [] })}
                style={{
                  padding: '5px 12px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                  background: docViewMode === 'list' ? '#0f172a' : 'transparent',
                  color: docViewMode === 'list' ? '#fff' : '#64748b',
                }}
              >
                ☰ List view
              </button>
            </div>

            {/* Archive Button */}
            <button
              onClick={() => host._goToView('archive')}
              style={{
                background: '#d97706', color: '#fff', border: 'none', borderRadius: 8,
                padding: '7px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6,
              }}
            >
              🗑 Archive
            </button>

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
                  background: listViewSelectedFiles.size > 0 ? '#fff5f5' : '#f8fafc',
                  color: listViewSelectedFiles.size > 0 ? '#ef4444' : '#cbd5e1',
                  border: `1px solid ${listViewSelectedFiles.size > 0 ? '#fca5a5' : '#e2e8f0'}`,
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
            {docViewMode === 'folder' && visibleCurrentFolderFiles.length > 0 && (
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
                  background: folderViewSelectedFiles.size > 0 ? '#fff5f5' : '#f8fafc',
                  color: folderViewSelectedFiles.size > 0 ? '#ef4444' : '#cbd5e1',
                  border: `1px solid ${folderViewSelectedFiles.size > 0 ? '#fca5a5' : '#e2e8f0'}`,
                  borderRadius: 8, padding: '7px 16px', fontSize: 13, fontWeight: 600,
                  cursor: folderViewSelectedFiles.size > 0 ? 'pointer' : 'not-allowed',
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                }}
                title={folderViewSelectedFiles.size > 0 ? `Delete ${folderViewSelectedFiles.size} selected file(s)` : 'Select files to delete'}
              >
                🗑 Delete{folderViewSelectedFiles.size > 0 ? ` (${folderViewSelectedFiles.size})` : ''}
              </button>
            )}

            {/* Top-Right Upload Button */}
            <label style={{
              background: '#0284c7', color: '#fff', border: 'none', borderRadius: 8,
              padding: '7px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 4px rgba(2,132,199,0.2)',
            }}>
              <input
                type="file"
                style={{ display: 'none' }}
                onChange={async e => {
                  const file = e.target.files?.[0];
                  if (!file) return;
                  // Build a unique folderKey that includes the parent path so that
                  // "To be Classified" under "Technical & Crewing" is distinct from
                  // "To be Classified" under "Commercial & Chartering".
                  const folderKey = (docMainFolder && currentFolderName && currentFolderName !== docMainFolder)
                    ? `${currentVesselNameFromStack || ''} > ${docMainFolder} > ${currentFolderName}`
                    : (docMainFolder ? `${currentVesselNameFromStack || ''} > ${docMainFolder}` : currentFolderName || 'Documents');
                  const topFolderId = currentFolderNode && !/^(sf_|common|vessels_root|specific_vessels|kaizen_root)/.test(currentFolderNode.id) ? currentFolderNode.id : '';
                  const currentVessel = currentVesselNameFromStack || (vessels.length > 0 ? vessels[0].name : 'Bow Fighter');
                  let subFolderPath: string;
                if (vesselStackIdx !== -1) {
  // docMainFolder (e.g. "Technical & Crewing") is tracked separately from
  // folderPathStack, so it must be spliced back in here or it silently
  // disappears from the breadcrumb used to resolve the SPO upload path.
  const afterVesselItems = folderPathStack.slice(vesselStackIdx + 1).map(n => n.name);
  subFolderPath = [currentVessel, ...(docMainFolder ? [docMainFolder] : []), ...afterVesselItems].join(' > ');
} else {
                    subFolderPath = currentFolderNode
                      ? `${currentVessel} > ${docMainFolder || ''} > ${currentFolderNode.name}`
                      : `${currentVessel} > ${docMainFolder || ''}`;
                  }

                  // Map display folder names to actual SharePoint folder path names
                  const mainFolderPathMap: Record<string, string> = {
                    'Technical & Crewing': 'Technical & Crewing',
                    'Commercial & Chartering': 'Commercial & Chartering',
                    'Insurance': 'Insurance',
                    'Kaizen - Knowledge Bank': 'Kaizen - Knowledge Bank',
                    'Knowledge Bank': 'Kaizen - Knowledge Bank',
                  };
                  const spoMainFolder = (docMainFolder && mainFolderPathMap[docMainFolder]) || 'Technical & Crewing';

                  const liveId = host._getLiveSharePointFolderId(subFolderPath);
                  // Try to find a real backend uploadFolderId from existing rows for host vessel+folder
                  const matchingRow = host.state.rows.find(r =>
                    r.vesselName === currentVessel &&
                    r.uploadFolderId &&
                    !r.uploadFolderId.includes('/') &&
                    (docMainFolder ? r.group.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0]) : true)
                  );
                  // Build fallback path that includes the subfolder (category/sub-category) if we're in one
                  // This ensures uploads go to the exact folder shown in the folder view, not just the main folder
                  const fallbackPathParts = [host.VESSEL_ROOT, 'Specific Vessels', currentVessel];
                  if (docMainFolder) fallbackPathParts.push(spoMainFolder);
                  // Include the current folder name if it's a subfolder (not the main folder itself)
                  if (currentFolderNode && currentFolderName && currentFolderName !== docMainFolder && currentFolderName !== 'Documents') {
                    fallbackPathParts.push(currentFolderName);
                  }
                  const fallbackPath = fallbackPathParts.join('/');
                  const resolvedFolderId = topFolderId || liveId || (matchingRow?.uploadFolderId) || fallbackPath;

                  host.setState({ docUploadMsg: null });
                  try {
                    const { fileId, statusPending, folderId, isGraphUpload } = await host._uploadFileToFolder(
                      resolvedFolderId,
                      subFolderPath,
                      currentVessel,
                      file
                    );

                    const msg = statusPending
                      ? `"${file.name}" submitted for approval.`
                      : `"${file.name}" uploaded successfully to ${folderKey}!`;

                                       const newUpload = { name: file.name, size: `${(file.size / 1024).toFixed(1)} KB`, date: 'Just now', pending: statusPending, id: fileId, uploadedAt: Date.now() };
                    const liveFolderId = folderId || topFolderId;

                    // If the upload returned a confirmed Graph folder ID, patch the
                    // folderPathStack so resolvedCurrentFolderId uses the correct ID.
                    if (liveFolderId && currentFolderNode && currentFolderNode.id !== liveFolderId) {
                      const patchedStack = host.state.folderPathStack.map(n =>
                        n.name === currentFolderNode.name && n.id === currentFolderNode.id
                          ? { ...n, id: liveFolderId }
                          : n
                      );
                      host.setState({ folderPathStack: patchedStack });
                    }

                    host.setState(prev => {
                      // Use functional setState to avoid stale closure
                      const uploadedFilesByFolder = prev.uploadedFilesByFolder;
                      const rows = (prev.rows || []) as FlatRow[];

                      // Find matching row from latest state — try folder ID match first,
                      // then fall back to breadcrumb match (rows from DB have DB IDs, not Graph IDs)
                    const matchingRowForUpload = rows.find(r =>
                    r.uploadFolderId === (liveFolderId || resolvedFolderId) ||
                      r.uploadFolderId === resolvedFolderId
) || rows.find(r =>
  r.vesselName === currentVessel &&
  (docMainFolder ? r.group === docMainFolder : true) &&
  (currentFolderNode && currentFolderNode.id !== 'vessels_root' && currentFolderNode.id !== 'specific_vessels'
    ? (r.subCategory === currentFolderNode.name || r.category === currentFolderNode.name)
    : true) &&
  (parentInStack && parentInStack.name !== currentVessel
    ? (r.category === parentInStack.name || r.subCategory === parentInStack.name)
    : true)
);

                      const updatedByFolder: Record<string, any[]> = { ...uploadedFilesByFolder };
                      const normSub = (subFolderPath || '').trim().toLowerCase();
                      // Always store under both the confirmed live ID and the resolved ID so
                      // resolvedCurrentFolderId (whichever ID it ends up being) finds the files.
                      // Only write to liveFolderId if isGraphUpload = true (real Graph drive item ID)
                      const keysToSet = [folderKey, resolvedFolderId, normSub, matchingRowForUpload?.groupKey].filter(Boolean) as string[];
                      if (isGraphUpload === true && liveFolderId) {
                        keysToSet.push(liveFolderId);
                      }
                      for (const key of keysToSet) {
                        updatedByFolder[key] = [...(uploadedFilesByFolder[key] || []).filter((f: any) => f.name !== file.name), newUpload];
                      }

                      console.log('[VesselDMS] Top-level upload optimistic update:', {
                        folderKey,
                        liveFolderId,
                        resolvedFolderId,
                        matchingRowGroupKey: matchingRowForUpload?.groupKey,
                        fileName: file.name,
                        updatedKeys: Object.keys(updatedByFolder),
                      });

                      return { uploadedFilesByFolder: updatedByFolder, docUploadMsg: msg, docListPage: 0 };
                    }, () => {
                      // Refresh the actual folder contents so the newly uploaded file and
                      // every file already in this SharePoint folder appear together.
                      // Single refresh - _refreshFolderFiles now handles Graph consistency window
                      // by preserving local uploads when Graph returns empty.
                      const refreshId = liveFolderId || resolvedFolderId;
                      // Use isGraphUpload from _uploadFileToFolder (true = Graph drive item ID)
                      const isRealGraphId = isGraphUpload === true;
                      // Find matching row for refresh after state updates
                                           const matchingRow = host.state.rows.find(r =>
                        r.uploadFolderId === (liveFolderId || resolvedFolderId) ||
                        r.uploadFolderId === resolvedFolderId
                      ) || host.state.rows.find(r =>
                        r.vesselName === currentVessel &&
                        (docMainFolder ? r.group === docMainFolder : true) &&
                        (currentFolderNode && currentFolderNode.id !== 'vessels_root' && currentFolderNode.id !== 'specific_vessels'
                          ? (r.subCategory === currentFolderNode.name || r.category === currentFolderNode.name)
                          : true) &&
                        (parentInStack && parentInStack.name !== currentVessel
                          ? (r.category === parentInStack.name || r.subCategory === parentInStack.name)
                          : true)
                      );
                      const gKey = matchingRow?.groupKey || refreshId;
                      // Always refresh using the confirmed live folder ID (folderId from upload)
                      // to avoid fetching from the wrong folder when _getLiveSharePointFolderId
                      // resolves to a different node than where the file actually landed.
                      void host._refreshFolderFiles(refreshId, gKey, true, isRealGraphId).catch(() => undefined);
                    });
                  } catch (err: any) {
                    host.setState({ docUploadMsg: `Upload failed: ${err?.message || 'Error'}` });
                  }
                }}
              />
              <span>⬆</span> Upload
            </label>
          </div>
        </div>

        {/* ── Filter Toolbar ── */}
        <div style={{ background: '#fff', borderRadius: 10, padding: '10px 12px', border: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', gap: 10 }}>
          {/* 3 Scope Options: 1st Option = Specific Vessels, 2nd Option = Common for all vessels, 3rd Option = Kaizen - Knowledge Bank */}
          {docViewMode === 'list' && (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 8, paddingBottom: 6, borderBottom: '1px solid #f1f5f9' }}>
              <div style={{ display: 'inline-flex', background: '#f8fafc', border: '1px solid #cbd5e1', borderRadius: 8, padding: 3, gap: 4 }}>
                <button
                  onClick={() => host.setState({ docScopeType: 'vessels', vesselFilter: 'all', docGroupFilter: 'all', catFilter: 'all', docListPage: 0 })}
                  style={{
                    padding: '6px 14px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 700, cursor: 'pointer',
                    background: docScopeType === 'vessels' ? '#0284c7' : 'transparent',
                    color: docScopeType === 'vessels' ? '#fff' : '#475569',
                    display: 'inline-flex', alignItems: 'center', gap: 6, transition: 'all 0.15s ease',
                  }}
                >
                  <span>🚢</span> 1. Specific Vessels
                </button>
                <button
                  onClick={() => host.setState({ docScopeType: 'common', vesselFilter: 'all', docGroupFilter: 'all', catFilter: 'all', docListPage: 0 })}
                  style={{
                    padding: '6px 14px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 700, cursor: 'pointer',
                    background: docScopeType === 'common' ? '#0284c7' : 'transparent',
                    color: docScopeType === 'common' ? '#fff' : '#475569',
                    display: 'inline-flex', alignItems: 'center', gap: 6, transition: 'all 0.15s ease',
                  }}
                >
                  <span>📁</span> 2. Common for all vessels
                </button>
                <button
                  onClick={() => host.setState({ docScopeType: 'kaizen', vesselFilter: 'all', docGroupFilter: 'all', catFilter: 'all', docListPage: 0 })}
                  style={{
                    padding: '6px 14px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 700, cursor: 'pointer',
                    background: docScopeType === 'kaizen' ? '#0284c7' : 'transparent',
                    color: docScopeType === 'kaizen' ? '#fff' : '#475569',
                    display: 'inline-flex', alignItems: 'center', gap: 6, transition: 'all 0.15s ease',
                  }}
                >
                  <span>📚</span> 3. Kaizen - Knowledge Bank
                </button>
              </div>

              <span style={{ fontSize: 12, color: '#64748b' }}>
                {docScopeType === 'vessels' && `Showing documents for ${vesselFilter !== 'all' ? vesselFilter : 'individual vessels'}`}
                {docScopeType === 'common' && 'Showing documents shared across all vessels (Common for all ships)'}
                {docScopeType === 'kaizen' && 'Showing global Kaizen - Knowledge Bank documents'}
              </span>
            </div>
          )}

          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
            <div style={{ position: 'relative', flex: '1 1 180px', minWidth: 160 }}>
              <span style={{ position: 'absolute', left: 9, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', fontSize: 12 }}>🔍</span>
              <input
                type="text"
                placeholder="Filter by vessel, group, category, path..."
                value={textFilter}
                onChange={e => host.setState({ textFilter: e.target.value, docListPage: 0 })}
                style={{ width: '100%', padding: '6px 10px 6px 28px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, outline: 'none', boxSizing: 'border-box' }}
              />
            </div>
            {docScopeType === 'vessels' && (
              <select
                value={vesselFilter}
                onChange={e => host.setState({ vesselFilter: e.target.value, docListPage: 0 })}
                style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 140 }}
              >
                <option value="all">All vessels</option>
                {vessels.map(v => (
                  <option key={v.id || v.name} value={v.name}>{v.name}</option>
                ))}
              </select>
            )}
            {docScopeType !== 'kaizen' && (
              <select
                value={docGroupFilter}
                onChange={e => host.setState({ docGroupFilter: e.target.value, catFilter: 'all', docListPage: 0 })}
                style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 160 }}
              >
                <option value="all">All groups</option>
                {allGroups.map(g => (
                  <option key={g} value={g}>{g}</option>
                ))}
              </select>
            )}
            <select
              value={catFilter}
              onChange={e => host.setState({ catFilter: e.target.value, docListPage: 0 })}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 160 }}
            >
              <option value="all">All categories</option>
              {allCategories.map(c => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
            <select
              value={attachmentFilter}
              onChange={e => host.setState({ attachmentFilter: e.target.value as any, docListPage: 0 })}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 180 }}
            >
              <option value="all">All attachment status</option>
              <option value="attached">Attachment Available</option>
              <option value="not_attached">Attachment Required</option>
            </select>
            <select
              value={docListSort}
              onChange={e => host.setState({ docListSort: e.target.value as any, docListPage: 0 })}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none' }}
            >
              <option value="default">Default order</option>
              <option value="name_az">Name A–Z</option>
              <option value="newest">Newest</option>
            </select>

            {docScopeType === 'vessels' && documentVesselCount < vessels.length && (
              <button
                onClick={() => void host._loadMoreDocumentVessels()}
                disabled={documentVesselsLoadingMore}
                style={{
                  padding: '6px 12px', borderRadius: 8, border: '1px solid #0284c7', fontSize: 12,
                  background: documentVesselsLoadingMore ? '#e2e8f0' : '#0284c7', color: documentVesselsLoadingMore ? '#64748b' : '#fff',
                  fontWeight: 600, cursor: documentVesselsLoadingMore ? 'wait' : 'pointer', outline: 'none',
                }}
              >
                {documentVesselsLoadingMore ? 'Loading vessels...' : `More vessels (+${Math.min(8, vessels.length - documentVesselCount)})`}
              </button>
            )}
            <div style={{ display: 'flex', border: '1px solid #cbd5e1', borderRadius: 8, overflow: 'hidden', marginLeft: 'auto' }}>
              <button
                onClick={() => host.setState({ docViewMode: 'folder' })}
                style={{ padding: '5px 10px', background: docViewMode === 'folder' ? '#e2e8f0' : '#fff', border: 'none', cursor: 'pointer', fontSize: 13 }}
                title="Folder view"
              >::</button>
              <button
                onClick={() => host.setState({ docViewMode: 'list', docMainFolder: null, folderPathStack: [] })}
                style={{ padding: '5px 10px', background: docViewMode === 'list' ? '#e2e8f0' : '#fff', border: 'none', borderLeft: '1px solid #cbd5e1', cursor: 'pointer', fontSize: 13 }}
                title="List view"
              >☰</button>
            </div>
          </div>
        </div>

        {/* Success Message Banner */}
        {docUploadMsg && (
          <div style={{ background: '#f0fdf4', border: '1px solid #86efac', borderRadius: 8, padding: '8px 14px', fontSize: 12, color: '#16a34a', display: 'flex', justifyContent: 'space-between' }}>
            <span>✓ {docUploadMsg}</span>
            <button onClick={() => host.setState({ docUploadMsg: null })} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#16a34a', fontWeight: 700 }}>✕</button>
          </div>
        )}

        {docViewMode === 'folder' ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 24, marginTop: 4 }}>

            {/* Level 0: Vessels + Kaizen - Knowledge Bank */}
            {stackLevel === 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                  {[
                    { id: 'vessels_root', name: 'Vessels', emoji: '🚢', bg: '#e0f2fe', sub: `${vessels.length} vessels` },
                    { id: 'common', name: 'Common for all vessels', emoji: '📁', bg: '#fef3c7', sub: 'Shared documents' },
                    { id: 'kaizen_root', name: 'Kaizen - Knowledge Bank', emoji: '📚', bg: '#ede9fe', sub: 'Knowledge base' },
                  ].map(item => (
                    <div
                      key={item.id}
                      onClick={() => {
                        if (item.id === 'kaizen_root') {
                          host._pushFolderNav([{ id: 'kaizen_root', name: 'Kaizen - Knowledge Bank' }], 'Kaizen - Knowledge Bank');
                          host.setState({ vesselFilter: 'all' });
                        } else if (item.id === 'common') {
                          host._pushFolderNav([{ id: 'common', name: 'Common for all vessels' }], null);
                          host.setState({ docScopeType: 'common', vesselFilter: 'all' });
                        } else {
                          host._pushFolderNav([{ id: 'vessels_root', name: 'Vessels' }], null);
                          host.setState({ docScopeType: 'vessels', vesselFilter: 'all' });
                        }
                      }}
                      style={{
                        background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18,
                        display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                        boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                      }}
                    >
                      <div style={{ width: 44, height: 44, borderRadius: 10, background: item.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>{item.emoji}</div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.name}</div>
                        <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>{item.sub}</div>
                      </div>
                      <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                    </div>
                  ))}
                </div>

                {/* Anomalies at root level */}
                {(() => {
                  const mainAnomalies = (host.state.folderAnomalies || []).filter(a => a.anomaly_type === 'main_folder_unmatched');
                  if (mainAnomalies.length === 0) return null;
                  const mainFolders = mainAnomalies.filter(a => a.item_type === 'folder');
                  const mainFiles   = mainAnomalies.filter(a => a.item_type === 'file');
                  return (
                    <div style={{ marginTop: 24, display: 'flex', flexDirection: 'column', gap: 16 }}>
                      {mainFolders.length > 0 && (
                        <div style={{ background: 'linear-gradient(135deg, #fffbeb 0%, #fff9e6 100%)', border: '2px solid #f59e0b', borderRadius: 14, padding: 20 }}>
                          <h4 style={{ margin: '0 0 6px', fontSize: 15, fontWeight: 700, color: '#92400e', display: 'flex', alignItems: 'center', gap: 8 }}>
                            ⚠️ Folders Created Outside Standard Main Folders ({mainFolders.length})
                          </h4>
                          <p style={{ margin: '0 0 12px', fontSize: 12, color: '#a16207' }}>
                            These folders were created directly in SharePoint Online at the main folder root level outside standard category structures.
                          </p>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                            {mainFolders.map(item => (
                              <div key={item.id} style={{ background: '#fff', borderRadius: 10, border: '1px solid #fde68a', padding: '12px 16px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 200 }}>
                                  <span style={{ fontSize: 24 }}>📁</span>
                                  <div>
                                    <div style={{ fontWeight: 700, fontSize: 13, color: '#1f1f1f' }}>{item.name}</div>
                                    <div style={{ fontSize: 11, color: '#78716c', fontFamily: 'monospace' }}>{item.spo_path}</div>
                                  </div>
                                </div>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                  <button onClick={() => host.setState({ spoClassifyDialog: { anomaly: item, provisioning: false, done: false, error: null } })} style={{ background: 'linear-gradient(135deg, #f59e0b, #d97706)', color: '#fff', border: 'none', borderRadius: 6, padding: '6px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                    🔍 Classify
                                  </button>
                                  <button onClick={() => host._dismissAnomaly(item.id)} style={{ background: '#fff', color: '#78716c', border: '1px solid #d6d3d1', borderRadius: 6, padding: '6px 10px', fontSize: 11, cursor: 'pointer' }}>
                                    ✕ Dismiss
                                  </button>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                      {mainFiles.length > 0 && (
                        <div style={{ background: 'linear-gradient(135deg, #faf5ff 0%, #f3e8ff 100%)', border: '2px solid #a855f7', borderRadius: 14, padding: 20 }}>
                          <h4 style={{ margin: '0 0 6px', fontSize: 15, fontWeight: 700, color: '#6b21a8', display: 'flex', alignItems: 'center', gap: 8 }}>
                            📄 Files Uploaded Outside Main Category Structure ({mainFiles.length})
                          </h4>
                          <p style={{ margin: '0 0 12px', fontSize: 12, color: '#7c3aed' }}>
                            These files were uploaded directly to the main folder root in SharePoint Online.
                          </p>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                            {mainFiles.map(item => {
                              const ext = item.name.split('.').pop()?.toUpperCase() || 'FILE';
                              return (
                                <div key={item.id} style={{ background: '#fff', borderRadius: 10, border: '1px solid #e9d5ff', padding: '10px 14px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                                    <span style={{ background: '#f3e8ff', color: '#6b21a8', borderRadius: 4, padding: '2px 6px', fontSize: 10, fontWeight: 700 }}>{ext}</span>
                                    <div>
                                      <span style={{ fontWeight: 600, fontSize: 13, color: '#1e293b' }}>{item.name}</span>
                                      <span style={{ fontSize: 11, color: '#64748b', marginLeft: 8, fontFamily: 'monospace' }}>{item.spo_path}</span>
                                    </div>
                                  </div>
                                  <button onClick={() => host._dismissAnomaly(item.id)} style={{ background: '#fff', color: '#6b21a8', border: '1px solid #e9d5ff', borderRadius: 6, padding: '4px 10px', fontSize: 11, cursor: 'pointer' }}>
                                    ✕ Dismiss
                                  </button>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })()}
              </div>

            ) : atVesselsRoot ? (

              /* Level 1: Inside Vessels - Specific Vessels */
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                {[
                  { id: 'specific_vessels', name: 'Specific Vessels', emoji: '🚢', bg: '#e0f2fe', sub: `${vessels.length} vessels` },
                ].map(item => (
                  <div
                    key={item.id}
                    onClick={() => host._pushFolderNav([...folderPathStack, { id: item.id, name: item.name }], null)}
                    style={{ background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' }}
                  >
                    <div style={{ width: 44, height: 44, borderRadius: 10, background: item.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>{item.emoji}</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a' }}>{item.name}</div>
                      <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>{item.sub}</div>
                    </div>
                    <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                  </div>
                ))}
              </div>

            ) : atSpecificVessels ? (

              /* Level 2: Vessel list under Specific Vessels */
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                {displayVessels.map(v => (
                  <div
                    key={v.id}
                    onClick={() => {
                      host._pushFolderNav([...folderPathStack, { id: v.id, name: v.name }], null);
                      host.setState({ vesselFilter: v.name });
                      host._loadFilesForVessel(v.name).catch(() => undefined);
                    }}
                    style={{ background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' }}
                  >
                    <div style={{ width: 44, height: 44, borderRadius: 10, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7' }}>🚢</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{v.name}</div>
                      <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>Vessel</div>
                    </div>
                    <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                  </div>
                ))}
              </div>

              ) : atCommonShips && !docMainFolder ? (

              /* Main folder selection inside "Common for all ships" */
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                {VESSEL_MAIN_FOLDERS.map(mf => (
                  <div
                    key={mf.key}
                    onClick={() => {
                      host._pushFolderNav(folderPathStack, mf.key);
                      host.setState({ vesselFilter: 'all' });
                    }}
                    style={{
                      background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18,
                      display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                      boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                    }}
                  >
                    <div style={{ width: 44, height: 44, borderRadius: 10, background: mf.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>{mf.emoji}</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{mf.key}</div>
                      <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>Common folder</div>
                    </div>
                    <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                  </div>
                ))}
              </div>

            ) : atVesselMainFolderSelect ? (

              /* Level 3: Main folder selection inside a vessel */
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                {VESSEL_MAIN_FOLDERS.map(mf => (
                  <div
                    key={mf.key}
                    onClick={() => {
                      host._pushFolderNav(folderPathStack, mf.key);
                      host.setState({ vesselFilter: currentVesselNameFromStack || 'all' });
                    }}
                    style={{
                      background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18,
                      display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                      boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                    }}
                  >
                    <div style={{ width: 44, height: 44, borderRadius: 10, background: mf.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>{mf.emoji}</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{mf.key}</div>
                      <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>Main folder</div>
                    </div>
                    <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                  </div>
                ))}
              </div>
            ) : subfolderNames.length > 0 ? (
              /* Subfolders Grid */
              <div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 }}>
                  {subfolderNames.map((sfName, idx) => (
                    <div
                      key={sfName + idx}
                      onClick={() => {
                        const vesselName = currentVesselNameFromStack || (folderPathStack.length > 0 ? folderPathStack[0].name : null);
                        const candidateSubPath = vesselName ? `${vesselName} > ${docMainFolder || ''} > ${sfName}` : '';
                        const liveFolderId = candidateSubPath ? host._getLiveSharePointFolderId(candidateSubPath) : null;
                        const matchingSubRow = vesselName ? host.state.rows.find(r =>
                          r.vesselName === vesselName &&
                          r.uploadFolderId &&
                          !r.uploadFolderId.includes('/') &&
                          (r.subCategory === sfName || r.category === sfName) &&
                          (docMainFolder ? r.group === docMainFolder : true) &&
                          (currentFolderNode && currentFolderNode.name !== vesselName
                            ? (r.category === currentFolderNode.name || r.group === currentFolderNode.name)
                            : true)
                        ) : null;
                        const realFolderId = liveFolderId || matchingSubRow?.uploadFolderId || `sf_${idx}`;
                        const newStack = [...folderPathStack, { id: realFolderId, name: sfName }];
                        host._pushFolderNav(newStack, docMainFolder);
                        // Immediately refresh files for this folder from live SPO
                        if (!/^sf_/.test(realFolderId)) {
                          const subGroupKey = matchingSubRow?.groupKey || realFolderId;
                          void host._refreshFolderFiles(realFolderId, subGroupKey, true).catch(() => undefined);
                        } else if (currentFolderNode && !/^(sf_|category_|common)/.test(currentFolderNode.id)) {
                          const parentGroupKey = currentFolderMatchRow?.groupKey || currentFolderNode.id;
                          host._refreshFolderFiles(currentFolderNode.id, parentGroupKey).catch(() => undefined);
                        }
                      }}
                      style={{
                        background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18,
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
                        <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{sfName}</div>
                        <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>Folder</div>
                      </div>
                      <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                    </div>
                  ))}
                </div>

                {allCurrentFolderFiles.length > 0 && (
                  <div style={{ marginTop: 24, background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                      <thead>
                        <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
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
                          return (
                            <tr key={file.name + idx} style={{ borderBottom: '1px solid #f1f5f9' }}>
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
                              <td style={{ padding: '12px 16px', fontWeight: 600, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 }}>
                                <span style={{ fontSize: 18 }}>{(file as any).pending ? '⏳' : '📄'}</span>
                                <span
                                  onClick={() => {
                                    if ((file as any).pending) {
                                      alert(`File "${file.name}" is pending — it will be available after approval.`);
                                      return;
                                    }
                                    const currentPath = currentVesselNameFromStack && docMainFolder && currentFolderNode
                                      ? `${currentVesselNameFromStack} > ${docMainFolder} > ${currentFolderNode.name}`
                                      : (currentFolderNode?.name || '');
                                    void host._openDocumentFile(fileId, file.name, currentPath);
                                  }}
                                  style={{ cursor: 'pointer', color: '#0284c7', textDecoration: 'underline' }}
                                  title={`Click to view/download ${file.name}`}
                                >
                                  {file.name}
                                </span>
                                {(file as any).pending && <span style={{ fontSize: 11, color: '#d97706', fontWeight: 600, background: '#fef3c7', borderRadius: 4, padding: '1px 6px' }}>Pending Approval</span>}
                              </td>
                              <td style={{ padding: '12px 16px', color: '#64748b' }}>{file.size}</td>
                              <td style={{ padding: '12px 16px', color: '#64748b' }}>{file.date}</td>
                              <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                                <div style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
                                  <button
                                    onClick={() => {
                                      if ((file as any).pending) {
                                        alert(`File "${file.name}" is pending — it will be available after approval.`);
                                        return;
                                      }
                                      const currentPath = currentVesselNameFromStack && docMainFolder && currentFolderNode
                                        ? `${currentVesselNameFromStack} > ${docMainFolder} > ${currentFolderNode.name}`
                                        : (currentFolderNode?.name || '');
                                      void host._openDocumentFile(fileId, file.name, currentPath);
                                    }}
                                    style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#0078d4' }}
                                    title={`View or download ${file.name}`}
                                  >
                                    View / Download
                                  </button>
                                  <button
                                    onClick={() => host._openFileDeleteDialog([{ id: fileId, name: file.name, folderId: currentFolderNode?.id || '', folderPath: currentFolderNode?.name || '' }])}
                                    style={{ border: '1px solid #fca5a5', background: '#fff5f5', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#ef4444' }}
                                    title="Delete file"
                                  >
                                    🗑 Delete
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            ) : visibleCurrentFolderFiles.length > 0 ? (
              /* Folder File Items List */
              <div style={{ background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
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
                      return (
                        <tr key={file.name + idx} style={{ borderBottom: '1px solid #f1f5f9' }}>
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
                          <td style={{ padding: '12px 16px', fontWeight: 600, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 }}>
                            <span style={{ fontSize: 18 }}>{(file as any).pending ? '⏳' : '📄'}</span>
                            <span
                              onClick={() => {
                                if ((file as any).pending) {
                                  alert(`File "${file.name}" is pending — it will be available after approval.`);
                                  return;
                                }
                                const currentPath = currentVesselNameFromStack && docMainFolder && currentFolderNode
                                  ? `${currentVesselNameFromStack} > ${docMainFolder} > ${currentFolderNode.name}`
                                  : (currentFolderNode?.name || '');
                                void host._openDocumentFile(fileId, file.name, currentPath);
                              }}
                              style={{ cursor: 'pointer', color: '#0284c7', textDecoration: 'underline' }}
                              title={`Click to view/download ${file.name}`}
                            >
                              {file.name}
                            </span>
                            {(file as any).pending && <span style={{ fontSize: 11, color: '#d97706', fontWeight: 600, background: '#fef3c7', borderRadius: 4, padding: '1px 6px' }}>Pending Approval</span>}
                          </td>
                          <td style={{ padding: '12px 16px', color: '#64748b' }}>{file.size}</td>
                          <td style={{ padding: '12px 16px', color: '#64748b' }}>{file.date}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                            <div style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
                              <button
                                onClick={() => {
                                  if ((file as any).pending) {
                                    alert(`File "${file.name}" is pending — it will be available after approval.`);
                                    return;
                                  }
                                  const currentPath = currentVesselNameFromStack && docMainFolder && currentFolderNode
                                    ? `${currentVesselNameFromStack} > ${docMainFolder} > ${currentFolderNode.name}`
                                    : (currentFolderNode?.name || '');
                                  void host._openDocumentFile(fileId, file.name, currentPath);
                                }}
                                style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#0078d4' }}
                                title={`View or download ${file.name}`}
                              >
                                View / Download
                              </button>
                              <button
                                onClick={() => host._openFileDeleteDialog([{ id: fileId, name: file.name, folderId: currentFolderNode?.id || '', folderPath: currentFolderNode?.name || '' }])}
                                style={{ border: '1px solid #fca5a5', background: '#fff5f5', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#ef4444' }}
                                title="Delete file"
                              >
                                🗑 Delete
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
                  <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#0f172a' }}>This folder is empty</h3>
                  <p style={{ margin: 0, fontSize: 13, color: '#64748b', maxWidth: 320, lineHeight: 1.5 }}>
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
                          background: '#fff', borderRadius: 8, border: '1px solid #fed7aa', padding: '10px 14px',
                          display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                          <span style={{ fontSize: 18 }}>{item.item_type === 'folder' ? '📁' : '📄'}</span>
                          <div>
                            <span style={{ fontWeight: 600, fontSize: 13, color: '#1e293b' }}>{item.name}</span>
                            <span style={{ fontSize: 11, color: '#64748b', marginLeft: 8 }}>Path: {item.spo_path}</span>
                          </div>
                        </div>
                        <button
                          onClick={() => host._dismissAnomaly(item.id)}
                          style={{ background: '#fff', color: '#c2410c', border: '1px solid #fed7aa', borderRadius: 6, padding: '4px 10px', fontSize: 11, cursor: 'pointer' }}
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
            <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', overflowX: 'auto', width: '100%' }}>
              <table style={{ width: '100%', minWidth: 950, borderCollapse: 'collapse', fontSize: 12 }}>
                <thead>
                  <tr style={{ background: '#f8fafc', borderBottom: '2px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
                    <th style={{ padding: '10px 10px', width: 44, textAlign: 'center' }}>SR.</th>
                    <th style={{ padding: '10px 12px' }}>VESSEL NAME</th>
                    <th style={{ padding: '10px 12px' }}>GROUP</th>
                    <th style={{ padding: '10px 12px' }}>CATEGORY</th>
                    <th style={{ padding: '10px 12px' }}>SUB-CATEGORY</th>
                    <th style={{ padding: '10px 12px' }}>FOLDER PATH</th>
                    <th style={{ padding: '10px 12px' }}>FILE NAME</th>
                    <th style={{ padding: '10px 12px' }}>SIZE</th>
                    <th style={{ padding: '10px 12px' }}>DATE &amp; TIME UPLOADED</th>
                    <th style={{ padding: '10px 12px', textAlign: 'right' }}>ATTACHMENT</th>
                  </tr>
                </thead>
                <tbody>
                  {pageGroupedRows.length === 0 ? (
                    <tr>
                      <td colSpan={10} style={{ padding: '36px 16px', textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>
                        No documents found. {textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || catFilter !== 'all' ? 'Try clearing the filters.' : ''}
                      </td>
                    </tr>
                  ) : pageGroupedRows.map((r, idx) => {
                    const globalIdx = safePage * PAGE_ROWS + idx + 1;
                    const isUploading = docUploadRowKey === r.groupKey && docUploadBusy;
                    const hasFiles = r.files.length > 0;
                    const rowFileIds = r.files.map(f => f.id);
                    const rowSelectedCount = rowFileIds.filter(id => listViewSelectedFiles.has(id)).length;

                    return (
                      <tr key={`${r.groupKey}-${idx}`} style={{ borderBottom: '1px solid #f1f5f9', transition: 'background 0.1s' }}
                        onMouseEnter={e => (e.currentTarget.style.background = '#f8fafc')}
                        onMouseLeave={e => (e.currentTarget.style.background = '')}
                      >
                        <td style={{ padding: '10px 10px', color: '#94a3b8', fontSize: 11, fontFamily: 'monospace', textAlign: 'center' }}>{globalIdx}</td>
                        <td style={{ padding: '10px 12px', fontWeight: 700, color: '#0f172a' }}>
                          {r.vesselName === 'Common for all vessels' ? (
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: '#fef3c7', color: '#92400e', borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 700 }}>
                              <span>📁</span> Common for all vessels
                            </span>
                          ) : r.vesselName === 'Kaizen - Knowledge Bank' ? (
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: '#ede9fe', color: '#6b21a8', borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 700 }}>
                              <span>📚</span> Kaizen - Knowledge Bank
                            </span>
                          ) : (
                            r.vesselName
                          )}
                        </td>
                        <td style={{ padding: '10px 12px' }}>
                          <span style={{
                            display: 'inline-block', borderRadius: 8, padding: '3px 8px', fontSize: 11, fontWeight: 600,
                            background: r.group === 'Kaizen - Knowledge Bank' ? '#ede9fe' : (r.group === 'Insurance' ? '#fef3c7' : (r.group === 'Commercial & Chartering' ? '#dcfce7' : '#eff6ff')),
                            color: r.group === 'Kaizen - Knowledge Bank' ? '#6b21a8' : (r.group === 'Insurance' ? '#b45309' : (r.group === 'Commercial & Chartering' ? '#15803d' : '#2563eb')),
                          }}>
                            {r.group}
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', fontWeight: 600, color: '#334155' }}>{r.category}</td>
                        <td style={{ padding: '10px 12px', fontWeight: 600, color: '#1e293b' }}>{r.subCategory || r.category}</td>
                        <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 11 }} title={r.subFolderPath}>
                          {r.subFolderPath}
                        </td>
                        <td style={{ padding: '10px 12px' }}>
                          {hasFiles ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              {r.files.map(file => (
                                <div key={file.name} style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
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
                                    style={{ width: 14, height: 14, accentColor: '#ef4444', cursor: 'pointer', flexShrink: 0 }}
                                  />
                                  <span style={{ fontSize: 14 }}>📄</span>
                                  <span
                                    onClick={() => {
                                      if (file.id && !file.id.startsWith('file_')) {
                                        void host._openDocumentFile(file.id, file.name);
                                      } else {
                                        alert(`File "${file.name}" is pending — it will be available after approval.`);
                                      }
                                    }}
                                    style={{ color: '#0284c7', textDecoration: 'underline', fontWeight: 600, cursor: 'pointer' }}
                                    title={file.id && /^\d+$/.test(file.id) ? `${file.name} (pending approval - click to preview staged copy)` : `Click to open ${file.name}`}
                                  >
                                    {file.name}{file.id && /^\d+$/.test(file.id) ? ' ⏳' : ''}
                                  </span>
                                </div>
                              ))}
                            </div>
                          ) : (
                            <span style={{ color: '#94a3b8', fontStyle: 'italic', fontSize: 11 }}>—</span>
                          )}
                        </td>
                        <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 11, whiteSpace: 'nowrap' }}>
                          {hasFiles ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              {r.files.map(file => <span key={file.id}>{file.size || '—'}</span>)}
                            </div>
                          ) : '—'}
                        </td>
                        <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 11, whiteSpace: 'nowrap' }}>
                          {hasFiles ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              {r.files.map(file => (
                                <span key={file.id} title={file.uploadedAt ? new Date(file.uploadedAt).toISOString() : undefined}>
                                  {file.uploadedAt ? new Date(file.uploadedAt).toLocaleString() : '—'}
                                </span>
                              ))}
                            </div>
                          ) : '—'}
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right' }}>
                          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 4 }}>
                            <div style={{ display: 'inline-flex', gap: 4, alignItems: 'center' }}>
                            <label style={{
                              background: isUploading ? '#f1f5f9' : '#fff', border: '1px solid #cbd5e1', borderRadius: 6,
                              padding: '3px 8px', fontSize: 11, fontWeight: 600,
                              color: isUploading ? '#94a3b8' : '#334155',
                              cursor: isUploading ? 'not-allowed' : 'pointer',
                              display: 'inline-flex', alignItems: 'center', gap: 4, whiteSpace: 'nowrap',
                            }}>
                              <input
                                type="file"
                                style={{ display: 'none' }}
                                disabled={isUploading}
                                onChange={async e => {
                                  const file = e.target.files?.[0];
                                  if (!file) return;
                                  host.setState({ docUploadRowKey: r.groupKey, docUploadBusy: true, docUploadMsg: null });
                                  try {
                                    const { fileId, statusPending, folderId, isGraphUpload } = await host._uploadFileToFolder(
                                      r.uploadFolderId,
                                      r.subFolderPath,
                                      r.vesselName,
                                      file,
                                      r.monthDriven
                                    );

                                    const msg = statusPending
                                      ? `"${file.name}" submitted for approval.`
                                      : `"${file.name}" uploaded successfully!`;

                                    if (statusPending) {
                                      const newApproval: ApprovalItem = {
                                        id: fileId || `a_${Date.now()}`,
                                        documentName: file.name,
                                        vessel: r.vesselName,
                                        requestedBy: host.props.userDisplayName || host.props.userEmail || 'You',
                                        requestedOn: new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }),
                                        status: 'Pending',
                                      };
                                      host.setState(prev => ({ approvalsList: [newApproval, ...prev.approvalsList] }));
                                    }

                                    const newUpload = { name: file.name, size: `${(file.size / 1024).toFixed(1)} KB`, date: 'Just now', pending: statusPending, id: fileId, uploadedAt: Date.now() };
                                    const liveFolderId = folderId || r.uploadFolderId;
                                    const vesselRecord = host.state.vessels.find(v => v.name === r.vesselName);
                                    const folderTargetStack = [
                                      { id: 'vessels_root', name: 'Vessels' },
                                      { id: 'specific_vessels', name: 'Specific Vessels' },
                                      { id: vesselRecord?.id || r.vesselName, name: r.vesselName },
                                      { id: `category_${r.category}`, name: r.category },
                                      ...(r.subCategory && r.subCategory !== r.category
                                        ? [{ id: liveFolderId, name: r.subCategory }]
                                        : []),
                                    ];

                                    host.setState(prev => {
                                      // Use functional setState to avoid stale closure on rows/uploadedFilesByFolder
                                      const uploadedFilesByFolder = prev.uploadedFilesByFolder;
                                      const rows = (prev.rows || []) as FlatRow[];
                                      const baseRows = rows.length > 0 ? rows : (prev.rows || []) as FlatRow[];

                                      const normSub = (r.subFolderPath || '').trim().toLowerCase();
                                      const dedupeKey = (r.subFolderPath || r.groupKey).trim().toLowerCase();
                                      const updatedByFolder: Record<string, any[]> = {
                                        ...uploadedFilesByFolder,
                                        [r.groupKey]: [...(uploadedFilesByFolder[r.groupKey] || []).filter((f: any) => f.name !== file.name), newUpload],
                                      };
                                      // Write to liveFolderId only if it's a real Graph drive item ID (isGraphUpload = true)
                                      // Do NOT write to r.uploadFolderId — it may be a backend DB ID shared across rows
                                      if (isGraphUpload === true && liveFolderId) {
                                        updatedByFolder[liveFolderId] = [...(uploadedFilesByFolder[liveFolderId] || []).filter((f: any) => f.name !== file.name), newUpload];
                                      }
                                      if (normSub) {
                                        updatedByFolder[normSub] = [...(uploadedFilesByFolder[normSub] || []).filter((f: any) => f.name !== file.name), newUpload];
                                      }
                                      if (dedupeKey && dedupeKey !== normSub) {
                                        updatedByFolder[dedupeKey] = [...(uploadedFilesByFolder[dedupeKey] || []).filter((f: any) => f.name !== file.name), newUpload];
                                      }

                                      // Find the matching row; if it already has a file, add a new row instead of overwriting
                                      const existingRowIdx = baseRows.findIndex((row: FlatRow) => row.groupKey === r.groupKey && !row.fileName);
                                      let updatedRows: FlatRow[];
                                      if (existingRowIdx !== -1) {
                                        updatedRows = baseRows.map((row: FlatRow, i: number) =>
                                          i === existingRowIdx
                                            ? { ...row, fileName: file.name, fileId: fileId || `file_${Date.now()}`, filePending: statusPending, fileUploadedAt: newUpload.uploadedAt }
                                            : row
                                        );
                                      } else {
                                        // All rows for this groupKey already have files — append a new row
                                        const refRow = baseRows.find((row: FlatRow) => row.groupKey === r.groupKey) || r;
                                        const newRow = { ...refRow, fileName: file.name, fileId: fileId || `file_${Date.now()}`, filePending: statusPending, fileUploadedAt: newUpload.uploadedAt };
                                        updatedRows = [...baseRows, newRow];
                                      }

                                      console.log('[VesselDMS] Upload optimistic update:', {
                                        groupKey: r.groupKey,
                                        uploadFolderId: r.uploadFolderId,
                                        liveFolderId,
                                        fileName: file.name,
                                        existingRowIdx,
                                        updatedRowsCount: updatedRows.length,
                                        rowsBefore: baseRows.length,
                                      });

                                      return {
                                        docUploadBusy: false,
                                        docUploadRowKey: null,
                                        docUploadMsg: msg,
                                        uploadedFilesByFolder: updatedByFolder,
                                        rows: updatedRows,
                                        folderPathStack: folderTargetStack,
                                        docMainFolder: r.group as 'Technical & Crewing' | 'Commercial & Chartering' | 'Insurance' | 'Kaizen - Knowledge Bank' | 'Knowledge Bank',
                                        docListPage: 0,
                                      };
                                    }, () => {
                                      // Call _refreshFolderFiles in setState callback to ensure state is updated first
                                      // Refresh the live folder, rather than replacing this
                                      // vessel's rows with the DB-only response. The latter can
                                      // temporarily hide files that were already in SharePoint.
                                      // Single refresh - _refreshFolderFiles now handles Graph consistency window
                                      // by preserving local uploads when Graph returns empty.
                                      const refreshFolderId = folderId || r.uploadFolderId;
                                      // Use isGraphUpload flag from _uploadFileToFolder return value
                                      // (true = folderId is a real Graph drive item ID, false = backend DB ID)
                                      const isRealGraphId = isGraphUpload === true;
                                      if (refreshFolderId && !/^f\d+$/.test(refreshFolderId)) {
                                        void host._refreshFolderFiles(refreshFolderId, r.groupKey, true, isRealGraphId).catch(() => undefined);
                                      }
                                    });
                                  } catch (err: any) {
                                    host.setState({ docUploadBusy: false, docUploadRowKey: null, docUploadMsg: `Upload failed: ${err?.message || 'Error'}` });
                                  }
                                }}
                              />
                              {isUploading ? '⏳...' : '↑ Upload'}
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
                                border: `1px solid ${rowSelectedCount > 0 ? '#fca5a5' : '#e2e8f0'}`,
                                background: rowSelectedCount > 0 ? '#fff5f5' : '#f8fafc',
                                borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 600,
                                cursor: rowSelectedCount > 0 ? 'pointer' : 'not-allowed',
                                color: rowSelectedCount > 0 ? '#ef4444' : '#cbd5e1',
                                display: 'inline-flex', alignItems: 'center', gap: 4, whiteSpace: 'nowrap',
                              }}
                              title={rowSelectedCount > 0 ? `Delete ${rowSelectedCount} selected file(s)` : 'Select files to delete'}
                            >
                              🗑{rowSelectedCount > 0 ? ` Delete (${rowSelectedCount})` : ' Delete'}
                            </button>
                            </div>
                            <button
                              type="button"
                              onClick={() => void host._openSharePointFolder(r)}
                              title="Open this folder in SharePoint"
                              aria-label={`Open ${r.subFolderPath} in SharePoint`}
                              style={{
                                width: 28, height: 27, padding: 0, borderRadius: 6,
                                border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8',
                                cursor: 'pointer', fontSize: 16, fontWeight: 700, lineHeight: 1,
                                display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                              }}
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

              {/* Pagination footer */}
              <div style={{ padding: '10px 14px', color: '#64748b', fontSize: 11, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#f8fafc', borderTop: '1px solid #e2e8f0' }}>
                <span>
                  Showing {filtered.length === 0 ? 0 : safePage * PAGE_ROWS + 1}–{Math.min((safePage + 1) * PAGE_ROWS, filtered.length)} of {filtered.length} rows
                </span>
                <div style={{ display: 'flex', gap: 3, alignItems: 'center' }}>
                  <button
                    onClick={() => host.setState({ docListPage: Math.max(0, safePage - 1) })}
                    disabled={safePage === 0}
                    style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 4, padding: '3px 8px', fontSize: 11, cursor: safePage === 0 ? 'not-allowed' : 'pointer', opacity: safePage === 0 ? 0.4 : 1 }}
                  >‹</button>
                  {Array.from({ length: totalPages }, (_, i) => (
                    <button
                      key={i}
                      onClick={() => host.setState({ docListPage: i })}
                      style={{
                        border: i === safePage ? 'none' : '1px solid #cbd5e1',
                        background: i === safePage ? '#0078d4' : '#fff',
                        color: i === safePage ? '#fff' : '#334155',
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
                    style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 4, padding: '3px 8px', fontSize: 11, cursor: safePage >= totalPages - 1 ? 'not-allowed' : 'pointer', opacity: safePage >= totalPages - 1 ? 0.4 : 1 }}
                  >›</button>
                </div>
              </div>
            </div>
          </div>
        )}
        {/* Classify Dialog Modal */}
        {renderClassifyDialog(host)}
      </div>
    );



}
