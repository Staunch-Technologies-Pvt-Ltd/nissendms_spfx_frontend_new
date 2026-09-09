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

export function renderDocumentsPage(host: VesselEmail): React.ReactElement {

    const {
      textFilter, vesselFilter, catFilter, attachmentFilter, docViewMode, showAllVesselsInFolderView,
      vessels, rows, docListPage, docListSort, docGroupFilter, docCategoryFilter, docGroupLevelFilter, docLeafCategoryFilter, docSubCategoryFilter,
      docScopeType,
      documentVesselCount, documentVesselsLoadingMore, documentFilesLoading, vesselLoadingName,
      docUploadRowKey, docUploadBusy, docUploadMsg, documentsList,
      folderPathStack, uploadedFilesByFolder, docMainFolder,
      folderNavHistory, folderNavIndex, listViewSelectedFiles, folderViewSelectedFiles,
      windowWidth,
    } = host.state;

    const vesselColumns = windowWidth <= 767 ? 1 : windowWidth <= 1024 ? 2 : 4;

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
    const atKaizenRoot = stackLevel >= 1 && (folderPathStack[0]?.id === 'kaizen_root' || folderPathStack[0]?.name === 'Kaizen - Knowledge Bank');
    const atMainDepartment = stackLevel >= 1 && !atKaizenRoot && MAIN_FOLDERS.some(mf => mf.key === folderPathStack[0]?.name);
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
    const vesselStackIdx = vesselNodeInStack ? 1 : -1;
    const currentVesselNameFromStack = atKaizenRoot
      ? 'Kaizen - Knowledge Bank'
      : (atCommonShips ? 'Common for all vessels' : (vesselNodeInStack?.name || null));

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

    if (currentVesselNameFromStack && !host._filesLoadedForVessels.has(currentVesselNameFromStack)) {
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

    // Determine subfolders for current depth
    const mainsDefaultsSource = atCommonShips ? COMMON_DEFAULT_MAINS : (atKaizenRoot ? folderNamesByMainFolder() : DEFAULT_VESSEL_MAINS);
    const scopedKey = docMainFolder && currentFolderName ? `${docMainFolder} > ${currentFolderName}` : null;
    const isAtCategoryLevel =
      (atKaizenRoot && stackLevel === 1) ||
      (atCommonShips && stackLevel === 2) ||
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

        // If key is a groupKey or path, ensure it strictly belongs to current vessel and department
        if (vesselName) {
          const vNorm = vesselName.toLowerCase();
          const kNorm = key.toLowerCase();
          if (!kNorm.includes(vNorm)) return;
        }
        if (docMainFolder) {
          const dNorm = docMainFolder.toLowerCase();
          const kNorm = key.toLowerCase();
          if (!kNorm.includes(dNorm)) return;
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
      ? allCurrentFolderFiles.filter(file => file.name.toLowerCase().includes(textFilter.trim().toLowerCase()))
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

    const allRows: FlatRow[] = [
      ...effectiveVesselRows,
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

    // Scope rows for dropdown populating
    const scopeVesselRows = activeVesselName
      ? scopeRows.filter(r => r.vesselName.trim().toLowerCase() === activeVesselName.trim().toLowerCase())
      : scopeRows;

    const allGroups = Array.from(new Set(scopeVesselRows.map(r => r.group))).sort();
    const filteredCategoryRows = scopeVesselRows.filter(r =>
      docGroupFilter === 'all' || r.group.trim().toLowerCase() === docGroupFilter.trim().toLowerCase()
    );
    const hierarchyForRow = (row: FlatRow): { section: string; group: string; category: string; subCategory: string } => {
      const parts = (row.subFolderPath || '').split('>').map(part => part.trim()).filter(Boolean);
      const knownContainers = new Set([
        'vessels', 'specific vessels', 'documents', 'shared documents',
        'common for all ships', 'common', 'kaizen - knowledge bank', 'kaizen',
        (row.vesselName || '').trim().toLowerCase(),
        (row.group || '').trim().toLowerCase(),
        'technical & crewing', 'commercial & chartering', 'insurance', 'knowledge bank'
      ]);

      const subLevels = parts.filter(p => !knownContainers.has(p.toLowerCase()));

      const section = subLevels[0] || (row.category && !knownContainers.has(row.category.toLowerCase()) ? row.category : '') || '';
      const group = subLevels[1] || (row.category && row.category !== section && !knownContainers.has(row.category.toLowerCase()) ? row.category : '') || '';
      const category = subLevels[2] || (row.subCategory && row.subCategory !== section && row.subCategory !== group ? row.subCategory : '') || '';
      const subCategory = subLevels[3] || '';

      return { section, group, category, subCategory };
    };
    const hierarchyRows = filteredCategoryRows.map(row => ({ row, levels: hierarchyForRow(row) }));
    const documentSectionOptions = Array.from(new Set(
      hierarchyRows
        .map(item => item.levels.section)
        .filter(s => s && !allGroups.some(g => g.toLowerCase() === s.toLowerCase()))
    )).sort();
    const groupLevelRows = hierarchyRows.filter(item =>
      docCategoryFilter === 'all' || item.levels.section.trim().toLowerCase() === docCategoryFilter.trim().toLowerCase()
    );
    const groupLevelOptions = Array.from(new Set(groupLevelRows.map(item => item.levels.group).filter(Boolean))).sort();
    const categoryRows = groupLevelRows.filter(item =>
      docGroupLevelFilter === 'all' || item.levels.group.trim().toLowerCase() === docGroupLevelFilter.trim().toLowerCase()
    );
    const categoryOptions = Array.from(new Set(categoryRows.map(item => item.levels.category).filter(Boolean))).sort();
    const subCategoryRows = categoryRows.filter(item =>
      docLeafCategoryFilter === 'all' || item.levels.category.trim().toLowerCase() === docLeafCategoryFilter.trim().toLowerCase()
    );
    const subCategoryOptions = Array.from(new Set(subCategoryRows.map(item => item.levels.subCategory).filter(Boolean))).sort();
    const vesselFilterOptions = mainFolderPage && docMainFolder
      ? vessels.filter(v => scopeRows.some(r =>
          r.vesselName.trim().toLowerCase() === v.name.trim().toLowerCase() &&
          r.group.trim().toLowerCase() === docMainFolder.trim().toLowerCase()
        ))
      : vessels;

    // ── mainFolderGroupMap: which groups belong to which main folder (for list view filtering) ──
    const mainFolderGroupMap = folderNamesByMainFolder(docScopeType === 'common');

    // Vessels are returned newest first. Start with four and extend in pages
    // of eight when the user chooses "More vessels".
    const visibleVesselNames = new Set(vessels.slice(0, documentVesselCount).map(v => v.name));

    const scopedVesselInFolderView =
      docViewMode === 'folder' && docScopeType === 'vessels' && currentVesselNameFromStack
        ? currentVesselNameFromStack.trim().toLowerCase()
        : null;

    let filtered = scopeRows.filter(r => {
      // 1. Vessel filter
      if (docScopeType === 'vessels') {
        // When user has drilled into a vessel in Folder view, always scope to
        // that vessel even if the top dropdown still says "all".
        if (scopedVesselInFolderView) {
          if ((r.vesselName || '').trim().toLowerCase() !== scopedVesselInFolderView) return false;
        } else if (!activeVesselName) {
          if (!visibleVesselNames.has(r.vesselName)) return false;
        } else if (r.vesselName.trim().toLowerCase() !== activeVesselName.trim().toLowerCase()) {
          return false;
        }
      }

      // 2. Main folder / group filter
      const effectiveGroupFilter = docGroupFilter !== 'all' ? docGroupFilter : (docViewMode === 'folder' && docMainFolder ? docMainFolder : null);
      if (effectiveGroupFilter && effectiveGroupFilter !== 'all') {
        const normGroup = (r.group || '').trim().toLowerCase();
        const normFilter = effectiveGroupFilter.trim().toLowerCase();
        const groupMatch = normGroup === normFilter ||
          normGroup.includes(normFilter.split(' ')[0]) ||
          (r.subFolderPath || '').toLowerCase().includes(normFilter.split(' ')[0]);
        if (!groupMatch) return false;
      }

      // 3. Category / Document section filter
      const levels = hierarchyForRow(r);
      const effectiveSectionFilter = docCategoryFilter !== 'all' ? docCategoryFilter : (catFilter !== 'all' ? catFilter : 'all');
      if (effectiveSectionFilter !== 'all') {
        const normFilter = effectiveSectionFilter.trim().toLowerCase();
        const sectionMatch = (levels.section || '').trim().toLowerCase() === normFilter ||
          (r.category || '').trim().toLowerCase() === normFilter ||
          (r.subFolderPath || '').toLowerCase().includes(`> ${normFilter}`) ||
          (r.subFolderPath || '').toLowerCase().includes(`${normFilter} >`);
        if (!sectionMatch) return false;
      }
      if (docGroupLevelFilter !== 'all') {
        const normFilter = docGroupLevelFilter.trim().toLowerCase();
        const groupMatch = (levels.group || '').trim().toLowerCase() === normFilter ||
          (r.subFolderPath || '').toLowerCase().includes(`> ${normFilter}`);
        if (!groupMatch) return false;
      }
      if (docLeafCategoryFilter !== 'all') {
        const normFilter = docLeafCategoryFilter.trim().toLowerCase();
        const catMatch = (levels.category || '').trim().toLowerCase() === normFilter ||
          (r.category || '').trim().toLowerCase() === normFilter ||
          (r.subCategory || '').trim().toLowerCase() === normFilter ||
          (r.subFolderPath || '').toLowerCase().includes(`> ${normFilter}`);
        if (!catMatch) return false;
      }
      if (docSubCategoryFilter !== 'all') {
        const normFilter = docSubCategoryFilter.trim().toLowerCase();
        const subMatch = (levels.subCategory || '').trim().toLowerCase() === normFilter ||
          (r.subCategory || '').trim().toLowerCase() === normFilter ||
          (r.subFolderPath || '').toLowerCase().includes(`> ${normFilter}`);
        if (!subMatch) return false;
      }

      // 4. Text search filter
      if (textFilter) {
        const q = textFilter.trim().toLowerCase();
        return (r.vesselName || '').toLowerCase().includes(q) ||
          (r.group || '').toLowerCase().includes(q) ||
          (r.category || '').toLowerCase().includes(q) ||
          (r.subCategory || '').toLowerCase().includes(q) ||
          (r.subFolderPath || '').toLowerCase().includes(q) ||
          (r.fileName || '').toLowerCase().includes(q);
      }
      return true;
    });

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

      if (vesselName) {
        const vNorm = vesselName.toLowerCase();
        if (!(key || '').toLowerCase().includes(vNorm)) return;
      }
      if (docMainFolder) {
        const dNorm = docMainFolder.toLowerCase();
        if (!(key || '').toLowerCase().includes(dNorm)) return;
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
        } else {
          if (!subfolderFileCountMap.has(targetSfKey)) subfolderFileCountMap.set(targetSfKey, 0);
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
      const dedupeKey = canonicalRowKey || (r.subFolderPath || r.groupKey).trim().toLowerCase();
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
      let nextScope: 'vessels' | 'common' | 'kaizen' = 'vessels';
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
      } else {
        host.setState({ docViewMode: 'folder' });
      }
    };

    const openFolderViewFromListContext = (): void => {
      if (docViewMode !== 'list') {
        host.setState({ docViewMode: 'folder' });
        return;
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

      let nextScope: 'vessels' | 'common' | 'kaizen' = host.state.docScopeType || 'vessels';
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

      host.setState({
        docViewMode: 'list',
        docScopeType: nextScope,
        vesselFilter: nextVesselFilter || 'all',
        docGroupFilter: nextGroupFilter || 'all',
        docCategoryFilter: level0,
        docGroupLevelFilter: level1,
        docLeafCategoryFilter: level2,
        docSubCategoryFilter: level3,
        catFilter: 'all',
        textFilter: '',
        docListPage: 0,
      });
    };

    // In List View: trigger on-demand live file refresh for visible page rows if not loaded yet
    if (docViewMode === 'list') {
      pageGroupedRows.forEach(row => {
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
    // Deduplicate by name and exclude any vessel whose name matches a navigation folder
    const displayVessels = (() => {
      const baseList = vessels.filter(
        (v, i, arr) =>
          arr.findIndex(x => x.name.trim().toLowerCase() === v.name.trim().toLowerCase()) === i &&
          !NAV_FOLDER_NAMES.has(v.name.trim().toLowerCase())
      );
      if (vesselFilter !== 'all') {
        const matching = baseList.filter(v => v.name.trim().toLowerCase() === vesselFilter.trim().toLowerCase());
        if (matching.length > 0) return matching;
        return [{ id: vesselFilter, name: vesselFilter }];
      }
      return baseList.slice(0, documentVesselCount);
    })();

    const handleDocumentsPageDrop = async (e: React.DragEvent): Promise<void> => {
      e.preventDefault();
      e.stopPropagation();
      if (!e.dataTransfer) return;
      const extracted = await extractFilesFromDataTransfer(e.dataTransfer);
      if (!extracted.length) return;

      const topFolderId = currentFolderNode && !/^(sf_|category_|common|vessels_root|specific_vessels|kaizen_root)/.test(currentFolderNode.id) ? currentFolderNode.id : '';
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
      const resolvedFolderId = topFolderId || liveId || (matchingRow?.uploadFolderId) || fallbackPath;

      host._openBulkUpload(extracted, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode);
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
        {/* Breadcrumb Navigation Trail */}
        <div style={{ fontSize: 12, color: '#64748b', display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
          {/* Back / Forward navigation buttons */}
          <button
            onClick={goBack}
            disabled={!canGoBack}
            aria-label="Go back"
            title="Go back (Left Arrow)"
            className="dms-nav-arrow"
            style={{
              width: 34, height: 34, borderRadius: 10, border: canGoBack ? '1px solid #38bdf8' : '1px solid #cbd5e1',
              background: canGoBack ? 'linear-gradient(135deg, #0ea5e9, #2563eb)' : '#f1f5f9',
              color: canGoBack ? '#fff' : '#cbd5e1',
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
              width: 34, height: 34, borderRadius: 10, border: canGoForward ? '1px solid #c084fc' : '1px solid #cbd5e1',
              background: canGoForward ? 'linear-gradient(135deg, #8b5cf6, #ec4899)' : '#f1f5f9',
              color: canGoForward ? '#fff' : '#cbd5e1',
              cursor: canGoForward ? 'pointer' : 'not-allowed',
              fontSize: 20, fontWeight: 800, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              transition: 'transform 160ms ease, filter 160ms ease, box-shadow 160ms ease',
              flexShrink: 0,
            }}
          >→</button>
          <span style={{ cursor: 'pointer', color: '#0284c7' }} onClick={() => host._pushFolderNav([], null)}>Home</span>
          <span>›</span>
          <span
            style={{ cursor: stackLevel === 0 && !docMainFolder ? 'default' : 'pointer', color: stackLevel === 0 && !docMainFolder ? '#0f172a' : '#0284c7', fontWeight: stackLevel === 0 && !docMainFolder ? 600 : 400 }}
            onClick={() => host._pushFolderNav([], null)}
          >
            Documents
          </span>
         {folderPathStack.map((item, idx) => {
            const isLast = idx === folderPathStack.length - 1;
            return (
              <React.Fragment key={item.id + idx}>
                <span>›</span>
                <span
                  onClick={() => {
                    if (atKaizenRoot && idx === 0) {
                      host._pushFolderNav(folderPathStack.slice(0, 1), 'Kaizen - Knowledge Bank');
                    } else if (idx === 0) {
                      const mainFolder = item.name as MainFolderKey;
                      host._pushFolderNav(folderPathStack.slice(0, 1), mainFolder);
                      if (mainFolder === 'Technical & Crewing' || mainFolder === 'Commercial & Chartering' || mainFolder === 'Insurance') {
                        host.setState({
                          vesselFilter: 'all',
                          docScopeType: 'vessels',
                          docCategoryFilter: 'all',
                          docGroupLevelFilter: 'all',
                          docLeafCategoryFilter: 'all',
                          docSubCategoryFilter: 'all',
                          catFilter: 'all',
                          docListPage: 0,
                        });
                      }
                    } else if (idx === 1 && atMainDepartment && !atCommonShips) {
                      host._pushFolderNav(folderPathStack.slice(0, idx + 1), docMainFolder);
                      host.setState({
                        vesselFilter: item.name,
                        docScopeType: 'vessels',
                        docCategoryFilter: 'all',
                        docGroupLevelFilter: 'all',
                        docLeafCategoryFilter: 'all',
                        docSubCategoryFilter: 'all',
                        catFilter: 'all',
                        docListPage: 0,
                      });
                    } else {
                      host._pushFolderNav(folderPathStack.slice(0, idx + 1), docMainFolder);
                    }
                  }}
                  style={{ cursor: isLast ? 'default' : 'pointer', color: isLast ? '#0f172a' : '#0284c7', fontWeight: isLast ? 600 : 400 }}
                >
                  {item.name}
                </span>
              </React.Fragment>
            );
          })}
        </div>

        {/* Module Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 20 }}>📁</span>
              {currentFolderNode ? currentFolderNode.name : 'Documents'}
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>
              {docViewMode === 'list'
                ? `${filtered.length} rows · flattened list view`
                : stackLevel === 0
                  ? `${MAIN_FOLDERS.length} main folders`
                  : atDepartmentVesselList
                    ? `${vessels.length} vessels · Common for all ships`
                    : `${subfolderNames.length} sections · ${allCurrentFolderFiles.length} files`}
            </p>
            {docViewMode === 'folder' && currentFolderNode && subfolderNames.length === 0 && (
              <div style={{ marginTop: 6, fontSize: 12, fontWeight: 600, color: allCurrentFolderFiles.length > 0 ? '#15803d' : '#64748b' }}>
                {allCurrentFolderFiles.length > 0 ? '✅ Attached' : '⚪ Not Attached'}
              </div>
            )}
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            {/* Folder view / List view Pill Toggle */}
            <div style={{ display: 'inline-flex', background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8, padding: 3, gap: 2 }}>
              <button
                onClick={openFolderViewFromListContext}
                style={{
                  padding: '5px 12px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                  background: docViewMode === 'folder' ? '#0f172a' : 'transparent',
                  color: docViewMode === 'folder' ? '#fff' : '#64748b',
                }}
              >
                ⊞ Folder view
              </button>
              <button
                onClick={openListViewFromFolderContext}
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
              disabled={(docViewMode === 'list' ? listViewSelectedFiles.size : folderViewSelectedFiles.size) === 0}
              onClick={() => {
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
                background: (docViewMode === 'list' ? listViewSelectedFiles.size : folderViewSelectedFiles.size) > 0 ? '#d97706' : '#f1f5f9',
                color: (docViewMode === 'list' ? listViewSelectedFiles.size : folderViewSelectedFiles.size) > 0 ? '#fff' : '#94a3b8',
                border: 'none', borderRadius: 8,
                padding: '7px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6,
              }}
            >
              📦 Archive{((docViewMode === 'list' ? listViewSelectedFiles.size : folderViewSelectedFiles.size) > 0) ? ` (${docViewMode === 'list' ? listViewSelectedFiles.size : folderViewSelectedFiles.size})` : ''}
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

            {/* Top-Right Action Buttons — Always available in Folder View and List View (Root, Department, Vessel, Category, Subfolder) */}
            {(() => {
              const resolveDocPageUploadTarget = (): {
                currentVessel: string;
                subFolderPath: string;
                fallbackPath: string;
                resolvedFolderId: string;
              } => {
                const topFolderId = currentFolderNode && !/^(sf_|category_|common|vessels_root|specific_vessels|kaizen_root)/.test(currentFolderNode.id) ? currentFolderNode.id : '';
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

                return { currentVessel, subFolderPath, fallbackPath, resolvedFolderId };
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
                        const { currentVessel, subFolderPath, resolvedFolderId } = resolveDocPageUploadTarget();
                        const bulkFiles = filesList.map(f => ({ file: f, relativePath: f.name }));
                        host._openBulkUpload(bulkFiles, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode);
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
                        const { currentVessel, subFolderPath, resolvedFolderId } = resolveDocPageUploadTarget();
                        const bulkFiles = filesList.map(f => ({
                          file: f,
                          relativePath: (f as any).webkitRelativePath || f.name,
                        }));
                        host._openBulkUpload(bulkFiles, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode);
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
        </div>

        {/* ── Filter Toolbar ── */}
        <div style={{ background: '#fff', borderRadius: 10, padding: '10px 12px', border: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', gap: 10 }}>
          {/* 3 Scope Options: 1st Option = Specific Vessels, 2nd Option = Common for all vessels, 3rd Option = Kaizen - Knowledge Bank */}
          {docViewMode === 'list' && (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 8, paddingBottom: 6, borderBottom: '1px solid #f1f5f9' }}>
              <div style={{ display: 'inline-flex', background: '#f8fafc', border: '1px solid #cbd5e1', borderRadius: 8, padding: 3, gap: 4 }}>
                <button
                  onClick={() => host.setState({ docScopeType: 'vessels', vesselFilter: 'all', docGroupFilter: 'all', docCategoryFilter: 'all', docGroupLevelFilter: 'all', docLeafCategoryFilter: 'all', docSubCategoryFilter: 'all', catFilter: 'all', docListPage: 0, folderPathStack: [], docMainFolder: null })}
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
                  onClick={() => host.setState({ docScopeType: 'common', vesselFilter: 'all', docGroupFilter: 'all', docCategoryFilter: 'all', docGroupLevelFilter: 'all', docLeafCategoryFilter: 'all', docSubCategoryFilter: 'all', catFilter: 'all', docListPage: 0, folderPathStack: [], docMainFolder: null })}
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
                  onClick={() => {
                    host.setState({
                      docScopeType: 'kaizen',
                      vesselFilter: 'all',
                      docGroupFilter: 'all',
                      docCategoryFilter: 'all',
                      docGroupLevelFilter: 'all',
                      docLeafCategoryFilter: 'all',
                      docSubCategoryFilter: 'all',
                      catFilter: 'all',
                      docListPage: 0,
                      folderPathStack: [{ id: 'kaizen_root', name: 'Kaizen - Knowledge Bank' }],
                      docMainFolder: 'Kaizen - Knowledge Bank',
                    });
                    void host._mergeLiveSharePointFiles(['Kaizen - Knowledge Bank']).catch(() => undefined);
                  }}
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
                placeholder="Filter by vessel, main folder, document section, path..."
                value={textFilter}
                onChange={e => host.setState({ textFilter: e.target.value, docListPage: 0 })}
                style={{ width: '100%', padding: '6px 10px 6px 28px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, outline: 'none', boxSizing: 'border-box' }}
              />
            </div>
            {docScopeType !== 'kaizen' && (
              <select
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
                style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 160 }}
              >
                <option value="all">All main folders</option>
                {allGroups.map(g => (
                  <option key={g} value={g}>{g}</option>
                ))}
              </select>
            )}
            {docScopeType === 'vessels' && (mainFolderPage || docViewMode === 'list') && (
              <select
                value={mainFolderPage ? (vesselFilter === 'all' ? '' : vesselFilter) : vesselFilter}
                onChange={e => {
                  const val = e.target.value;
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
                  host.setState({ vesselFilter: val || 'all', docListPage: 0 });
                  if (val !== 'all') {
                    void host._loadVesselRowsFromApi(val).catch(() => undefined);
                    void host._loadFilesForVessel(val).catch(() => undefined);
                  }
                }}
                style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 140 }}
              >
                {!mainFolderPage && <option value="all">All vessels ({Math.min(documentVesselCount, vessels.length)})</option>}
                {mainFolderPage && <option value="">Select vessel</option>}
                {vesselFilterOptions.map(v => (
                  <option key={v.id || v.name} value={v.name}>{v.name}</option>
                ))}
              </select>
            )}
            <select
              value={docCategoryFilter}
              onChange={e => host.setState({ docCategoryFilter: e.target.value, docGroupLevelFilter: 'all', docLeafCategoryFilter: 'all', docSubCategoryFilter: 'all', catFilter: e.target.value, docListPage: 0 })}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 180 }}
            >
              <option value="all">All document sections</option>
              {documentSectionOptions.map(section => <option key={section} value={section}>{section}</option>)}
            </select>
            <select
              value={docGroupLevelFilter}
              onChange={e => host.setState({ docGroupLevelFilter: e.target.value, docLeafCategoryFilter: 'all', docSubCategoryFilter: 'all', docListPage: 0 })}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 180 }}
            >
              <option value="all">All groups</option>
              {groupLevelOptions.map(group => <option key={group} value={group}>{group}</option>)}
            </select>
            <select
              value={docLeafCategoryFilter}
              onChange={e => host.setState({ docLeafCategoryFilter: e.target.value, docSubCategoryFilter: 'all', docListPage: 0 })}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 180 }}
            >
              <option value="all">All categories</option>
              {categoryOptions.map(category => <option key={category} value={category}>{category}</option>)}
            </select>
            <select
              value={docSubCategoryFilter}
              onChange={e => host.setState({ docSubCategoryFilter: e.target.value, docListPage: 0 })}
              style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 180 }}
            >
              <option value="all">All sub-categories</option>
              {subCategoryOptions.map(subCategory => <option key={subCategory} value={subCategory}>{subCategory}</option>)}
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

            {(documentFilesLoading || vesselLoadingName || documentVesselsLoadingMore) && (
              <div style={{
                display: 'inline-flex', alignItems: 'center', gap: 6, padding: '4px 10px',
                borderRadius: 8, background: '#eff6ff', border: '1px solid #bfdbfe',
                fontSize: 11, color: '#1d4ed8', fontWeight: 600,
              }}>
                <span style={{ display: 'inline-block', animation: 'spin 1s linear infinite' }}>⏳</span>
                <span>{vesselLoadingName ? `Loading ${vesselLoadingName}...` : (documentVesselsLoadingMore ? 'Loading vessels...' : 'Syncing live files...')}</span>
              </div>
            )}
            <div style={{ display: 'flex', border: '1px solid #cbd5e1', borderRadius: 8, overflow: 'hidden', marginLeft: 'auto' }}>
              <button
                onClick={openFolderViewFromListContext}
                style={{ padding: '5px 10px', background: docViewMode === 'folder' ? '#e2e8f0' : '#fff', border: 'none', cursor: 'pointer', fontSize: 13 }}
                title="Folder view"
              >::</button>
              <button
                onClick={openListViewFromFolderContext}
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

            {/* Level 0: Main Departments + Kaizen - Knowledge Bank */}
            {stackLevel === 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                  {MAIN_FOLDERS.filter(item => {
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
                          host.setState({ vesselFilter: 'all', docScopeType: 'kaizen' });
                        } else {
                          host._pushFolderNav([{ id: item.key, name: item.key }], item.key);
                          host.setState({ vesselFilter: 'all', docScopeType: 'vessels' });
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
                        <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.key}</div>
                        <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>
                          {item.key === 'Kaizen - Knowledge Bank' ? 'Knowledge base' : `${vessels.length} vessels`}
                        </div>
                      </div>
                      <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                    </div>
                  ))}
                </div>
                {(textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || docCategoryFilter !== 'all' || docGroupLevelFilter !== 'all' || docLeafCategoryFilter !== 'all' || docSubCategoryFilter !== 'all' || attachmentFilter !== 'all') && (
                  <div style={{ marginTop: 24, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 12, overflow: 'hidden' }}>
                    <div style={{ padding: '14px 16px', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12 }}>
                      <strong style={{ color: '#0f172a', fontSize: 14 }}>Filtered document results</strong>
                      <span style={{ color: '#64748b', fontSize: 12 }}>{groupedList.length} folder{groupedList.length === 1 ? '' : 's'}</span>
                    </div>
                    {groupedList.length === 0 ? (
                      <div style={{ padding: 28, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No documents match the selected filters.</div>
                    ) : (
                      <div style={{ display: 'flex', flexDirection: 'column' }}>
                        {groupedList.slice(0, PAGE_ROWS).map(result => (
                          <button
                            key={result.groupKey}
                            type="button"
                            onClick={() => host._pushFolderNav([{ id: result.groupKey, name: result.subFolderPath.split(' > ').pop() || result.subCategory }], result.group as 'Technical & Crewing' | 'Commercial & Chartering' | 'Insurance' | 'Kaizen - Knowledge Bank' | 'Knowledge Bank')}
                            style={{ border: 0, borderBottom: '1px solid #f1f5f9', background: '#fff', padding: '12px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 16, textAlign: 'left', cursor: 'pointer' }}
                          >
                            <span style={{ minWidth: 0 }}>
                              <strong style={{ display: 'block', color: '#0f172a', fontSize: 13 }}>{result.subFolderPath.split(' > ').pop() || result.subCategory}</strong>
                              <span style={{ display: 'block', color: '#64748b', fontSize: 11, marginTop: 3 }}>{result.vesselName} · {result.group} · {result.category}</span>
                            </span>
                            <span style={{ color: '#0284c7', fontSize: 12, whiteSpace: 'nowrap' }}>{result.files.length} file{result.files.length === 1 ? '' : 's'} ›</span>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>

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
                            background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18,
                            display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                            boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                          }}
                        >
                          <div style={{ width: 44, height: 44, borderRadius: 10, background: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>📁</div>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{commonDisplayName}</div>
                            <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>Shared {docMainFolder} documents</div>
                          </div>
                          <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
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
                    .map(v => (
                    <div
                      key={v.id}
                      onClick={() => {
                        host._pushFolderNav([...folderPathStack, { id: v.id, name: v.name }], docMainFolder);
                        host.setState({ vesselFilter: v.name, docScopeType: 'vessels' });
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

                    {documentVesselCount < vessels.length && (
                      <div
                        onClick={() => {
                          if (!documentVesselsLoadingMore) {
                            void host._loadMoreDocumentVessels();
                          }
                        }}
                        style={{
                          background: documentVesselsLoadingMore ? '#f8fafc' : '#f0f9ff',
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
                          <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>
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
                        <div style={{ fontSize: 12, color: '#64748b', marginTop: 3, display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                          <span>{isAtCategoryLevel ? 'Document Section' : (stackLevel === 3 ? 'Category' : 'Sub-Category')}</span>
                          {(() => {
                            const fc = subfolderFileCountMap.get(sfName) ?? subfolderFileCountMap.get(sfName.trim().toLowerCase()) ?? 0;
                            const hasLiveData = resolvedCurrentFolderId && host.state.spoFolderMap.has(resolvedCurrentFolderId);
                            const badgeColor = fc > 0 ? '#0284c7' : (hasLiveData ? '#94a3b8' : '#cbd5e1');
                            return (
                              <span style={{
                                display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                background: badgeColor, color: '#fff',
                                borderRadius: 20, padding: '1px 8px',
                                fontSize: 10, fontWeight: 700, lineHeight: '16px',
                                whiteSpace: 'nowrap',
                              }}>
                                {fc} {fc === 1 ? 'file' : 'files'}
                              </span>
                            );
                          })()}
                        </div>
                        {showDebugKeys && (
                          <div style={{ marginTop: 4, fontSize: 10, color: '#7c3aed', fontFamily: 'monospace', wordBreak: 'break-all' }}>
                            key: {fullBreadcrumbWithSf} | live:{folderDebugLiveId}
                          </div>
                        )}
                      </div>
                      <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                    </div>
                    );
                  })}
                </div>

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
                      const currentFolderSharePointPath = atKaizenRoot
                        ? ['Kaizen - Knowledge Bank', ...folderPathStack.slice(1).map(n => n.name)].join(' > ')
                        : (atCommonShips
                          ? ['Common for all vessels', ...(docMainFolder ? [docMainFolder] : []), ...folderPathStack.slice(1).map(n => n.name)].join(' > ')
                          : (currentVesselNameFromStack && docMainFolder
                            ? [currentVesselNameFromStack, docMainFolder, ...folderPathStack.slice(2).map(n => n.name)].filter(Boolean).join(' > ')
                            : (currentFolderNode?.name || '')));
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
                          <td style={{ padding: '12px 16px', fontWeight: 600, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
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
                                  const currentPath = atKaizenRoot
                                    ? ['Kaizen - Knowledge Bank', ...folderPathStack.slice(1).map(n => n.name)].join(' > ')
                                    : (atCommonShips
                                      ? ['Common for all vessels', ...(docMainFolder ? [docMainFolder] : []), ...folderPathStack.slice(1).map(n => n.name)].join(' > ')
                                      : (currentVesselNameFromStack && docMainFolder
                                        ? [currentVesselNameFromStack, docMainFolder, ...folderPathStack.slice(2).map(n => n.name)].filter(Boolean).join(' > ')
                                        : (currentFolderNode?.name || '')));
                                  void host._openDocumentFile(fileId, file.name, currentPath);
                                }}
                                style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#0078d4' }}
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
                    <th style={{ padding: '10px 12px' }}>MAIN FOLDER</th>
                    <th style={{ padding: '10px 12px' }}>DOCUMENT SECTION</th>
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
                      <td colSpan={11} style={{ padding: '48px 16px', textAlign: 'center' }}>
                        {(host.state.loading || Boolean(vesselLoadingName) || (documentFilesLoading && filtered.length === 0) || documentVesselsLoadingMore) ? (
                          <div style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
                            <div style={{
                              width: 32, height: 32, border: '3px solid #e0f2fe',
                              borderTop: '3px solid #0284c7', borderRadius: '50%',
                              animation: 'spin 0.8s linear infinite',
                            }} />
                            <div style={{ fontWeight: 700, fontSize: 13, color: '#0f172a' }}>
                              {vesselLoadingName
                                ? `Loading documents and attachments for ${vesselLoadingName}...`
                                : (documentVesselsLoadingMore ? 'Loading more vessels...' : 'Loading vessel documents from SharePoint...')}
                            </div>
                            <div style={{ fontSize: 11, color: '#64748b', maxWidth: 360 }}>
                              Please wait while folder structures and live files are loaded.
                            </div>
                          </div>
                        ) : (
                          <span style={{ color: '#94a3b8', fontSize: 13 }}>
                            No documents found. {textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || docCategoryFilter !== 'all' || docGroupLevelFilter !== 'all' || docLeafCategoryFilter !== 'all' || docSubCategoryFilter !== 'all' ? 'Try clearing the filters.' : ''}
                          </span>
                        )}
                      </td>
                    </tr>
                  ) : pageGroupedRows.map((r, idx) => {
                    const globalIdx = safePage * PAGE_ROWS + idx + 1;
                    const isUploading = docUploadRowKey === r.groupKey && docUploadBusy;
                    const hasFiles = r.files.length > 0;
                    const rowFileIds = r.files.map(f => f.id);
                    const rowSelectedCount = rowFileIds.filter(id => listViewSelectedFiles.has(id)).length;
                    const rowTail = getFolderTailSegments(r.subFolderPath, r.vesselName, r.group);
                    const documentSectionLabel = rowTail[0] || 'Drawings and Manuals';
                    const categoryLabel = r.category || rowTail[2] || rowTail[1] || 'To be Classified';

                    return (
                      <tr key={`${r.groupKey}-${idx}`} style={{ borderBottom: '1px solid #f1f5f9', transition: 'background 0.1s' }}
                        onMouseEnter={e => (e.currentTarget.style.background = '#f8fafc')}
                        onMouseLeave={e => (e.currentTarget.style.background = '')}
                      >
                        <td style={{ padding: '10px 10px', color: '#94a3b8', fontSize: 11, fontFamily: 'monospace', textAlign: 'center', verticalAlign: 'top' }}>{globalIdx}</td>
                        <td style={{ padding: '10px 12px', fontWeight: 700, color: '#0f172a', verticalAlign: 'top' }}>
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
                        <td style={{ padding: '10px 12px', verticalAlign: 'top' }}>
                          <span style={{
                            display: 'inline-block', borderRadius: 8, padding: '3px 8px', fontSize: 11, fontWeight: 600,
                            background: r.group === 'Kaizen - Knowledge Bank' ? '#ede9fe' : (r.group === 'Insurance' ? '#fef3c7' : (r.group === 'Commercial & Chartering' ? '#dcfce7' : '#eff6ff')),
                            color: r.group === 'Kaizen - Knowledge Bank' ? '#6b21a8' : (r.group === 'Insurance' ? '#b45309' : (r.group === 'Commercial & Chartering' ? '#15803d' : '#2563eb')),
                          }}>
                            {r.group}
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', fontWeight: 600, color: '#334155', verticalAlign: 'top' }}>{documentSectionLabel}</td>
                        <td style={{ padding: '10px 12px', fontWeight: 700, color: '#1e293b', verticalAlign: 'top' }}>{categoryLabel}</td>
                        <td style={{ padding: '10px 12px', fontWeight: 600, color: '#1e293b', verticalAlign: 'top' }}>{r.subCategory || r.category}</td>
                        <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 11, verticalAlign: 'top' }} title={r.subFolderPath}>
                          <button
                            type="button"
                            onClick={() => openFolderViewForListRow(r)}
                            style={{
                              border: 'none',
                              background: 'transparent',
                              color: '#0284c7',
                              cursor: 'pointer',
                              textAlign: 'left',
                              padding: 0,
                              fontSize: 11,
                              maxWidth: '100%',
                              textDecoration: 'underline',
                            }}
                            title="Open this path in Folder view"
                          >
                            {r.subFolderPath}
                          </button>
                          {showDebugKeys && (
                            <div style={{ marginTop: 4, color: '#7c3aed', fontSize: 10, fontFamily: 'monospace', wordBreak: 'break-all' }}>
                              {getRowDebugKey(r)}
                            </div>
                          )}
                        </td>
                        <td style={{ padding: '10px 12px', verticalAlign: 'top' }}>
                          {hasFiles ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              {r.files.map(file => (
                                <div key={file.name} style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
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
                                  {(() => {
                                    const detectedVessel = resolveDetectedVesselForFile(file, host, r.vesselName);
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
                                            host._openVesselSuggestions([new File([], file.name)], r.vesselName || '');
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
                                </div>
                              ))}
                            </div>
                          ) : (
                            <span style={{ color: '#94a3b8', fontStyle: 'italic', fontSize: 11 }}>—</span>
                          )}
                        </td>
                        <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 11, whiteSpace: 'nowrap', verticalAlign: 'top' }}>
                          {hasFiles ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              {r.files.map(file => <span key={file.id}>{file.size || '—'}</span>)}
                            </div>
                          ) : '—'}
                        </td>
                        <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 11, whiteSpace: 'nowrap', verticalAlign: 'top' }}>
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
                        <td style={{ padding: '10px 12px', textAlign: 'right', verticalAlign: 'top' }}>
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
      </div>
    );



}
