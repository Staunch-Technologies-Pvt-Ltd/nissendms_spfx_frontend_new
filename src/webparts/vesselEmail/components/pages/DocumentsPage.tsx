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

export function renderDocumentsPage(host: VesselEmail): React.ReactElement {

    const {
      textFilter, vesselFilter, catFilter, docViewMode, showAllVesselsInFolderView,
      vessels, rows, docListPage, docListSort, docGroupFilter,
      docUploadRowKey, docUploadBusy, docUploadMsg, documentsList,
      folderPathStack, uploadedFilesByFolder, docMainFolder,
    } = host.state;

    const PAGE_ROWS = 10;

    // ── Main folder definitions ──
    type MainFolderKey = 'Technical & Crewing' | 'Commercial & Chartering' | 'Insurance' | 'Kaizen - Knowledge Bank' | 'Knowledge Bank';
    const MAIN_FOLDERS: Array<{ key: MainFolderKey; icon: string; emoji: string; color: string; bg: string }> = [
      { key: 'Technical & Crewing', icon: '⚙️', emoji: '⚙️', color: '#dc2626', bg: '#fee2e2' },
      { key: 'Commercial & Chartering', icon: '💼', emoji: '💼', color: '#16a34a', bg: '#dcfce7' },
      { key: 'Insurance', icon: '🛡️', emoji: '🛡️', color: '#d97706', bg: '#fef3c7' },
      { key: 'Kaizen - Knowledge Bank', icon: '📚', emoji: '📚', color: '#7c3aed', bg: '#ede9fe' },
    ];

    // ── Per main-folder subfolder structure (aligned with vesselFolderTemplate.ts) ──
    const DEFAULT_VESSEL_MAINS: Record<MainFolderKey, string[]> = {
      'Technical & Crewing': ['Month End Reports', 'Service Agreements', 'Registration', 'Drawings and Manuals', 'PO & Invoice', 'Incidents', 'Crewing', 'To be Classified'],
      'Commercial & Chartering': ['Agreements', 'Invoices & Payments', 'Claims & Disputes', 'To be Classified'],
      'Insurance': ['P&I', 'H&M', 'War Risk', 'Flag and MPA'],
      'Kaizen - Knowledge Bank': ['Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance'],
      'Knowledge Bank': ['Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance'],
    };

    const SUBFOLDERS_MAP: Record<string, string[]> = {
      'Month End Reports': ['Month Wise Folder'],
      'Drawings and Manuals': [
        'Flag & MPA', 'Ship Builder', 'Rules & Telecom', 'Crewing & SMOU', 'Novation', 'Automation and To be Classified',
        'Drawings', 'Manuals',
      ],
      'Drawings': ['Archive', 'Basic', 'Electrical', 'Engine', 'Hull', 'Other Drawings', 'Safety'],
      'Manuals': ['Automation', 'Auxiliary Engine', 'Boiler', 'Cargo', 'Deck Machinery', 'Electrical', 'Main Engine', 'Other Manuals', 'Pollution', 'Propulsion', 'Refrigeration', 'Safety', 'Shafting', 'Steering Gear', 'Thrusters'],
      'Invoices & Payments': ['Month Wise Folder'],
      'Claims & Disputes': ['Month Wise Folder'],
      'Circulars and Guidance': ['Equipment Maker', 'Class', 'Flag ⁄ Port State', 'SIRE⁄OCIMF⁄RightShip', 'Shipyard'],
      'Common for all ships': ['Vendor & Service Agreements', 'Vendor Management'],
      'Common Agreements (Not Ship Specific)': ['Agreements', 'To be Classified'],
      'Common (Not Ship Specific)': ['Agreements', 'Miscellaneous'],
    };

    const currentFolderNode = folderPathStack.length > 0 ? folderPathStack[folderPathStack.length - 1] : null;
    const currentFolderName = currentFolderNode ? currentFolderNode.name : null;
    const currentVesselNameFromStack = folderPathStack.length > 0 ? folderPathStack[0].name : null;

    if (currentVesselNameFromStack && !host._filesLoadedForVessels.has(currentVesselNameFromStack)) {
      host._filesLoadedForVessels.add(currentVesselNameFromStack);
      setTimeout(() => host._loadFilesForVessel(currentVesselNameFromStack!).catch(() => undefined), 0);
    }


    // Determine subfolders for current depth
    const subfolderNames: string[] = currentFolderName
      ? (SUBFOLDERS_MAP[currentFolderName] || (
          // If we are at vessel-level depth (stack[0] is a vessel), show main folder's default top-level folders
          folderPathStack.length === 1 && docMainFolder
            ? DEFAULT_VESSEL_MAINS[docMainFolder]
            : []
        ))
      : [];


    const currentFolderFiles = currentFolderName ? (uploadedFilesByFolder[currentFolderName] || []).filter((f: any) => !f.pending) : [];

    // Also include approved files from backend rows that match the current folder node
    const backendFolderFiles: Array<{ name: string; size: string; date: string }> = [];
    if (currentFolderNode && !/^(sf_|common)/.test(currentFolderNode.id)) {
      rows.filter(r => r.uploadFolderId === currentFolderNode.id && r.fileName && !r.filePending)
        .forEach(r => backendFolderFiles.push({ name: r.fileName!, size: '—', date: '—' }));
    }
    const allCurrentFolderFiles = [
      ...backendFolderFiles,
      ...currentFolderFiles.filter(f => !backendFolderFiles.some(b => b.name === f.name)),
    ];

    // ── Build source rows from API only — no mock fallback ──
    const allRows: FlatRow[] = rows || [];

    // Subfolders lookup map for Folder View levels
    // (subfolder resolution already computed above via SUBFOLDERS_MAP / DEFAULT_VESSEL_MAINS)


    // Filtered list rows for List view
    const activeVesselName = vesselFilter !== 'all' ? vesselFilter : null;
    const allGroups = Array.from(new Set(allRows.map(r => r.group))).sort();
    const allCategories = Array.from(
      new Set(allRows.filter(r => docGroupFilter === 'all' || r.group === docGroupFilter).map(r => r.category))
    ).sort();

    // ── mainFolderGroupMap: which groups belong to which main folder (for list view filtering) ──
    // These must match the actual template group names returned by the backend API.
    const mainFolderGroupMap: Record<string, string[]> = {
      'Technical & Crewing': [
        'Month End Reports', 'Service Agreements', 'Registration',
        'Drawings and Manuals', 'PO & Invoice', 'Incidents', 'Crewing', 'To be Classified',
      ],
      'Commercial & Chartering': [
        'Agreements', 'Invoices & Payments', 'Claims & Disputes', 'To be Classified',
      ],
      'Insurance': ['P&I', 'H&M', 'War Risk', 'Flag - MPA', 'USA Related', 'Flag and MPA'],
      'Kaizen - Knowledge Bank': [
        'Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance',
      ],
      'Knowledge Bank': [
        'Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance',
      ],
    };

    let filtered = allRows.filter(r => {
      // Filter by active main folder using template-based group membership
      if (docMainFolder) {
        const allowedGroups = mainFolderGroupMap[docMainFolder];
        // If allowedGroups is defined, filter to only matching groups.
        // If the group is not in the map (e.g. flat folder or unknown), still show all rows.
        if (allowedGroups && allowedGroups.length > 0) {
          if (!allowedGroups.includes(r.group) && !r.subFolderPath.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0].toLowerCase())) {
            return false;
          }
        }
      }
      if (activeVesselName && r.vesselName !== activeVesselName) return false;
      if (docGroupFilter !== 'all' && r.group !== docGroupFilter) return false;
      if (catFilter !== 'all' && r.category !== catFilter && r.group !== catFilter) return false;
      if (textFilter) {
        const q = textFilter.toLowerCase();
        return r.vesselName.toLowerCase().includes(q) || r.group.toLowerCase().includes(q) ||
          r.category.toLowerCase().includes(q) || (r.subFolderPath || '').toLowerCase().includes(q) ||
          (r.fileName || '').toLowerCase().includes(q);
      }
      return true;
    });

    // ── Group filtered rows by groupKey (category/folder level) ──
    const groupedMap = new Map<string, GroupedRow>();
    for (const r of filtered) {
      const folderUploads = (uploadedFilesByFolder[r.groupKey] || [])
        .concat(uploadedFilesByFolder[r.uploadFolderId] || [])
        .concat(uploadedFilesByFolder[r.category] || []);

      const existing = groupedMap.get(r.groupKey);
      if (!existing) {
        const files: Array<{ id: string; name: string }> = [];
        if (r.fileName) {
          files.push({ id: r.fileId || r.fileName, name: r.fileName });
        }
        for (const f of folderUploads) {
          if (!files.some(ex => ex.name === f.name)) {
            files.push({ id: (f as any).id || f.name, name: f.name });
          }
        }
        groupedMap.set(r.groupKey, {
          srNo: r.srNo,
          vesselName: r.vesselName,
          group: r.group,
          category: r.category,
          subFolderPath: r.subFolderPath,
          groupKey: r.groupKey,
          uploadFolderId: r.uploadFolderId,
          monthDriven: r.monthDriven,
          canUpload: r.canUpload,
          files,
        });
      } else {
        if (r.fileName && !existing.files.some(f => f.name === r.fileName)) {
          existing.files.push({ id: r.fileId || r.fileName, name: r.fileName });
        }
        for (const f of folderUploads) {
          if (!existing.files.some(ex => ex.name === f.name)) {
            existing.files.push({ id: (f as any).id || f.name, name: f.name });
          }
        }
      }
    }

    let groupedList = Array.from(groupedMap.values());
    if (docListSort === 'name_az') groupedList.sort((a, b) => a.category.localeCompare(b.category));
    else if (docListSort === 'newest') groupedList.reverse();

    const totalPages = Math.max(1, Math.ceil(groupedList.length / PAGE_ROWS));
    const safePage = Math.min(docListPage, totalPages - 1);
    const pageGroupedRows = groupedList.slice(safePage * PAGE_ROWS, (safePage + 1) * PAGE_ROWS);

    const displayVessels = showAllVesselsInFolderView ? vessels : vessels.slice(0, 4);

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        {/* Breadcrumb Navigation Trail */}
        <div style={{ fontSize: 12, color: '#64748b', display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
          <span style={{ cursor: 'pointer', color: '#0284c7' }} onClick={() => host.setState({ docMainFolder: null, folderPathStack: [] })}>Home</span>
          <span>›</span>
          <span
            style={{ cursor: 'pointer', color: !docMainFolder ? '#0f172a' : '#0284c7', fontWeight: !docMainFolder ? 600 : 400 }}
            onClick={() => host.setState({ docMainFolder: null, folderPathStack: [] })}
          >
            Documents
          </span>
          {docMainFolder && (
            <>
              <span>›</span>
              <span
                style={{ cursor: folderPathStack.length === 0 ? 'default' : 'pointer', color: folderPathStack.length === 0 ? '#0f172a' : '#0284c7', fontWeight: folderPathStack.length === 0 ? 600 : 400 }}
                onClick={() => host.setState({ folderPathStack: [] })}
              >
                {docMainFolder}
              </span>
            </>
          )}
          {folderPathStack.map((item, idx) => {
            const isLast = idx === folderPathStack.length - 1;
            return (
              <React.Fragment key={item.id + idx}>
                <span>›</span>
                <span
                  onClick={() => host.setState({ folderPathStack: folderPathStack.slice(0, idx + 1) })}
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
              {currentFolderName || docMainFolder || 'Documents'}
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>
              {docViewMode === 'list'
                ? `${filtered.length} rows · flattened list view`
                : !docMainFolder
                  ? `${MAIN_FOLDERS.length} main folders`
                  : folderPathStack.length === 0
                    ? `${vessels.length} vessels · folder view`
                    : `${subfolderNames.length} folders · ${allCurrentFolderFiles.length} files`}
            </p>
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
                onClick={() => host.setState({ docViewMode: 'list' })}
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
                  const folderKey = currentFolderName || (docMainFolder as string) || 'Documents';
                  const topFolderId = currentFolderNode && !/^(sf_|common)/.test(currentFolderNode.id) ? currentFolderNode.id : '';
                  const currentVessel = folderPathStack.length > 0 ? folderPathStack[0].name : (vessels.length > 0 ? vessels[0].name : 'Bow Fighter');
                  const subFolderPath = currentFolderNode ? `${currentVessel} > ${docMainFolder || ''} > ${currentFolderNode.name}` : `${currentVessel} > ${docMainFolder || ''}`;

                  // Map display folder names to actual SharePoint folder path names
                  const mainFolderPathMap: Record<string, string> = {
                    'Technical & Crewing': 'Technical & Crewing',
                    'Commercial & Chartering': 'Commercial & Chartering',
                    'Insurance': 'Insurance',
                    'Kaizen - Knowledge Bank': 'Kaizen - Knowledge Bank',
                    'Knowledge Bank': 'Kaizen - Knowledge Bank',
                  };
                  const spoMainFolder = (docMainFolder && mainFolderPathMap[docMainFolder]) || 'Technical & Crewing';

                  // Try to find a real backend uploadFolderId from existing rows for host vessel+folder
                  const matchingRow = host.state.rows.find(r =>
                    r.vesselName === currentVessel &&
                    r.uploadFolderId &&
                    !r.uploadFolderId.includes('/') &&
                    (docMainFolder ? r.group.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0]) : true)
                  );
                  const resolvedFolderId = topFolderId || (matchingRow?.uploadFolderId) || `${host.VESSEL_ROOT}/${spoMainFolder}/${currentVessel}`;

                  host.setState({ docUploadMsg: null });
                  try {
                    const { fileId, statusPending } = await host._uploadFileToFolder(
                      resolvedFolderId,
                      subFolderPath,
                      currentVessel,
                      file
                    );

                    const msg = statusPending
                      ? `"${file.name}" submitted for approval.`
                      : `"${file.name}" uploaded successfully to ${folderKey}!`;

                    const newUpload = { name: file.name, size: `${(file.size / 1024).toFixed(1)} KB`, date: 'Just now', pending: statusPending, id: fileId };
                    const existing = uploadedFilesByFolder[folderKey] || [];
                    host.setState({
                      uploadedFilesByFolder: {
                        ...uploadedFilesByFolder,
                        [folderKey]: [...existing.filter(f => f.name !== file.name), newUpload],
                      },
                      docUploadMsg: msg,
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

        {/* Filter Toolbar */}
        <div style={{ background: '#fff', borderRadius: 10, padding: '10px 12px', border: '1px solid #e2e8f0', display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
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
          <select
            value={docGroupFilter}
            onChange={e => host.setState({ docGroupFilter: e.target.value, catFilter: 'all', docListPage: 0 })}
            style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 140 }}
          >
            <option value="all">All groups</option>
            {allGroups.map(g => (
              <option key={g} value={g}>{g}</option>
            ))}
          </select>
          <select
            value={catFilter}
            onChange={e => host.setState({ catFilter: e.target.value, docListPage: 0 })}
            style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 140 }}
          >
            <option value="all">All categories</option>
            {allCategories.map(c => (
              <option key={c} value={c}>{c}</option>
            ))}
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
          <div style={{ display: 'flex', border: '1px solid #cbd5e1', borderRadius: 8, overflow: 'hidden', marginLeft: 'auto' }}>
            <button
              onClick={() => host.setState({ docViewMode: 'folder' })}
              style={{ padding: '5px 10px', background: docViewMode === 'folder' ? '#e2e8f0' : '#fff', border: 'none', cursor: 'pointer', fontSize: 13 }}
              title="Folder view"
            >::</button>
            <button
              onClick={() => host.setState({ docViewMode: 'list' })}
              style={{ padding: '5px 10px', background: docViewMode === 'list' ? '#e2e8f0' : '#fff', border: 'none', borderLeft: '1px solid #cbd5e1', cursor: 'pointer', fontSize: 13 }}
              title="List view"
            >☰</button>
          </div>
        </div>

        {/* Success Message Banner */}
        {docUploadMsg && (
          <div style={{ background: '#f0fdf4', border: '1px solid #86efac', borderRadius: 8, padding: '8px 14px', fontSize: 12, color: '#16a34a', display: 'flex', justifyContent: 'space-between' }}>
            <span>✓ {docUploadMsg}</span>
            <button onClick={() => host.setState({ docUploadMsg: null })} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#16a34a', fontWeight: 700 }}>✕</button>
          </div>
        )}

        {/* View Mode Content */}
        {docViewMode === 'folder' ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 24, marginTop: 4 }}>

            {/* ── Level 0: Four Main Folders ── */}
            {!docMainFolder ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                  {MAIN_FOLDERS.map(mf => (
                    <div
                      key={mf.key}
                      onClick={() => host.setState({ docMainFolder: mf.key, folderPathStack: [], vesselFilter: 'all' })}
                      style={{
                        background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18,
                        display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                        boxShadow: '0 1px 3px rgba(0,0,0,0.05)', transition: 'transform 0.15s, box-shadow 0.15s',
                      }}
                    >
                      <div style={{ width: 44, height: 44, borderRadius: 10, background: mf.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>
                        {mf.emoji}
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{mf.key}</div>
                        <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>{vessels.length} vessels</div>
                      </div>
                      <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                    </div>
                  ))}
                </div>

                {/* ── Segregated Section: Main Folder Level Unmatched Items ── */}
                {(() => {
                  const mainAnomalies = (host.state.folderAnomalies || []).filter(a => a.anomaly_type === 'main_folder_unmatched');
                  if (mainAnomalies.length === 0) return null;

                  const mainFolders = mainAnomalies.filter(a => a.item_type === 'folder');
                  const mainFiles   = mainAnomalies.filter(a => a.item_type === 'file');

                  return (
                    <div style={{ marginTop: 24, display: 'flex', flexDirection: 'column', gap: 16 }}>
                      {/* Unrecognised Folders Section */}
                      {mainFolders.length > 0 && (
                        <div style={{ background: 'linear-gradient(135deg, #fffbeb 0%, #fff9e6 100%)', border: '2px solid #f59e0b', borderRadius: 14, padding: 20 }}>
                          <h4 style={{ margin: '0 0 6px', fontSize: 15, fontWeight: 700, color: '#92400e', display: 'flex', alignItems: 'center', gap: 8 }}>
                            ⚠️ Folders Created Outside Standard Main Folders ({mainFolders.length})
                          </h4>
                          <p style={{ margin: '0 0 12px', fontSize: 12, color: '#a16207' }}>
                            These folders were created directly in SharePoint Online at the main folder root level (`Vessel Management`) outside standard category structures.
                          </p>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                            {mainFolders.map(item => (
                              <div
                                key={item.id}
                                style={{
                                  background: '#fff', borderRadius: 10, border: '1px solid #fde68a', padding: '12px 16px',
                                  display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
                                }}
                              >
                                <div style={{ display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 200 }}>
                                  <span style={{ fontSize: 24 }}>📁</span>
                                  <div>
                                    <div style={{ fontWeight: 700, fontSize: 13, color: '#1f1f1f' }}>{item.name}</div>
                                    <div style={{ fontSize: 11, color: '#78716c', fontFamily: 'monospace' }}>{item.spo_path}</div>
                                  </div>
                                </div>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                  <button
                                    onClick={() => host.setState({ spoClassifyDialog: { anomaly: item, provisioning: false, done: false, error: null } })}
                                    style={{
                                      background: 'linear-gradient(135deg, #f59e0b, #d97706)',
                                      color: '#fff', border: 'none', borderRadius: 6, padding: '6px 12px',
                                      fontSize: 12, fontWeight: 700, cursor: 'pointer',
                                      display: 'inline-flex', alignItems: 'center', gap: 4,
                                    }}
                                  >
                                    🔍 Classify
                                  </button>
                                  <button
                                    onClick={() => host._dismissAnomaly(item.id)}
                                    style={{ background: '#fff', color: '#78716c', border: '1px solid #d6d3d1', borderRadius: 6, padding: '6px 10px', fontSize: 11, cursor: 'pointer' }}
                                  >
                                    ✕ Dismiss
                                  </button>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}

                      {/* Main Level Uploaded Files Section */}
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
                                <div
                                  key={item.id}
                                  style={{
                                    background: '#fff', borderRadius: 10, border: '1px solid #e9d5ff', padding: '10px 14px',
                                    display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
                                  }}
                                >
                                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                                    <span style={{ background: '#f3e8ff', color: '#6b21a8', borderRadius: 4, padding: '2px 6px', fontSize: 10, fontWeight: 700 }}>{ext}</span>
                                    <div>
                                      <span style={{ fontWeight: 600, fontSize: 13, color: '#1e293b' }}>{item.name}</span>
                                      <span style={{ fontSize: 11, color: '#64748b', marginLeft: 8, fontFamily: 'monospace' }}>{item.spo_path}</span>
                                    </div>
                                  </div>
                                  <button
                                    onClick={() => host._dismissAnomaly(item.id)}
                                    style={{ background: '#fff', color: '#6b21a8', border: '1px solid #e9d5ff', borderRadius: 6, padding: '4px 10px', fontSize: 11, cursor: 'pointer' }}
                                  >
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


            ) : folderPathStack.length === 0 ? (

              /* ── Level 1: Vessel list inside the selected main folder ── */
              <>
                {/* Common Documents section */}
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 12 }}>COMMON AGREEMENTS / DOCUMENTS</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 340px))', gap: 16 }}>
                    <div
                      onClick={() => host.setState({ folderPathStack: [{ id: 'common', name: 'Common for all ships' }] })}
                      style={{ background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' }}
                    >
                      <div style={{ width: 44, height: 44, borderRadius: 10, background: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#d97706' }}>📁</div>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a' }}>Common for all ships</div>
                        <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>Common</div>
                      </div>
                      <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                    </div>
                  </div>
                </div>

                {/* Vessels section */}
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 12 }}>VESSELS</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                    {displayVessels.map(v => (
                      <div
                        key={v.id}
                        onClick={() => {
                          host.setState({ vesselFilter: v.name, folderPathStack: [{ id: v.id, name: v.name }] });
                          // Pre-load files & scan live SPO structure for host vessel so new SPO folders show immediately
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
                  {vessels.length > 4 && (
                    <div style={{ display: 'flex', justifyContent: 'center', marginTop: 20 }}>
                      <button
                        onClick={() => host.setState({ showAllVesselsInFolderView: !showAllVesselsInFolderView })}
                        style={{ background: '#fff', border: '1px solid #cbd5e1', borderRadius: 20, padding: '8px 20px', fontSize: 13, fontWeight: 600, color: '#334155', cursor: 'pointer' }}
                      >
                        {showAllVesselsInFolderView ? 'Show Less ⌃' : `More Vessels (${vessels.length - 4}) ⌄`}
                      </button>
                    </div>
                  )}
                </div>
              </>
            ) : subfolderNames.length > 0 ? (
              /* Subfolders Grid (Screenshot 1) */
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 }}>
                {subfolderNames.map((sfName, idx) => (
                  <div
                    key={sfName + idx}
                    onClick={() => {
                      const newStack = [...folderPathStack, { id: `sf_${idx}`, name: sfName }];
                      host.setState({ folderPathStack: newStack });
                      // If the current folder node has a real backend ID, refresh its files
                      if (currentFolderNode && !/^(sf_|common)/.test(currentFolderNode.id)) {
                        host._refreshFolderFiles(currentFolderNode.id, currentFolderNode.id).catch(() => undefined);
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
            ) : allCurrentFolderFiles.length > 0 ? (
              /* Folder File Items List */
              <div style={{ background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
                      <th style={{ padding: '10px 16px' }}>FILE NAME</th>
                      <th style={{ padding: '10px 16px' }}>SIZE</th>
                      <th style={{ padding: '10px 16px' }}>DATE UPLOADED</th>
                      <th style={{ padding: '10px 16px', textAlign: 'right' }}>ACTION</th>
                    </tr>
                  </thead>
                  <tbody>
                    {allCurrentFolderFiles.map((file, idx) => (
                      <tr key={file.name + idx} style={{ borderBottom: '1px solid #f1f5f9' }}>
                        <td style={{ padding: '12px 16px', fontWeight: 600, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 }}>
                          <span style={{ fontSize: 18 }}>📄</span> {file.name}
                        </td>
                        <td style={{ padding: '12px 16px', color: '#64748b' }}>{file.size}</td>
                        <td style={{ padding: '12px 16px', color: '#64748b' }}>{file.date}</td>
                        <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                          <button style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#0078d4' }}>
                            View / Download
                          </button>
                        </td>
                      </tr>
                    ))}
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
              const currentVessel = folderPathStack.length > 0 ? folderPathStack[0].name : null;
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
              <table style={{ width: '100%', minWidth: 850, borderCollapse: 'collapse', fontSize: 12 }}>
                <thead>
                  <tr style={{ background: '#f8fafc', borderBottom: '2px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
                    <th style={{ padding: '10px 10px', width: 44, textAlign: 'center' }}>SR.</th>
                    <th style={{ padding: '10px 12px' }}>VESSEL NAME</th>
                    <th style={{ padding: '10px 12px' }}>GROUP</th>
                    <th style={{ padding: '10px 12px' }}>SUB-CATEGORY</th>
                    <th style={{ padding: '10px 12px' }}>FOLDER PATH</th>
                    <th style={{ padding: '10px 12px' }}>FILE NAME</th>
                    <th style={{ padding: '10px 12px', textAlign: 'right' }}>ATTACHMENT</th>
                  </tr>
                </thead>
                <tbody>
                  {pageGroupedRows.length === 0 ? (
                    <tr>
                      <td colSpan={7} style={{ padding: '36px 16px', textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>
                        No documents found. {textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || catFilter !== 'all' ? 'Try clearing the filters.' : ''}
                      </td>
                    </tr>
                  ) : pageGroupedRows.map((r, idx) => {
                    const globalIdx = safePage * PAGE_ROWS + idx + 1;
                    const isUploading = docUploadRowKey === r.groupKey && docUploadBusy;
                    const hasFiles = r.files.length > 0;

                    return (
                      <tr key={`${r.groupKey}-${idx}`} style={{ borderBottom: '1px solid #f1f5f9', transition: 'background 0.1s' }}
                        onMouseEnter={e => (e.currentTarget.style.background = '#f8fafc')}
                        onMouseLeave={e => (e.currentTarget.style.background = '')}
                      >
                        <td style={{ padding: '10px 10px', color: '#94a3b8', fontSize: 11, fontFamily: 'monospace', textAlign: 'center' }}>{globalIdx}</td>
                        <td style={{ padding: '10px 12px', fontWeight: 700, color: '#0f172a' }}>{r.vesselName}</td>
                        <td style={{ padding: '10px 12px' }}>
                          <span style={{ display: 'inline-block', borderRadius: 8, padding: '3px 8px', fontSize: 11, fontWeight: 600, background: '#eff6ff', color: '#2563eb' }}>
                            {r.group}
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', fontWeight: 600, color: '#1e293b' }}>{r.category}</td>
                        <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 11 }} title={r.subFolderPath}>
                          {r.subFolderPath}
                        </td>
                        <td style={{ padding: '10px 12px' }}>
                          {hasFiles ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              {r.files.map(file => (
                                <div key={file.name} style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                                  <span style={{ fontSize: 14 }}>📄</span>
                                  <span
                                    onClick={() => {
                                      if (file.id && !file.id.startsWith('file_')) {
                                        // Numeric IDs are approval DB row IDs (pending files) — backend
                                        // now resolves them to the staged drive_item_id for preview.
                                        window.open(`${host._base()}/api/files/${file.id}/content`, '_blank');
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
                        <td style={{ padding: '10px 12px', textAlign: 'right' }}>
                          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 4 }}>
                            {hasFiles ? (
                              <button
                                onClick={() => {
                                  const firstFile = r.files[0];
                                  if (firstFile.id && !firstFile.id.startsWith('file_')) {
                                    window.open(`${host._base()}/api/files/${firstFile.id}/content`, '_blank');
                                  } else {
                                    alert(`File "${firstFile.name}" in folder: ${r.subFolderPath}`);
                                  }
                                }}
                                style={{
                                  background: '#eff6ff', color: '#1d4ed8', border: '1px solid #bfdbfe',
                                  borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 600, cursor: 'pointer',
                                  display: 'inline-flex', alignItems: 'center', gap: 4,
                                }}
                              >
                                ↗ Open ({r.files.length})
                              </button>
                            ) : (
                              <span style={{ fontSize: 11, color: '#94a3b8', fontStyle: 'italic' }}>No attachment</span>
                            )}

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
                                    const { fileId, statusPending } = await host._uploadFileToFolder(
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

                                    const newUpload = { name: file.name, size: `${(file.size / 1024).toFixed(1)} KB`, date: 'Just now', pending: statusPending, id: fileId };
                                    const updatedByFolder = {
                                      ...uploadedFilesByFolder,
                                      [r.groupKey]: [...(uploadedFilesByFolder[r.groupKey] || []).filter((f: any) => f.name !== file.name), newUpload],
                                      [r.category]: [...(uploadedFilesByFolder[r.category] || []).filter((f: any) => f.name !== file.name), newUpload],
                                      [r.group]: [...(uploadedFilesByFolder[r.group] || []).filter((f: any) => f.name !== file.name), newUpload],
                                      [r.uploadFolderId]: [...(uploadedFilesByFolder[r.uploadFolderId] || []).filter((f: any) => f.name !== file.name), newUpload],
                                    };

                                    const baseRows = (rows && rows.length > 0) ? rows : allRows;
                                    const updatedRows = baseRows.map(row =>
                                      row.groupKey === r.groupKey
                                        ? { ...row, fileName: file.name, fileId: fileId || row.fileId || `file_${Date.now()}`, filePending: statusPending }
                                        : row
                                    );

                                    host.setState({
                                      docUploadBusy: false,
                                      docUploadRowKey: null,
                                      docUploadMsg: msg,
                                      uploadedFilesByFolder: updatedByFolder,
                                      rows: updatedRows,
                                    });

                                    if (statusPending && r.uploadFolderId && !/^f\d+$/.test(r.uploadFolderId)) {
                                      setTimeout(() => host._refreshFolderFiles(r.uploadFolderId, r.groupKey, true).catch(() => undefined), 1500);
                                    }
                                  } catch (err: any) {
                                    host.setState({ docUploadBusy: false, docUploadRowKey: null, docUploadMsg: `Upload failed: ${err?.message || 'Error'}` });
                                  }
                                }}
                              />
                              {isUploading ? '⏳...' : '↑ Upload'}
                            </label>
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
