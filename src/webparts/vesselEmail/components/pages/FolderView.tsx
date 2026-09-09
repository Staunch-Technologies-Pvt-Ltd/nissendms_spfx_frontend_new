import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { renderClassifyDialog } from './VesselsPage';
import { MAIN_FOLDERS, folderNamesByMainFolder, subfolderNamesByFolder } from '../vesselFolderTemplate';
import { extractFilesFromDataTransfer } from '../BulkUploadModal';
import { resolveDetectedVesselForFile } from '../constants';

export type MainFolderKey =
  | 'Technical & Crewing'
  | 'Commercial & Chartering'
  | 'Insurance'
  | 'Kaizen - Knowledge Bank'
  | 'Knowledge Bank';

export const VESSEL_MAIN_FOLDERS: Array<{ key: MainFolderKey; emoji: string; color: string; bg: string }> = [
  { key: 'Technical & Crewing',      emoji: '⚙️', color: '#dc2626', bg: '#fee2e2' },
  { key: 'Commercial & Chartering',  emoji: '💼', color: '#16a34a', bg: '#dcfce7' },
  { key: 'Insurance',                emoji: '🛡️', color: '#d97706', bg: '#fef3c7' },
];

export const DEFAULT_VESSEL_MAINS = folderNamesByMainFolder();
export const SUBFOLDERS_MAP = subfolderNamesByFolder();

// ── File table used in both subfolder and leaf views ──────────────────────────
function FileTable({ files, onDelete, selectedIds, onToggleSelect, host, vesselName, sharePointPath }: {
  files: Array<{ name: string; size: string; date: string; pending?: boolean; uploading?: boolean; id?: string }>;
  onDelete?: (file: { id: string; name: string }) => void;
  selectedIds?: Set<string>;
  onToggleSelect?: (id: string) => void;
  host?: VesselEmail;
  vesselName?: string | null;
  sharePointPath?: string;
}): React.ReactElement {
  return (
    <div style={{ background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
        <thead>
          <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
            {onToggleSelect && <th style={{ padding: '10px 16px', width: 40, textAlign: 'center' }}></th>}
            <th style={{ padding: '10px 16px' }}>FILE NAME</th>
            <th style={{ padding: '10px 16px' }}>SIZE</th>
            <th style={{ padding: '10px 16px' }}>DATE UPLOADED</th>
            <th style={{ padding: '10px 16px', textAlign: 'right' }}>ACTION</th>
          </tr>
        </thead>
        <tbody>
          {files.map((file, idx) => {
            const fileId = (file as any).id || file.name;
            const isSelected = selectedIds?.has(fileId) ?? false;
            return (
              <tr key={file.name + idx} style={{ borderBottom: '1px solid #f1f5f9' }}>
                {onToggleSelect && (
                  <td style={{ padding: '12px 16px', textAlign: 'center', width: 40 }}>
                    <input
                      type="checkbox"
                      checked={isSelected}
                      onChange={() => onToggleSelect(fileId)}
                      style={{ width: 16, height: 16, accentColor: '#ef4444', cursor: 'pointer' }}
                    />
                  </td>
                )}
                <td style={{ padding: '12px 16px', fontWeight: 600, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                  <span style={{ fontSize: 18 }}>{file.uploading ? '⏳' : (file.pending ? '🕒' : '📄')}</span>
                  <span>{file.name}</span>
                  {file.uploading && <span style={{ fontSize: 11, color: '#0369a1', fontWeight: 600, background: '#e0f2fe', borderRadius: 4, padding: '1px 6px' }}>Uploading...</span>}
                  {file.pending && <span style={{ fontSize: 11, color: '#d97706', fontWeight: 600, background: '#fef3c7', borderRadius: 4, padding: '1px 6px' }}>Pending Approval</span>}
                  {(() => {
                    const detectedVessel = resolveDetectedVesselForFile(file, host, vesselName);
                    const isUnidentified = (host?.state?.ocrUnidentifiedFiles || []).some(
                      n => (n || '').trim().toLowerCase() === file.name.trim().toLowerCase()
                    );

                    if (!host) return null;

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
                            borderRadius: 6, padding: '1px 7px', fontSize: 10, color: '#0284c7',
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
                            host._openVesselSuggestions([new File([], file.name)], vesselName || '');
                          }}
                          title="Vessel name not detected by OCR — click to assign vessel"
                          style={{
                            background: '#fff7ed', border: '1px solid #fed7aa',
                            borderRadius: 6, padding: '1px 7px', fontSize: 10, color: '#c2410c',
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
                  <div style={{ display: 'inline-flex', gap: 6, alignItems: 'center', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                    <button style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#0078d4' }}>
                      View / Download
                    </button>
                    {/* OCR Re-classify button for real (non-pending) files */}
                    {host && !file.pending && !file.uploading && fileId && !fileId.startsWith('file_') && !/^\d+$/.test(fileId) && (
                      <button
                        type="button"
                        title={`Run OCR & Re-classify "${file.name}" into correct folder`}
                        onClick={() => {
                          host._goToView('templates');
                          setTimeout(() => {
                            host.setState({ ocrPendingItemId: fileId, ocrPendingFilename: file.name });
                          }, 100);
                        }}
                        style={{
                          border: '1px solid #86efac', background: '#f0fdf4', borderRadius: 6,
                          padding: '4px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer',
                          color: '#15803d', display: 'inline-flex', alignItems: 'center', gap: 3,
                        }}
                      >
                        🔍 OCR
                      </button>
                    )}
                    {onDelete && (
                      <button
                        onClick={() => onDelete({ id: fileId, name: file.name })}
                        style={{ border: '1px solid #fca5a5', background: '#fff5f5', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#ef4444' }}
                        title="Delete file"
                      >
                        🗑 Delete
                      </button>
                    )}
                    {host && sharePointPath && (
                      <button
                        type="button"
                        onClick={() => void host._openSharePointFolder({ subFolderPath: sharePointPath } as any)}
                        title="Open this folder in SharePoint"
                        aria-label={`Open ${sharePointPath} in SharePoint`}
                        style={{ width: 28, height: 27, padding: 0, borderRadius: 6, border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8', cursor: 'pointer', fontSize: 16, fontWeight: 700, lineHeight: 1, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
                      >↗</button>
                    )}
                  </div>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ── Folder card used at every level ──────────────────────────────────────────
function FolderCard({ name, sub, emoji, bg, onClick }: {
  name: string; sub: string; emoji: string; bg: string; onClick: () => void;
}): React.ReactElement {
  return (
    <div
      onClick={onClick}
      style={{ background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}
    >
      <div style={{ width: 44, height: 44, borderRadius: 10, background: bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22, flexShrink: 0 }}>{emoji}</div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{name}</div>
        <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>{sub}</div>
      </div>
      <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
    </div>
  );
}

// ── Main export ───────────────────────────────────────────────────────────────
export function renderFolderView(
  host: VesselEmail,
  ctx: {
    stackLevel: number;
    atVesselsRoot: boolean;
    atSpecificVessels: boolean;
    atVesselMainFolderSelect: boolean;
    vesselNodeInStack: { id: string; name: string } | null;
    currentVesselNameFromStack: string | null;
    currentFolderNode: { id: string; name: string } | null;
    currentFolderName: string | null;
    vesselStackIdx: number;
    subfolderNames: string[];
    allCurrentFolderFiles: Array<{ name: string; size: string; date: string; pending?: boolean; id?: string }>;
    displayVessels: Array<{ id: string; name: string }>;
  }
): React.ReactElement {
  const {
    stackLevel, atVesselsRoot, atSpecificVessels, atVesselMainFolderSelect,
    vesselNodeInStack, currentVesselNameFromStack, currentFolderNode, currentFolderName,
    vesselStackIdx, subfolderNames, allCurrentFolderFiles, displayVessels,
  } = ctx;

  const handleFileDelete = (file: { id: string; name: string }): void => {
    host._openFileDeleteDialog([{
      id: file.id,
      name: file.name,
      folderId: currentFolderNode?.id || '',
      folderPath: currentFolderNode?.name || '',
    }]);
  };

  const handleToggleSelect = (id: string): void => {
    host.setState(prev => {
      const next = new Set(prev.folderViewSelectedFiles);
      if (next.has(id)) next.delete(id); else next.add(id);
      return { folderViewSelectedFiles: next };
    });
  };

  const { folderPathStack, docMainFolder, vessels, folderViewSelectedFiles, textFilter, vesselFilter, docGroupFilter, catFilter } = host.state;

  const resolveTargetUploadInfo = (): { currentVessel: string; subFolderPath: string; resolvedFolderId: string } => {
    const atKaizenRoot = stackLevel >= 1 && (folderPathStack[0]?.id === 'kaizen_root' || folderPathStack[0]?.name === 'Kaizen - Knowledge Bank');
    const atCommonShips = folderPathStack[1]?.id === 'common' || folderPathStack[1]?.name === 'Common for all ships' || folderPathStack[1]?.name === 'Common for all vessels';
    const topFolderId = currentFolderNode && !/^(sf_|category_|common|vessels_root|specific_vessels|kaizen_root)/.test(currentFolderNode.id) ? currentFolderNode.id : '';
    let currentVessel: string;
    let subFolderPath: string;
    let fallbackPath: string;

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
    } else {
      currentVessel = (vesselFilter !== 'all' ? vesselFilter : '') || currentVesselNameFromStack || '';
      subFolderPath = currentFolderNode && currentFolderNode.name !== docMainFolder && currentFolderNode.name !== 'Documents'
        ? `${docMainFolder || ''}${currentVessel ? ` > ${currentVessel}` : ''} > ${currentFolderNode.name}`
        : `${docMainFolder || ''}${currentVessel ? ` > ${currentVessel}` : ''}`;
      fallbackPath = [docMainFolder || 'Technical & Crewing', ...(currentVessel ? [currentVessel] : []), ...(currentFolderNode && currentFolderNode.name !== docMainFolder && currentFolderNode.name !== 'Documents' ? [currentFolderNode.name] : [])].join('/');
    }

    const liveId = host._getLiveSharePointFolderId(subFolderPath);
    const matchingRow = host.state.rows.find(r =>
      r.vesselName === currentVessel &&
      r.uploadFolderId &&
      !r.uploadFolderId.includes('/') &&
      (docMainFolder ? r.group.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0]) : true)
    );
    const resolvedFolderId = topFolderId || liveId || (matchingRow?.uploadFolderId) || fallbackPath;
    return { currentVessel, subFolderPath, resolvedFolderId };
  };

  const handleDrop = async (e: React.DragEvent): Promise<void> => {
    e.preventDefault();
    e.stopPropagation();
    if (!e.dataTransfer) return;
    const extracted = await extractFilesFromDataTransfer(e.dataTransfer);
    if (!extracted.length) return;
    const { currentVessel, subFolderPath, resolvedFolderId } = resolveTargetUploadInfo();
    host._openBulkUpload(extracted, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode);
  };
  const currentSharePointPath = resolveTargetUploadInfo().subFolderPath;
  const templateSubfolderSet = new Set((subfolderNames || []).map(name => (name || '').trim().toLowerCase()));

  const dynamicSubfolderSet = new Set<string>();
  const pathLower = (folderPathStack || []).map(n => (n.name || '').trim().toLowerCase()).filter(Boolean);
  const vesselInPathLower = pathLower.find(n => n !== 'vessels' && n !== 'specific vessels' && n !== 'common for all ships' && n !== 'kaizen - knowledge bank');
  const mainNorm = (docMainFolder || '').trim().toLowerCase();
  const currentNorm = (currentFolderName || currentFolderNode?.name || '').trim().toLowerCase();

  const forbiddenContainers = new Set([
    'vessels', 'specific vessels', 'documents', 'shared documents', 'root',
    'common for all ships', 'common for all vessels', 'common',
    'common (not ship specific)', 'common agreements (not ship specific)',
    'kaizen', 'kaizen - knowledge bank', 'knowledge bank', 'technical & crewing',
    'commercial & chartering', 'insurance', 'empty list'
  ]);

  (host.state.rows || []).forEach(r => {
    const candidate = (r.subCategory || r.category || '').trim();
    if (!candidate) return;
    if (candidate.includes('||')) return;
    if (/^01[A-Za-z0-9]{15,}$/.test(candidate) || /^[A-Fa-f0-9]{24,}$/.test(candidate)) return;
    if (forbiddenContainers.has(candidate.toLowerCase())) return;
    if (vessels.some(v => (v.name || '').trim().toLowerCase() === candidate.toLowerCase())) return;

    if (!r.uploadFolderId) return;
    if (r.uploadFolderId.includes('/')) return;
    if (host._appDeletedItemIds.has(r.uploadFolderId)) return;
    if (mainNorm && (r.group || '').trim().toLowerCase() !== mainNorm) return;

    const vesselNorm = (r.vesselName || '').trim().toLowerCase();
    const rowCatNorm = (r.category || '').trim().toLowerCase();
    if (vesselInPathLower && vesselNorm && vesselNorm !== vesselInPathLower) return;
    if (currentNorm && rowCatNorm && rowCatNorm !== currentNorm) return;

    dynamicSubfolderSet.add(candidate);
  });

  const uploadedOnlySubfolderNames = Array.from(dynamicSubfolderSet)
    .filter(name => !templateSubfolderSet.has((name || '').trim().toLowerCase()));

  // ── Level 0: Root — Main Departments + Kaizen ────────────────────────────
  if (stackLevel === 0) {
    const mainAnomalies = (host.state.folderAnomalies || []).filter(a => a.anomaly_type === 'main_folder_unmatched');
    const mainFolderAnomalies = mainAnomalies.filter(a => a.item_type === 'folder');

    const allMainFolders = [
      ...VESSEL_MAIN_FOLDERS,
      { key: 'Kaizen - Knowledge Bank' as MainFolderKey, emoji: '📚', color: '#7c3aed', bg: '#ede9fe' },
    ];

    const visibleMainFolders = allMainFolders.filter(item => {
      if (docGroupFilter && docGroupFilter !== 'all') {
        const itemNorm = item.key.toLowerCase();
        const filterNorm = docGroupFilter.toLowerCase();
        if (!itemNorm.includes(filterNorm) && !filterNorm.includes(itemNorm)) return false;
      }
      if (textFilter) {
        const q = textFilter.toLowerCase();
        if (!item.key.toLowerCase().includes(q)) return false;
      }
      return true;
    });

    return (
      <div
        onDragOver={e => e.preventDefault()}
        onDrop={handleDrop}
        style={{ display: 'flex', flexDirection: 'column', gap: 24 }}
      >
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
          {visibleMainFolders.map(item => (
            <FolderCard
              key={item.key}
              name={item.key}
              sub={item.key === 'Kaizen - Knowledge Bank' ? 'Knowledge base' : `${vessels.length} vessels`}
              emoji={item.emoji}
              bg={item.bg}
              onClick={() => {
                if (item.key === 'Kaizen - Knowledge Bank') {
                  host._pushFolderNav([{ id: 'kaizen_root', name: 'Kaizen - Knowledge Bank' }], 'Kaizen - Knowledge Bank');
                  host.setState({ vesselFilter: 'all', docScopeType: 'kaizen' });
                } else {
                  host._pushFolderNav([{ id: item.key, name: item.key }], item.key);
                  host.setState({ vesselFilter: 'all', docScopeType: 'vessels' });
                }
              }}
            />
          ))}
        </div>

        {/* Uploaded Documents Table at Root Level */}
        {allCurrentFolderFiles && allCurrentFolderFiles.length > 0 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <div style={{ fontSize: 15, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ fontSize: 18 }}>📄</span> Uploaded Documents ({allCurrentFolderFiles.length})
              </div>
            </div>
            <FileTable
              files={allCurrentFolderFiles}
              onDelete={handleFileDelete}
              selectedIds={folderViewSelectedFiles}
              onToggleSelect={handleToggleSelect}
              host={host}
              vesselName={currentVesselNameFromStack}
              sharePointPath={currentSharePointPath}
            />
          </div>
        )}

        {/* Anomaly panels */}
        {mainAnomalies.length > 0 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            {mainFolderAnomalies.length > 0 && (
              <div style={{ background: 'linear-gradient(135deg,#fffbeb,#fff9e6)', border: '2px solid #f59e0b', borderRadius: 14, padding: 20 }}>
                <h4 style={{ margin: '0 0 6px', fontSize: 15, fontWeight: 700, color: '#92400e' }}>
                  ⚠️ Folders Created Outside Standard Main Folders ({mainFolderAnomalies.length})
                </h4>
                <p style={{ margin: '0 0 12px', fontSize: 12, color: '#a16207' }}>
                  These folders were created directly in SharePoint Online at the main folder root level outside standard category structures.
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                  {mainFolderAnomalies.map(item => (
                    <div key={item.id} style={{ background: '#fff', borderRadius: 10, border: '1px solid #fde68a', padding: '12px 16px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 200 }}>
                        <span style={{ fontSize: 24 }}>📁</span>
                        <div>
                          <div style={{ fontWeight: 700, fontSize: 13, color: '#1f1f1f' }}>{item.name}</div>
                          <div style={{ fontSize: 11, color: '#78716c', fontFamily: 'monospace' }}>{item.spo_path}</div>
                        </div>
                      </div>
                      <div style={{ display: 'flex', gap: 8 }}>
                        <button onClick={() => host.setState({ spoClassifyDialog: { anomaly: item, provisioning: false, done: false, error: null } })}
                          style={{ background: 'linear-gradient(135deg,#f59e0b,#d97706)', color: '#fff', border: 'none', borderRadius: 6, padding: '6px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                          🔍 Classify
                        </button>
                        <button onClick={() => host._dismissAnomaly(item.id)}
                          style={{ background: '#fff', color: '#78716c', border: '1px solid #d6d3d1', borderRadius: 6, padding: '6px 10px', fontSize: 11, cursor: 'pointer' }}>
                          ✕ Dismiss
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

          </div>
        )}
      </div>
    );
  }

  // ── Level 1: Inside Main Department — Vessels list + Common folder ──────
  const isInsideMainDepartment = stackLevel === 1 && folderPathStack[0]?.id !== 'kaizen_root';
  if (isInsideMainDepartment) {
    // Derive the correct common folder name for this department from the template
    const activeMfTemplate = MAIN_FOLDERS.find(mf => mf.name === docMainFolder);
    const commonFolderDisplayName = (activeMfTemplate as any)?.commonFolderName ?? 'Common for all ships';
    const filteredDisplayVessels = displayVessels.filter(v => {
      if (vesselFilter && vesselFilter !== 'all') {
        if (v.name.trim().toLowerCase() !== vesselFilter.trim().toLowerCase()) return false;
      }
      if (textFilter) {
        if (!v.name.toLowerCase().includes(textFilter.trim().toLowerCase())) return false;
      }
      return true;
    });

    return (
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
        {(!vesselFilter || vesselFilter === 'all') && (
          <FolderCard
            name={commonFolderDisplayName}
            sub={`Shared ${docMainFolder || ''} documents`}
            emoji="📁"
            bg="#fef3c7"
            onClick={() => {
              host._pushFolderNav([...folderPathStack, { id: 'common', name: commonFolderDisplayName }], docMainFolder);
              host.setState({ vesselFilter: 'all', docScopeType: 'common' });
            }}
          />
        )}
        {filteredDisplayVessels.map(v => (
          <FolderCard
            key={v.id}
            name={v.name}
            sub="Vessel"
            emoji="🚢"
            bg="#e0f2fe"
            onClick={() => {
              host._pushFolderNav([...folderPathStack, { id: v.id, name: v.name }], docMainFolder);
              host.setState({ vesselFilter: v.name, docScopeType: 'vessels' });
              host._loadFilesForVessel(v.name).catch(() => undefined);
            }}
          />
        ))}
        {host.state.documentVesselCount < host.state.vessels.length && (
          <div
            onClick={() => {
              if (!host.state.documentVesselsLoadingMore) {
                void host._loadMoreDocumentVessels();
              }
            }}
            style={{
              background: host.state.documentVesselsLoadingMore ? '#f8fafc' : '#f0f9ff',
              borderRadius: 14,
              border: '2px dashed #0284c7',
              padding: 18,
              display: 'flex',
              alignItems: 'center',
              gap: 14,
              cursor: host.state.documentVesselsLoadingMore ? 'wait' : 'pointer',
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
              {host.state.documentVesselsLoadingMore ? '⏳' : '+'}
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: '#0284c7', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                {host.state.documentVesselsLoadingMore ? 'Loading vessels...' : `More vessels (+${Math.min(8, host.state.vessels.length - host.state.documentVesselCount)})`}
              </div>
              <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>
                {host.state.documentVesselsLoadingMore ? 'Please wait...' : `Load next batch (${host.state.documentVesselCount} of ${host.state.vessels.length} shown)`}
              </div>
            </div>
            <span style={{ color: '#0284c7', fontSize: 18, fontWeight: 700 }}>›</span>
          </div>
        )}
      </div>
    );
  }

  // ── Subfolders level ───────────────────────────────────────────────────────
  if (subfolderNames.length > 0) {
    const isAtCategoryLevel =
      (folderPathStack[0]?.id === 'kaizen_root' && stackLevel === 1) ||
      (folderPathStack[1]?.id === 'common' && stackLevel === 2) ||
      (vesselNodeInStack && stackLevel === 2);

    const subAnomalies = (host.state.folderAnomalies || []).filter(a =>
      a.anomaly_type === 'subfolder_unmatched' &&
      (!currentVesselNameFromStack || a.vessel_name === currentVesselNameFromStack)
    );

    const activeSectionFilter = (host.state.docCategoryFilter && host.state.docCategoryFilter !== 'all')
      ? host.state.docCategoryFilter
      : (catFilter && catFilter !== 'all' ? catFilter : null);

    const visibleSubfolderNames = subfolderNames.filter(sf => {
      const sfNorm = sf.trim().toLowerCase();
      if (isAtCategoryLevel && activeSectionFilter) {
        if (sfNorm !== activeSectionFilter.trim().toLowerCase()) return false;
      }
      if (stackLevel === 3 && host.state.docGroupLevelFilter && host.state.docGroupLevelFilter !== 'all') {
        if (sfNorm !== host.state.docGroupLevelFilter.trim().toLowerCase()) return false;
      }
      if (stackLevel === 4 && host.state.docLeafCategoryFilter && host.state.docLeafCategoryFilter !== 'all') {
        if (sfNorm !== host.state.docLeafCategoryFilter.trim().toLowerCase()) return false;
      }
      if (stackLevel >= 5 && host.state.docSubCategoryFilter && host.state.docSubCategoryFilter !== 'all') {
        if (sfNorm !== host.state.docSubCategoryFilter.trim().toLowerCase()) return false;
      }
      if (textFilter) {
        if (!sf.toLowerCase().includes(textFilter.trim().toLowerCase())) return false;
      }
      return true;
    });

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 }}>
          {visibleSubfolderNames.map((sfName, idx) => (
            <FolderCard
              key={sfName + idx}
              name={sfName}
              sub={isAtCategoryLevel ? 'Document Section' : (stackLevel === 3 ? 'Category' : 'Sub-Category')}
              emoji="📁"
              bg="#e0f2fe"
              onClick={() => {
                const vesselName = currentVesselNameFromStack || (folderPathStack.length > 0 ? folderPathStack[0].name : null);
                const stackNames = folderPathStack.map(n => n.name).filter(n => n !== 'Vessels' && n !== 'Specific Vessels');
                const fullBreadcrumbWithSf = [...stackNames, sfName].join(' > ');
                const candidateSubPath = vesselName ? `${vesselName} > ${docMainFolder || ''} > ${sfName}` : '';
                const liveFolderId = host._getLiveSharePointFolderId(fullBreadcrumbWithSf) ||
                  (candidateSubPath ? host._getLiveSharePointFolderId(candidateSubPath) : null);
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
                const realFolderId = liveFolderId || matchingSubRow?.uploadFolderId || fullBreadcrumbWithSf;
                host._pushFolderNav([...folderPathStack, { id: realFolderId, name: sfName }], docMainFolder);
                const subGroupKey = matchingSubRow?.groupKey || fullBreadcrumbWithSf;
                void host._refreshFolderFiles(realFolderId, subGroupKey, true).catch(() => undefined);
              }} />
          ))}
        </div>

        {allCurrentFolderFiles.length > 0 && (
        <div style={{ marginTop: 8 }}>
            <FileTable
              files={allCurrentFolderFiles}
              onDelete={handleFileDelete}
              selectedIds={folderViewSelectedFiles}
              onToggleSelect={handleToggleSelect}
              host={host}
              sharePointPath={currentSharePointPath}
            />
          </div>
        )}

        {subAnomalies.length > 0 && (
          <div style={{ background: '#fff7ed', border: '1px solid #ffedd5', borderRadius: 14, padding: 18 }}>
            <h4 style={{ margin: '0 0 8px', fontSize: 14, fontWeight: 700, color: '#c2410c' }}>
              ⚠️ Other / Unclassified Items Inside Vessel Tree ({subAnomalies.length})
            </h4>
            <p style={{ margin: '0 0 12px', fontSize: 12, color: '#9a3412' }}>
              These items were added inside the vessel folder in SharePoint but are not part of the standard template structure.
            </p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {subAnomalies.map(item => (
                <div key={item.id} style={{ background: '#fff', borderRadius: 8, border: '1px solid #fed7aa', padding: '10px 14px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                    <span style={{ fontSize: 18 }}>{item.item_type === 'folder' ? '📁' : '📄'}</span>
                    <div>
                      <span style={{ fontWeight: 600, fontSize: 13, color: '#1e293b' }}>{item.name}</span>
                      <span style={{ fontSize: 11, color: '#64748b', marginLeft: 8 }}>Path: {item.spo_path}</span>
                    </div>
                  </div>
                  <button onClick={() => host._dismissAnomaly(item.id)}
                    style={{ background: '#fff', color: '#c2410c', border: '1px solid #fed7aa', borderRadius: 6, padding: '4px 10px', fontSize: 11, cursor: 'pointer' }}>
                    ✕ Dismiss
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  }

  // ── Leaf: files only ───────────────────────────────────────────────────────
  if (allCurrentFolderFiles.length > 0) {
    return (
      <div
        onDragOver={e => e.preventDefault()}
        onDrop={handleDrop}
        style={{ display: 'flex', flexDirection: 'column', gap: 12 }}
      >
        <FileTable
          files={allCurrentFolderFiles}
          onDelete={handleFileDelete}
          selectedIds={folderViewSelectedFiles}
          onToggleSelect={handleToggleSelect}
          host={host}
          vesselName={currentVesselNameFromStack}
          sharePointPath={currentSharePointPath}
        />
      </div>
    );
  }

  // ── Empty folder ───────────────────────────────────────────────────────────
  return (
    <div
      onDragOver={e => e.preventDefault()}
      onDrop={handleDrop}
      style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', padding: '50px 20px' }}
    >
      <div style={{
        background: '#fff', border: '2px dashed #93c5fd', borderRadius: 24,
        padding: '40px 36px', maxWidth: 520, width: '100%', textAlign: 'center',
        display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 14,
        boxShadow: '0 4px 20px rgba(2,132,199,0.06)',
      }}>
        <div style={{ width: 64, height: 64, borderRadius: 18, background: '#eff6ff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 30 }}>
          ☁️
        </div>
        <div>
          <h3 style={{ margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0f172a' }}>This folder is empty</h3>
          <p style={{ margin: 0, fontSize: 13, color: '#64748b', lineHeight: 1.5 }}>
            Drag and drop files or entire folders here, or use the buttons below.
          </p>
        </div>

        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'center', marginTop: 6 }}>
          <label style={{
            background: '#0284c7', color: '#fff', border: 'none', borderRadius: 8,
            padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
            display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 6px rgba(2,132,199,0.25)',
          }}>
            <input
              type="file"
              multiple
              style={{ display: 'none' }}
              onChange={e => {
                const filesList = Array.from(e.target.files || []);
                if (filesList.length === 0) return;
                e.target.value = '';
                const { currentVessel, subFolderPath, resolvedFolderId } = resolveTargetUploadInfo();
                const bulkFiles = filesList.map(f => ({ file: f, relativePath: f.name }));
                host._openBulkUpload(bulkFiles, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode);
              }}
            />
            <span>⬆</span> Upload Files
          </label>

          <label style={{
            background: '#059669', color: '#fff', border: 'none', borderRadius: 8,
            padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
            display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 6px rgba(5,150,105,0.25)',
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
                const { currentVessel, subFolderPath, resolvedFolderId } = resolveTargetUploadInfo();
                const bulkFiles = filesList.map(f => ({
                  file: f,
                  relativePath: (f as any).webkitRelativePath || f.name,
                }));
                host._openBulkUpload(bulkFiles, resolvedFolderId, subFolderPath, currentVessel, currentFolderNode);
              }}
            />
            <span>📁</span> Upload Folder
          </label>

          {uploadedOnlySubfolderNames.length > 0 && (
            <button
              onClick={() => {
                const { currentVessel, subFolderPath } = resolveTargetUploadInfo();
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
                padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 6px rgba(225,29,72,0.15)',
              }}
              title="Delete uploaded folders (non-template only)"
            >
              <span>🗑</span> Delete Uploaded Folder
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
