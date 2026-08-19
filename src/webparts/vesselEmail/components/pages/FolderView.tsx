import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { renderClassifyDialog } from './VesselsPage';
import { MAIN_FOLDERS, folderNamesByMainFolder, subfolderNamesByFolder } from '../vesselFolderTemplate';

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
function FileTable({ files, onDelete, selectedIds, onToggleSelect }: {
  files: Array<{ name: string; size: string; date: string; pending?: boolean; id?: string }>;
  onDelete?: (file: { id: string; name: string }) => void;
  selectedIds?: Set<string>;
  onToggleSelect?: (id: string) => void;
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
                <td style={{ padding: '12px 16px', fontWeight: 600, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 }}>
                  <span style={{ fontSize: 18 }}>{file.pending ? '⏳' : '📄'}</span>
                  {file.name}
                  {file.pending && <span style={{ fontSize: 11, color: '#d97706', fontWeight: 600, background: '#fef3c7', borderRadius: 4, padding: '1px 6px' }}>Pending Approval</span>}
                </td>
                <td style={{ padding: '12px 16px', color: '#64748b' }}>{file.size}</td>
                <td style={{ padding: '12px 16px', color: '#64748b' }}>{file.date}</td>
                <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                  <div style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
                    <button style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#0078d4' }}>
                      View / Download
                    </button>
                    {onDelete && (
                      <button
                        onClick={() => onDelete({ id: fileId, name: file.name })}
                        style={{ border: '1px solid #fca5a5', background: '#fff5f5', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#ef4444' }}
                        title="Delete file"
                      >
                        🗑 Delete
                      </button>
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

  const { folderPathStack, docMainFolder, vessels, folderViewSelectedFiles } = host.state;

  // ── Level 0: Root — Vessels + Kaizen ──────────────────────────────────────
  if (stackLevel === 0) {
    const mainAnomalies = (host.state.folderAnomalies || []).filter(a => a.anomaly_type === 'main_folder_unmatched');
    const mainFolderAnomalies = mainAnomalies.filter(a => a.item_type === 'folder');
    const mainFileAnomalies   = mainAnomalies.filter(a => a.item_type === 'file');

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
          <FolderCard name="Vessels" sub={`${vessels.length} vessels`} emoji="🚢" bg="#e0f2fe"
            onClick={() => host._pushFolderNav([{ id: 'vessels_root', name: 'Vessels' }], null)} />
          <FolderCard name="Kaizen - Knowledge Bank" sub="Knowledge base" emoji="📚" bg="#ede9fe"
            onClick={() => { host._pushFolderNav([{ id: 'kaizen_root', name: 'Kaizen - Knowledge Bank' }], 'Kaizen - Knowledge Bank'); host.setState({ vesselFilter: 'all' }); }} />
        </div>

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
            {mainFileAnomalies.length > 0 && (
              <div style={{ background: 'linear-gradient(135deg,#faf5ff,#f3e8ff)', border: '2px solid #a855f7', borderRadius: 14, padding: 20 }}>
                <h4 style={{ margin: '0 0 6px', fontSize: 15, fontWeight: 700, color: '#6b21a8' }}>
                  📄 Files Uploaded Outside Main Category Structure ({mainFileAnomalies.length})
                </h4>
                <p style={{ margin: '0 0 12px', fontSize: 12, color: '#7c3aed' }}>
                  These files were uploaded directly to the main folder root in SharePoint Online.
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                  {mainFileAnomalies.map(item => {
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
                        <button onClick={() => host._dismissAnomaly(item.id)}
                          style={{ background: '#fff', color: '#6b21a8', border: '1px solid #e9d5ff', borderRadius: 6, padding: '4px 10px', fontSize: 11, cursor: 'pointer' }}>
                          ✕ Dismiss
                        </button>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    );
  }

  // ── Level 1: Vessels root — Specific Vessels + Common ─────────────────────
  if (atVesselsRoot) {
    return (
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
        {[
          { id: 'specific_vessels', name: 'Specific Vessels',    emoji: '🚢', bg: '#e0f2fe', sub: `${vessels.length} vessels` },
          { id: 'common',           name: 'Common for all ships', emoji: '📁', bg: '#fef3c7', sub: 'Shared documents' },
        ].map(item => (
          <FolderCard key={item.id} name={item.name} sub={item.sub} emoji={item.emoji} bg={item.bg}
            onClick={() => host._pushFolderNav([...folderPathStack, { id: item.id, name: item.name }], null)} />
        ))}
      </div>
    );
  }

  // ── Level 2: Specific Vessels — vessel list ────────────────────────────────
  if (atSpecificVessels) {
    return (
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
        {displayVessels.map(v => (
          <FolderCard key={v.id} name={v.name} sub="Vessel" emoji="🚢" bg="#e0f2fe"
            onClick={() => {
              host._pushFolderNav([...folderPathStack, { id: v.id, name: v.name }], null);
              host.setState({ vesselFilter: v.name });
              host._loadFilesForVessel(v.name).catch(() => undefined);
            }} />
        ))}
      </div>
    );
  }

  // ── Level 3: Main folder selection inside a vessel ─────────────────────────
  if (atVesselMainFolderSelect) {
    return (
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
        {MAIN_FOLDERS.map(mf => {
          const visual = VESSEL_MAIN_FOLDERS.find(item => item.key === mf.name);
          return (
          <FolderCard key={mf.name} name={mf.name} sub="Main folder" emoji={visual?.emoji || '📁'} bg={visual?.bg || '#e0f2fe'}
            onClick={() => {
              host._pushFolderNav(folderPathStack, mf.name as MainFolderKey);
              host.setState({ vesselFilter: currentVesselNameFromStack || 'all' });
            }} />
          );
        })}
      </div>
    );
  }

  // ── Subfolders level ───────────────────────────────────────────────────────
  if (subfolderNames.length > 0) {
    const subAnomalies = (host.state.folderAnomalies || []).filter(a =>
      a.anomaly_type === 'subfolder_unmatched' &&
      (!currentVesselNameFromStack || a.vessel_name === currentVesselNameFromStack)
    );

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 }}>
          {subfolderNames.map((sfName, idx) => (
            <FolderCard key={sfName + idx} name={sfName} sub="Folder" emoji="📁" bg="#e0f2fe"
              onClick={() => {
                const vesselName = currentVesselNameFromStack || (folderPathStack.length > 0 ? folderPathStack[0].name : null);
                const candidateSubPath = vesselName ? `${vesselName} > ${docMainFolder || ''} > ${sfName}` : '';
                const liveFolderId = candidateSubPath ? host._getLiveSharePointFolderId(candidateSubPath) : null;
                const realFolderId = liveFolderId || (vesselName
                  ? (host.state.rows.find(r =>
                      r.vesselName === vesselName &&
                      r.uploadFolderId &&
                      !r.uploadFolderId.includes('/') &&
                      (r.subCategory === sfName || r.category === sfName) &&
                      (docMainFolder ? r.group === docMainFolder : true) &&
                      (currentFolderNode && currentFolderNode.name !== vesselName
                        ? (r.category === currentFolderNode.name || r.group === currentFolderNode.name)
                        : true)
                    )?.uploadFolderId || `sf_${idx}`)
                  : `sf_${idx}`);
                host._pushFolderNav([...folderPathStack, { id: realFolderId, name: sfName }], docMainFolder);
                if (!/^sf_/.test(realFolderId)) {
                  void host._refreshFolderFiles(realFolderId, realFolderId, true).catch(() => undefined);
                } else if (currentFolderNode && !/^(sf_|common)/.test(currentFolderNode.id)) {
                  host._refreshFolderFiles(currentFolderNode.id, currentFolderNode.id).catch(() => undefined);
                }
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
      <FileTable
        files={allCurrentFolderFiles}
        onDelete={handleFileDelete}
        selectedIds={folderViewSelectedFiles}
        onToggleSelect={handleToggleSelect}
      />
    );
  }

  // ── Empty folder ───────────────────────────────────────────────────────────
  return (
    <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', padding: '60px 20px' }}>
      <div style={{ background: 'rgba(240,249,255,0.6)', border: '1px solid #e0f2fe', borderRadius: 24, padding: '48px 40px', maxWidth: 500, width: '100%', textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
        <div style={{ width: 56, height: 56, borderRadius: 14, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 24, color: '#0284c7' }}>📁</div>
        <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#0f172a' }}>This folder is empty</h3>
        <p style={{ margin: 0, fontSize: 13, color: '#64748b', maxWidth: 320, lineHeight: 1.5 }}>
          Use the Upload button in the top-right to add a document.
        </p>
      </div>
    </div>
  );
}
