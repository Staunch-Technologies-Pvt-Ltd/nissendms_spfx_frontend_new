import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import { isMobileWidth } from '../responsive';

/** One folder/file entry as returned by `_loadAndCacheSiteFolderChildren` —
 *  a raw Microsoft Graph driveItem. `.folder` is only present on folders. */
interface GraphChildItem {
  id: string;
  name: string;
  folder?: unknown;
}

interface FolderNodeState {
  loading: boolean;
  error: boolean;
  folders: GraphChildItem[];
  files: GraphChildItem[];
}

interface SelectedFile {
  id: string;
  name: string;
  /** "<root folder name>/.../<parent folder name>" — informational only,
   *  matches the shape _archiveDocumentFile's other callers pass. */
  folderPath: string;
}

export type ArchivePickerDialogState = {
  siteId: string;
  driveId: string;
  folderId: string;
  folderName: string;
  department: string;
  vesselName: string;
};

/** A folder tree the size of a real SharePoint library can be very deep —
 *  cap how many files a single "select whole folder" recursive scan will
 *  pull in so one click can't hang the popup or the archive run after it. */
const MAX_FOLDER_SCAN_FILES = 500;

export function renderArchivePickerModal(host: VesselEmail): React.ReactElement | null {
  const dialog = host.state.archivePickerDialog;
  if (!dialog) return null;
  return <ArchivePickerDialog key={`${dialog.siteId}::${dialog.driveId}::${dialog.folderId}`} host={host} dialog={dialog} />;
}

function ArchivePickerDialog({ host, dialog }: { host: VesselEmail; dialog: ArchivePickerDialogState }): React.ReactElement {
  const { siteId, driveId, folderId: rootFolderId, folderName: rootFolderName, department, vesselName } = dialog;
  const isMobile = isMobileWidth(host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200));

  const [nodes, setNodes] = React.useState<Record<string, FolderNodeState>>({});
  const [expanded, setExpanded] = React.useState<Set<string>>(() => new Set([rootFolderId]));
  const [selected, setSelected] = React.useState<Map<string, SelectedFile>>(new Map());
  // Folder ids that are fully included via "select whole folder" — used only
  // to render that folder's own checkbox as checked; per-file checkboxes are
  // the source of truth for what actually gets archived.
  const [includedFolders, setIncludedFolders] = React.useState<Set<string>>(new Set());
  const [scanningFolders, setScanningFolders] = React.useState<Set<string>>(new Set());
  const [filter, setFilter] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const [archiveProgress, setArchiveProgress] = React.useState<{ done: number; total: number } | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [resultMsg, setResultMsg] = React.useState<string | null>(null);

  const loadFolder = React.useCallback((folderId: string) => {
    setNodes(prev => ({
      ...prev,
      [folderId]: { loading: true, error: false, folders: prev[folderId]?.folders || [], files: prev[folderId]?.files || [] },
    }));
    host._loadAndCacheSiteFolderChildren(siteId, driveId, folderId)
      .then((items: GraphChildItem[]) => {
        const folders = (items || []).filter(it => !!it.folder);
        const files = (items || []).filter(it => !it.folder);
        setNodes(prev => ({ ...prev, [folderId]: { loading: false, error: false, folders, files } }));
      })
      .catch(() => {
        setNodes(prev => ({ ...prev, [folderId]: { loading: false, error: true, folders: [], files: [] } }));
      });
  }, [siteId, driveId, host]);

  React.useEffect(() => {
    loadFolder(rootFolderId);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rootFolderId]);

  const toggleExpanded = (folderId: string): void => {
    setExpanded(prev => {
      const next = new Set(prev);
      if (next.has(folderId)) {
        next.delete(folderId);
      } else {
        next.add(folderId);
        if (!nodes[folderId]) loadFolder(folderId);
      }
      return next;
    });
  };

  const toggleFile = (file: SelectedFile): void => {
    setSelected(prev => {
      const next = new Map(prev);
      if (next.has(file.id)) next.delete(file.id); else next.set(file.id, file);
      return next;
    });
  };

  /** Recursively walk a folder (using the same cache-aware loader the tree
   *  view uses) and select every file under it, up to MAX_FOLDER_SCAN_FILES. */
  const includeFolder = async (folderId: string, folderPath: string): Promise<void> => {
    setScanningFolders(prev => new Set(prev).add(folderId));
    const found: SelectedFile[] = [];
    let truncated = false;
    const walk = async (fid: string, path: string): Promise<void> => {
      if (truncated) return;
      const items = await host._loadAndCacheSiteFolderChildren(siteId, driveId, fid);
      for (const it of items || []) {
        if (truncated) return;
        if (it.folder) {
          await walk(it.id, `${path}/${it.name}`);
        } else {
          found.push({ id: it.id, name: it.name, folderPath: path });
          if (found.length >= MAX_FOLDER_SCAN_FILES) { truncated = true; return; }
        }
      }
    };
    try {
      await walk(folderId, folderPath);
    } catch {
      // best-effort — whatever was found before the failure still gets selected
    }
    setSelected(prev => {
      const next = new Map(prev);
      for (const f of found) next.set(f.id, f);
      return next;
    });
    setIncludedFolders(prev => new Set(prev).add(folderId));
    setScanningFolders(prev => { const next = new Set(prev); next.delete(folderId); return next; });
    if (truncated) {
      setError(`Only the first ${MAX_FOLDER_SCAN_FILES} files under "${folderPath.split('/').pop()}" were selected — open its sub-folders to pick the rest individually.`);
    }
  };

  const excludeFolder = (folderId: string, folderPath: string): void => {
    setIncludedFolders(prev => { const next = new Set(prev); next.delete(folderId); return next; });
    setSelected(prev => {
      const next = new Map(prev);
      prev.forEach((f, id) => {
        if (f.folderPath === folderPath || f.folderPath.startsWith(`${folderPath}/`)) next.delete(id);
      });
      return next;
    });
  };

  const matchesFilter = (name: string): boolean =>
    !filter.trim() || name.toLowerCase().includes(filter.trim().toLowerCase());

  const renderNode = (folderId: string, folderName: string, folderPath: string, depth: number): React.ReactElement => {
    const node = nodes[folderId];
    const isExpanded = expanded.has(folderId);
    const isIncluded = includedFolders.has(folderId);
    const isScanning = scanningFolders.has(folderId);

    return (
      <div key={folderId}>
        <div style={{
          display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px',
          paddingLeft: 8 + depth * 20, borderRadius: 6,
          background: isIncluded ? '#fffbeb' : 'transparent',
        }}>
          <button
            type="button"
            onClick={() => toggleExpanded(folderId)}
            aria-label={isExpanded ? 'Collapse folder' : 'Expand folder'}
            style={{
              width: 20, height: 20, border: 'none', background: 'transparent', cursor: 'pointer',
              color: 'var(--vdms-text-muted)', fontSize: 12, display: 'flex', alignItems: 'center', justifyContent: 'center',
            }}
          >
            {isExpanded ? <Icon iconName="ChevronDown" aria-hidden="true" style={{ fontSize: 12 }} /> : <Icon iconName="ChevronRight" aria-hidden="true" style={{ fontSize: 12 }} />}
          </button>
          <input
            type="checkbox"
            checked={isIncluded}
            disabled={isScanning || busy}
            onChange={() => {
              if (isIncluded) excludeFolder(folderId, folderPath);
              else void includeFolder(folderId, folderPath);
            }}
            style={{ width: 15, height: 15, accentColor: '#d97706', cursor: isScanning ? 'wait' : 'pointer' }}
          />
          <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 14 }} />
          <span style={{ fontSize: 13, fontWeight: 600, color: 'var(--vdms-text)', flex: 1 }}>{folderName}</span>
          {isScanning && <span style={{ fontSize: 11, color: 'var(--vdms-text-faint)' }}>Scanning…</span>}
        </div>

        {isExpanded && (
          <div>
            {!node || node.loading ? (
              <div style={{ paddingLeft: 28 + depth * 20, fontSize: 12, color: 'var(--vdms-text-faint)', padding: '4px 0' }}>Loading…</div>
            ) : node.error ? (
              <div style={{ paddingLeft: 28 + depth * 20, fontSize: 12, color: '#dc2626', padding: '4px 0' }}>
                Couldn't load this folder.
                <button type="button" onClick={() => loadFolder(folderId)} style={{ marginLeft: 8, border: 'none', background: 'none', color: '#0284c7', cursor: 'pointer', fontSize: 12, fontWeight: 600 }}>Retry</button>
              </div>
            ) : node.folders.length === 0 && node.files.length === 0 ? (
              <div style={{ paddingLeft: 28 + depth * 20, fontSize: 12, color: 'var(--vdms-text-faint)', padding: '4px 0' }}>Empty folder</div>
            ) : (
              <>
                {node.folders.filter(f => matchesFilter(f.name)).map(f =>
                  renderNode(f.id, f.name, `${folderPath}/${f.name}`, depth + 1)
                )}
                {node.files.filter(f => matchesFilter(f.name)).map(f => {
                  const sel = selected.has(f.id);
                  return (
                    <label key={f.id} style={{
                      display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px',
                      paddingLeft: 28 + depth * 20, borderRadius: 6, cursor: 'pointer',
                      background: sel ? '#fff7ed' : 'transparent',
                    }}>
                      <input
                        type="checkbox"
                        checked={sel}
                        disabled={busy}
                        onChange={() => toggleFile({ id: f.id, name: f.name, folderPath })}
                        style={{ width: 15, height: 15, accentColor: '#d97706', cursor: 'pointer' }}
                      />
                      <Icon iconName="Page" aria-hidden="true" style={{ fontSize: 13 }} />
                      <span style={{ fontSize: 12.5, color: 'var(--vdms-text)', flex: 1, wordBreak: 'break-all' }}>{f.name}</span>
                    </label>
                  );
                })}
              </>
            )}
          </div>
        )}
      </div>
    );
  };

  const close = (): void => {
    if (busy) return;
    host.setState({ archivePickerDialog: null });
  };

  const confirmArchive = async (): Promise<void> => {
    const files = Array.from(selected.values());
    if (files.length === 0) return;
    setBusy(true);
    setError(null);
    setArchiveProgress({ done: 0, total: files.length });

    let archivedCount = 0;
    const failures: string[] = [];
    for (const file of files) {
      try {
        const ok = await host._archiveDocumentFile(file.id, file.name, file.folderPath, department, vesselName);
        if (ok) archivedCount += 1; else failures.push(file.name);
      } catch (err: any) {
        failures.push(`${file.name}: ${err?.message || 'failed'}`);
      }
      setArchiveProgress(prev => (prev ? { ...prev, done: prev.done + 1 } : prev));
    }

    // Refresh whatever folder(s) were touched so archived files disappear
    // from the browser immediately instead of waiting for the next
    // unrelated navigation to refetch them: the root folder itself, plus
    // the immediate parent folder each archived file actually lived in.
    void host._refreshSiteFolder(siteId, driveId, rootFolderId).catch(() => undefined);
    const refreshedFolderIds = new Set<string>([rootFolderId]);
    for (const file of files) {
      const parentId = Object.keys(nodes).find(id => (nodes[id].files || []).some(f => f.id === file.id));
      if (parentId && !refreshedFolderIds.has(parentId)) {
        refreshedFolderIds.add(parentId);
        void host._refreshSiteFolder(siteId, driveId, parentId).catch(() => undefined);
      }
    }

    setBusy(false);
    if (failures.length > 0) {
      setError(`${archivedCount} of ${files.length} archived. Failed: ${failures.join(', ')}`);
      return;
    }
    setResultMsg(`${archivedCount} file${archivedCount === 1 ? '' : 's'} archived successfully.`);
    setTimeout(() => host.setState({ archivePickerDialog: null }), 1400);
  };

  const selectedCount = selected.size;

  // Rendered via a portal straight onto <body> — the SharePoint workbench/page
  // canvas that hosts this web part sets a `transform` on an ancestor section
  // (for its drag-and-drop chrome), which redefines the containing block for
  // any `position: fixed` descendant. Without the portal this dialog's
  // backdrop and header end up clipped to that section instead of the real
  // viewport, so it renders too short and overlaps the page's own toolbar
  // (the "Archive files" header colliding with the site selector/buttons)
  // instead of covering and dimming the whole screen.
  // The --vdms-* custom properties this dialog's inline styles read (surface,
  // border, text colors) are only defined by the [data-vessel-theme="..."]
  // CSS rule in AppLayout.tsx, scoped to the .vessel-dms-app root. Now that
  // the portal mounts this dialog directly on <body> — outside that root —
  // those variables no longer cascade down to it and every var(--vdms-*)
  // falls back to nothing, so the card had no background/border/text color
  // and effectively rendered invisible. Re-declaring the same attribute here
  // re-establishes the variables for this subtree.
  const themeMode = host.state.themeMode === 'night' ? 'night' : 'light';

  return ReactDOM.createPortal(
    <div data-vessel-theme={themeMode} style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 100001,
      display: 'flex', alignItems: 'center', justifyContent: 'center', padding: isMobile ? 10 : 20,
    }} onClick={close}>
      <div
        style={{
          background: 'var(--vdms-surface)', borderRadius: 16, padding: isMobile ? '16px 14px' : '24px 28px',
          width: isMobile ? '95vw' : 560, maxWidth: '95vw', maxHeight: '85vh', display: 'flex', flexDirection: 'column',
          boxShadow: '0 8px 40px rgba(0,0,0,0.18)', fontFamily: "'Segoe UI', sans-serif",
        }}
        onClick={e => e.stopPropagation()}
      >
        {resultMsg ? (
          <div style={{ textAlign: 'center', padding: '20px 0 4px' }}>
            <div style={{ fontSize: 44, marginBottom: 10 }}><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 44 }} /></div>
            <div style={{ fontSize: 17, fontWeight: 700, color: '#059669', marginBottom: 6 }}>Archived</div>
            <p style={{ fontSize: 13, color: 'var(--vdms-text-muted)', margin: 0 }}>{resultMsg}</p>
          </div>
        ) : (
          <>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 12, flexShrink: 0 }}>
              <div style={{ width: 40, height: 40, borderRadius: 10, background: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}><Icon iconName="Package" aria-hidden="true" style={{ fontSize: 20 }} /></div>
              <div style={{ minWidth: 0, flex: 1 }}>
                <div style={{ fontWeight: 700, fontSize: 16, color: 'var(--vdms-text)' }}>Archive files</div>
                <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginTop: 2, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  Check folders or files under "{rootFolderName}" to archive
                </div>
              </div>
            </div>

            <input
              type="text"
              value={filter}
              onChange={e => setFilter(e.target.value)}
              placeholder="Filter by name…"
              disabled={busy}
              style={{
                width: '100%', boxSizing: 'border-box', borderRadius: 8, border: '1px solid var(--vdms-border)',
                padding: '8px 10px', fontSize: 13, marginBottom: 10, background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text)',
                flexShrink: 0,
              }}
            />

            <div style={{
              flex: 1, overflowY: 'auto', border: '1px solid var(--vdms-border)', borderRadius: 10,
              padding: 6, marginBottom: 12, minHeight: 180,
            }}>
              {renderNode(rootFolderId, rootFolderName, rootFolderName, 0)}
            </div>

            {archiveProgress && busy && (
              <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginBottom: 10 }}>
                Archiving {archiveProgress.done} of {archiveProgress.total}…
              </div>
            )}

            {error && (
              <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '8px 12px', fontSize: 12, color: '#dc2626', marginBottom: 12 }}>
                <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 12 }} /> {error}
              </div>
            )}

            <div style={{ background: '#fff7ed', border: '1px solid #fed7aa', borderRadius: 8, padding: '8px 12px', fontSize: 12, color: '#c2410c', marginBottom: 16, flexShrink: 0 }}>
              <Icon iconName="Package" aria-hidden="true" style={{ fontSize: 12 }} /> {selectedCount} file{selectedCount !== 1 ? 's' : ''} selected to archive.
            </div>

            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', flexDirection: isMobile ? 'column' : 'row', flexShrink: 0 }}>
              <button
                onClick={close}
                disabled={busy}
                style={{ minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 20px', borderRadius: 8, border: '1px solid var(--vdms-border)', background: 'var(--vdms-surface)', fontSize: 13, fontWeight: 600, cursor: 'pointer', color: 'var(--vdms-text)' }}
              >
                Cancel
              </button>
              <button
                onClick={() => void confirmArchive()}
                disabled={busy || selectedCount === 0}
                style={{
                  minHeight: 44, width: isMobile ? '100%' : 'auto', padding: '8px 20px', borderRadius: 8, border: 'none', fontSize: 13, fontWeight: 600,
                  cursor: busy || selectedCount === 0 ? 'not-allowed' : 'pointer',
                  background: busy || selectedCount === 0 ? '#fcd34d' : '#d97706', color: '#fff',
                }}
              >
                {busy ? 'Archiving…' : `Archive${selectedCount > 0 ? ` (${selectedCount})` : ''}`}
              </button>
            </div>
          </>
        )}
      </div>
    </div>,
    document.body
  );
}
