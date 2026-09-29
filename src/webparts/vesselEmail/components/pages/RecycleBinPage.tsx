import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import type { DeletedNode } from '../types/ui';
import { isMobileWidth } from '../responsive';

type LocationDetails = {
  folderPath: string;
  vessel: string;
  category: string;
  subCategory: string;
};

type PermanentDeleteProgress = {
  id: string;
  name: string;
  status: 'waiting' | 'deleting' | 'success' | 'failed';
  message?: string;
};

const MAIN_FOLDERS = [
  'Technical & Crewing', 'Commercial & Chartering', 'Insurance',
  'Kaizen - Knowledge Bank', 'Knowledge Bank',
];

function getLocationDetails(item: DeletedNode): LocationDetails {
  const parts = (item.original_path || '')
    .split('/')
    .map(part => part.trim())
    .filter(Boolean);
  if (parts[parts.length - 1] === item.name) parts.pop();

  const mainIndex = parts.findIndex(part => MAIN_FOLDERS.some(main => main.toLowerCase() === part.toLowerCase()));
  let vessel = item.vessel_name || '';
  let category = item.category || '';
  let subCategory = item.sub_category || '';
  let remainder: string[] = [];

  if (mainIndex >= 0) {
    const beforeMain = parts.slice(0, mainIndex);
    const afterMain = parts.slice(mainIndex + 1);
    // Supports both legacy "Main folder/Vessel/..." and current "Vessels/Vessel/Main folder/..." paths.
    if (beforeMain[0] === 'Vessels') {
      vessel = vessel || beforeMain[1] || '';
      remainder = afterMain;
    } else {
      vessel = vessel || afterMain[0] || beforeMain[0] || '';
      remainder = afterMain.slice(1);
    }
    category = category || remainder[0] || parts[mainIndex] || '';
  } else {
    const vesselsIndex = parts.findIndex(part => part.toLowerCase() === 'vessels');
    if (vesselsIndex >= 0) {
      vessel = vessel || parts[vesselsIndex + 1] || '';
      remainder = parts.slice(vesselsIndex + 2);
    } else {
      vessel = vessel || parts[0] || '';
      remainder = parts.slice(1);
    }
    category = category || remainder[0] || '';
  }
  subCategory = subCategory || (remainder.length > 1 ? remainder[remainder.length - 1] : '');

  return {
    folderPath: parts.join(' / ') || '—',
    vessel: vessel || '—',
    category: category || item.main_folder || '—',
    subCategory: subCategory || '—',
  };
}

function isVessel(item: DeletedNode): boolean {
  if (item.kind === 'vessel' || item.item_type === 'vessel') return true;
  const path = (item.original_path || '').replace(/\\/g, '/').toLowerCase();
  if (path.includes('vessels/specific vessels') && !path.includes('drawings') && !path.includes('manuals') && !path.includes('commercial') && !path.includes('technical') && !path.includes('insurance')) {
    const parts = path.split('/').filter(Boolean);
    if (parts.length <= 3 && parts[parts.length - 1] === item.name.toLowerCase()) {
      return true;
    }
  }
  return false;
}

function isFile(item: DeletedNode): boolean {
  if (isVessel(item)) return false;
  if (item.kind === 'file' || item.item_type === 'file' || / file$/i.test(item.item_type || '')) return true;
  if (item.kind === 'folder' || item.item_type === 'folder' || item.item_type === 'File folder') return false;
  return /\.[a-zA-Z0-9]{1,8}$/.test(item.name);
}

export function renderRecycleBinPage(host: VesselEmail): React.ReactElement {
  return <RecycleBinContent host={host} />;
}

function RecycleBinContent({ host }: { host: VesselEmail }): React.ReactElement {
  const { recycleBin, panelLoading } = host.state;
  const viewportWidth = host.state.windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200);
  const isMobile = isMobileWidth(viewportWidth);
  const [selectedIds, setSelectedIds] = React.useState<Set<string>>(new Set());
  const [activeTab, setActiveTab] = React.useState<'vessels' | 'folders' | 'files'>('vessels');
  const [searchQuery, setSearchQuery] = React.useState('');
  const [deletedByFilter, setDeletedByFilter] = React.useState('');

  React.useEffect(() => {
    void host._loadRecycleBin();
  }, []);

  // Permanent Delete modal state
  const [permanentDeleteItems, setPermanentDeleteItems] = React.useState<DeletedNode[] | null>(null);
  const [permanentDeleteBusy, setPermanentDeleteBusy] = React.useState(false);
  const [permanentDeleteProgress, setPermanentDeleteProgress] = React.useState<PermanentDeleteProgress[]>([]);
  const [deleteElapsedSeconds, setDeleteElapsedSeconds] = React.useState(0);
  const [autoCloseSeconds, setAutoCloseSeconds] = React.useState<number | null>(null);

  const openPermanentDeleteModal = (items: DeletedNode[]): void => {
    if (!items.length) return;
    setPermanentDeleteItems(items);
    setPermanentDeleteProgress([]);
    setDeleteElapsedSeconds(0);
    setAutoCloseSeconds(null);
  };

  const cancelPermanentDelete = (): void => {
    if (permanentDeleteBusy) return;
    setPermanentDeleteItems(null);
    setPermanentDeleteProgress([]);
    setAutoCloseSeconds(null);
  };

  // Restore modal state
  const [restoreItems, setRestoreItems] = React.useState<DeletedNode[] | null>(null);
  const [restoreBusy, setRestoreBusy] = React.useState(false);
  const [restoreProgress, setRestoreProgress] = React.useState<Array<{ id: string; name: string; status: 'waiting' | 'restoring' | 'success' | 'failed'; message?: string }>>([]);
  const [restoreElapsedSeconds, setRestoreElapsedSeconds] = React.useState(0);
  const [restoreAutoCloseSeconds, setRestoreAutoCloseSeconds] = React.useState<number | null>(null);

  const vesselItems = recycleBin.filter(isVessel);
  const fileItems = recycleBin.filter(item => !isVessel(item) && isFile(item));
  const folderItems = recycleBin.filter(item => !isVessel(item) && !isFile(item));

  const toggleOne = (id: string): void => {
    setSelectedIds(previous => {
      const next = new Set(previous);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  const toggleAllInSection = (items: DeletedNode[]): void => {
    const ids = items.map(item => item.id);
    const selected = ids.length > 0 && ids.every(id => selectedIds.has(id));
    setSelectedIds(previous => {
      const next = new Set(previous);
      ids.forEach(id => selected ? next.delete(id) : next.add(id));
      return next;
    });
  };

  React.useEffect(() => {
    setSelectedIds(new Set());
  }, [activeTab]);

  // ── Restore Flow ────────────────────────────────────────────────────────────
  const openRestoreModal = (items: DeletedNode[]): void => {
    if (items.length) {
      setRestoreItems(items);
      setRestoreProgress([]);
      setRestoreElapsedSeconds(0);
      setRestoreAutoCloseSeconds(null);
    }
  };

  const cancelRestore = (): void => {
    if (!restoreBusy) {
      setRestoreItems(null);
      setRestoreProgress([]);
      setRestoreAutoCloseSeconds(null);
    }
  };

  const confirmRestore = async (): Promise<void> => {
    if (!restoreItems?.length) return;
    setRestoreBusy(true);
    setRestoreElapsedSeconds(0);
    setRestoreAutoCloseSeconds(null);
    setRestoreProgress(restoreItems.map(item => ({ id: item.id, name: item.name, status: 'waiting' })));
    const startedAt = Date.now();

    for (const item of restoreItems) {
      setRestoreProgress(prev => prev.map(p => p.id === item.id ? { ...p, status: 'restoring' } : p));
      const result = await host._restoreFromRecycleBin(item);
      setRestoreProgress(prev => prev.map(p => p.id === item.id ? {
        ...p,
        status: result.ok ? 'success' : 'failed',
        message: result.message,
      } : p));
      setRestoreElapsedSeconds(Math.max(1, Math.round((Date.now() - startedAt) / 1000)));
    }

    setRestoreBusy(false);
    setSelectedIds(new Set());
    setRestoreAutoCloseSeconds(10);
    void host._goToView('recycle').catch(() => undefined);
  };

  React.useEffect(() => {
    if (!restoreBusy) return undefined;
    const startedAt = Date.now() - restoreElapsedSeconds * 1000;
    const timer = window.setInterval(() => setRestoreElapsedSeconds(Math.floor((Date.now() - startedAt) / 1000)), 250);
    return () => window.clearInterval(timer);
  }, [restoreBusy]);

  React.useEffect(() => {
    if (restoreAutoCloseSeconds === null) return undefined;
    if (restoreAutoCloseSeconds <= 0) {
      setRestoreItems(null);
      setRestoreProgress([]);
      setRestoreAutoCloseSeconds(null);
      return undefined;
    }
    const timer = window.setTimeout(() => setRestoreAutoCloseSeconds(v => v === null ? null : v - 1), 1000);
    return () => window.clearTimeout(timer);
  }, [restoreAutoCloseSeconds]);

  const restoreSelected = (): void => {
    const items = recycleBin.filter(item => selectedIds.has(item.id));
    if (!items.length) return;
    openRestoreModal(items);
  };

  // ── Permanent Delete Flow ───────────────────────────────────────────────────
  const permanentlyDeleteSelected = (): void => {
    const items = recycleBin.filter(item => selectedIds.has(item.id));
    if (!items.length) return;
    openPermanentDeleteModal(items);
  };

  const confirmPermanentDelete = async (): Promise<void> => {
    if (!permanentDeleteItems?.length) return;
    setPermanentDeleteBusy(true);
    setDeleteElapsedSeconds(0);
    setAutoCloseSeconds(null);
    setPermanentDeleteProgress(permanentDeleteItems.map(item => ({ id: item.id, name: item.name, status: 'waiting' })));

    for (const item of permanentDeleteItems) {
      setPermanentDeleteProgress(previous => previous.map(progress =>
        progress.id === item.id ? { ...progress, status: 'deleting' } : progress
      ));

      const result = await host._permanentDeleteFromRecycleBin(item, true);

      setPermanentDeleteProgress(previous => previous.map(progress =>
        progress.id === item.id
          ? { ...progress, status: result.ok ? 'success' : 'failed', message: result.message }
          : progress
      ));
    }

    setPermanentDeleteBusy(false);
    setSelectedIds(new Set());
    await host._loadRecycleBin();
    setAutoCloseSeconds(10);
  };

  React.useEffect(() => {
    if (!permanentDeleteBusy) return undefined;
    const startedAt = Date.now() - deleteElapsedSeconds * 1000;
    const timer = window.setInterval(() => setDeleteElapsedSeconds(Math.floor((Date.now() - startedAt) / 1000)), 250);
    return () => window.clearInterval(timer);
  }, [permanentDeleteBusy]);

  React.useEffect(() => {
    if (autoCloseSeconds === null) return undefined;
    if (autoCloseSeconds <= 0) {
      setPermanentDeleteItems(null);
      setPermanentDeleteProgress([]);
      setAutoCloseSeconds(null);
      return undefined;
    }
    const timer = window.setTimeout(() => setAutoCloseSeconds(value => value === null ? null : value - 1), 1000);
    return () => window.clearTimeout(timer);
  }, [autoCloseSeconds]);

  // ── Render Restore Modal ────────────────────────────────────────────────────
  const renderRestoreModal = (): React.ReactElement | null => {
    if (!restoreItems?.length) return null;
    const itemCount = restoreItems.length;
    const completed = !restoreBusy && restoreProgress.length > 0;
    const successCount = restoreProgress.filter(item => item.status === 'success').length;
    const failedCount = restoreProgress.filter(item => item.status === 'failed').length;

    return (
      <div
        role="presentation"
        onClick={cancelRestore}
        style={{
          position: 'fixed', inset: 0, zIndex: 10000,
          background: 'rgba(15, 23, 42, 0.55)', backdropFilter: 'blur(3px)',
          display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
        }}
      >
        <div
          role="dialog"
          aria-modal="true"
          aria-labelledby="restore-modal-title"
          onClick={event => event.stopPropagation()}
          style={{
            width: 520, maxWidth: '100%', background: '#fff',
            borderRadius: 16, boxShadow: '0 25px 60px rgba(15, 23, 42, 0.35)',
            overflow: 'hidden', border: '1px solid #d1fae5', maxHeight: '90vh',
          }}
        >
          {/* Header */}
          <div style={{
            padding: '20px 24px 16px',
            background: completed ? '#ecfdf5' : restoreBusy ? '#f0fdf4' : '#f8fafc',
            borderBottom: '1px solid #e2e8f0',
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <span style={{ fontSize: 24 }}>{completed ? (failedCount ? '⚠️' : '🎉') : restoreBusy ? '⏳' : '♻️'}</span>
              <div>
                <div style={{ fontSize: 18, fontWeight: 700, color: completed ? (failedCount ? '#b45309' : '#059669') : '#0f172a' }} id="restore-modal-title">
                  {completed
                    ? (failedCount ? 'Restored with Some Warnings' : 'Restored Successfully!')
                    : restoreBusy
                    ? `Restoring ${itemCount} Item${itemCount === 1 ? '' : 's'}…`
                    : `Restore ${itemCount} Item${itemCount === 1 ? '' : 's'} to SharePoint?`}
                </div>
                <p style={{ margin: '4px 0 0', color: '#475569', fontSize: 13, lineHeight: 1.4 }}>
                  {completed
                    ? `${successCount} item${successCount === 1 ? '' : 's'} moved from SharePoint Recycle Bin back to original folder path.`
                    : restoreBusy
                    ? `Moving folder and file items back to original SharePoint locations. Elapsed time: ${restoreElapsedSeconds}s.`
                    : `Items will be restored from the SharePoint Recycle Bin to their respective department folders and reactivated in the DMS.`}
                </p>
              </div>
            </div>
          </div>

          {/* Body / Item List */}
          {restoreProgress.length ? (
            <div style={{ margin: '16px 24px', maxHeight: 210, overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: 10 }}>
              {restoreProgress.map(p => (
                <div key={p.id} style={{ display: 'flex', gap: 10, alignItems: 'center', padding: '10px 12px', borderBottom: '1px solid #f1f5f9', fontSize: 13 }}>
                  <span style={{
                    width: 22, height: 22, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center',
                    background: p.status === 'success' ? '#dcfce7' : p.status === 'failed' ? '#fee2e2' : p.status === 'restoring' ? '#e0f2fe' : '#f1f5f9',
                    color: p.status === 'success' ? '#15803d' : p.status === 'failed' ? '#dc2626' : p.status === 'restoring' ? '#0284c7' : '#94a3b8',
                    fontWeight: 700, fontSize: 12,
                  }}>
                    {p.status === 'success' ? '✓' : p.status === 'failed' ? '✗' : p.status === 'restoring' ? '⏳' : '○'}
                  </span>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ color: '#1e293b', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{p.name}</div>
                    {p.message && <div style={{ fontSize: 11, color: p.status === 'success' ? '#059669' : '#b91c1c' }}>{p.message}</div>}
                  </div>
                  <span style={{
                    fontSize: 12, fontWeight: 600,
                    color: p.status === 'success' ? '#15803d' : p.status === 'failed' ? '#dc2626' : p.status === 'restoring' ? '#0284c7' : '#94a3b8',
                  }}>
                    {p.status === 'success' ? 'Restored' : p.status === 'failed' ? 'Failed' : p.status === 'restoring' ? 'Restoring…' : 'Waiting'}
                  </span>
                </div>
              ))}
            </div>
          ) : (
            <div style={{ margin: '16px 24px', maxHeight: 200, overflowY: 'auto' }}>
              {restoreItems.map(item => {
                const location = getLocationDetails(item);
                return (
                  <div key={item.id} style={{ padding: '10px 14px', background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 10, marginBottom: 8, fontSize: 12 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 4 }}>
                      <strong style={{ fontSize: 13, color: '#0f172a' }}>{item.name}</strong>
                      <span style={{ background: '#dbeafe', color: '#1d4ed8', borderRadius: 6, padding: '2px 8px', fontSize: 11, fontWeight: 700, textTransform: 'capitalize' }}>
                        {item.kind || item.item_type}
                      </span>
                    </div>
                    <div style={{ color: '#64748b' }}>
                      <strong>Original Location:</strong> <span style={{ color: '#0369a1' }}>{location.folderPath}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Footer Buttons */}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, padding: '12px 24px 20px', background: '#fff' }}>
            {completed ? (
              <button
                onClick={cancelRestore}
                style={{
                  background: 'linear-gradient(135deg,#10b981,#059669)',
                  color: '#fff', border: 'none', borderRadius: 8, padding: '9px 20px',
                  fontSize: 13, fontWeight: 700, cursor: 'pointer',
                  boxShadow: '0 3px 8px rgba(16,185,129,0.3)',
                }}
              >
                Close now{restoreAutoCloseSeconds !== null ? ` (${restoreAutoCloseSeconds}s)` : ''}
              </button>
            ) : (
              <>
                <button
                  onClick={cancelRestore}
                  disabled={restoreBusy}
                  style={buttonStyle('#fff', '#475569', '#cbd5e1')}
                >
                  Cancel
                </button>
                <button
                  onClick={confirmRestore}
                  disabled={restoreBusy}
                  style={{
                    background: 'linear-gradient(135deg,#10b981,#059669)',
                    color: '#fff', border: 'none', borderRadius: 8, padding: '8px 20px',
                    fontSize: 13, fontWeight: 700, cursor: restoreBusy ? 'not-allowed' : 'pointer',
                    boxShadow: '0 3px 8px rgba(16,185,129,0.3)',
                    opacity: restoreBusy ? 0.65 : 1,
                  }}
                >
                  {restoreBusy ? 'Restoring…' : `♻️ Restore ${itemCount === 1 ? 'Item' : `${itemCount} Items`}`}
                </button>
              </>
            )}
          </div>
        </div>
      </div>
    );
  };

  // ── Render Permanent Delete Modal ───────────────────────────────────────────
  const renderPermanentDeleteModal = (): React.ReactElement | null => {
    if (!permanentDeleteItems?.length) return null;
    const itemCount = permanentDeleteItems.length;
    const itemNames = permanentDeleteItems.slice(0, 3).map(item => item.name);
    const extraCount = itemCount - itemNames.length;
    const completed = !permanentDeleteBusy && permanentDeleteProgress.length > 0;
    const successCount = permanentDeleteProgress.filter(item => item.status === 'success').length;
    const failedCount = permanentDeleteProgress.filter(item => item.status === 'failed').length;
    return (
      <div
        role="presentation"
        onClick={cancelPermanentDelete}
        style={{ position: 'fixed', inset: 0, zIndex: 10000, background: 'rgba(15, 23, 42, 0.52)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}
      >
        <div role="dialog" aria-modal="true" aria-labelledby="permanent-delete-title" onClick={event => event.stopPropagation()} style={{ width: 500, maxWidth: '100%', maxHeight: '90vh', background: '#fff', borderRadius: 12, boxShadow: '0 20px 50px rgba(15, 23, 42, 0.3)', overflow: 'hidden' }}>
          <div style={{ padding: '20px 22px 14px', borderBottom: '1px solid #fee2e2' }}>
            <div style={{ fontSize: 17, fontWeight: 700, color: completed ? '#107c10' : '#991b1b' }} id="permanent-delete-title">
              {completed ? 'Permanent deletion completed' : permanentDeleteBusy ? 'Permanently deleting items' : `Permanently delete ${itemCount} item${itemCount === 1 ? '' : 's'}?`}
            </div>
            <p style={{ margin: '10px 0 0', color: '#475569', fontSize: 13, lineHeight: 1.5 }}>
              {completed ? `${successCount} item${successCount === 1 ? '' : 's'} permanently deleted${failedCount ? `; ${failedCount} could not be deleted.` : '.'}` : permanentDeleteBusy ? `Deleting one item at a time. Elapsed time: ${deleteElapsedSeconds}s.` : `This action cannot be undone. The selected item${itemCount === 1 ? '' : 's'} will be permanently removed from the Recycle Bin.`}
            </p>
          </div>
          {permanentDeleteProgress.length ? (
            <div style={{ margin: '14px 22px', maxHeight: 190, overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: 7 }}>
              {permanentDeleteProgress.map(progress => <div key={progress.id} style={{ display: 'flex', gap: 8, alignItems: 'center', padding: '9px 10px', borderBottom: '1px solid #f1f5f9', fontSize: 12 }}>
                <span style={{ width: 18, textAlign: 'center' }}>{progress.status === 'success' ? '✓' : progress.status === 'failed' ? '!' : progress.status === 'deleting' ? '…' : '○'}</span>
                <span style={{ flex: 1, color: '#334155' }}>{progress.name}</span>
                <span style={{ color: progress.status === 'success' ? '#107c10' : progress.status === 'failed' ? '#a4262c' : '#64748b' }}>{progress.status === 'success' ? 'Deleted' : progress.status === 'failed' ? progress.message || 'Failed' : progress.status === 'deleting' ? 'Deleting...' : 'Waiting'}</span>
              </div>)}
            </div>
          ) : <div style={{ margin: '14px 22px', padding: '10px 12px', background: '#f8fafc', borderRadius: 7, color: '#475569', fontSize: 12 }}><strong style={{ color: '#334155' }}>{itemNames.join(', ')}</strong>{extraCount > 0 ? ` and ${extraCount} more` : ''}</div>}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, padding: '0 22px 20px' }}>
            {completed ? <button onClick={cancelPermanentDelete} style={buttonStyle('#fff', '#475569', '#cbd5e1')}>Close now{autoCloseSeconds !== null ? ` (${autoCloseSeconds}s)` : ''}</button> : <><button onClick={cancelPermanentDelete} disabled={permanentDeleteBusy} style={buttonStyle('#fff', '#475569', '#cbd5e1')}>Cancel</button><button onClick={confirmPermanentDelete} disabled={permanentDeleteBusy} style={{ ...buttonStyle('#dc2626', '#fff', '#dc2626'), opacity: permanentDeleteBusy ? 0.65 : 1 }}>{permanentDeleteBusy ? 'Deleting...' : 'Delete permanently'}</button></>}
          </div>
        </div>
      </div>
    );
  };

  const actionCell = (item: DeletedNode): React.ReactElement => (
    <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
      <button onClick={() => openRestoreModal([item])} style={buttonStyle('#dff6dd', '#107c10', '#86efac')}>Restore</button>
      <button onClick={() => openPermanentDeleteModal([item])} style={buttonStyle('#fde7e9', '#a4262c', '#fca5a5')}>Delete permanently</button>
    </div>
  );

  const renderTable = (title: string, items: DeletedNode[], columns: string[], cells: (item: DeletedNode) => React.ReactNode[], empty: string): React.ReactElement => {
    let filteredItems = searchQuery.trim()
      ? items.filter(item => {
          const q = searchQuery.toLowerCase();
          const nameMatch = item.name.toLowerCase().includes(q);
          const pathMatch = (item.original_path || '').toLowerCase().includes(q);
          const vesselMatch = (item.vessel_name || '').toLowerCase().includes(q);
          const siteMatch = (item.site_name || item.site_key || '').toLowerCase().includes(q);
          const deletedByMatch = ((item.deleted_by_name || '') + ' ' + (item.deleted_by_email || '')).toLowerCase().includes(q);
          return nameMatch || pathMatch || vesselMatch || siteMatch || deletedByMatch;
        })
      : items;
    if (deletedByFilter.trim()) {
      const q = deletedByFilter.trim().toLowerCase();
      filteredItems = filteredItems.filter(item =>
        ((item.deleted_by_name || '') + ' ' + (item.deleted_by_email || '')).toLowerCase().includes(q)
      );
    }

    const sectionAllSelected = filteredItems.length > 0 && filteredItems.every(item => selectedIds.has(item.id));
    return (
      <section style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 10 }}>
          <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: '#1e293b' }}>
            {title} ({filteredItems.length}{searchQuery.trim() && filteredItems.length !== items.length ? ` of ${items.length}` : ''})
          </h3>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <input
              type="text"
              placeholder={`Search ${title.toLowerCase()}...`}
              value={searchQuery}
              onChange={e => setSearchQuery(e.target.value)}
              style={{
                padding: '6px 12px', borderRadius: 8, border: '1px solid #cbd5e1',
                fontSize: 12, width: 220, outline: 'none',
              }}
            />
            <input
              type="text"
              placeholder="Filter by Deleted By..."
              value={deletedByFilter}
              onChange={e => setDeletedByFilter(e.target.value)}
              style={{
                padding: '6px 12px', borderRadius: 8, border: '1px solid #cbd5e1',
                fontSize: 12, width: 180, outline: 'none',
              }}
            />
          </div>
        </div>

        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 12, overflowX: 'auto', boxShadow: '0 1px 3px rgba(0,0,0,0.03)' }}>
          {isMobile ? (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: 10 }}>
              {!filteredItems.length ? (
                <div style={{ padding: 24, color: '#94a3b8', textAlign: 'center', fontSize: 13 }}>{empty}</div>
              ) : filteredItems.map(item => {
                const location = getLocationDetails(item);
                return (
                  <div key={item.id} style={{ border: '1px solid #e2e8f0', borderRadius: 10, padding: 12, background: selectedIds.has(item.id) ? '#f0f9ff' : '#fff' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
                      <div style={{ fontWeight: 700, color: '#0f172a', fontSize: 13 }}>{item.name}</div>
                      <input type="checkbox" checked={selectedIds.has(item.id)} onChange={() => toggleOne(item.id)} />
                    </div>
                    <div style={{ marginTop: 8, fontSize: 12, color: '#475569', lineHeight: 1.5 }}>
                      <div><strong>Path:</strong> {location.folderPath}</div>
                      <div><strong>Site:</strong> {item.site_name || item.site_key || '—'}</div>
                      <div><strong>Vessel:</strong> {location.vessel}</div>
                      <div><strong>Category:</strong> {location.category}</div>
                      <div><strong>Sub-category:</strong> {location.subCategory}</div>
                      <div><strong>Deleted by:</strong> {item.deleted_by_name || item.deleted_by_email || '—'}</div>
                      {item.reason && <div><strong>Reason:</strong> {item.reason}</div>}
                    </div>
                    <div style={{ marginTop: 10, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                      <button onClick={() => openRestoreModal([item])} style={buttonStyle('#dff6dd', '#107c10', '#86efac')}>Restore</button>
                      <button onClick={() => openPermanentDeleteModal([item])} style={buttonStyle('#fde7e9', '#a4262c', '#fca5a5')}>Delete</button>
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
          <table style={{ width: '100%', minWidth: 760, borderCollapse: 'collapse', fontSize: 13 }}>
            <thead><tr style={{ background: '#f8fafc', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left', borderBottom: '1px solid #e2e8f0' }}>
              <th style={{ padding: 12, width: 34, textAlign: 'center' }}>
                <input type="checkbox" checked={sectionAllSelected} onChange={() => toggleAllInSection(filteredItems)} disabled={!filteredItems.length} />
              </th>
              {columns.map(column => <th key={column} style={{ padding: 12 }}>{column}</th>)}
              <th style={{ padding: 12 }}>Deleted By</th>
              <th style={{ padding: 12 }}>Reason for deletion</th>
              <th style={{ padding: 12, textAlign: 'right' }}>Actions</th>
            </tr></thead>
            <tbody>
              {!filteredItems.length ? (
                <tr><td colSpan={columns.length + 4} style={{ padding: 36, color: '#94a3b8', textAlign: 'center', fontSize: 13 }}>{empty}</td></tr>
              ) : filteredItems.map(item => (
                <tr key={item.id} style={{ borderTop: '1px solid #f1f5f9', background: selectedIds.has(item.id) ? '#f0f9ff' : '#fff' }}>
                  <td style={{ padding: 12, textAlign: 'center' }}>
                    <input type="checkbox" checked={selectedIds.has(item.id)} onChange={() => toggleOne(item.id)} />
                  </td>
                  {cells(item).map((cell, index) => <td key={index} style={{ padding: 12, color: index === 0 ? '#0f172a' : '#475569', fontWeight: index === 0 ? 600 : 400 }}>{cell}</td>)}
                  <td style={{ padding: 12, color: '#475569' }} title={item.deleted_by_email || ''}>
                    {item.deleted_by_name || item.deleted_by_email || '—'}
                  </td>
                  <td style={{ padding: 12, color: item.reason ? '#475569' : '#cbd5e1', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={item.reason || ''}>
                    {item.reason || '—'}
                  </td>
                  <td style={{ padding: 12 }}>{actionCell(item)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          )}
        </div>
      </section>
    );
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      {/* Header Banner */}
      <div style={{ background: '#fffbe6', border: '1px solid #f59e0b', borderRadius: 12, padding: '18px 22px', display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, color: '#92400e', fontWeight: 800 }}>Recycle Bin</h2>
          <p style={{ margin: '4px 0 0', color: '#92400e', fontSize: 13 }}>Deleted vessels, folders and files remain here until restored or permanently deleted.</p>
        </div>
        <button onClick={() => host._goToView('recycle').catch(() => undefined)} style={buttonStyle('#fff', '#475569', '#cbd5e1')}>🔄 Refresh</button>
      </div>

      {/* Compact Section Tabs */}
      <div style={{
        display: 'flex',
        flexWrap: 'wrap',
        gap: 8,
        background: '#fff',
        border: '1px solid #e2e8f0',
        borderRadius: 12,
        padding: 8,
      }}>
        <button
          onClick={() => { setActiveTab('vessels'); setSearchQuery(''); }}
          style={{
            ...buttonStyle(activeTab === 'vessels' ? '#0284c7' : '#fff', activeTab === 'vessels' ? '#fff' : '#334155', activeTab === 'vessels' ? '#0284c7' : '#cbd5e1'),
            padding: '8px 12px',
            borderRadius: 10,
            fontSize: 12,
            fontWeight: 700,
          }}
        >
          🚢 Deleted Vessels ({vesselItems.length})
        </button>

        <button
          onClick={() => { setActiveTab('folders'); setSearchQuery(''); }}
          style={{
            ...buttonStyle(activeTab === 'folders' ? '#0284c7' : '#fff', activeTab === 'folders' ? '#fff' : '#334155', activeTab === 'folders' ? '#0284c7' : '#cbd5e1'),
            padding: '8px 12px',
            borderRadius: 10,
            fontSize: 12,
            fontWeight: 700,
          }}
        >
          📁 Deleted Normal Folders ({folderItems.length})
        </button>

        <button
          onClick={() => { setActiveTab('files'); setSearchQuery(''); }}
          style={{
            ...buttonStyle(activeTab === 'files' ? '#0284c7' : '#fff', activeTab === 'files' ? '#fff' : '#334155', activeTab === 'files' ? '#0284c7' : '#cbd5e1'),
            padding: '8px 12px',
            borderRadius: 10,
            fontSize: 12,
            fontWeight: 700,
          }}
        >
          📄 Deleted Individual Files ({fileItems.length})
        </button>
      </div>

      {/* Bulk Action Bar */}
      {selectedIds.size > 0 && (
        <div style={{ background: '#0f172a', color: '#fff', borderRadius: 10, padding: '12px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
          <span style={{ fontWeight: 600 }}>{selectedIds.size} item{selectedIds.size === 1 ? '' : 's'} selected</span>
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={restoreSelected} style={buttonStyle('#10b981', '#fff', '#10b981')}>♻️ Restore selected</button>
            <button onClick={permanentlyDeleteSelected} style={buttonStyle('#ef4444', '#fff', '#ef4444')}>🗑️ Delete permanently</button>
          </div>
        </div>
      )}

      {/* Active Section Table */}
      {panelLoading ? (
        <div style={{ padding: 48, textAlign: 'center', color: '#64748b', fontSize: 14 }}>
          ⏳ Loading Recycle Bin items...
        </div>
      ) : (
        <>
          {activeTab === 'vessels' && renderTable(
            'Deleted vessels',
            vesselItems,
            ['Vessel', 'Site', 'Original path', 'Deleted'],
            item => [item.name, item.site_name || item.site_key || '—', item.original_path || '—', item.deleted_at ? new Date(item.deleted_at).toLocaleString() : '—'],
            'No deleted vessels found in Recycle Bin.'
          )}

          {activeTab === 'folders' && renderTable(
            'Deleted normal folders',
            folderItems,
            ['Folder', 'Site', 'Original path', 'Vessel', 'Category', 'Sub-category'],
            item => {
              const location = getLocationDetails(item);
              return [item.name, item.site_name || item.site_key || '—', location.folderPath, location.vessel, location.category, location.subCategory];
            },
            'No deleted normal folders found in Recycle Bin.'
          )}

          {activeTab === 'files' && renderTable(
            'Deleted individual files',
            fileItems,
            ['File', 'Site', 'Folder path', 'Vessel', 'Category', 'Sub-category'],
            item => {
              const location = getLocationDetails(item);
              return [item.name, item.site_name || item.site_key || '—', location.folderPath, location.vessel, location.category, location.subCategory];
            },
            'No deleted individual files found in Recycle Bin.'
          )}
        </>
      )}

      {renderRestoreModal()}
      {renderPermanentDeleteModal()}
    </div>
  );
}

function buttonStyle(background: string, color: string, border: string): React.CSSProperties {
  return { background, color, border: `1px solid ${border}`, borderRadius: 6, minHeight: 44, padding: '6px 10px', cursor: 'pointer', fontSize: 12, fontWeight: 600, whiteSpace: 'nowrap' };
}

