import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import type { DeletedNode } from '../types/ui';

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
  return item.kind === 'vessel' || item.item_type === 'vessel';
}

function isFile(item: DeletedNode): boolean {
  return item.kind === 'file' || / file$/i.test(item.item_type || '');
}

export function renderRecycleBinPage(host: VesselEmail): React.ReactElement {
  return <RecycleBinContent host={host} />;
}

function RecycleBinContent({ host }: { host: VesselEmail }): React.ReactElement {
  const { recycleBin, panelLoading } = host.state;
  const [selectedIds, setSelectedIds] = React.useState<Set<string>>(new Set());
  const [permanentDeleteItems, setPermanentDeleteItems] = React.useState<DeletedNode[] | null>(null);
  const [permanentDeleteBusy, setPermanentDeleteBusy] = React.useState(false);
  const [permanentDeleteProgress, setPermanentDeleteProgress] = React.useState<PermanentDeleteProgress[]>([]);
  const [deleteElapsedSeconds, setDeleteElapsedSeconds] = React.useState(0);
  const [autoCloseSeconds, setAutoCloseSeconds] = React.useState<number | null>(null);
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

  const restoreSelected = async (): Promise<void> => {
    const items = recycleBin.filter(item => selectedIds.has(item.id));
    for (const item of items) await host._restoreFromRecycleBin(item);
    setSelectedIds(new Set());
  };

  const permanentlyDeleteSelected = async (): Promise<void> => {
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
    const startedAt = Date.now();
    for (const item of permanentDeleteItems) {
      setPermanentDeleteProgress(previous => previous.map(progress => progress.id === item.id ? { ...progress, status: 'deleting' } : progress));
      const result = await host._permanentDeleteFromRecycleBin(item, true);
      setPermanentDeleteProgress(previous => previous.map(progress => progress.id === item.id ? {
        ...progress,
        status: result.ok ? 'success' : 'failed',
        message: result.message,
      } : progress));
      setDeleteElapsedSeconds(Math.max(1, Math.round((Date.now() - startedAt) / 1000)));
    }
    setPermanentDeleteBusy(false);
    setSelectedIds(new Set());
    setAutoCloseSeconds(10);
    void host._goToView('recycle').catch(() => undefined);
  };

  const cancelPermanentDelete = (): void => {
    if (!permanentDeleteBusy) {
      setPermanentDeleteItems(null);
      setPermanentDeleteProgress([]);
      setAutoCloseSeconds(null);
    }
  };

  const openPermanentDeleteModal = (items: DeletedNode[]): void => {
    if (items.length) {
      setPermanentDeleteItems(items);
      setPermanentDeleteProgress([]);
      setDeleteElapsedSeconds(0);
      setAutoCloseSeconds(null);
    }
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
        <div role="dialog" aria-modal="true" aria-labelledby="permanent-delete-title" onClick={event => event.stopPropagation()} style={{ width: 500, maxWidth: '100%', background: '#fff', borderRadius: 12, boxShadow: '0 20px 50px rgba(15, 23, 42, 0.3)', overflow: 'hidden' }}>
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
      <button onClick={() => host._restoreFromRecycleBin(item)} style={buttonStyle('#dff6dd', '#107c10', '#86efac')}>Restore</button>
      <button onClick={() => openPermanentDeleteModal([item])} style={buttonStyle('#fde7e9', '#a4262c', '#fca5a5')}>Delete permanently</button>
    </div>
  );

  const renderTable = (title: string, items: DeletedNode[], columns: string[], cells: (item: DeletedNode) => React.ReactNode[], empty: string): React.ReactElement => {
    const sectionAllSelected = items.length > 0 && items.every(item => selectedIds.has(item.id));
    return (
      <section>
        <h3 style={{ margin: '0 0 10px', fontSize: 14, color: '#334155' }}>{title} ({items.length})</h3>
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflowX: 'auto' }}>
          <table style={{ width: '100%', minWidth: 760, borderCollapse: 'collapse', fontSize: 13 }}>
            <thead><tr style={{ background: '#f8fafc', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
              <th style={{ padding: 12, width: 34, textAlign: 'center' }}><input type="checkbox" checked={sectionAllSelected} onChange={() => toggleAllInSection(items)} disabled={!items.length} /></th>
              {columns.map(column => <th key={column} style={{ padding: 12 }}>{column}</th>)}
              <th style={{ padding: 12, textAlign: 'right' }}>Actions</th>
            </tr></thead>
            <tbody>
              {!items.length ? <tr><td colSpan={columns.length + 2} style={{ padding: 22, color: '#94a3b8', textAlign: 'center' }}>{empty}</td></tr> : items.map(item => (
                <tr key={item.id} style={{ borderTop: '1px solid #f1f5f9', background: selectedIds.has(item.id) ? '#f0f9ff' : '#fff' }}>
                  <td style={{ padding: 12, textAlign: 'center' }}><input type="checkbox" checked={selectedIds.has(item.id)} onChange={() => toggleOne(item.id)} /></td>
                  {cells(item).map((cell, index) => <td key={index} style={{ padding: 12, color: index === 0 ? '#0f172a' : '#475569', fontWeight: index === 0 ? 600 : 400 }}>{cell}</td>)}
                  <td style={{ padding: 12 }}>{actionCell(item)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    );
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      <div style={{ background: '#fffbe6', border: '1px solid #f59e0b', borderRadius: 12, padding: '18px 22px', display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
        <div><h2 style={{ margin: 0, fontSize: 20, color: '#92400e' }}>Recycle Bin</h2><p style={{ margin: '4px 0 0', color: '#92400e', fontSize: 13 }}>Deleted vessels, folders and files remain here until restored or permanently deleted.</p></div>
        <button onClick={() => host._goToView('recycle').catch(() => undefined)} style={buttonStyle('#fff', '#475569', '#cbd5e1')}>Refresh</button>
      </div>

      {selectedIds.size > 0 && <div style={{ background: '#0f172a', color: '#fff', borderRadius: 10, padding: '12px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
        <span>{selectedIds.size} selected</span>
        <div style={{ display: 'flex', gap: 8 }}><button onClick={restoreSelected} style={buttonStyle('#10b981', '#fff', '#10b981')}>Restore selected</button><button onClick={permanentlyDeleteSelected} style={buttonStyle('#ef4444', '#fff', '#ef4444')}>Delete permanently</button></div>
      </div>}

      {panelLoading ? <div style={{ padding: 48, textAlign: 'center', color: '#64748b' }}>Loading Recycle Bin...</div> : <>
        {renderTable('Deleted vessels', vesselItems, ['Vessel', 'Original path', 'Deleted'], item => [item.name, item.original_path || '—', item.deleted_at ? new Date(item.deleted_at).toLocaleString() : '—'], 'No deleted vessels.')}
        {renderTable('Deleted normal folders', folderItems, ['Folder', 'Original path', 'Vessel', 'Category', 'Sub-category'], item => { const location = getLocationDetails(item); return [item.name, location.folderPath, location.vessel, location.category, location.subCategory]; }, 'No deleted normal folders.')}
        {renderTable('Deleted individual files', fileItems, ['File', 'Folder path', 'Vessel', 'Category', 'Sub-category'], item => { const location = getLocationDetails(item); return [item.name, location.folderPath, location.vessel, location.category, location.subCategory]; }, 'No deleted individual files.')}
      </>}
      {renderPermanentDeleteModal()}
    </div>
  );
}

function buttonStyle(background: string, color: string, border: string): React.CSSProperties {
  return { background, color, border: `1px solid ${border}`, borderRadius: 6, padding: '6px 10px', cursor: 'pointer', fontSize: 12, fontWeight: 600, whiteSpace: 'nowrap' };
}
