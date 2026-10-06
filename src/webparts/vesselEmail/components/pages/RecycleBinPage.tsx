import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import type { DeletedNode } from '../types/ui';
import { isMobileWidth } from '../responsive';
import { Icon } from '@fluentui/react/lib/Icon';
import { clay } from '../clayTheme';
import {
  DmsPageHeader, dmsBtn, dmsRowBtn, dmsControlStyle, DMS_ON_ACCENT,
  DMS_TABLE_CARD, DMS_TABLE, DMS_TH, DMS_TR, DMS_TD,
} from '../dmsDesignSystem';

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
            width: 520, maxWidth: '100%', background: 'var(--vdms-surface)',
            borderRadius: 16, boxShadow: clay.shadowRaised,
            overflow: 'hidden', border: '1px solid var(--vdms-line)', maxHeight: '90vh',
          }}
        >
          {/* Header */}
          <div style={{
            padding: '20px 24px 16px',
            background: completed ? (failedCount ? clay.pillWarnBg : clay.pillActiveBg) : restoreBusy ? clay.accentSoft : 'var(--vdms-surface-alt)',
            borderBottom: '1px solid var(--vdms-line)',
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <span style={{ fontSize: 24 }}>{completed ? (failedCount ? <Icon iconName="Warning" aria-hidden="true" /> : <Icon iconName="Completed" aria-hidden="true" />) : restoreBusy ? <Icon iconName="Sync" aria-hidden="true" /> : <Icon iconName="RecycleBin" aria-hidden="true" />}</span>
              <div>
                <div style={{ fontSize: 18, fontWeight: 700, color: completed ? (failedCount ? clay.pillWarnText : clay.pillActiveText) : 'var(--vdms-text)' }} id="restore-modal-title">
                  {completed
                    ? (failedCount ? 'Restored with Some Warnings' : 'Restored Successfully!')
                    : restoreBusy
                    ? `Restoring ${itemCount} Item${itemCount === 1 ? '' : 's'}…`
                    : `Restore ${itemCount} Item${itemCount === 1 ? '' : 's'} to SharePoint?`}
                </div>
                <p style={{ margin: '4px 0 0', color: 'var(--vdms-text-muted)', fontSize: 13, lineHeight: 1.4 }}>
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
            <div style={{ margin: '16px 24px', maxHeight: 210, overflowY: 'auto', border: '1px solid var(--vdms-line)', borderRadius: 10 }}>
              {restoreProgress.map(p => (
                <div key={p.id} style={{ display: 'flex', gap: 10, alignItems: 'center', padding: '10px 12px', borderBottom: '1px solid var(--vdms-border-soft)', fontSize: 13 }}>
                  <span style={{
                    width: 22, height: 22, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center',
                    background: p.status === 'success' ? clay.pillActiveBg : p.status === 'failed' ? clay.pillDangerBg : p.status === 'restoring' ? clay.accentSoft : 'var(--vdms-surface-alt)',
                    color: p.status === 'success' ? clay.pillActiveText : p.status === 'failed' ? clay.pillDangerText : p.status === 'restoring' ? clay.accent : 'var(--vdms-text-faint)',
                    fontWeight: 700, fontSize: 12,
                  }}>
                    {p.status === 'success' ? <Icon iconName="CheckMark" aria-hidden="true" /> : p.status === 'failed' ? <Icon iconName="ErrorBadge" aria-hidden="true" /> : p.status === 'restoring' ? <Icon iconName="Sync" aria-hidden="true" /> : <Icon iconName="CircleRing" aria-hidden="true" />}
                  </span>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ color: 'var(--vdms-text)', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{p.name}</div>
                    {p.message && <div style={{ fontSize: 11, color: p.status === 'success' ? clay.pillActiveText : clay.pillDangerText }}>{p.message}</div>}
                  </div>
                  <span style={{
                    fontSize: 12, fontWeight: 600,
                    color: p.status === 'success' ? clay.pillActiveText : p.status === 'failed' ? clay.pillDangerText : p.status === 'restoring' ? clay.accent : 'var(--vdms-text-faint)',
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
                  <div key={item.id} style={{ padding: '10px 14px', background: 'var(--vdms-surface-alt)', border: '1px solid var(--vdms-line)', borderRadius: 10, marginBottom: 8, fontSize: 12 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 4 }}>
                      <strong style={{ fontSize: 13, color: 'var(--vdms-text)' }}>{item.name}</strong>
                      <span style={{ background: clay.accentSoft, color: clay.accentDark, borderRadius: 6, padding: '2px 8px', fontSize: 11, fontWeight: 700, textTransform: 'capitalize' }}>
                        {item.kind || item.item_type}
                      </span>
                    </div>
                    <div style={{ color: 'var(--vdms-text-muted)' }}>
                      <strong>Original Location:</strong> <span style={{ color: clay.accent }}>{location.folderPath}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Footer Buttons */}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, padding: '12px 24px 20px', background: 'var(--vdms-surface)' }}>
            {completed ? (
              <button
                onClick={cancelRestore}
                style={dmsBtn('primary', true)}
              >
                Close now{restoreAutoCloseSeconds !== null ? ` (${restoreAutoCloseSeconds}s)` : ''}
              </button>
            ) : (
              <>
                <button
                  onClick={cancelRestore}
                  disabled={restoreBusy}
                  style={dmsBtn('secondary', !restoreBusy)}
                >
                  Cancel
                </button>
                <button
                  onClick={confirmRestore}
                  disabled={restoreBusy}
                  style={dmsBtn('primary', !restoreBusy)}
                >
                  {restoreBusy ? 'Restoring…' : <><Icon iconName="RecycleBin" aria-hidden="true" style={{ fontSize: 13 }} /> {`Restore ${itemCount === 1 ? 'Item' : `${itemCount} Items`}`}</>}
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
        <div role="dialog" aria-modal="true" aria-labelledby="permanent-delete-title" onClick={event => event.stopPropagation()} style={{ width: 500, maxWidth: '100%', maxHeight: '90vh', background: 'var(--vdms-surface)', borderRadius: 12, boxShadow: clay.shadowRaised, overflow: 'hidden', border: '1px solid var(--vdms-line)' }}>
          <div style={{ padding: '20px 22px 14px', borderBottom: `1px solid ${clay.pillDangerBg}` }}>
            <div style={{ fontSize: 17, fontWeight: 700, color: completed ? clay.pillActiveText : clay.pillDangerText }} id="permanent-delete-title">
              {completed ? 'Permanent deletion completed' : permanentDeleteBusy ? 'Permanently deleting items' : `Permanently delete ${itemCount} item${itemCount === 1 ? '' : 's'}?`}
            </div>
            <p style={{ margin: '10px 0 0', color: 'var(--vdms-text-muted)', fontSize: 13, lineHeight: 1.5 }}>
              {completed ? `${successCount} item${successCount === 1 ? '' : 's'} permanently deleted${failedCount ? `; ${failedCount} could not be deleted.` : '.'}` : permanentDeleteBusy ? `Deleting one item at a time. Elapsed time: ${deleteElapsedSeconds}s.` : `This action cannot be undone. The selected item${itemCount === 1 ? '' : 's'} will be permanently removed from the Recycle Bin.`}
            </p>
          </div>
          {permanentDeleteProgress.length ? (
            <div style={{ margin: '14px 22px', maxHeight: 190, overflowY: 'auto', border: '1px solid var(--vdms-line)', borderRadius: 7 }}>
              {permanentDeleteProgress.map(progress => <div key={progress.id} style={{ display: 'flex', gap: 8, alignItems: 'center', padding: '9px 10px', borderBottom: '1px solid var(--vdms-border-soft)', fontSize: 12 }}>
                <span style={{ width: 18, textAlign: 'center' }}>{progress.status === 'success' ? <Icon iconName="CheckMark" aria-hidden="true" /> : progress.status === 'failed' ? '!' : progress.status === 'deleting' ? '…' : <Icon iconName="CircleRing" aria-hidden="true" />}</span>
                <span style={{ flex: 1, color: 'var(--vdms-text)' }}>{progress.name}</span>
                <span style={{ color: progress.status === 'success' ? clay.pillActiveText : progress.status === 'failed' ? clay.pillDangerText : 'var(--vdms-text-muted)' }}>{progress.status === 'success' ? 'Deleted' : progress.status === 'failed' ? progress.message || 'Failed' : progress.status === 'deleting' ? 'Deleting...' : 'Waiting'}</span>
              </div>)}
            </div>
          ) : <div style={{ margin: '14px 22px', padding: '10px 12px', background: 'var(--vdms-surface-alt)', borderRadius: 7, color: 'var(--vdms-text-muted)', fontSize: 12 }}><strong style={{ color: 'var(--vdms-text)' }}>{itemNames.join(', ')}</strong>{extraCount > 0 ? ` and ${extraCount} more` : ''}</div>}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, padding: '0 22px 20px' }}>
            {completed ? <button onClick={cancelPermanentDelete} style={dmsBtn('secondary', true)}>Close now{autoCloseSeconds !== null ? ` (${autoCloseSeconds}s)` : ''}</button> : <><button onClick={cancelPermanentDelete} disabled={permanentDeleteBusy} style={dmsBtn('secondary', !permanentDeleteBusy)}>Cancel</button><button onClick={confirmPermanentDelete} disabled={permanentDeleteBusy} style={dmsBtn('danger', !permanentDeleteBusy)}>{permanentDeleteBusy ? 'Deleting...' : 'Delete permanently'}</button></>}
          </div>
        </div>
      </div>
    );
  };

  const actionCell = (item: DeletedNode): React.ReactElement => (
    <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
      <button onClick={() => openRestoreModal([item])} style={dmsRowBtn('success')}>Restore</button>
      <button onClick={() => openPermanentDeleteModal([item])} style={dmsRowBtn('danger')}>Delete permanently</button>
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
          <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: 'var(--vdms-text)' }}>
            {title} ({filteredItems.length}{searchQuery.trim() && filteredItems.length !== items.length ? ` of ${items.length}` : ''})
          </h3>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <input
              type="text"
              placeholder={`Search ${title.toLowerCase()}...`}
              value={searchQuery}
              onChange={e => setSearchQuery(e.target.value)}
              style={{ ...dmsControlStyle(), width: 220 }}
            />
            <input
              type="text"
              placeholder="Filter by Deleted By..."
              value={deletedByFilter}
              onChange={e => setDeletedByFilter(e.target.value)}
              style={{ ...dmsControlStyle(), width: 180 }}
            />
          </div>
        </div>

        <div style={{ ...DMS_TABLE_CARD, overflowX: 'auto' }}>
          {isMobile ? (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: 10 }}>
              {!filteredItems.length ? (
                <div style={{ padding: 24, color: 'var(--vdms-text-faint)', textAlign: 'center', fontSize: 13 }}>{empty}</div>
              ) : filteredItems.map(item => {
                const location = getLocationDetails(item);
                return (
                  <div key={item.id} style={{ border: '1px solid var(--vdms-line)', borderRadius: 10, padding: 12, background: selectedIds.has(item.id) ? clay.accentSoft : 'var(--vdms-surface)' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
                      <div style={{ fontWeight: 700, color: 'var(--vdms-text)', fontSize: 13 }}>{item.name}</div>
                      <input type="checkbox" checked={selectedIds.has(item.id)} onChange={() => toggleOne(item.id)} />
                    </div>
                    <div style={{ marginTop: 8, fontSize: 12, color: 'var(--vdms-text-muted)', lineHeight: 1.5 }}>
                      <div><strong>Path:</strong> {location.folderPath}</div>
                      <div><strong>Site:</strong> {item.site_name || item.site_key || '—'}</div>
                      <div><strong>Vessel:</strong> {location.vessel}</div>
                      <div><strong>Category:</strong> {location.category}</div>
                      <div><strong>Sub-category:</strong> {location.subCategory}</div>
                      <div><strong>Deleted by:</strong> {item.deleted_by_name || item.deleted_by_email || '—'}</div>
                      {item.reason && <div><strong>Reason:</strong> {item.reason}</div>}
                    </div>
                    <div style={{ marginTop: 10, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                      <button onClick={() => openRestoreModal([item])} style={dmsRowBtn('success')}>Restore</button>
                      <button onClick={() => openPermanentDeleteModal([item])} style={dmsRowBtn('danger')}>Delete</button>
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
          <table style={{ ...DMS_TABLE, minWidth: 760 }}>
            <thead><tr>
              <th style={{ ...DMS_TH, width: 34, textAlign: 'center' }}>
                <input type="checkbox" checked={sectionAllSelected} onChange={() => toggleAllInSection(filteredItems)} disabled={!filteredItems.length} />
              </th>
              {columns.map(column => <th key={column} style={DMS_TH}>{column}</th>)}
              <th style={DMS_TH}>Deleted By</th>
              <th style={DMS_TH}>Reason for deletion</th>
              <th style={{ ...DMS_TH, textAlign: 'right' }}>Actions</th>
            </tr></thead>
            <tbody>
              {!filteredItems.length ? (
                <tr><td colSpan={columns.length + 4} style={{ padding: 36, color: 'var(--vdms-text-faint)', textAlign: 'center', fontSize: 13 }}>{empty}</td></tr>
              ) : filteredItems.map(item => (
                <tr key={item.id} style={{ ...DMS_TR, background: selectedIds.has(item.id) ? clay.accentSoft : 'transparent' }}>
                  <td style={{ ...DMS_TD, textAlign: 'center' }}>
                    <input type="checkbox" checked={selectedIds.has(item.id)} onChange={() => toggleOne(item.id)} />
                  </td>
                  {cells(item).map((cell, index) => <td key={index} style={{ padding: '10px 14px', color: index === 0 ? 'var(--vdms-text)' : 'var(--vdms-text-muted)', fontWeight: index === 0 ? 600 : 400, fontSize: 12.5 }}>{cell}</td>)}
                  <td style={DMS_TD} title={item.deleted_by_email || ''}>
                    {item.deleted_by_name || item.deleted_by_email || '—'}
                  </td>
                  <td style={{ ...DMS_TD, color: item.reason ? 'var(--vdms-text-muted)' : 'var(--vdms-text-faint)', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={item.reason || ''}>
                    {item.reason || '—'}
                  </td>
                  <td style={DMS_TD}>{actionCell(item)}</td>
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
      {/* Header */}
      <DmsPageHeader title="Recycle Bin" subtitle="Deleted vessels, folders and files remain here until restored or permanently deleted.">
        <button onClick={() => host._goToView('recycle').catch(() => undefined)} style={dmsBtn('secondary', true)}><Icon iconName="Refresh" aria-hidden="true" style={{ fontSize: 12 }} /> Refresh</button>
      </DmsPageHeader>

      {/* Compact Section Tabs */}
      <div style={{
        display: 'flex',
        flexWrap: 'wrap',
        gap: 8,
        background: 'var(--vdms-surface)',
        border: '1px solid var(--vdms-line)',
        borderRadius: 12,
        padding: 8,
        boxShadow: clay.shadowRaised,
      }}>
        <button
          onClick={() => { setActiveTab('vessels'); setSearchQuery(''); }}
          style={{ ...dmsBtn('primary', true), background: activeTab === 'vessels' ? clay.accent : 'transparent', color: activeTab === 'vessels' ? DMS_ON_ACCENT : 'var(--vdms-text)', boxShadow: activeTab === 'vessels' ? clay.shadowButton : 'none', borderRadius: 10 }}
        >
          <Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 12 }} /> Deleted Vessels ({vesselItems.length})
        </button>

        <button
          onClick={() => { setActiveTab('folders'); setSearchQuery(''); }}
          style={{ ...dmsBtn('primary', true), background: activeTab === 'folders' ? clay.accent : 'transparent', color: activeTab === 'folders' ? DMS_ON_ACCENT : 'var(--vdms-text)', boxShadow: activeTab === 'folders' ? clay.shadowButton : 'none', borderRadius: 10 }}
        >
          <Icon iconName="FabricFolder" aria-hidden="true" style={{ fontSize: 12 }} /> Deleted Normal Folders ({folderItems.length})
        </button>

        <button
          onClick={() => { setActiveTab('files'); setSearchQuery(''); }}
          style={{ ...dmsBtn('primary', true), background: activeTab === 'files' ? clay.accent : 'transparent', color: activeTab === 'files' ? DMS_ON_ACCENT : 'var(--vdms-text)', boxShadow: activeTab === 'files' ? clay.shadowButton : 'none', borderRadius: 10 }}
        >
          <Icon iconName="Page" aria-hidden="true" style={{ fontSize: 12 }} /> Deleted Individual Files ({fileItems.length})
        </button>
      </div>

      {/* Bulk Action Bar */}
      {selectedIds.size > 0 && (
        <div style={{ background: clay.accentDeep, color: DMS_ON_ACCENT, borderRadius: 10, padding: '12px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
          <span style={{ fontWeight: 600 }}>{selectedIds.size} item{selectedIds.size === 1 ? '' : 's'} selected</span>
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={restoreSelected} style={dmsRowBtn('success')}><Icon iconName="RecycleBin" aria-hidden="true" style={{ fontSize: 12 }} /> Restore selected</button>
            <button onClick={permanentlyDeleteSelected} style={dmsRowBtn('danger')}><Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 12 }} /> Delete permanently</button>
          </div>
        </div>
      )}

      {/* Active Section Table */}
      {panelLoading ? (
        <div style={{ padding: 48, textAlign: 'center', color: 'var(--vdms-text-muted)', fontSize: 14 }}>
          <Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 14 }} /> Loading Recycle Bin items...
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

