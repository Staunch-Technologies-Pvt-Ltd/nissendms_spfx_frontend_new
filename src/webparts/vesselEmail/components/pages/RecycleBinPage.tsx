import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { badge } from '../constants';
import type { DeletedNode } from '../types/ui';

export function renderRecycleBinPage(host: VesselEmail): React.ReactElement {
  return <RecycleBinContent host={host} />;
}

function RecycleBinContent({ host }: { host: VesselEmail }): React.ReactElement {
  const { recycleBin, panelLoading } = host.state;
  const [selectedIds, setSelectedIds] = React.useState<Set<string>>(new Set());

  const vesselItems = recycleBin.filter(r => r.kind === 'vessel' || r.item_type === 'vessel');
  const otherItems  = recycleBin.filter(r => r.kind !== 'vessel' && r.item_type !== 'vessel');

  const toggleSelectOne = (id: string) => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  const handleBulkRestore = async () => {
    const selectedItems = recycleBin.filter(r => selectedIds.has(r.id));
    if (selectedItems.length === 0) return;

    for (const item of selectedItems) {
      await host._restoreFromRecycleBin(item);
    }
    setSelectedIds(new Set());
  };

  const handleBulkPermanentDelete = async () => {
    const selectedItems = recycleBin.filter(r => selectedIds.has(r.id));
    if (selectedItems.length === 0) return;

    if (!window.confirm(`Permanently delete ${selectedItems.length} selected item(s)? This action cannot be undone.`)) {
      return;
    }

    for (const item of selectedItems) {
      await host._permanentDeleteFromRecycleBin(item, true);
    }
    setSelectedIds(new Set());
  };

  const renderTable = (items: DeletedNode[], emptyMsg: string): React.ReactElement => {
    const sectionIds = items.map(i => i.id);
    const isSectionAllSelected = sectionIds.length > 0 && sectionIds.every(id => selectedIds.has(id));

    const toggleSectionSelectAll = () => {
      setSelectedIds(prev => {
        const next = new Set(prev);
        if (isSectionAllSelected) {
          sectionIds.forEach(id => next.delete(id));
        } else {
          sectionIds.forEach(id => next.add(id));
        }
        return next;
      });
    };

    return (
      <div style={{ background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', overflow: 'hidden', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead>
            <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
              <th style={{ padding: '12px 16px', width: 40, textAlign: 'center' }}>
                {items.length > 0 && (
                  <input
                    type="checkbox"
                    checked={isSectionAllSelected}
                    onChange={toggleSectionSelectAll}
                    style={{ cursor: 'pointer', width: 16, height: 16 }}
                    title="Select / Deselect all in this section"
                  />
                )}
              </th>
              <th style={{ padding: '12px 16px' }}>Name</th>
              <th style={{ padding: '12px 16px' }}>Type</th>
              <th style={{ padding: '12px 16px' }}>Original Path</th>
              <th style={{ padding: '12px 16px' }}>Deleted At</th>
              <th style={{ padding: '12px 16px', textAlign: 'right' }}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {items.length === 0 ? (
              <tr><td colSpan={6} style={{ padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>{emptyMsg}</td></tr>
            ) : items.map(item => {
              const isChecked = selectedIds.has(item.id);
              const isVessel = item.kind === 'vessel' || item.item_type === 'vessel';

              return (
                <tr
                  key={item.id}
                  style={{
                    borderBottom: '1px solid #f1f5f9',
                    background: isChecked ? '#f0f9ff' : 'transparent',
                    transition: 'background 0.15s',
                  }}
                  onMouseEnter={e => { if (!isChecked) e.currentTarget.style.background = '#f8fafc'; }}
                  onMouseLeave={e => { if (!isChecked) e.currentTarget.style.background = 'transparent'; }}
                >
                  <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                    <input
                      type="checkbox"
                      checked={isChecked}
                      onChange={() => toggleSelectOne(item.id)}
                      style={{ cursor: 'pointer', width: 16, height: 16 }}
                    />
                  </td>
                  <td style={{ padding: '12px 16px', fontWeight: 600, color: '#0f172a' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                      <span style={{ fontSize: 18 }}>{isVessel ? '🚢' : item.kind === 'folder' ? '📁' : '📄'}</span>
                      <span>{item.name}</span>
                    </div>
                  </td>
                  <td style={{ padding: '12px 16px' }}>
                    {badge(isVessel ? 'blue' : item.kind === 'folder' ? 'orange' : 'default', isVessel ? 'Vessel' : item.kind || item.item_type || 'File')}
                  </td>
                  <td style={{ padding: '12px 16px', color: '#64748b', fontSize: 11, fontFamily: 'monospace', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={item.original_path || '—'}>
                    {item.original_path || '—'}
                  </td>
                  <td style={{ padding: '12px 16px', color: '#64748b', fontSize: 12, whiteSpace: 'nowrap' }}>
                    {item.deleted_at ? new Date(item.deleted_at).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : '—'}
                  </td>
                  <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                    <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                      <button
                        onClick={() => host._restoreFromRecycleBin(item)}
                        style={{ background: '#dff6dd', color: '#107c10', border: '1px solid #86efac', borderRadius: 6, padding: '5px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}
                        title="Restore item to active status"
                      >
                        ↩ Restore
                      </button>
                      <button
                        onClick={() => host._permanentDeleteFromRecycleBin(item)}
                        style={{ background: '#fde7e9', color: '#a4262c', border: '1px solid #fca5a5', borderRadius: 6, padding: '5px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}
                        title="Delete permanently from system"
                      >
                        🗑 Delete Permanently
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      {/* ── Top Header Card ── */}
      <div style={{ background: 'linear-gradient(135deg, #fffbe6 0%, #fef3c7 100%)', border: '1px solid #f59e0b', borderRadius: 12, padding: '20px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#92400e', display: 'flex', alignItems: 'center', gap: 10 }}>
            🗑️ Recycle Bin
            {recycleBin.length > 0 && (
              <span style={{ background: '#f59e0b', color: '#fff', borderRadius: 20, padding: '2px 10px', fontSize: 12, fontWeight: 700 }}>
                {recycleBin.length}
              </span>
            )}
          </h2>
          <p style={{ margin: '4px 0 0', fontSize: 13, color: '#92400e', opacity: 0.9 }}>
            Select items using checkboxes to restore or permanently delete them from DMS.
          </p>
        </div>

        <div style={{ display: 'flex', gap: 10 }}>
          <button
            onClick={() => host._loadData()}
            style={{ background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8, padding: '8px 16px', fontSize: 13, fontWeight: 600, color: '#475569', cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}
          >
            🔄 Refresh
          </button>
        </div>
      </div>

      {/* ── Bulk Action Bar (when 1 or more checkboxes are selected) ── */}
      {selectedIds.size > 0 && (
        <div style={{
          background: '#0f172a', color: '#fff', borderRadius: 12, padding: '14px 24px',
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 12,
          boxShadow: '0 8px 24px rgba(0,0,0,0.2)',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <span style={{ background: '#0284c7', color: '#fff', padding: '4px 12px', borderRadius: 20, fontWeight: 700, fontSize: 13 }}>
              ☑️ {selectedIds.size} selected
            </span>
            <span style={{ fontSize: 13, opacity: 0.8 }}>Bulk actions for checked item(s):</span>
          </div>

          <div style={{ display: 'flex', gap: 10 }}>
            <button
              onClick={handleBulkRestore}
              style={{
                background: '#10b981', color: '#fff', border: 'none', borderRadius: 8,
                padding: '8px 18px', fontSize: 13, fontWeight: 700, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6,
              }}
            >
              ↩ Restore Selected ({selectedIds.size})
            </button>
            <button
              onClick={handleBulkPermanentDelete}
              style={{
                background: '#ef4444', color: '#fff', border: 'none', borderRadius: 8,
                padding: '8px 18px', fontSize: 13, fontWeight: 700, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6,
              }}
            >
              🗑 Delete Permanently ({selectedIds.size})
            </button>
            <button
              onClick={() => setSelectedIds(new Set())}
              style={{
                background: 'transparent', color: '#94a3b8', border: '1px solid #475569', borderRadius: 8,
                padding: '8px 14px', fontSize: 13, cursor: 'pointer',
              }}
            >
              ✕ Deselect
            </button>
          </div>
        </div>
      )}

      {/* ── Table List ── */}
      {panelLoading ? (
        <div style={{ padding: 48, textAlign: 'center', color: '#64748b', fontSize: 14 }}>⏳ Loading recycle bin items…</div>
      ) : recycleBin.length === 0 ? (
        <div style={{ padding: 48, textAlign: 'center', background: '#fff', borderRadius: 12, border: '2px dashed #e2e8f0' }}>
          <div style={{ fontSize: 44, marginBottom: 12 }}>🗑️</div>
          <div style={{ fontWeight: 700, fontSize: 16, color: '#0f172a' }}>Recycle Bin is Empty</div>
          <div style={{ fontSize: 13, color: '#64748b', marginTop: 4 }}>No deleted vessels, folders, or files found.</div>
        </div>
      ) : (
        <>
          {/* Deleted Vessels section */}
          {vesselItems.length > 0 && (
            <div>
              <div style={{ fontSize: 12, fontWeight: 700, letterSpacing: '0.05em', color: '#0369a1', textTransform: 'uppercase', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
                <span>🚢</span><span>Deleted Vessels ({vesselItems.length})</span>
              </div>
              {renderTable(vesselItems, 'No deleted vessels in recycle bin.')}
            </div>
          )}

          {/* Other Deleted Items section */}
          {otherItems.length > 0 && (
            <div style={{ marginTop: vesselItems.length > 0 ? 12 : 0 }}>
              <div style={{ fontSize: 12, fontWeight: 700, letterSpacing: '0.05em', color: '#475569', textTransform: 'uppercase', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
                <span>📁</span><span>Other Deleted Folders & Files ({otherItems.length})</span>
              </div>
              {renderTable(otherItems, 'No other deleted items.')}
            </div>
          )}
        </>
      )}
    </div>
  );
}
