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

export function ArchivePage({ host }: { host: VesselEmail }): React.ReactElement {
    const { archiveList, archiveSearch, panelLoading } = host.state;
    const [vesselFilter, setVesselFilter] = React.useState('all');
    const [siteFilter, setSiteFilter] = React.useState('all');
    const [fromDate, setFromDate] = React.useState('');
    const [toDate, setToDate] = React.useState('');

    React.useEffect(() => {
      void host._loadArchiveList();
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    const formatSize = (size?: number | null): string => {
      if (size == null) return '—';
      if (size < 1024) return `${size} B`;
      if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
      if (size < 1024 * 1024 * 1024) return `${(size / (1024 * 1024)).toFixed(1)} MB`;
      return `${(size / (1024 * 1024 * 1024)).toFixed(1)} GB`;
    };
    const formatDate = (value?: string | null): string => value ? new Date(value).toLocaleString() : '—';

    const vesselOptions = Array.from(new Set(archiveList.map(i => i.vessel_name).filter(Boolean) as string[])).sort();
    const siteOptions = Array.from(new Set(archiveList.map(i => i.site_name).filter(Boolean) as string[])).sort();

    const q = (archiveSearch || '').trim().toLowerCase();
    const fromMs = fromDate ? new Date(fromDate).getTime() : null;
    const toMs = toDate ? new Date(toDate).getTime() + 24 * 60 * 60 * 1000 - 1 : null; // inclusive end of day

    const filteredList = archiveList.filter(item => {
      if (vesselFilter !== 'all' && (item.vessel_name || '') !== vesselFilter) return false;
      if (siteFilter !== 'all' && (item.site_name || '') !== siteFilter) return false;
      if (fromMs != null || toMs != null) {
        const ts = item.archived_at ? new Date(item.archived_at).getTime() : NaN;
        if (isNaN(ts)) return false;
        if (fromMs != null && ts < fromMs) return false;
        if (toMs != null && ts > toMs) return false;
      }
      if (q) {
        const hay = [item.name, item.vessel_name, item.original_path, item.site_name, item.category, item.sub_category]
          .filter(Boolean).join(' \u0001 ').toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });

    const activeFilterCount = [archiveSearch, vesselFilter !== 'all', siteFilter !== 'all', !!fromDate, !!toDate].filter(Boolean).length;

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ background: '#ede8f5', border: '1px solid #b4a0d4', borderRadius: 8, padding: '16px 20px' }}>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#5c2d91' }}>📦 Archive</h2>
          <p style={{ margin: '4px 0 0', fontSize: 12, color: '#5c2d91', opacity: 0.9 }}>
            Vessels, folders and files moved out of the working Documents view without being deleted — restore them here or move them to the Recycle Bin.
          </p>
        </div>

        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', padding: 12, display: 'flex', flexWrap: 'wrap', gap: 10, alignItems: 'center' }}>
          <input
            type="text"
            placeholder="Search file, vessel, path..."
            value={archiveSearch}
            onChange={e => host.setState({ archiveSearch: e.target.value })}
            style={{ padding: '6px 10px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12, outline: 'none', width: 220 }}
          />
          <select value={vesselFilter} onChange={e => setVesselFilter(e.target.value)} style={{ padding: '6px 10px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12 }}>
            <option value="all">All vessels</option>
            {vesselOptions.map(v => <option key={v} value={v}>{v}</option>)}
          </select>
          <select value={siteFilter} onChange={e => setSiteFilter(e.target.value)} style={{ padding: '6px 10px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12 }}>
            <option value="all">All sites</option>
            {siteOptions.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <label style={{ fontSize: 11, color: '#64748b', display: 'flex', alignItems: 'center', gap: 4 }}>
            From
            <input type="date" value={fromDate} onChange={e => setFromDate(e.target.value)} style={{ padding: '5px 8px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12 }} />
          </label>
          <label style={{ fontSize: 11, color: '#64748b', display: 'flex', alignItems: 'center', gap: 4 }}>
            To
            <input type="date" value={toDate} onChange={e => setToDate(e.target.value)} style={{ padding: '5px 8px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12 }} />
          </label>
          {activeFilterCount > 0 && (
            <button
              type="button"
              onClick={() => { host.setState({ archiveSearch: '' }); setVesselFilter('all'); setSiteFilter('all'); setFromDate(''); setToDate(''); }}
              style={{ background: 'none', border: 'none', color: '#5c2d91', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}
            >
              ✕ Clear filters ({activeFilterCount})
            </button>
          )}
          <span style={{ marginLeft: 'auto', fontSize: 11, color: '#64748b' }}>
            {panelLoading ? 'Loading…' : `${filteredList.length} of ${archiveList.length} item${archiveList.length === 1 ? '' : 's'}`}
          </span>
        </div>

        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'auto' }}>
          <table style={{ width: '100%', minWidth: 1650, borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                <th style={{ padding: '10px 16px' }}>File</th>
                <th style={{ padding: '10px 16px' }}>Site</th>
                <th style={{ padding: '10px 16px' }}>Location</th>
                <th style={{ padding: '10px 16px' }}>Vessel Name</th>
                <th style={{ padding: '10px 16px' }}>Document Section</th>
                <th style={{ padding: '10px 16px' }}>Group</th>
                <th style={{ padding: '10px 16px' }}>Category</th>
                <th style={{ padding: '10px 16px' }}>Sub-category</th>
                <th style={{ padding: '10px 16px' }}>Type</th>
                <th style={{ padding: '10px 16px' }}>Extension</th>
                <th style={{ padding: '10px 16px' }}>Size</th>
                <th style={{ padding: '10px 16px' }}>Modified</th>
                <th style={{ padding: '10px 16px' }}>Archived At</th>
                <th style={{ padding: '10px 16px', textAlign: 'right' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredList.length === 0 ? (
                <tr>
                  <td colSpan={14} style={{ padding: 20, textAlign: 'center', color: '#64748b' }}>
                    {archiveList.length === 0 ? 'No items archived.' : 'No archived items match these filters.'}
                  </td>
                </tr>
              ) : (
                filteredList.map(item => (
                  <tr key={item.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                    <td style={{ padding: '10px 16px', fontWeight: 600 }}>
                      <div>{item.name}</div>
                      <div style={{ color: '#64748b', fontSize: 11, fontWeight: 400 }}>ID: {item.id}</div>
                    </td>
                    <td style={{ padding: '10px 16px', color: '#475569' }}>{item.site_name || '—'}</td>
                    <td style={{ padding: '10px 16px', color: '#475569' }}>{item.original_path || item.main_folder || '—'}</td>
                    <td style={{ padding: '10px 16px' }}>{item.vessel_name || '—'}</td>
                    <td style={{ padding: '10px 16px' }}>{item.document_section || '—'}</td>
                    <td style={{ padding: '10px 16px' }}>{item.group || '—'}</td>
                    <td style={{ padding: '10px 16px' }}>{item.category || '—'}</td>
                    <td style={{ padding: '10px 16px' }}>{item.sub_category || '—'}</td>
                    <td style={{ padding: '10px 16px' }}>{item.kind}</td>
                    <td style={{ padding: '10px 16px' }}>{item.ext ? `.${item.ext}` : '—'}</td>
                    <td style={{ padding: '10px 16px' }}>{formatSize(item.size)}</td>
                    <td style={{ padding: '10px 16px' }}>{formatDate(item.modified)}</td>
                    <td style={{ padding: '10px 16px' }}>{formatDate(item.archived_at)}</td>
                    <td style={{ padding: '10px 16px', textAlign: 'right', whiteSpace: 'nowrap' }}>
                      <button
                        type="button"
                        onClick={() => void host._restoreArchivedItem(item)}
                        style={{ border: '1px solid #86efac', background: '#f0fdf4', color: '#15803d', borderRadius: 6, padding: '4px 9px', cursor: 'pointer', fontWeight: 700, fontSize: 11, marginRight: 6 }}
                      >Restore</button>
                      <button
                        type="button"
                        onClick={() => {
                          if (!confirm(`Move "${item.name}" to the Recycle Bin? This deletes it from SharePoint (recoverable from the Recycle Bin).`)) return;
                          void host._moveArchivedItemToRecycleBin(item);
                        }}
                        style={{ border: '1px solid #fca5a5', background: '#fff5f5', color: '#ef4444', borderRadius: 6, padding: '4px 9px', cursor: 'pointer', fontWeight: 700, fontSize: 11 }}
                      >🗑 Recycle Bin</button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    );

}
