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
import { Icon } from '@fluentui/react/lib/Icon';
import { clay } from '../clayTheme';
import {
  DmsPageHeader, DMS_SECTION_CARD, DMS_TABLE_CARD, DMS_TABLE, DMS_TH, DMS_TR, DMS_TD,
  dmsControlStyle, dmsRowBtn,
} from '../dmsDesignSystem';

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
        <DmsPageHeader
          title={<><Icon iconName="Package" aria-hidden="true" style={{ fontSize: 22, marginRight: 8 }} /> Archive</>}
          subtitle="Vessels, folders and files moved out of the working Documents view without being deleted — restore them here or move them to the Recycle Bin."
        />

        <div style={{ ...DMS_SECTION_CARD, flexDirection: 'row', flexWrap: 'wrap', padding: 12, gap: 10, alignItems: 'center' }}>
          <input
            type="text"
            placeholder="Search file, vessel, path..."
            value={archiveSearch}
            onChange={e => host.setState({ archiveSearch: e.target.value })}
            style={{ ...dmsControlStyle(), padding: '6px 10px', width: 220 }}
          />
          <select value={vesselFilter} onChange={e => setVesselFilter(e.target.value)} style={{ ...dmsControlStyle(), padding: '6px 10px' }}>
            <option value="all">All vessels</option>
            {vesselOptions.map(v => <option key={v} value={v}>{v}</option>)}
          </select>
          <select value={siteFilter} onChange={e => setSiteFilter(e.target.value)} style={{ ...dmsControlStyle(), padding: '6px 10px' }}>
            <option value="all">All sites</option>
            {siteOptions.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <label style={{ fontSize: 11, color: 'var(--vdms-text-muted)', display: 'flex', alignItems: 'center', gap: 4 }}>
            From
            <input type="date" value={fromDate} onChange={e => setFromDate(e.target.value)} style={{ ...dmsControlStyle(), padding: '5px 8px' }} />
          </label>
          <label style={{ fontSize: 11, color: 'var(--vdms-text-muted)', display: 'flex', alignItems: 'center', gap: 4 }}>
            To
            <input type="date" value={toDate} onChange={e => setToDate(e.target.value)} style={{ ...dmsControlStyle(), padding: '5px 8px' }} />
          </label>
          {activeFilterCount > 0 && (
            <button
              type="button"
              onClick={() => { host.setState({ archiveSearch: '' }); setVesselFilter('all'); setSiteFilter('all'); setFromDate(''); setToDate(''); }}
              style={{ background: 'none', border: 'none', color: clay.accent, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}
            >
              <Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 12 }} /> Clear filters ({activeFilterCount})
            </button>
          )}
          <span style={{ marginLeft: 'auto', fontSize: 11, color: 'var(--vdms-text-muted)' }}>
            {panelLoading ? 'Loading…' : `${filteredList.length} of ${archiveList.length} item${archiveList.length === 1 ? '' : 's'}`}
          </span>
        </div>

        <div style={{ ...DMS_TABLE_CARD, overflow: 'auto' }}>
          <table style={{ ...DMS_TABLE, minWidth: 1650 }}>
            <thead>
              <tr>
                <th style={DMS_TH}>File</th>
                <th style={DMS_TH}>Site</th>
                <th style={DMS_TH}>Location</th>
                <th style={DMS_TH}>Vessel Name</th>
                <th style={DMS_TH}>Document Section</th>
                <th style={DMS_TH}>Group</th>
                <th style={DMS_TH}>Category</th>
                <th style={DMS_TH}>Sub-category</th>
                <th style={DMS_TH}>Type</th>
                <th style={DMS_TH}>Extension</th>
                <th style={DMS_TH}>Size</th>
                <th style={DMS_TH}>Modified</th>
                <th style={DMS_TH}>Archived At</th>
                <th style={{ ...DMS_TH, textAlign: 'right' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredList.length === 0 ? (
                <tr>
                  <td colSpan={14} style={{ ...DMS_TD, padding: 20, textAlign: 'center' }}>
                    {archiveList.length === 0 ? 'No items archived.' : 'No archived items match these filters.'}
                  </td>
                </tr>
              ) : (
                filteredList.map(item => (
                  <tr key={item.id} style={DMS_TR}>
                    <td style={{ ...DMS_TD, fontWeight: 600, color: 'var(--vdms-text)' }}>
                      <div>{item.name}</div>
                      <div style={{ color: 'var(--vdms-text-muted)', fontSize: 11, fontWeight: 400 }}>ID: {item.id}</div>
                    </td>
                    <td style={DMS_TD}>{item.site_name || '—'}</td>
                    <td style={DMS_TD}>{item.original_path || item.main_folder || '—'}</td>
                    <td style={DMS_TD}>{item.vessel_name || '—'}</td>
                    <td style={DMS_TD}>{item.document_section || '—'}</td>
                    <td style={DMS_TD}>{item.group || '—'}</td>
                    <td style={DMS_TD}>{item.category || '—'}</td>
                    <td style={DMS_TD}>{item.sub_category || '—'}</td>
                    <td style={DMS_TD}>{item.kind}</td>
                    <td style={DMS_TD}>{item.ext ? `.${item.ext}` : '—'}</td>
                    <td style={DMS_TD}>{formatSize(item.size)}</td>
                    <td style={DMS_TD}>{formatDate(item.modified)}</td>
                    <td style={DMS_TD}>{formatDate(item.archived_at)}</td>
                    <td style={{ ...DMS_TD, textAlign: 'right', whiteSpace: 'nowrap' }}>
                      <button
                        type="button"
                        onClick={() => void host._restoreArchivedItem(item)}
                        style={{ ...dmsRowBtn('success'), marginRight: 6 }}
                      >Restore</button>
                      <button
                        type="button"
                        onClick={() => {
                          if (!confirm(`Move "${item.name}" to the Recycle Bin? This deletes it from SharePoint (recoverable from the Recycle Bin).`)) return;
                          void host._moveArchivedItemToRecycleBin(item);
                        }}
                        style={dmsRowBtn('danger')}
                      ><Icon iconName="Delete" aria-hidden="true" style={{ fontSize: 11 }} /> Recycle Bin</button>
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
