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

export function renderArchivePage(host: VesselEmail): React.ReactElement {
    const { archiveList } = host.state;
    const formatSize = (size?: number | null): string => {
      if (size == null) return '—';
      if (size < 1024) return `${size} B`;
      if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
      if (size < 1024 * 1024 * 1024) return `${(size / (1024 * 1024)).toFixed(1)} MB`;
      return `${(size / (1024 * 1024 * 1024)).toFixed(1)} GB`;
    };
    const formatDate = (value?: string | null): string => value ? new Date(value).toLocaleString() : '—';
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ background: '#ede8f5', border: '1px solid #b4a0d4', borderRadius: 8, padding: '16px 20px' }}>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#5c2d91' }}>📦 Archive</h2>
          <p style={{ margin: '4px 0 0', fontSize: 12, color: '#5c2d91', opacity: 0.9 }}>
            Archived items stored for historical record-keeping.
          </p>
        </div>

        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'auto' }}>
          <table style={{ width: '100%', minWidth: 1550, borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                <th style={{ padding: '10px 16px' }}>File</th>
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
              {archiveList.length === 0 ? (
                <tr>
                  <td colSpan={13} style={{ padding: 20, textAlign: 'center', color: '#64748b' }}>No items archived.</td>
                </tr>
              ) : (
                archiveList.map(item => (
                  <tr key={item.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                    <td style={{ padding: '10px 16px', fontWeight: 600 }}>
                      <div>{item.name}</div>
                      <div style={{ color: '#64748b', fontSize: 11, fontWeight: 400 }}>ID: {item.id}</div>
                    </td>
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
                    <td style={{ padding: '10px 16px', textAlign: 'right' }}>
                      <button
                        type="button"
                        onClick={() => void host._restoreArchivedItem(item)}
                        style={{ border: '1px solid #86efac', background: '#f0fdf4', color: '#15803d', borderRadius: 6, padding: '4px 9px', cursor: 'pointer', fontWeight: 700, fontSize: 11 }}
                      >Restore</button>
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
