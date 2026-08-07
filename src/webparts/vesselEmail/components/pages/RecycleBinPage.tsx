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

export function renderRecycleBinPage(host: VesselEmail): React.ReactElement {
    const { recycleBin, panelLoading } = host.state;

    const vesselItems = recycleBin.filter(r => r.kind === 'vessel' || r.item_type === 'vessel');
    const otherItems  = recycleBin.filter(r => r.kind !== 'vessel' && r.item_type !== 'vessel');

    const renderTable = (items: DeletedNode[], emptyMsg: string): React.ReactElement => (
      <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead>
            <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
              <th style={{ padding: '10px 16px' }}>Name</th>
              <th style={{ padding: '10px 16px' }}>Type</th>
              <th style={{ padding: '10px 16px' }}>Original Path</th>
              <th style={{ padding: '10px 16px' }}>Deleted At</th>
              <th style={{ padding: '10px 16px', textAlign: 'right' }}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {items.length === 0 ? (
              <tr><td colSpan={5} style={{ padding: 20, textAlign: 'center', color: '#94a3b8', fontSize: 12 }}>{emptyMsg}</td></tr>
            ) : items.map(item => (
              <tr key={item.id} style={{ borderBottom: '1px solid #f1f5f9' }}
                onMouseEnter={e => (e.currentTarget.style.background = '#f8fafc')}
                onMouseLeave={e => (e.currentTarget.style.background = '')}>
                <td style={{ padding: '10px 16px', fontWeight: 600, color: '#0f172a' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <span>{item.kind === 'vessel' || item.item_type === 'vessel' ? 'ðŸš¢' : item.kind === 'folder' ? 'ðŸ“' : 'ðŸ“„'}</span>
                    {item.name}
                  </div>
                </td>
                <td style={{ padding: '10px 16px' }}>
                  {badge(item.kind === 'vessel' || item.item_type === 'vessel' ? 'blue' : item.kind === 'folder' ? 'orange' : 'default', item.kind || item.item_type || 'â€”')}
                </td>
                <td style={{ padding: '10px 16px', color: '#64748b', fontSize: 11, maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={item.original_path || 'â€”'}>
                  {item.original_path || 'â€”'}
                </td>
                <td style={{ padding: '10px 16px', color: '#64748b', fontSize: 12, whiteSpace: 'nowrap' }}>
                  {item.deleted_at ? new Date(item.deleted_at).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : 'â€”'}
                </td>
                <td style={{ padding: '10px 16px', textAlign: 'right' }}>
                  <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                    <button
                      onClick={() => host._restoreFromRecycleBin(item)}
                      style={{ background: '#dff6dd', color: '#107c10', border: 'none', borderRadius: 6, padding: '5px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}
                    >
                      â†© Restore
                    </button>
                    <button
                      onClick={() => host._permanentDeleteFromRecycleBin(item)}
                      style={{ background: '#fde7e9', color: '#a4262c', border: 'none', borderRadius: 6, padding: '5px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}
                    >
                      ðŸ—‘ Delete
                    </button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ background: '#fff4ce', border: '1px solid #f7cf72', borderRadius: 8, padding: '16px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#8a5700' }}>ðŸ—‘ï¸ Recycle Bin</h2>
            <p style={{ margin: '4px 0 0', fontSize: 12, color: '#8a5700', opacity: 0.9 }}>
              {recycleBin.length} item{recycleBin.length !== 1 ? 's' : ''} â€” Restore to reactivate or permanently delete.
            </p>
          </div>
          <button
            onClick={() => host._goToView('recycle')}
            style={{ background: '#fff', border: '1px solid #f7cf72', borderRadius: 6, padding: '6px 14px', fontSize: 12, fontWeight: 600, color: '#8a5700', cursor: 'pointer' }}
          >
            ðŸ”„ Refresh
          </button>
        </div>

        {panelLoading ? (
          <div style={{ padding: 32, textAlign: 'center', color: '#64748b', fontSize: 13 }}>â³ Loading recycle bin...</div>
        ) : (
          <>
            {vesselItems.length > 0 && (
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 8 }}>Deleted Vessels</div>
                {renderTable(vesselItems, 'No deleted vessels.')}
              </div>
            )}
            <div>
              {vesselItems.length > 0 && (
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 8 }}>Other Items</div>
              )}
              {renderTable(otherItems, recycleBin.length === 0 ? 'Recycle bin is empty.' : 'No other items.')}
            </div>
          </>
        )}
      </div>
    );
}
