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
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ background: '#ede8f5', border: '1px solid #b4a0d4', borderRadius: 8, padding: '16px 20px' }}>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#5c2d91' }}>📦 Archive</h2>
          <p style={{ margin: '4px 0 0', fontSize: 12, color: '#5c2d91', opacity: 0.9 }}>
            Archived items stored for historical record-keeping.
          </p>
        </div>

        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                <th style={{ padding: '10px 16px' }}>Name</th>
                <th style={{ padding: '10px 16px' }}>Type</th>
                <th style={{ padding: '10px 16px' }}>Archived At</th>
              </tr>
            </thead>
            <tbody>
              {archiveList.length === 0 ? (
                <tr>
                  <td colSpan={3} style={{ padding: 20, textAlign: 'center', color: '#64748b' }}>No items archived.</td>
                </tr>
              ) : (
                archiveList.map(item => (
                  <tr key={item.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                    <td style={{ padding: '10px 16px', fontWeight: 600 }}>{item.name}</td>
                    <td style={{ padding: '10px 16px' }}>{item.kind}</td>
                    <td style={{ padding: '10px 16px' }}>{item.archived_at || '—'}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    );
}
