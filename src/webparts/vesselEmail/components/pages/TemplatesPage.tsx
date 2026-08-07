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

export function renderTemplatesPage(host: VesselEmail): React.ReactElement {
    const { templatesList } = host.state;
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Templates</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Create and manage document templates.</p>
          </div>
          <button
            onClick={() => alert('New Template dialog')}
            style={{
              background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
              padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              display: 'inline-flex', alignItems: 'center', gap: 6,
            }}
          >
            ＋ New Template
          </button>
        </div>

        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                <th style={{ padding: '10px 16px' }}>Name</th>
                <th style={{ padding: '10px 16px' }}>Type</th>
                <th style={{ padding: '10px 16px' }}>Description</th>
                <th style={{ padding: '10px 16px' }}>Modified</th>
              </tr>
            </thead>
            <tbody>
              {templatesList.map(tpl => (
                <tr key={tpl.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '12px 16px', fontWeight: 600, color: '#1e293b', display: 'flex', alignItems: 'center', gap: 8 }}>
                    <span style={{ color: '#2563eb' }}>📝</span>
                    {tpl.name}
                  </td>
                  <td style={{ padding: '12px 16px' }}>{badge('blue', tpl.type)}</td>
                  <td style={{ padding: '12px 16px', color: '#475569' }}>{tpl.description}</td>
                  <td style={{ padding: '12px 16px', color: '#64748b' }}>{tpl.modified}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
}
