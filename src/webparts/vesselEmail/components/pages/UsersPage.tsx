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

export function renderUsersPage(host: VesselEmail): React.ReactElement {
    const { usersList, userSearch } = host.state;
    const filtered = usersList.filter(u =>
      u.name.toLowerCase().indexOf(userSearch.toLowerCase()) !== -1 ||
      u.email.toLowerCase().indexOf(userSearch.toLowerCase()) !== -1
    );

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>User Management</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Manage users and their roles.</p>
          </div>
          <button
            onClick={() => alert('Add User Modal')}
            style={{
              background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
              padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              display: 'inline-flex', alignItems: 'center', gap: 6,
            }}
          >
            ＋ Add User
          </button>
        </div>

        <div style={{ background: '#fff', borderRadius: 8, padding: 12, border: '1px solid #e2e8f0', display: 'flex', gap: 12 }}>
          <input
            type="text"
            placeholder="Search users"
            value={userSearch}
            onChange={e => host.setState({ userSearch: e.target.value })}
            style={{ padding: '6px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', width: 240 }}
          />
        </div>

        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                <th style={{ padding: '10px 16px' }}>Name</th>
                <th style={{ padding: '10px 16px' }}>Email</th>
                <th style={{ padding: '10px 16px' }}>Role</th>
                <th style={{ padding: '10px 16px' }}>Status</th>
                <th style={{ padding: '10px 16px' }}>Last Login</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(usr => (
                <tr key={usr.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '12px 16px', fontWeight: 600, color: '#1e293b', display: 'flex', alignItems: 'center', gap: 10 }}>
                    <div style={{ width: 28, height: 28, borderRadius: '50%', background: '#e2e8f0', color: '#475569', fontSize: 12, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                      {usr.name.charAt(0)}
                    </div>
                    {usr.name}
                  </td>
                  <td style={{ padding: '12px 16px', color: '#475569' }}>{usr.email}</td>
                  <td style={{ padding: '12px 16px', color: '#334155' }}>{usr.role}</td>
                  <td style={{ padding: '12px 16px' }}>
                    {badge(usr.status === 'Active' ? 'green' : 'red', usr.status)}
                  </td>
                  <td style={{ padding: '12px 16px', color: '#64748b' }}>{usr.lastLogin}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
}
