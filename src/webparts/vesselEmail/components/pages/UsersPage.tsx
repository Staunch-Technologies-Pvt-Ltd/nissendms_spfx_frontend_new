/* eslint-disable @typescript-eslint/no-unused-vars */
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
  ApprovalItem, NotificationItem, UserItem, UserSitePermissionItem,
} from '../types/ui';
import { getVesselImageForId, pickRandomVesselImage, resolveImgUrl } from '../vesselImagePool';

// Reaching this page at all means /api/users (admin-gated server-side via
// require_admin_session) already returned data for the signed-in user, so
// every row rendered here belongs to an admin session -- no separate
// client-side admin check is needed before showing the edit controls.

const emptyPermission = (siteKey: string): UserSitePermissionItem => ({
  site_key: siteKey, can_view: true, can_upload: false, can_tag_on_upload: false,
});

function PermissionsEditor({ host, usr }: { host: VesselEmail; usr: UserItem }): React.ReactElement {
  const draft = host.state.usersPermissionsDraft;
  const busy = host.state.usersPermissionsBusy;
  const knownSites = (host.state.documentSites || []).map(s => ({
    key: s.site_key || s.site_id || '',
    label: s.sp_site_name || s.site_key || s.site_id || '',
  })).filter(s => s.key);

  const setDraft = (next: UserSitePermissionItem[]): void => host.setState({ usersPermissionsDraft: next });

  const updateRow = (siteKey: string, patch: Partial<UserSitePermissionItem>): void => {
    const next = draft.map(p => {
      if (p.site_key !== siteKey) return p;
      const merged = { ...p, ...patch };
      // Mirror the server-side dependency (require_site_permission /
      // update_user_permissions in backend/app/main.py): upload implies
      // view, tag-on-upload implies upload -- keep the checkboxes honest
      // rather than letting the admin set a combination the save will
      // silently narrow anyway.
      if (!merged.can_view) { merged.can_upload = false; merged.can_tag_on_upload = false; }
      if (!merged.can_upload) { merged.can_tag_on_upload = false; }
      return merged;
    });
    setDraft(next);
  };

  const addSite = (siteKey: string): void => {
    if (!siteKey || draft.some(p => p.site_key === siteKey)) return;
    setDraft([...draft, emptyPermission(siteKey)]);
  };

  const removeSite = (siteKey: string): void => setDraft(draft.filter(p => p.site_key !== siteKey));

  const [manualSiteKey, setManualSiteKey] = React.useState('');

  return (
    <div style={{ padding: '14px 16px', background: '#f8fafc', borderTop: '1px solid #e2e8f0', borderBottom: '1px solid #e2e8f0' }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: '#334155', marginBottom: 8 }}>
        SharePoint site access for {usr.email}
      </div>
      {draft.length === 0 && (
        <div style={{ fontSize: 12, color: '#94a3b8', marginBottom: 10 }}>No site access granted yet.</div>
      )}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 10 }}>
        {draft.map(p => (
          <div key={p.site_key} style={{ display: 'flex', alignItems: 'center', gap: 16, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 6, padding: '8px 12px' }}>
            <span style={{ fontSize: 12, fontWeight: 600, color: '#1e293b', minWidth: 160, flex: '1 1 160px' }}>
              {knownSites.find(s => s.key === p.site_key)?.label || p.site_key}
            </span>
            <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, color: '#475569' }}>
              <input type="checkbox" checked={p.can_view} onChange={e => updateRow(p.site_key, { can_view: e.target.checked })} />
              View
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, color: '#475569' }}>
              <input type="checkbox" checked={p.can_upload} disabled={!p.can_view} onChange={e => updateRow(p.site_key, { can_upload: e.target.checked })} />
              Upload
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, color: '#475569' }}>
              <input type="checkbox" checked={p.can_tag_on_upload} disabled={!p.can_upload} onChange={e => updateRow(p.site_key, { can_tag_on_upload: e.target.checked })} />
              Tag on upload
            </label>
            <button
              type="button"
              onClick={() => removeSite(p.site_key)}
              style={{ marginLeft: 'auto', border: 'none', background: 'none', color: '#ef4444', fontSize: 12, cursor: 'pointer', fontWeight: 600 }}
            >
              Remove
            </button>
          </div>
        ))}
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12, flexWrap: 'wrap' }}>
        {knownSites.length > 0 ? (
          <select
            value=""
            onChange={e => { if (e.target.value) addSite(e.target.value); }}
            style={{ padding: '5px 8px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12 }}
          >
            <option value="">+ Add a site...</option>
            {knownSites.filter(s => !draft.some(p => p.site_key === s.key)).map(s => (
              <option key={s.key} value={s.key}>{s.label}</option>
            ))}
          </select>
        ) : (
          <>
            <input
              type="text"
              placeholder="site_key (e.g. from Sites page)"
              value={manualSiteKey}
              onChange={e => setManualSiteKey(e.target.value)}
              style={{ padding: '5px 8px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12, width: 200 }}
            />
            <button
              type="button"
              onClick={() => { addSite(manualSiteKey.trim().toLowerCase()); setManualSiteKey(''); }}
              style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '5px 10px', fontSize: 12, cursor: 'pointer' }}
            >
              Add
            </button>
          </>
        )}
      </div>

      <div style={{ display: 'flex', gap: 8 }}>
        <button
          type="button"
          disabled={busy}
          onClick={() => void host._updateUserSitePermissions(usr.email, draft)}
          style={{
            background: busy ? '#93c5fd' : '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
            padding: '7px 16px', fontSize: 12, fontWeight: 600, cursor: busy ? 'not-allowed' : 'pointer',
          }}
        >
          {busy ? 'Saving...' : 'Save access'}
        </button>
        <button
          type="button"
          disabled={busy}
          onClick={() => host.setState({ usersExpandedEmail: null })}
          style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '7px 16px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}
        >
          Cancel
        </button>
      </div>
    </div>
  );
}

export function renderUsersPage(host: VesselEmail): React.ReactElement {
    const { usersList, userSearch, usersExpandedEmail } = host.state;
    const filtered = usersList.filter(u =>
      u.name.toLowerCase().indexOf(userSearch.toLowerCase()) !== -1 ||
      u.email.toLowerCase().indexOf(userSearch.toLowerCase()) !== -1
    );

    const toggleExpanded = (usr: UserItem): void => {
      if (usersExpandedEmail === usr.email) {
        host.setState({ usersExpandedEmail: null });
        return;
      }
      host.setState({
        usersExpandedEmail: usr.email,
        usersPermissionsDraft: (usr.permissions || []).map(p => ({ ...p })),
      });
    };

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>User Management</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>
              Manage roles and per-site access. Users appear here once they've logged in at least once.
            </p>
          </div>
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
                <th style={{ padding: '10px 16px' }}>Site access</th>
                <th style={{ padding: '10px 16px' }} />
              </tr>
            </thead>
            <tbody>
              {filtered.map(usr => (
                <React.Fragment key={usr.id}>
                  <tr style={{ borderBottom: '1px solid #f1f5f9' }}>
                    <td style={{ padding: '12px 16px', fontWeight: 600, color: '#1e293b', display: 'flex', alignItems: 'center', gap: 10 }}>
                      <div style={{ width: 28, height: 28, borderRadius: '50%', background: '#e2e8f0', color: '#475569', fontSize: 12, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                        {usr.name.charAt(0)}
                      </div>
                      {usr.name}
                    </td>
                    <td style={{ padding: '12px 16px', color: '#475569' }}>{usr.email}</td>
                    <td style={{ padding: '12px 16px', color: '#334155' }}>
                      <select
                        value={usr.role}
                        disabled={host.state.usersPermissionsBusy}
                        onChange={e => {
                          const nextRole = e.target.value as UserItem['role'];
                          if (nextRole === usr.role) return;
                          if (!confirm(`Change ${usr.email}'s role to ${nextRole}?`)) return;
                          void host._updateUserRole(usr.email, nextRole);
                        }}
                        style={{ padding: '4px 8px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff' }}
                      >
                        <option value="User">User</option>
                        <option value="Admin">Admin</option>
                      </select>
                    </td>
                    <td style={{ padding: '12px 16px' }}>
                      {badge(usr.status === 'Active' ? 'green' : 'red', usr.status)}
                    </td>
                    <td style={{ padding: '12px 16px', color: '#64748b' }}>{usr.lastLogin}</td>
                    <td style={{ padding: '12px 16px', color: '#64748b' }}>
                      {usr.role === 'Admin'
                        ? <span style={{ fontSize: 11, color: '#94a3b8' }}>All sites (Admin)</span>
                        : `${(usr.permissions || []).length} site${(usr.permissions || []).length === 1 ? '' : 's'}`}
                    </td>
                    <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                      {usr.role !== 'Admin' && (
                        <button
                          type="button"
                          onClick={() => toggleExpanded(usr)}
                          style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '5px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer', color: '#0078d4' }}
                        >
                          {usersExpandedEmail === usr.email ? 'Close' : 'Manage access'}
                        </button>
                      )}
                    </td>
                  </tr>
                  {usersExpandedEmail === usr.email && (
                    <tr>
                      <td colSpan={7} style={{ padding: 0 }}>
                        <PermissionsEditor host={host} usr={usr} />
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              ))}
              {filtered.length === 0 && (
                <tr>
                  <td colSpan={7} style={{ padding: '32px 16px', textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>
                    No users found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    );
}
