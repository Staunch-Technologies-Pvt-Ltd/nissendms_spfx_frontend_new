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
import {
  DmsPageHeader, DMS_SECTION_CARD, DMS_TABLE_CARD, DMS_TABLE, DMS_TH, DMS_TR, DMS_TD, DMS_TD_NAME, DMS_NAME_CELL,
  dmsBtn, dmsControlStyle, dmsRowBtn, dmsTone,
} from '../dmsDesignSystem';

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
    <div style={{ padding: '14px 16px', background: 'var(--vdms-surface-alt)', borderTop: '1px solid var(--vdms-line)', borderBottom: '1px solid var(--vdms-line)' }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: 'var(--vdms-text)', marginBottom: 8 }}>
        SharePoint site access for {usr.email}
      </div>
      {draft.length === 0 && (
        <div style={{ fontSize: 12, color: 'var(--vdms-text-faint)', marginBottom: 10 }}>No site access granted yet.</div>
      )}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 10 }}>
        {draft.map(p => (
          <div key={p.site_key} style={{ display: 'flex', alignItems: 'center', gap: 16, background: 'var(--vdms-surface)', border: '1px solid var(--vdms-line)', borderRadius: 6, padding: '8px 12px' }}>
            <span style={{ fontSize: 12, fontWeight: 600, color: 'var(--vdms-text)', minWidth: 160, flex: '1 1 160px' }}>
              {knownSites.find(s => s.key === p.site_key)?.label || p.site_key}
            </span>
            <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, color: 'var(--vdms-text-muted)' }}>
              <input type="checkbox" checked={p.can_view} onChange={e => updateRow(p.site_key, { can_view: e.target.checked })} />
              View
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, color: 'var(--vdms-text-muted)' }}>
              <input type="checkbox" checked={p.can_upload} disabled={!p.can_view} onChange={e => updateRow(p.site_key, { can_upload: e.target.checked })} />
              Upload
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, color: 'var(--vdms-text-muted)' }}>
              <input type="checkbox" checked={p.can_tag_on_upload} disabled={!p.can_upload} onChange={e => updateRow(p.site_key, { can_tag_on_upload: e.target.checked })} />
              Tag on upload
            </label>
            <button
              type="button"
              onClick={() => removeSite(p.site_key)}
              style={{ marginLeft: 'auto', border: 'none', background: 'none', color: dmsTone('danger').fg, fontSize: 12, cursor: 'pointer', fontWeight: 600 }}
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
            style={{ ...dmsControlStyle(), padding: '5px 8px' }}
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
              style={{ ...dmsControlStyle(), padding: '5px 8px', width: 200 }}
            />
            <button
              type="button"
              onClick={() => { addSite(manualSiteKey.trim().toLowerCase()); setManualSiteKey(''); }}
              style={dmsBtn('secondary')}
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
          style={dmsBtn('primary', !busy)}
        >
          {busy ? 'Saving...' : 'Save access'}
        </button>
        <button
          type="button"
          disabled={busy}
          onClick={() => host.setState({ usersExpandedEmail: null })}
          style={dmsBtn('secondary', !busy)}
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
        <DmsPageHeader
          title="User Management"
          subtitle="Manage roles and per-site access. Users appear here once they've logged in at least once."
        />

        <div style={{ ...DMS_SECTION_CARD, flexDirection: 'row', padding: 12, gap: 12 }}>
          <input
            type="text"
            placeholder="Search users"
            value={userSearch}
            onChange={e => host.setState({ userSearch: e.target.value })}
            style={{ ...dmsControlStyle(), padding: '6px 12px', fontSize: 13, width: 240 }}
          />
        </div>

        <div style={DMS_TABLE_CARD}>
          <table style={DMS_TABLE}>
            <thead>
              <tr>
                <th style={DMS_TH}>Name</th>
                <th style={DMS_TH}>Email</th>
                <th style={DMS_TH}>Role</th>
                <th style={DMS_TH}>Status</th>
                <th style={DMS_TH}>Last Login</th>
                <th style={DMS_TH}>Site access</th>
                <th style={DMS_TH} />
              </tr>
            </thead>
            <tbody>
              {filtered.map(usr => (
                <React.Fragment key={usr.id}>
                  <tr style={DMS_TR}>
                    <td style={DMS_TD_NAME}>
                      <div style={DMS_NAME_CELL}>
                        <div style={{ width: 28, height: 28, borderRadius: '50%', background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text-muted)', fontSize: 12, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                          {usr.name.charAt(0)}
                        </div>
                        {usr.name}
                      </div>
                    </td>
                    <td style={DMS_TD}>{usr.email}</td>
                    <td style={DMS_TD}>
                      <select
                        value={usr.role}
                        disabled={host.state.usersPermissionsBusy}
                        onChange={e => {
                          const nextRole = e.target.value as UserItem['role'];
                          if (nextRole === usr.role) return;
                          if (!confirm(`Change ${usr.email}'s role to ${nextRole}?`)) return;
                          void host._updateUserRole(usr.email, nextRole);
                        }}
                        style={{ ...dmsControlStyle(), padding: '4px 8px' }}
                      >
                        <option value="User">User</option>
                        <option value="Admin">Admin</option>
                      </select>
                    </td>
                    <td style={DMS_TD}>
                      {badge(usr.status === 'Active' ? 'green' : 'red', usr.status)}
                    </td>
                    <td style={DMS_TD}>{usr.lastLogin}</td>
                    <td style={DMS_TD}>
                      {usr.role === 'Admin'
                        ? <span style={{ fontSize: 11, color: 'var(--vdms-text-faint)' }}>All sites (Admin)</span>
                        : `${(usr.permissions || []).length} site${(usr.permissions || []).length === 1 ? '' : 's'}`}
                    </td>
                    <td style={{ ...DMS_TD, textAlign: 'right' }}>
                      {usr.role !== 'Admin' && (
                        <button
                          type="button"
                          onClick={() => toggleExpanded(usr)}
                          style={dmsRowBtn('plain')}
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
                  <td colSpan={7} style={{ ...DMS_TD, padding: '32px 16px', textAlign: 'center' }}>
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
