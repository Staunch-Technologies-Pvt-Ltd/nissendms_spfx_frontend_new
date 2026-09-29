/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable @typescript-eslint/no-unused-expressions */
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
import { FolderStructureModeSection } from './FolderStructureModeSection';
import { TagConfigurationSection } from './TagConfigurationSection';
import { ModuleManagementSection } from './ModuleManagementSection';
import { FilterSearchManagementSection } from './FilterSearchManagementSection';
import { ColorManagementSection } from './ColorManagementSection';
import { SettingsTabManagementSection } from './SettingsTabManagementSection';

export function renderSettingsPage(host: VesselEmail): React.ReactElement {
  return <SettingsPageView host={host} />;
}

// Every tab that can be hidden via Settings → Settings Management. Keep in
// sync with SETTINGS_TAB_CATALOG in backend/app/settings_tab_api.py.
const HIDEABLE_SETTINGS_TABS = [
  'Site Management', 'Vessel Settings', 'Tag Configuration', 'Module Management', 'Filter Search Management', 'Color Management', 'Audit Logs'
];

function SettingsPageView({ host }: { host: VesselEmail }): React.ReactElement {
    const { settingsTab, hiddenSettingsTabs } = host.state;

    // 'Settings Management' is always shown — it's the only place to
    // un-hide everything else, so it can never be hidden itself (the
    // server also strips it from any hidden list it's saved with).
    const visibleTabs = [
      ...HIDEABLE_SETTINGS_TABS.filter(tab => hiddenSettingsTabs.indexOf(tab) === -1),
      'Settings Management',
    ];

    // If the currently-selected tab was just hidden (by this admin or
    // another one), land on the first tab that's still visible instead of
    // rendering a blank/unreachable pane.
    React.useEffect(() => {
      if (visibleTabs.indexOf(settingsTab) === -1) {
        host.setState({ settingsTab: visibleTabs[0] as any });
      }
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [settingsTab, hiddenSettingsTabs]);

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 18, width: '100%', minHeight: '100%', boxSizing: 'border-box', padding: '8px 18px 0 0' }}>
        <div style={{ paddingLeft: 2 }}>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Settings</h2>
          <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Configure application settings and preferences.</p>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '220px minmax(0, 1fr)', gap: 0, width: '100%', maxWidth: 1280, minHeight: 660, background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', overflow: 'hidden', boxSizing: 'border-box', alignSelf: 'stretch' }}>
          {/* Settings Left Nav */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4, borderRight: '1px solid #f1f5f9', background: '#f8fafc', padding: '18px 14px', boxSizing: 'border-box' }}>
            {visibleTabs.map(tab => (
              <button
                key={tab}
                onClick={() => host.setState({ settingsTab: tab as any })}
                style={{
                  border: 'none', background: settingsTab === tab ? '#e8f1ff' : 'transparent',
                  color: settingsTab === tab ? '#0a66d0' : '#475569', fontWeight: settingsTab === tab ? 700 : 500,
                  fontSize: 13, padding: '8px 12px', borderRadius: 6, textAlign: 'left', cursor: 'pointer',
                  boxShadow: settingsTab === tab ? 'inset 0 0 0 1px rgba(10,102,208,0.08)' : 'none'
                }}
              >
                {tab}
              </button>
            ))}
          </div>

          {/* Settings Content Area */}
          <div style={{ width: '100%', minWidth: 0, padding: '18px 22px 22px', boxSizing: 'border-box' }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' }}>{settingsTab}</h3>

            {settingsTab === 'Site Management' && (
              <SiteIntegrationInfo host={host} />
            )}

            {settingsTab === 'Vessel Settings' && (
              <FolderStructureModeSection host={host} />
            )}

            {settingsTab === 'Tag Configuration' && (
              <TagConfigurationSection host={host} />
            )}

            {settingsTab === 'Module Management' && (
              <ModuleManagementSection host={host} />
            )}

            {settingsTab === 'Filter Search Management' && (
              <FilterSearchManagementSection host={host} />
            )}

            {settingsTab === 'Color Management' && (
              <ColorManagementSection host={host} />
            )}

            {settingsTab === 'Audit Logs' && (
              <SessionAuditLog host={host} />
            )}

            {settingsTab === 'Settings Management' && (
              <SettingsTabManagementSection host={host} />
            )}
          </div>
        </div>
      </div>
    );
}

/**
 * Renders the signed-in user's session audit trail from GET /api/sessions/audit
 * (backend/app/main.py, list_session_audit). That endpoint filters strictly by
 * the X-User-Email header, so this shows the current user's own login/logout
 * history — not a tenant-wide admin audit log (no such endpoint exists yet).
 */
function SessionAuditLog({ host }: { host: VesselEmail }): React.ReactElement {
  const [entries, setEntries] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [err, setErr] = React.useState<string | null>(null);

  const loadAuditLog = React.useCallback(async (): Promise<void> => {
    setLoading(true);
    setErr(null);
    try {
      const response = await fetch(`${host._base()}/api/sessions/audit?limit=50`, { headers: host._headers() });
      const data = await response.json();
      if (!response.ok) throw new Error((data && data.detail) || 'Failed to load audit log');
      setEntries(Array.isArray(data) ? data : []);
    } catch (e: any) {
      setErr(e?.message || 'Failed to load audit log');
    } finally {
      setLoading(false);
    }
  }, [host]);

  React.useEffect(() => {
    void loadAuditLog();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [host.props.userEmail]);

  const formatTime = (iso: string | null): string => {
    if (!iso) return '—';
    const d = new Date(iso);
    return isNaN(d.getTime()) ? '—' : d.toLocaleString();
  };

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
        <p style={{ margin: 0, fontSize: 13, color: '#64748b' }}>
          Your recent sign-in and sign-out activity for this application.
        </p>
        <button
          onClick={() => void loadAuditLog()}
          disabled={loading}
          style={{
            border: '1px solid #e2e8f0', background: '#fff', color: '#334155', fontSize: 12, fontWeight: 600,
            padding: '6px 12px', borderRadius: 6, cursor: loading ? 'default' : 'pointer',
          }}
        >
          {loading ? 'Refreshing…' : '⟳ Refresh'}
        </button>
      </div>

      {err && (
        <div style={{ background: '#fef2f2', border: '1px solid #fecaca', color: '#b91c1c', fontSize: 12, borderRadius: 6, padding: '8px 12px', marginBottom: 12 }}>
          {err}
        </div>
      )}

      {loading && entries.length === 0 && !err ? (
        <div style={{ color: '#94a3b8', fontSize: 13, padding: '24px 0', textAlign: 'center' }}>Loading audit log…</div>
      ) : entries.length === 0 && !err ? (
        <div style={{ color: '#94a3b8', fontSize: 13, padding: '24px 0', textAlign: 'center' }}>
          No audit log entries yet. Sign-in / sign-out activity will appear here as it happens.
        </div>
      ) : (
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12.5 }}>
            <thead>
              <tr style={{ borderBottom: '2px solid #e2e8f0', textTransform: 'uppercase', fontSize: 10.5, color: '#64748b', textAlign: 'left', letterSpacing: '0.04em' }}>
                <th style={{ padding: '8px 10px' }}>Event</th>
                <th style={{ padding: '8px 10px' }}>Detail</th>
                <th style={{ padding: '8px 10px' }}>Status</th>
                <th style={{ padding: '8px 10px' }}>IP Address</th>
                <th style={{ padding: '8px 10px' }}>Browser</th>
                <th style={{ padding: '8px 10px' }}>Login Time</th>
                <th style={{ padding: '8px 10px' }}>Logout Time</th>
                <th style={{ padding: '8px 10px' }}>Duration</th>
              </tr>
            </thead>
            <tbody>
              {entries.map((e, idx) => (
                <tr key={e.session_id ? `${e.session_id}-${idx}` : idx} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '8px 10px', fontWeight: 600, color: '#1e293b', whiteSpace: 'nowrap' }}>{e.event || '—'}</td>
                  <td style={{ padding: '8px 10px', color: '#475569' }}>{e.detail || '—'}</td>
                  <td style={{ padding: '8px 10px', whiteSpace: 'nowrap' }}>
                    {badge(e.status === 'success' || e.status === 'active' ? 'green' : e.status === 'failed' ? 'red' : 'default', e.status || '—')}
                  </td>
                  <td style={{ padding: '8px 10px', color: '#475569', whiteSpace: 'nowrap' }}>{e.ip_address || '—'}</td>
                  <td style={{ padding: '8px 10px', color: '#475569', whiteSpace: 'nowrap', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis' }} title={e.browser || ''}>{e.browser || '—'}</td>
                  <td style={{ padding: '8px 10px', color: '#64748b', whiteSpace: 'nowrap' }}>{formatTime(e.login_time)}</td>
                  <td style={{ padding: '8px 10px', color: '#64748b', whiteSpace: 'nowrap' }}>{formatTime(e.logout_time)}</td>
                  <td style={{ padding: '8px 10px', color: '#64748b', whiteSpace: 'nowrap' }}>{e.active_duration_formatted || '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function SiteIntegrationInfo({ host }: { host: VesselEmail }): React.ReactElement {
  const [info, setInfo] = React.useState<any>(null);
  const [adminConfig, setAdminConfig] = React.useState<any>(null);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [err, setErr] = React.useState<string | null>(null);
  const [isAdmin, setIsAdmin] = React.useState<boolean>(false);

  // Default site state
  const [settingDefaultKey, setSettingDefaultKey] = React.useState<string | null>(null);
  const [defaultMsg, setDefaultMsg] = React.useState<string | null>(null);

  // Site visibility toggle state
  const [togglingSiteKey, setTogglingSiteKey] = React.useState<string | null>(null);
  const [removingSiteKey, setRemovingSiteKey] = React.useState<string | null>(null);
  const [visibilityMsg, setVisibilityMsg] = React.useState<string | null>(null);

  // Add-site state
  const [showAddSite, setShowAddSite] = React.useState<boolean>(false);
  const [addSiteTenantSites, setAddSiteTenantSites] = React.useState<any[]>([]);
  const [addSiteTenantLoading, setAddSiteTenantLoading] = React.useState<boolean>(false);
  const [addSiteSearch, setAddSiteSearch] = React.useState<string>('');
  const [addSiteSelected, setAddSiteSelected] = React.useState<any>(null);
  const [addSiteDrives, setAddSiteDrives] = React.useState<any[]>([]);
  const [addSiteDrivesLoading, setAddSiteDrivesLoading] = React.useState<boolean>(false);
  const [addSiteSelDrive, setAddSiteSelDrive] = React.useState<any>(null);
  const [addSiteDisplayName, setAddSiteDisplayName] = React.useState<string>('');
  const [addSiteKey, setAddSiteKey] = React.useState<string>('');
  const [addSiteSaving, setAddSiteSaving] = React.useState<boolean>(false);
  const [addSiteMsg, setAddSiteMsg] = React.useState<string | null>(null);

  // Audit changes
  const [siteChanges, setSiteChanges] = React.useState<any[]>([]);
  const [showChanges, setShowChanges] = React.useState<boolean>(false);

  // Tenant-site discovery
  const [tenantSitesLoading, setTenantSitesLoading] = React.useState<boolean>(false);
  const [siteSearch, setSiteSearch] = React.useState<string>('');

  const loadTenantSites = React.useCallback(async (base: string, headers: any) => {
    setTenantSitesLoading(true);
    try {
      const r = await fetch(`${base}/api/admin/discover-sites`, { headers });
      const d = await r.json();
      if (r.ok && Array.isArray(d.sites)) {
        setAddSiteTenantSites(d.sites);
      }
    } catch {
      // ignore
    } finally {
      setTenantSitesLoading(false);
    }
  }, []);

  const loadAdminConfig = React.useCallback(async () => {
    const base = host._base();
    try {
      const r = await fetch(`${base}/api/admin/site-configuration?include_hidden=true`, { headers: host._headers() });
      if (r.status === 403) { setIsAdmin(false); return; }
      const d = await r.json();
      if (d) {
        setAdminConfig(d);
        setIsAdmin(true);
      }
    } catch {
      // ignore
    }
  }, [host]);

  const loadSiteChanges = React.useCallback(async (): Promise<void> => {
    try {
      const response = await fetch(`${host._base()}/api/admin/site-changes?limit=10`, { headers: host._headers() });
      const data = await response.json();
      if (response.ok) setSiteChanges(data.changes || []);
    } catch {
      // Keep the existing audit list when refresh fails.
    }
  }, [host]);

  const handleSetDefaultSite = async (siteKey: string) => {
    setSettingDefaultKey(siteKey);
    setDefaultMsg(null);
    try {
      const r = await fetch(`${host._base()}/api/admin/set-default-site`, {
        method: 'POST',
        headers: { ...host._headers(), 'Content-Type': 'application/json' },
        body: JSON.stringify({ site_name: siteKey }),
      });
      const d = await r.json();
      if (!r.ok) throw new Error(d.detail || 'Failed to set default site');
      setDefaultMsg(`✓ Successfully set "${d.new_site_name || siteKey}" as the default site.`);
      setInfo((prev: any) => ({ ...(prev || {}), active_site: d.new_site, site_name: d.new_site_name }));
      setAdminConfig((prev: any) => prev ? { ...prev, current_site: d.new_site, current_site_name: d.new_site_name } : prev);
      void Promise.all([host._loadDocumentSites(), host._loadData()]).catch(() => undefined);
      setTimeout(() => setDefaultMsg(null), 4000);
    } catch (e: any) {
      setDefaultMsg(`⚠️ ${e?.message || 'Failed to set default site'}`);
      setTimeout(() => setDefaultMsg(null), 4000);
    } finally {
      setSettingDefaultKey(null);
    }
  };

  const handleToggleVisibility = async (siteKey: string, currentIsHidden: boolean) => {
    setTogglingSiteKey(siteKey);
    setVisibilityMsg(null);
    try {
      const r = await fetch(`${host._base()}/api/admin/site-configurations/${encodeURIComponent(siteKey)}/visibility`, {
        method: 'PATCH',
        headers: { ...host._headers(), 'Content-Type': 'application/json' },
        body: JSON.stringify({ is_hidden: !currentIsHidden }),
      });
      const d = await r.json();
      if (!r.ok) throw new Error(d.detail || 'Failed to update visibility');
      await Promise.all([loadAdminConfig(), loadSiteChanges()]);
      // Documents/Vessels' own site list, plus the Home dashboard's cached
      // counts. Deliberately UNforced (no force_refresh): the backend just
      // reconciled its site list in place when the visibility change above
      // invalidated the cache, so an unforced read already reflects this
      // instantly — forcing here would instead make this request itself
      // wait on a full live rescan of every site (20-30s+), which is the
      // "stuck for 20-30 seconds" symptom this replaces.
      void Promise.all([
        host._loadDocumentSites(),
        host._loadData(),
        host._loadDashboardStats(),
      ]).catch(() => undefined);
    } catch (e: any) {
      setVisibilityMsg(`⚠️ ${e?.message || 'Failed to update visibility'}`);
      setTimeout(() => setVisibilityMsg(null), 4000);
    } finally {
      setTogglingSiteKey(null);
    }
  };

  const handleRemoveSite = async (siteKey: string, displayName: string): Promise<void> => {
    if (!window.confirm(`Remove "${displayName}" from the Site Management list? The SharePoint site and its files will not be deleted, but this cannot be undone from here.`)) return;
    setRemovingSiteKey(siteKey);
    setVisibilityMsg(null);
    try {
      const r = await fetch(`${host._base()}/api/admin/site-configurations/${encodeURIComponent(siteKey)}`, {
        method: 'DELETE',
        headers: host._headers(),
      });
      const d = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error(d.detail || 'Failed to remove site');
      await Promise.all([loadAdminConfig(), loadSiteChanges()]);
      // Unforced — see the matching comment in handleToggleVisibility.
      void Promise.all([
        host._loadDocumentSites(),
        host._loadData(),
        host._loadDashboardStats(),
      ]).catch(() => undefined);
    } catch (e: any) {
      setVisibilityMsg(`⚠️ ${e?.message || 'Failed to remove site'}`);
    } finally {
      setRemovingSiteKey(null);
    }
  };

  React.useEffect(() => {
    const base = host._base();
    const userEmail = host.props.userEmail;

    // Fetch basic site info
    fetch(`${base}/api/config/site-info`, { headers: host._headers() })
      .then(r => r.json())
      .then(d => setInfo(d))
      .catch((e: any) => setErr(e?.message || 'Failed to load site information'));

    if (userEmail) {
      fetch(`${base}/api/admin/site-configuration?include_hidden=true`, { headers: host._headers() })
        .then(r => {
          if (r.status === 403) { setIsAdmin(false); setLoading(false); return null; }
          return r.json();
        })
        .then(d => {
          if (d) {
            setAdminConfig(d);
            setIsAdmin(true);

            void loadSiteChanges();

            void loadTenantSites(base, host._headers());
          }
          setLoading(false);
        })
        .catch(() => { setIsAdmin(false); setLoading(false); });
    } else {
      setLoading(false);
    }
  }, [host, host.props.userEmail, loadTenantSites, loadSiteChanges]);

  const filteredConfiguredSites = (adminConfig?.available_sites || []).filter((s: any) => {
    const q = siteSearch.toLowerCase();
    return !q || `${s.display_name} ${s.name}`.toLowerCase().includes(q);
  });

  if (loading) return <div style={{ color: '#64748b', fontSize: 13 }}>Loading site configuration...</div>;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14, width: '100%' }}>
      <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16, width: '100%', boxSizing: 'border-box' }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: '#0f172a', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
          <span>🌐</span> Default Site: {info?.site_name || info?.active_site}
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '170px minmax(0, 1fr)', columnGap: 16, rowGap: 8, fontSize: 12, width: '100%' }}>
          <span style={{ fontWeight: 600, color: '#64748b' }}>Site Key:</span>
          <span style={{ fontFamily: 'monospace', color: '#0f172a', fontWeight: 700, wordBreak: 'break-word' }}>{info?.active_site}</span>

          <span style={{ fontWeight: 600, color: '#64748b' }}>Database:</span>
          <span style={{ fontFamily: 'monospace', color: info?.db_configured ? '#15803d' : '#b45309', fontWeight: 600 }}>
            {info?.db_name || (info?.db_configured ? 'Connected' : 'In-Memory Stub')}
          </span>

          <span style={{ fontWeight: 600, color: '#64748b' }}>SharePoint Status:</span>
          <span style={{ color: info?.sp_configured ? '#15803d' : '#64748b', fontWeight: 600 }}>
            {info?.sp_configured ? '✅ Configured' : '⚪ Stub / Local Mode'}
          </span>

          <span style={{ fontWeight: 600, color: '#64748b' }}>Drive ID:</span>
          <span style={{ fontFamily: 'monospace', color: '#475569', fontSize: 11, wordBreak: 'break-all' }}>
            {info?.drive_id || '—'}
          </span>

          <span style={{ fontWeight: 600, color: '#64748b' }}>Backend Mode:</span>
          <span style={{ color: '#0284c7', fontWeight: 600 }}>{info?.mode}</span>
        </div>
      </div>

      {isAdmin && adminConfig && (
        <div style={{ background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 8, padding: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: '#166534', marginBottom: 6, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <span>👤 Configured SharePoint Sites & Default Site</span>
            <button
              onClick={() => {
                void loadAdminConfig();
                void loadTenantSites(host._base(), host._headers());
              }}
              disabled={tenantSitesLoading}
              title="Refresh all tenant and configured sites"
              style={{ padding: '4px 10px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: 5, fontSize: 11, fontWeight: 600, cursor: tenantSitesLoading ? 'wait' : 'pointer' }}
            >
              {tenantSitesLoading ? '⏳ Loading…' : '🔄 Refresh Sites'}
            </button>
          </div>
          <p style={{ margin: '0 0 10px', fontSize: 12, color: '#166534', lineHeight: 1.5 }}>
            Select the default site and manage site visibility. Hiding a site removes it from Sites, Documents, and Vessel Management modules. Hidden sites remain disabled here until you click Unhide.
          </p>

          {/* Search filter */}
          <input
            value={siteSearch}
            onChange={e => setSiteSearch(e.target.value)}
            placeholder="🔍 Search site by name or URL…"
            style={{ width: '100%', boxSizing: 'border-box', padding: '7px 10px', border: '1px solid #bbf7d0', borderRadius: 6, fontSize: 12, marginBottom: 8, outline: 'none' }}
          />

          {defaultMsg && (
            <div style={{ background: defaultMsg.startsWith('✓') ? '#dcfce7' : '#fee2e2', border: `1px solid ${defaultMsg.startsWith('✓') ? '#86efac' : '#fca5a5'}`, color: defaultMsg.startsWith('✓') ? '#166534' : '#991b1b', borderRadius: 6, padding: '6px 10px', fontSize: 12, marginBottom: 8 }}>
              {defaultMsg}
            </div>
          )}

          {visibilityMsg && (
            <div style={{ background: '#fee2e2', border: '1px solid #fca5a5', color: '#991b1b', borderRadius: 6, padding: '6px 10px', fontSize: 12, marginBottom: 8 }}>
              {visibilityMsg}
            </div>
          )}

          {/* ── Configured Sites card list ── */}
          <div style={{ marginBottom: 12 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: '#166534' }}>Configured Sites:</label>
              <span style={{ fontSize: 11, color: '#64748b' }}>Hidden sites remain listed here for unhide.</span>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              {filteredConfiguredSites.map((site: any) => {
                const isCurrent = site.name === adminConfig.current_site;
                const isHidden = !!site.is_hidden;
                const isToggling = togglingSiteKey === site.name;
                const isSettingDefault = settingDefaultKey === site.name;
                const isRemoving = removingSiteKey === site.name;

                return (
                  <div
                    key={`cfg-${site.name}`}
                    style={{
                      display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                      padding: '8px 12px', borderRadius: 8,
                      border: `1px solid ${isCurrent ? '#86efac' : isHidden ? '#cbd5e1' : '#d1fae5'}`,
                      background: isCurrent ? '#dcfce7' : isHidden ? '#f8fafc' : '#f0fdf4',
                      opacity: isHidden && !isCurrent ? 0.65 : 1,
                      transition: 'all 0.15s',
                    }}
                  >
                    <div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <span style={{ fontSize: 13, fontWeight: 600, color: isHidden ? '#64748b' : '#166534' }}>
                          {site.display_name || site.name}
                        </span>
                        {isCurrent && (
                          <span style={{ fontSize: 10, fontWeight: 700, background: '#16a34a', color: '#fff', padding: '2px 7px', borderRadius: 8 }}>
                            ★ DEFAULT
                          </span>
                        )}
                        {isHidden && (
                          <span style={{ fontSize: 10, fontWeight: 700, background: '#e2e8f0', color: '#475569', padding: '2px 7px', borderRadius: 8 }}>
                            HIDDEN
                          </span>
                        )}
                      </div>
                      <span style={{ display: 'block', fontSize: 10, color: '#64748b', fontFamily: 'monospace', marginTop: 2 }}>
                        {site.name}
                      </span>
                    </div>
                    <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                      {/* Set as Default button */}
                      {!isCurrent && !isHidden && (
                        <button
                          onClick={() => handleSetDefaultSite(site.name)}
                          disabled={isSettingDefault || isToggling}
                          title="Set this site as the default site"
                          style={{
                            padding: '4px 10px',
                            background: isSettingDefault ? '#93c5fd' : '#0284c7',
                            color: '#fff',
                            border: 'none',
                            borderRadius: 5,
                            fontSize: 11,
                            fontWeight: 600,
                            cursor: isSettingDefault ? 'wait' : 'pointer',
                          }}
                        >
                          {isSettingDefault ? '⏳ Setting…' : '★ Set as Default'}
                        </button>
                      )}
                      {/* Hidden rows stay visible here, but are disabled everywhere else. */}
                      {isHidden && (
                        <button
                          type="button"
                          disabled
                          style={{
                            padding: '4px 10px', background: '#e2e8f0', color: '#64748b',
                            border: '1px solid #cbd5e1', borderRadius: 5, fontSize: 11,
                            fontWeight: 600, cursor: 'not-allowed',
                          }}
                        >
                          HIDDEN
                        </button>
                      )}
                      <button
                        onClick={() => handleToggleVisibility(site.name, isHidden)}
                        disabled={isCurrent || isToggling || (!isHidden && isCurrent)}
                        title={
                          isCurrent
                            ? 'The default site cannot be hidden. Set another site as default first.'
                            : isHidden ? 'Make this site visible across all modules' : 'Hide this site from Sites, Documents, and Vessel Management'
                        }
                        style={{
                          padding: '4px 10px',
                          background: isHidden ? '#eff6ff' : isCurrent ? '#f1f5f9' : '#fff',
                          color: isHidden ? '#1d4ed8' : isCurrent ? '#94a3b8' : '#475569',
                          border: `1px solid ${isHidden ? '#bfdbfe' : isCurrent ? '#e2e8f0' : '#cbd5e1'}`,
                          borderRadius: 5,
                          fontSize: 11,
                          fontWeight: 600,
                          cursor: isCurrent ? 'not-allowed' : isToggling ? 'wait' : 'pointer',
                        }}
                      >
                        {isToggling ? '⏳…' : isHidden ? 'Unhide' : 'Hide'}
                      </button>
                      {/* Shown for every configured site, not just ones added via
                          "Add New Site" — an .env-discovered site (e.g. local/dev/prod)
                          has no DB row until it's removed or hidden, but it should
                          still be removable. */}
                      <button
                        type="button"
                        onClick={() => void handleRemoveSite(site.name, site.display_name || site.name)}
                        disabled={isCurrent || isToggling || isRemoving}
                        title={isCurrent ? 'Switch to a different site before removing this registration.' : 'Permanently remove this site from the Site Management list'}
                        style={{
                          padding: '4px 10px', background: '#fff', color: isCurrent ? '#94a3b8' : '#b91c1c',
                          border: `1px solid ${isCurrent ? '#e2e8f0' : '#fecaca'}`, borderRadius: 5,
                          fontSize: 11, fontWeight: 600, cursor: isCurrent ? 'not-allowed' : isRemoving ? 'wait' : 'pointer',
                        }}
                      >
                        {isRemoving ? '⏳ Removing…' : 'Remove'}
                      </button>
                    </div>
                  </div>
                );
              })}
              {filteredConfiguredSites.length === 0 && (
                <div style={{ fontSize: 12, color: '#64748b', padding: '8px 0' }}>
                  No configured sites match your search.
                </div>
              )}
            </div>
          </div>

          {/* ── Add New Site section ── */}
          <div style={{ marginBottom: 12, border: '1px solid #d1fae5', borderRadius: 8, overflow: 'hidden' }}>
            <button
              onClick={async () => {
                const next = !showAddSite;
                setShowAddSite(next);
                if (next && addSiteTenantSites.length === 0) {
                  setAddSiteTenantLoading(true);
                  try {
                    const r = await fetch(`${host._base()}/api/admin/discover-sites`, { headers: host._headers() });
                    const d = await r.json();
                    setAddSiteTenantSites(d.sites || []);
                  } catch { /* ignore */ } finally { setAddSiteTenantLoading(false); }
                }
              }}
              style={{ width: '100%', padding: '9px 14px', background: '#f0fdf4', border: 'none', color: '#166534', fontSize: 12, fontWeight: 700, cursor: 'pointer', textAlign: 'left' }}
            >
              {showAddSite ? '✕ Cancel Add Site' : '➕ Add New Site'}
            </button>
            {showAddSite && (
              <div style={{ padding: 14, background: '#fff' }}>
                <div style={{ fontSize: 12, color: '#475569', marginBottom: 10 }}>Search and select a SharePoint site to register as a new configured site.</div>
                <input
                  value={addSiteSearch}
                  onChange={e => setAddSiteSearch(e.target.value)}
                  placeholder="🔍 Search tenant sites…"
                  style={{ width: '100%', boxSizing: 'border-box', padding: '7px 10px', border: '1px solid #e2e8f0', borderRadius: 6, fontSize: 12, marginBottom: 8, outline: 'none' }}
                />
                {addSiteTenantLoading && <div style={{ fontSize: 12, color: '#64748b', marginBottom: 8 }}>⏳ Loading tenant sites…</div>}
                <div style={{ maxHeight: 160, overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: 6, marginBottom: 10 }}>
                  {addSiteTenantSites
                    .filter(s => !addSiteSearch || `${s.name} ${s.web_url}`.toLowerCase().includes(addSiteSearch.toLowerCase()))
                    .map(ts => (
                      <button
                        key={ts.id}
                        onClick={async () => {
                          setAddSiteSelected(ts);
                          setAddSiteSelDrive(null);
                          setAddSiteDrives([]);
                          setAddSiteDisplayName(ts.name || '');
                          const rawKey = (ts.name || '').toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
                          setAddSiteKey(rawKey);
                          setAddSiteDrivesLoading(true);
                          try {
                            const r = await fetch(`${host._base()}/api/admin/discover-sites/${encodeURIComponent(ts.id)}/drives`, { headers: host._headers() });
                            const d = await r.json();
                            const drives = r.ok ? (d.drives || []) : [];
                            setAddSiteDrives(drives);
                            if (drives.length === 1) setAddSiteSelDrive(drives[0]);
                          } catch { setAddSiteDrives([]); } finally { setAddSiteDrivesLoading(false); }
                        }}
                        style={{
                          display: 'block', width: '100%', textAlign: 'left', padding: '8px 12px',
                          border: 0, borderBottom: '1px solid #f1f5f9',
                          background: addSiteSelected?.id === ts.id ? '#eff6ff' : '#fff',
                          cursor: 'pointer', fontSize: 12, color: '#0f172a',
                        }}
                      >
                        <strong>{ts.name}</strong>
                        <span style={{ display: 'block', fontSize: 10, color: '#64748b' }}>{ts.web_url}</span>
                      </button>
                    ))}
                </div>

                {/* Drive picker */}
                {addSiteSelected && (
                  <div style={{ marginBottom: 10 }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: '#0f172a', marginBottom: 6 }}>Document Library:</div>
                    {addSiteDrivesLoading && <div style={{ fontSize: 11, color: '#64748b' }}>Loading libraries…</div>}
                    {!addSiteDrivesLoading && addSiteDrives.map(drv => (
                      <button
                        key={drv.id}
                        onClick={() => setAddSiteSelDrive(drv)}
                        style={{
                          display: 'block', width: '100%', textAlign: 'left', padding: '7px 10px', marginBottom: 4,
                          border: `1px solid ${addSiteSelDrive?.id === drv.id ? '#16a34a' : '#e2e8f0'}`,
                          borderRadius: 5, background: addSiteSelDrive?.id === drv.id ? '#dcfce7' : '#fff',
                          cursor: 'pointer', fontSize: 11,
                        }}
                      >
                        <strong>{drv.name}</strong>
                        <span style={{ display: 'block', fontSize: 10, color: '#64748b' }}>{drv.drive_type}</span>
                      </button>
                    ))}
                  </div>
                )}

                {/* Display name + site key inputs */}
                {addSiteSelected && addSiteSelDrive && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 10 }}>
                    <input
                      value={addSiteDisplayName}
                      onChange={e => setAddSiteDisplayName(e.target.value)}
                      placeholder="Display Name (e.g. Vessel Mgmt - Production)"
                      style={{ padding: '7px 10px', border: '1px solid #cbd5e1', borderRadius: 6, fontSize: 12, outline: 'none' }}
                    />
                    <input
                      value={addSiteKey}
                      onChange={e => setAddSiteKey(e.target.value.toLowerCase().replace(/[^a-z0-9_]/g, ''))}
                      placeholder="Internal site key (e.g. vessel_prod)"
                      style={{ padding: '7px 10px', border: '1px solid #cbd5e1', borderRadius: 6, fontSize: 12, fontFamily: 'monospace', outline: 'none' }}
                    />
                  </div>
                )}

                {addSiteMsg && (
                  <div style={{ background: addSiteMsg.startsWith('✓') ? '#dcfce7' : '#fee2e2', border: `1px solid ${addSiteMsg.startsWith('✓') ? '#86efac' : '#fca5a5'}`, color: addSiteMsg.startsWith('✓') ? '#166534' : '#991b1b', borderRadius: 6, padding: '6px 10px', fontSize: 12, marginBottom: 8 }}>
                    {addSiteMsg}
                  </div>
                )}

                <button
                  disabled={!addSiteSelected || !addSiteSelDrive || !addSiteDisplayName.trim() || !addSiteKey.trim() || addSiteSaving}
                  onClick={async () => {
                    if (!addSiteSelected || !addSiteSelDrive || !addSiteDisplayName.trim() || !addSiteKey.trim()) return;
                    setAddSiteSaving(true);
                    setAddSiteMsg(null);
                    try {
                      const r = await fetch(`${host._base()}/api/admin/site-configurations`, {
                        method: 'POST',
                        headers: { ...host._headers(), 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                          site_key: addSiteKey.trim(),
                          display_name: addSiteDisplayName.trim(),
                          site_name: addSiteSelected.name,
                          site_id: addSiteSelected.id,
                          drive_id: addSiteSelDrive.id,
                        }),
                      });
                      const d = await r.json();
                      if (!r.ok) throw new Error(d.detail || 'Failed to save site');
                      setAddSiteMsg(`✓ Site "${addSiteDisplayName.trim()}" registered successfully.`);
                      // Refresh the configured-sites list, and the Home
                      // dashboard's cached counts. Deliberately UNforced: the
                      // backend already reconciled the new site into the
                      // dashboard's site list the instant the save above
                      // invalidated the cache (see
                      // invalidate_dashboard_stats_cache), so it appears
                      // immediately with placeholder counts while its real
                      // scan runs in the background — forcing here would
                      // instead make this request block on a full live
                      // rescan of every site before the new one even shows up.
                      await Promise.all([loadAdminConfig(), loadSiteChanges()]);
                      void host._loadDocumentSites();
                      void host._loadDashboardStats().catch(() => undefined);
                      // Reset add-site form
                      setTimeout(() => {
                        setShowAddSite(false);
                        setAddSiteSelected(null); setAddSiteSelDrive(null);
                        setAddSiteDrives([]); setAddSiteDisplayName(''); setAddSiteKey('');
                        setAddSiteMsg(null); setAddSiteSearch('');
                      }, 2000);
                    } catch (e: any) {
                      setAddSiteMsg(`⚠️ ${e?.message || 'Failed to save'}`);
                    } finally {
                      setAddSiteSaving(false);
                    }
                  }}
                  style={{
                    padding: '8px 16px',
                    background: (!addSiteSelected || !addSiteSelDrive || !addSiteDisplayName.trim() || !addSiteKey.trim() || addSiteSaving) ? '#cbd5e1' : '#16a34a',
                    color: '#fff', border: 'none', borderRadius: 6, fontSize: 12, fontWeight: 600,
                    cursor: addSiteSaving ? 'wait' : 'pointer',
                  }}
                >
                  {addSiteSaving ? '⏳ Saving…' : '✅ Save & Register Site'}
                </button>
              </div>
            )}
          </div>

          {showChanges && siteChanges.length > 0 && (
            <div style={{ marginTop: 16, borderTop: '1px solid #bbf7d0', paddingTop: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: '#166534', marginBottom: 8 }}>Recent Changes:</div>
              <div style={{ maxHeight: 200, overflowY: 'auto', fontSize: 11, color: '#166534' }}>
                {siteChanges.map((change: any) => (
                  <div key={change.id} style={{ padding: '6px 0', borderBottom: '1px solid #dcfce7' }}>
                    <div>
                      <strong>
                        {change.action === 'hidden' ? 'Hidden' : change.action === 'unhidden' ? 'Unhidden' : change.action === 'removed' ? 'Removed' : `${change.previous_site} → ${change.new_site}`}
                      </strong>
                      {change.action && change.action !== 'success' ? ` ${change.new_site_name || change.new_site}` : ` (${change.status})`}
                      {change.action && change.action !== 'success' && <span> ({change.status})</span>}
                    </div>
                    <div style={{ fontSize: 10, color: '#15803d' }}>
                      {change.changed_by_email} • {new Date(change.created_at).toLocaleString()}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {siteChanges.length > 0 && (
            <button
              onClick={() => setShowChanges(!showChanges)}
              style={{ marginTop: 8, background: 'transparent', border: 'none', color: '#166534', fontSize: 11, fontWeight: 600, cursor: 'pointer', textDecoration: 'underline' }}
            >
              {showChanges ? '▼ Hide' : '▶ Show'} Recent Changes
            </button>
          )}
        </div>
      )}

      {err && (
        <div style={{ background: '#fee2e2', border: '1px solid #fca5a5', borderRadius: 6, padding: 12, fontSize: 12, color: '#991b1b' }}>
          ⚠️ {err}
        </div>
      )}
    </div>
  );
}

function VesselSiteProvisioningPanel({ host }: { host: VesselEmail }): React.ReactElement {
  const [sites, setSites] = React.useState<Array<{
    site_key: string;
    display_name: string;
    site_name: string;
    site_id: string;
    drive_id: string;
    is_available_for_provisioning: boolean;
    is_default_provisioning: boolean;
    is_hidden?: boolean;
  }>>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [saving, setSaving] = React.useState<boolean>(false);
  const [siteMsg, setSiteMsg] = React.useState<string | null>(null);
  const [siteErr, setSiteErr] = React.useState<string | null>(null);

  // Per-vessel provisioning
  const [vesselSearch, setVesselSearch] = React.useState<string>('');
  const [selectedVessel, setSelectedVessel] = React.useState<VesselRecord | null>(null);
  const [targetSiteKeys, setTargetSiteKeys] = React.useState<string[]>([]);
  const [provisioning, setProvisioning] = React.useState<boolean>(false);
  const [provMsg, setProvMsg] = React.useState<string | null>(null);
  const [provErr, setProvErr] = React.useState<string | null>(null);
  const [provResults, setProvResults] = React.useState<Record<string, any> | null>(null);

  const loadSites = () => {
    setLoading(true);
    fetch(`${host._base()}/api/admin/site-provisioning/sites`, { headers: host._headers() })
      .then(r => r.ok ? r.json() : Promise.reject(r))
      .then(d => {
        setSites(d.sites || []);
      })
      .catch((e: any) => {
        setSiteErr(e?.message || 'Failed to load site provisioning settings');
      })
      .finally(() => setLoading(false));
  };

  React.useEffect(() => {
    loadSites();
  }, []);

  const handleSaveSiteSettings = async () => {
    setSaving(true);
    setSiteMsg(null);
    setSiteErr(null);
    try {
      const res = await fetch(`${host._base()}/api/admin/site-provisioning/sites`, {
        method: 'PUT',
        headers: {
          ...host._headers(),
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          sites: sites.map(s => ({
            site_key: s.site_key,
            is_available_for_provisioning: s.is_available_for_provisioning,
            is_default_provisioning: s.is_default_provisioning,
            display_name: s.display_name,
          })),
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data?.detail || 'Failed to save settings');
      if (data.sites) setSites(data.sites);
      setSiteMsg('✓ Site provisioning settings saved successfully.');
      setTimeout(() => setSiteMsg(null), 3000);
    } catch (err: any) {
      setSiteErr(err?.message || 'Failed to save site settings');
    } finally {
      setSaving(false);
    }
  };

  const handleOpenVesselModal = (v: VesselRecord) => {
    setSelectedVessel(v);
    setTargetSiteKeys(v.provisioned_site_ids || []);
    setProvMsg(null);
    setProvErr(null);
    setProvResults(null);
  };

  const handleProvisionSites = async () => {
    if (!selectedVessel) return;
    setProvisioning(true);
    setProvMsg(null);
    setProvErr(null);
    setProvResults(null);
    try {
      const res = await fetch(`${host._base()}/api/vessels/${selectedVessel.id}/provision-sites`, {
        method: 'POST',
        headers: {
          ...host._headers(),
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ site_keys: targetSiteKeys }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data?.detail || 'Provisioning failed');
      setProvResults(data.results || {});
      const newlyProvisioned = data.provisioned_sites || targetSiteKeys;
      setProvMsg(`✓ Provisioning complete across target sites.`);

      const updatedVessels = host.state.vessels.map(v =>
        v.id === selectedVessel.id
          ? { ...v, is_provisioned: true, provisioned_site_ids: newlyProvisioned }
          : v
      );
      host.setState({ vessels: updatedVessels });
      setSelectedVessel({ ...selectedVessel, is_provisioned: true, provisioned_site_ids: newlyProvisioned });
    } catch (err: any) {
      setProvErr(err?.message || 'Provisioning request failed');
    } finally {
      setProvisioning(false);
    }
  };

  const vessels = host.state.vessels || [];
  const filteredVessels = vessels.filter(v => {
    const q = vesselSearch.trim().toLowerCase();
    if (!q) return true;
    return (v.name || '').toLowerCase().includes(q) || (v.imo || '').includes(q);
  });

  const siteNameMap: Record<string, string> = {};
  sites.forEach(s => { siteNameMap[s.site_key] = s.display_name; });

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      {/* Intro banner */}
      <div style={{ background: '#f0f9ff', border: '1px solid #bfdbfe', borderRadius: 10, padding: 14 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: '#1e40af', marginBottom: 4 }}>
          Multi-Site Vessel Provisioning & DMS Folder Structure
        </div>
        <div style={{ fontSize: 12, color: '#1e3a8a', lineHeight: 1.5 }}>
          Control which connected SharePoint sites are available for vessel DMS folder structures. Adding sites to an existing vessel creates standard folders on the newly selected site only, preserving existing sites intact.
        </div>
        <div style={{ marginTop: 8, fontSize: 11, color: '#475569' }}>
          <strong>Note:</strong> File uploads reside inside the targeted site document library and do not replicate files across sites.
        </div>
        <div style={{ marginTop: 10, display: 'flex', justifyContent: 'flex-start' }}>
        </div>
      </div>

      {siteMsg && <div style={{ background: '#dcfce7', border: '1px solid #86efac', color: '#15803d', padding: '10px 14px', borderRadius: 8, fontSize: 13 }}>{siteMsg}</div>}
      {siteErr && <div style={{ background: '#fee2e2', border: '1px solid #fca5a5', color: '#b91c1c', padding: '10px 14px', borderRadius: 8, fontSize: 13 }}>{siteErr}</div>}

      {/* Section 1: Connected Sites */}
      <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <div>
            <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: '#0f172a' }}>1. Connected Sites & Availability</h4>
            <span style={{ fontSize: 12, color: '#64748b' }}>Configure which sites can receive vessel folder structures.</span>
          </div>
          <button
            onClick={handleSaveSiteSettings}
            disabled={saving || loading}
            style={{
              padding: '8px 16px',
              borderRadius: 6,
              border: 'none',
              background: saving ? '#cbd5e1' : '#0078d4',
              color: '#fff',
              fontSize: 12,
              fontWeight: 600,
              cursor: saving ? 'wait' : 'pointer',
            }}
          >
            {saving ? 'Saving...' : 'Save Site Settings'}
          </button>
        </div>

        {loading ? (
          <div style={{ padding: 20, textAlign: 'center', color: '#64748b', fontSize: 13 }}>Loading sites...</div>
        ) : sites.length === 0 ? (
          <div style={{ padding: 20, textAlign: 'center', color: '#64748b', fontSize: 13 }}>No connected sites discovered.</div>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', textAlign: 'left', color: '#475569' }}>
                  <th style={{ padding: '10px 12px' }}>Site Display Name</th>
                  <th style={{ padding: '10px 12px' }}>Internal Key</th>
                  <th style={{ padding: '10px 12px', textAlign: 'center' }}>Available for Provisioning</th>
                  <th style={{ padding: '10px 12px', textAlign: 'center' }}>Default for New Vessels</th>
                </tr>
              </thead>
              <tbody>
                {sites.filter(s => !s.is_hidden).map(s => (
                  <tr key={s.site_key} style={{ borderBottom: '1px solid #f1f5f9' }}>
                    <td style={{ padding: '10px 12px', fontWeight: 600, color: '#0f172a' }}>
                      {s.display_name}
                    </td>
                    <td style={{ padding: '10px 12px', color: '#64748b' }}>
                      <code style={{ background: '#f1f5f9', padding: '2px 6px', borderRadius: 4, fontSize: 11 }}>{s.site_key}</code>
                    </td>
                    <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                      <input
                        type="checkbox"
                        checked={s.is_available_for_provisioning}
                        onChange={e => {
                          const checked = e.target.checked;
                          setSites(sites.map(item => item.site_key === s.site_key ? { ...item, is_available_for_provisioning: checked } : item));
                        }}
                        style={{ width: 16, height: 16, cursor: 'pointer', accentColor: '#0078d4' }}
                      />
                    </td>
                    <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                      <input
                        type="checkbox"
                        disabled={!s.is_available_for_provisioning}
                        checked={s.is_default_provisioning}
                        onChange={e => {
                          const checked = e.target.checked;
                          setSites(sites.map(item => item.site_key === s.site_key ? { ...item, is_default_provisioning: checked } : item));
                        }}
                        style={{ width: 16, height: 16, cursor: s.is_available_for_provisioning ? 'pointer' : 'default', accentColor: '#0078d4' }}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Section 2: Vessel Provisioning Status */}
      <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <div>
            <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: '#0f172a' }}>2. Vessel Site Allocation & Status</h4>
            <span style={{ fontSize: 12, color: '#64748b' }}>View provisioned sites per vessel and trigger provisioning for newly added sites.</span>
          </div>
          <input
            type="text"
            value={vesselSearch}
            onChange={e => setVesselSearch(e.target.value)}
            placeholder="Search vessels by name or IMO..."
            style={{ padding: '6px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 12, width: 220 }}
          />
        </div>

        <div style={{ maxHeight: 380, overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: 8 }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', textAlign: 'left', color: '#475569', position: 'sticky', top: 0, zIndex: 1 }}>
                <th style={{ padding: '10px 12px' }}>Vessel</th>
                <th style={{ padding: '10px 12px' }}>IMO</th>
                <th style={{ padding: '10px 12px' }}>Provisioned Sites</th>
                <th style={{ padding: '10px 12px', textAlign: 'right' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredVessels.length === 0 ? (
                <tr>
                  <td colSpan={4} style={{ padding: 20, textAlign: 'center', color: '#64748b' }}>No vessels match your search.</td>
                </tr>
              ) : (
                filteredVessels.map(v => {
                  const provSites = v.provisioned_site_ids || [];
                  return (
                    <tr key={v.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                      <td style={{ padding: '10px 12px', fontWeight: 600, color: '#0f172a' }}>
                        🚢 {v.name}
                      </td>
                      <td style={{ padding: '10px 12px', color: '#64748b' }}>
                        {v.imo || '—'}
                      </td>
                      <td style={{ padding: '10px 12px' }}>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                          {provSites.length === 0 ? (
                            <span style={{ fontSize: 11, background: '#fef3c7', color: '#92400e', padding: '2px 8px', borderRadius: 10 }}>
                              Default site only
                            </span>
                          ) : (
                            provSites.map(sk => (
                              <span
                                key={sk}
                                style={{
                                  fontSize: 11,
                                  fontWeight: 600,
                                  background: '#dcfce7',
                                  color: '#15803d',
                                  padding: '2px 8px',
                                  borderRadius: 10,
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: 4,
                                }}
                              >
                                <span>✅</span>
                                {siteNameMap[sk] || sk}
                              </span>
                            ))
                          )}
                        </div>
                      </td>
                      <td style={{ padding: '10px 12px', textAlign: 'right' }}>
                        <button
                          onClick={() => handleOpenVesselModal(v)}
                          style={{
                            padding: '5px 12px',
                            borderRadius: 6,
                            border: '1px solid #cbd5e1',
                            background: '#fff',
                            color: '#0f172a',
                            fontSize: 11,
                            fontWeight: 600,
                            cursor: 'pointer',
                          }}
                        >
                          Manage Sites
                        </button>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Per-Vessel Provisioning Modal */}
      {selectedVessel && (
        <div
          style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }}
          onClick={e => { if (e.target === e.currentTarget && !provisioning) setSelectedVessel(null); }}
        >
          <div style={{ background: '#fff', borderRadius: 16, padding: 24, width: 480, maxWidth: '95vw', boxShadow: '0 20px 50px rgba(0,0,0,0.25)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
              <div>
                <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#0f172a' }}>
                  Manage Sites: {selectedVessel.name}
                </h3>
                <span style={{ fontSize: 12, color: '#64748b' }}>IMO: {selectedVessel.imo || '—'}</span>
              </div>
              {!provisioning && (
                <button
                  onClick={() => setSelectedVessel(null)}
                  style={{ background: 'none', border: 'none', fontSize: 18, color: '#94a3b8', cursor: 'pointer' }}
                >✕</button>
              )}
            </div>

            {provMsg && <div style={{ background: '#dcfce7', border: '1px solid #86efac', color: '#15803d', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{provMsg}</div>}
            {provErr && <div style={{ background: '#fee2e2', border: '1px solid #fca5a5', color: '#b91c1c', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{provErr}</div>}

            <div style={{ fontSize: 12, color: '#475569', marginBottom: 12, lineHeight: 1.4 }}>
              Select additional sites to provision the DMS folder tree. Already provisioned sites will not be recreated.
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 20 }}>
              {sites.filter(s => s.is_available_for_provisioning && !s.is_hidden).map(site => {
                const isAlreadyProvisioned = (selectedVessel.provisioned_site_ids || []).includes(site.site_key);
                const isChecked = targetSiteKeys.includes(site.site_key);

                return (
                  <label
                    key={site.site_key}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      padding: '10px 12px',
                      borderRadius: 8,
                      background: isAlreadyProvisioned ? '#f8fafc' : isChecked ? '#eff6ff' : '#fff',
                      border: `1px solid ${isChecked ? '#93c5fd' : '#e2e8f0'}`,
                      cursor: isAlreadyProvisioned ? 'default' : 'pointer',
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                      <input
                        type="checkbox"
                        disabled={isAlreadyProvisioned || provisioning}
                        checked={isChecked}
                        onChange={e => {
                          const next = e.target.checked
                            ? [...targetSiteKeys, site.site_key]
                            : targetSiteKeys.filter(k => k !== site.site_key);
                          setTargetSiteKeys(next);
                        }}
                        style={{ width: 16, height: 16, accentColor: '#0078d4' }}
                      />
                      <div>
                        <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{site.display_name}</div>
                        <div style={{ fontSize: 11, color: '#64748b' }}>Key: {site.site_key}</div>
                      </div>
                    </div>

                    {isAlreadyProvisioned ? (
                      <span style={{ fontSize: 11, fontWeight: 700, color: '#15803d', background: '#dcfce7', padding: '2px 8px', borderRadius: 10 }}>
                        ✅ Provisioned
                      </span>
                    ) : isChecked ? (
                      <span style={{ fontSize: 11, fontWeight: 600, color: '#1d4ed8', background: '#dbeafe', padding: '2px 8px', borderRadius: 10 }}>
                        ➕ Will Provision
                      </span>
                    ) : (
                      <span style={{ fontSize: 11, color: '#94a3b8' }}>Not provisioned</span>
                    )}
                  </label>
                );
              })}
            </div>

            {/* Results breakdown if present */}
            {provResults && (
              <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: 10, marginBottom: 16, fontSize: 12 }}>
                <div style={{ fontWeight: 700, marginBottom: 6, color: '#0f172a' }}>Provisioning Results:</div>
                {Object.entries(provResults).map(([sk, res]: [string, any]) => (
                  <div key={sk} style={{ display: 'flex', justifyContent: 'space-between', padding: '3px 0' }}>
                    <span>{siteNameMap[sk] || sk}:</span>
                    <span style={{ color: res.status === 'success' ? '#15803d' : '#b91c1c', fontWeight: 600 }}>
                      {res.status === 'success' ? '✓ Provisioned' : `✗ Failed: ${res.error || 'Error'}`}
                    </span>
                  </div>
                ))}
              </div>
            )}

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
              <button
                onClick={() => setSelectedVessel(null)}
                disabled={provisioning}
                style={{ padding: '8px 16px', borderRadius: 6, border: '1px solid #cbd5e1', background: '#fff', color: '#475569', fontSize: 12, fontWeight: 500, cursor: 'pointer' }}
              >
                Close
              </button>
              <button
                onClick={handleProvisionSites}
                disabled={provisioning || targetSiteKeys.filter(k => !(selectedVessel.provisioned_site_ids || []).includes(k)).length === 0}
                style={{
                  padding: '8px 18px',
                  borderRadius: 6,
                  border: 'none',
                  background: provisioning || targetSiteKeys.filter(k => !(selectedVessel.provisioned_site_ids || []).includes(k)).length === 0 ? '#cbd5e1' : 'linear-gradient(135deg, #0d9488, #0f766e)',
                  color: '#fff',
                  fontSize: 12,
                  fontWeight: 600,
                  cursor: provisioning ? 'wait' : 'pointer',
                }}
              >
                {provisioning ? 'Provisioning Folders...' : 'Provision Missing Sites'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

