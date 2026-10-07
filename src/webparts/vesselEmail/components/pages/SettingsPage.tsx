/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable @typescript-eslint/no-unused-expressions */
import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { Icon } from '@fluentui/react/lib/Icon';
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
import { VesselFolderTemplateSection } from './VesselFolderTemplateSection';
import { TagConfigurationSection } from './TagConfigurationSection';
import { ModuleManagementSection } from './ModuleManagementSection';
import { FilterSearchManagementSection } from './FilterSearchManagementSection';
import { ColorManagementSection } from './ColorManagementSection';
import { SettingsTabManagementSection } from './SettingsTabManagementSection';
import { clay } from '../clayTheme';
import {
  DmsPageHeader, dmsBtn, dmsRowBtn, dmsControlStyle, DMS_ON_ACCENT, DMS_FONT_DISPLAY,
  DMS_TABLE_CARD, DMS_TABLE, DMS_TH, DMS_TR, DMS_TD, DMS_TD_NAME,
} from '../dmsDesignSystem';

export function renderSettingsPage(host: VesselEmail): React.ReactElement {
  return <SettingsPageView host={host} />;
}

// Every tab that can be hidden via Settings → Settings Management. Keep in
// sync with SETTINGS_TAB_CATALOG in backend/app/settings_tab_api.py.
// Note: 'Site Management' used to be listed here too, but that tab moved to
// the main Sites page (see SiteManagementSection.tsx), so Settings no
// longer has a pane for it.
const HIDEABLE_SETTINGS_TABS = [
  'Vessel Settings', 'Tag Configuration', 'Module Management', 'Filter Search Management', 'Color Management', 'Audit Logs'
];

function SettingsPageView({ host }: { host: VesselEmail }): React.ReactElement {
    const { settingsTab, hiddenSettingsTabs } = host.state;

    // 'Settings Management' is always shown — it's the only place to
    // un-hide everything else, so it can never be hidden itself (the
    // server also strips it from any hidden list it's saved with).
    // Only admins see the configuration tabs; everyone else gets Appearance and
    // Audit Logs (their own session history). Admin status comes from the same
    // is_admin flag the config endpoints already return (null = still loading,
    // in which case the admin tabs are shown so an admin doesn't get bounced).
    const [isAdmin, setIsAdmin] = React.useState<boolean | null>(null);
    React.useEffect(() => {
      let cancelled = false;
      fetch(`${host._base()}/api/settings-tab-settings/config`, { headers: host._headers() })
        .then(r => (r.ok ? r.json() : Promise.reject(new Error(String(r.status)))))
        .then(d => { if (!cancelled) setIsAdmin(!!d?.is_admin); })
        .catch(() => { if (!cancelled) setIsAdmin(false); });
      return () => { cancelled = true; };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);
    const showAdminTabs = isAdmin !== false;
    // Configuration tabs (admin only) stay as plain Settings tabs; Appearance is personal.
    const configTabs = showAdminTabs ? HIDEABLE_SETTINGS_TABS.filter(tab => tab !== 'Audit Logs' && hiddenSettingsTabs.indexOf(tab) === -1) : [];
    const topTabs = ['Appearance', ...configTabs];
    // Administration dropdown (admin only): two items open Settings tabs, two open their own pages.
    const hiddenMods = host.state.hiddenModules;
    const adminItems: Array<{ label: string; tab?: string; view?: 'sites' | 'users' }> = showAdminTabs ? [
      ...(hiddenMods.indexOf('sites') === -1 ? [{ label: 'Site Management', view: 'sites' as const }] : []),
      ...(hiddenMods.indexOf('users') === -1 ? [{ label: 'User Management', view: 'users' as const }] : []),
      ...(hiddenSettingsTabs.indexOf('Audit Logs') === -1 ? [{ label: 'Audit Logs', tab: 'Audit Logs' }] : []),
      { label: 'Settings Management', tab: 'Settings Management' },
    ] : [];
    const adminTabNames = adminItems.filter(i => !!i.tab).map(i => i.tab as string);
    const visibleTabs = [...topTabs, ...adminTabNames];
    const activeInAdmin = adminTabNames.indexOf(settingsTab) !== -1;
    const [adminOpen, setAdminOpen] = React.useState<boolean>(false);
    // Opening one of its tabs (e.g. directly) expands the dropdown so the active item is visible.
    React.useEffect(() => { if (activeInAdmin) setAdminOpen(true); }, [activeInAdmin]);

    // If the currently-selected tab was just hidden (by this admin or
    // another one), land on the first tab that's still visible instead of
    // rendering a blank/unreachable pane.
    React.useEffect(() => {
      if (isAdmin === null) return;
      if (visibleTabs.indexOf(settingsTab) === -1) {
        host.setState({ settingsTab: visibleTabs[0] as any });
      }
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [settingsTab, hiddenSettingsTabs, isAdmin]);

    const TAB_ICONS: Record<string, string> = {
      'Appearance': 'Color', 'Vessel Settings': 'Ferry', 'Tag Configuration': 'Tag', 'Module Management': 'GridViewMedium',
      'Filter Search Management': 'Filter', 'Color Management': 'Color', 'Audit Logs': 'History', 'Settings Management': 'Settings',
      'Site Management': 'SharepointLogo', 'User Management': 'People',
    };
    const navIcon = (name: string, active: boolean): React.ReactElement => (
      <Icon iconName={TAB_ICONS[name] || 'Settings'} aria-hidden="true" style={{ fontSize: 16, marginRight: 10, verticalAlign: '-2px', color: active ? DMS_ON_ACCENT : 'var(--vdms-text-muted)' }} />
    );
    const navGroupLabel = (text: string): React.ReactElement => (
      <div style={{ margin: '14px 10px 4px', fontSize: 11, fontWeight: 600, letterSpacing: '0.08em', textTransform: 'uppercase', color: 'var(--vdms-text-muted)' }}>{text}</div>
    );
    const navBtnStyle = (active: boolean, indent = 0): React.CSSProperties => ({
      border: 'none',
      background: active ? clay.accentGradient : 'transparent',
      color: active ? DMS_ON_ACCENT : 'var(--vdms-text)', fontWeight: active ? 800 : 600,
      fontSize: 14, padding: indent ? '9px 12px 9px 30px' : '10px 12px', borderRadius: 12, textAlign: 'left', cursor: 'pointer',
      fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
      boxShadow: active ? clay.shadowButton : 'none',
    });

    return (
      <div className="vdms-settings" style={{ display: 'flex', flexDirection: 'column', gap: 20, width: '100%', minHeight: '100%', boxSizing: 'border-box', padding: '8px 0 0 0' }}>
        <div style={{ paddingLeft: 4 }}>
          <DmsPageHeader title="Settings" subtitle="Configure application settings and preferences." />
        </div>

        <div className="vdms-settings-shell" style={{ display: 'grid', gridTemplateColumns: 'minmax(230px, 280px) minmax(0, 1fr)', gap: 0, width: '100%', minHeight: 'calc(100vh - 250px)', background: 'var(--vdms-glass)', backdropFilter: 'blur(20px) saturate(1.3)', WebkitBackdropFilter: 'blur(20px) saturate(1.3)', borderRadius: clay.radiusCard + 16, border: '1px solid var(--vdms-line)', boxShadow: clay.shadowRaised, overflow: 'hidden', boxSizing: 'border-box', alignSelf: 'stretch' }}>
          {/* Settings Left Nav */}
          <div className="vdms-settings-nav" style={{ display: 'flex', flexDirection: 'column', gap: 6, borderRight: '1px solid var(--vdms-line)', background: 'var(--vdms-surface-alt)', padding: '22px 16px', boxSizing: 'border-box' }}>
            {navGroupLabel('Personal')}
            {topTabs.filter(t => t === 'Appearance').map(tab => (
              <button key={tab} onClick={() => host.setState({ settingsTab: tab as any })} style={navBtnStyle(settingsTab === tab)}>
                {navIcon(tab, settingsTab === tab)}{tab}
              </button>
            ))}
            {configTabs.length > 0 && navGroupLabel('Configuration')}
            {configTabs.map(tab => (
              <button key={tab} onClick={() => host.setState({ settingsTab: tab as any })} style={navBtnStyle(settingsTab === tab)}>
                {navIcon(tab, settingsTab === tab)}{tab}
              </button>
            ))}
            {adminItems.length > 0 && (
              <>
                <button
                  type="button"
                  aria-expanded={adminOpen}
                  onClick={() => setAdminOpen(o => !o)}
                  style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', border: 'none', background: 'transparent', cursor: 'pointer', margin: '10px 0 0', padding: '12px 16px', borderRadius: 14, color: 'var(--vdms-text)', fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif", fontSize: 12, fontWeight: 700, letterSpacing: '0.08em', textTransform: 'uppercase' }}
                >
                  Administration
                  <Icon iconName={adminOpen ? 'ChevronUp' : 'ChevronDown'} style={{ fontSize: 12 }} />
                </button>
                {adminOpen && adminItems.map(item => (
                  <button
                    key={item.label}
                    onClick={() => { if (item.tab) host.setState({ settingsTab: item.tab as any }); else if (item.view) void host._goToView(item.view); }}
                    style={navBtnStyle(!!item.tab && settingsTab === item.tab, 1)}
                  >
                    {navIcon(item.label, !!item.tab && settingsTab === item.tab)}{item.label}
                  </button>
                ))}
              </>
            )}
          </div>

          {/* Settings Content Area */}
          <div className="vdms-settings-body" style={{ width: '100%', minWidth: 0, padding: '26px 32px 32px', boxSizing: 'border-box' }}>
            <h3 style={{ margin: '0 0 20px', fontFamily: DMS_FONT_DISPLAY, fontSize: 24, fontWeight: 800, letterSpacing: '-0.025em', color: 'var(--vdms-text)' }}>{settingsTab}</h3>

            {settingsTab === 'Vessel Settings' && (
              <>
                <FolderStructureModeSection host={host} />
                <VesselFolderTemplateSection host={host} />
              </>
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

            {(settingsTab as string) === 'Appearance' && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12, maxWidth: 520 }}>
                <div style={{ color: 'var(--vdms-text-muted)', fontSize: 14 }}>Choose how the application looks for you.</div>
                <div role="radiogroup" aria-label="Theme" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, padding: 4, borderRadius: 12, border: '1px solid var(--vdms-line)', background: 'var(--vdms-surface-alt)' }}>
                  {([['night', 'Night', 'ClearNight'], ['light', 'Light', 'Sunny']] as const).map(([mode, label, icon]) => {
                    const active = (host.state.themeMode === 'night') === (mode === 'night');
                    return (
                      <button
                        key={mode}
                        type="button"
                        role="radio"
                        aria-checked={active}
                        onClick={() => { if (!active) host._toggleThemeMode(); }}
                        style={{ border: 'none', borderRadius: 9, padding: '12px 0', cursor: 'pointer', fontWeight: 700, fontSize: 15, fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif", display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, background: active ? clay.accentGradient : 'transparent', color: active ? DMS_ON_ACCENT : 'var(--vdms-text)', boxShadow: active ? clay.shadowButton : 'none' }}
                      >
                        <Icon iconName={icon} /> {label}
                      </button>
                    );
                  })}
                </div>
              </div>
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
        <p style={{ margin: 0, fontSize: 13, color: 'var(--vdms-text-muted)' }}>
          Your recent sign-in and sign-out activity for this application.
        </p>
        <button
          onClick={() => void loadAuditLog()}
          disabled={loading}
          style={dmsBtn('secondary', !loading)}
        >
          {loading ? 'Refreshing…' : <><Icon iconName="Refresh" aria-hidden="true" style={{ fontSize: 12 }} /> Refresh</>}
        </button>
      </div>

      {err && (
        <div style={{ background: clay.pillDangerBg, border: '1px solid var(--vdms-line)', color: clay.pillDangerText, fontSize: 12, borderRadius: 6, padding: '8px 12px', marginBottom: 12 }}>
          {err}
        </div>
      )}

      {loading && entries.length === 0 && !err ? (
        <div style={{ color: 'var(--vdms-text-faint)', fontSize: 13, padding: '24px 0', textAlign: 'center' }}>Loading audit log…</div>
      ) : entries.length === 0 && !err ? (
        <div style={{ color: 'var(--vdms-text-faint)', fontSize: 13, padding: '24px 0', textAlign: 'center' }}>
          No audit log entries yet. Sign-in / sign-out activity will appear here as it happens.
        </div>
      ) : (
        <div style={{ ...DMS_TABLE_CARD, overflowX: 'auto' }}>
          <table style={DMS_TABLE}>
            <thead>
              <tr>
                <th style={DMS_TH}>Event</th>
                <th style={DMS_TH}>Detail</th>
                <th style={DMS_TH}>Status</th>
                <th style={DMS_TH}>IP Address</th>
                <th style={DMS_TH}>Browser</th>
                <th style={DMS_TH}>Login Time</th>
                <th style={DMS_TH}>Logout Time</th>
                <th style={DMS_TH}>Duration</th>
              </tr>
            </thead>
            <tbody>
              {entries.map((e, idx) => (
                <tr key={e.session_id ? `${e.session_id}-${idx}` : idx} style={DMS_TR}>
                  <td style={DMS_TD_NAME}>{e.event || '—'}</td>
                  <td style={DMS_TD}>{e.detail || '—'}</td>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap' }}>
                    {badge(e.status === 'success' || e.status === 'active' ? 'green' : e.status === 'failed' ? 'red' : 'default', e.status || '—')}
                  </td>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap' }}>{e.ip_address || '—'}</td>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis' }} title={e.browser || ''}>{e.browser || '—'}</td>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap' }}>{formatTime(e.login_time)}</td>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap' }}>{formatTime(e.logout_time)}</td>
                  <td style={{ ...DMS_TD, whiteSpace: 'nowrap' }}>{e.active_duration_formatted || '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
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
      <div style={{ background: clay.accentSoft, border: '1px solid var(--vdms-line)', borderRadius: clay.radiusCard, padding: 14 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: clay.accentDark, marginBottom: 4 }}>
          Multi-Site Vessel Provisioning & DMS Folder Structure
        </div>
        <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', lineHeight: 1.5 }}>
          Control which connected SharePoint sites are available for vessel DMS folder structures. Adding sites to an existing vessel creates standard folders on the newly selected site only, preserving existing sites intact.
        </div>
        <div style={{ marginTop: 8, fontSize: 11, color: 'var(--vdms-text-muted)' }}>
          <strong>Note:</strong> File uploads reside inside the targeted site document library and do not replicate files across sites.
        </div>
        <div style={{ marginTop: 10, display: 'flex', justifyContent: 'flex-start' }}>
        </div>
      </div>

      {siteMsg && <div style={{ background: clay.pillActiveBg, border: '1px solid var(--vdms-line)', color: clay.pillActiveText, padding: '10px 14px', borderRadius: 8, fontSize: 13 }}>{siteMsg}</div>}
      {siteErr && <div style={{ background: clay.pillDangerBg, border: '1px solid var(--vdms-line)', color: clay.pillDangerText, padding: '10px 14px', borderRadius: 8, fontSize: 13 }}>{siteErr}</div>}

      {/* Section 1: Connected Sites */}
      <div style={{ background: 'var(--vdms-surface)', border: '1px solid var(--vdms-line)', borderRadius: clay.radiusCard, padding: 16, boxShadow: clay.shadowRaised }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <div>
            <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: 'var(--vdms-text)' }}>1. Connected Sites & Availability</h4>
            <span style={{ fontSize: 12, color: 'var(--vdms-text-muted)' }}>Configure which sites can receive vessel folder structures.</span>
          </div>
          <button
            onClick={handleSaveSiteSettings}
            disabled={saving || loading}
            style={dmsBtn('primary', !(saving || loading))}
          >
            {saving ? 'Saving...' : 'Save Site Settings'}
          </button>
        </div>

        {loading ? (
          <div style={{ padding: 20, textAlign: 'center', color: 'var(--vdms-text-muted)', fontSize: 13 }}>Loading sites...</div>
        ) : sites.length === 0 ? (
          <div style={{ padding: 20, textAlign: 'center', color: 'var(--vdms-text-muted)', fontSize: 13 }}>No connected sites discovered.</div>
        ) : (
          <div style={{ ...DMS_TABLE_CARD, overflowX: 'auto' }}>
            <table style={DMS_TABLE}>
              <thead>
                <tr>
                  <th style={DMS_TH}>Site Display Name</th>
                  <th style={DMS_TH}>Internal Key</th>
                  <th style={{ ...DMS_TH, textAlign: 'center' }}>Available for Provisioning</th>
                  <th style={{ ...DMS_TH, textAlign: 'center' }}>Default for New Vessels</th>
                </tr>
              </thead>
              <tbody>
                {sites.filter(s => !s.is_hidden).map(s => (
                  <tr key={s.site_key} style={DMS_TR}>
                    <td style={DMS_TD_NAME}>
                      {s.display_name}
                    </td>
                    <td style={DMS_TD}>
                      <code style={{ background: 'var(--vdms-surface-alt)', padding: '2px 6px', borderRadius: 4, fontSize: 11 }}>{s.site_key}</code>
                    </td>
                    <td style={{ ...DMS_TD, textAlign: 'center' }}>
                      <input
                        type="checkbox"
                        checked={s.is_available_for_provisioning}
                        onChange={e => {
                          const checked = e.target.checked;
                          setSites(sites.map(item => item.site_key === s.site_key ? { ...item, is_available_for_provisioning: checked } : item));
                        }}
                        style={{ width: 16, height: 16, cursor: 'pointer', accentColor: clay.accent }}
                      />
                    </td>
                    <td style={{ ...DMS_TD, textAlign: 'center' }}>
                      <input
                        type="checkbox"
                        disabled={!s.is_available_for_provisioning}
                        checked={s.is_default_provisioning}
                        onChange={e => {
                          const checked = e.target.checked;
                          setSites(sites.map(item => item.site_key === s.site_key ? { ...item, is_default_provisioning: checked } : item));
                        }}
                        style={{ width: 16, height: 16, cursor: s.is_available_for_provisioning ? 'pointer' : 'default', accentColor: clay.accent }}
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
      <div style={{ background: 'var(--vdms-surface)', border: '1px solid var(--vdms-line)', borderRadius: clay.radiusCard, padding: 16, boxShadow: clay.shadowRaised }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <div>
            <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: 'var(--vdms-text)' }}>2. Vessel Site Allocation & Status</h4>
            <span style={{ fontSize: 12, color: 'var(--vdms-text-muted)' }}>View provisioned sites per vessel and trigger provisioning for newly added sites.</span>
          </div>
          <input
            type="text"
            value={vesselSearch}
            onChange={e => setVesselSearch(e.target.value)}
            placeholder="Search vessels by name or IMO..."
            style={{ ...dmsControlStyle(), width: 220 }}
          />
        </div>

        <div style={{ maxHeight: 380, overflowY: 'auto', border: '1px solid var(--vdms-line)', borderRadius: 8 }}>
          <table style={DMS_TABLE}>
            <thead>
              <tr style={{ position: 'sticky', top: 0, zIndex: 1 }}>
                <th style={DMS_TH}>Vessel</th>
                <th style={DMS_TH}>IMO</th>
                <th style={DMS_TH}>Provisioned Sites</th>
                <th style={{ ...DMS_TH, textAlign: 'right' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredVessels.length === 0 ? (
                <tr>
                  <td colSpan={4} style={{ padding: 20, textAlign: 'center', color: 'var(--vdms-text-muted)' }}>No vessels match your search.</td>
                </tr>
              ) : (
                filteredVessels.map(v => {
                  const provSites = v.provisioned_site_ids || [];
                  return (
                    <tr key={v.id} style={DMS_TR}>
                      <td style={DMS_TD_NAME}>
                        <Icon iconName="Ferry" aria-hidden="true" style={{ fontSize: 13 }} /> {v.name}
                      </td>
                      <td style={DMS_TD}>
                        {v.imo || '—'}
                      </td>
                      <td style={DMS_TD}>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                          {provSites.length === 0 ? (
                            <span style={{ fontSize: 11, background: clay.pillWarnBg, color: clay.pillWarnText, padding: '2px 8px', borderRadius: 10 }}>
                              Default site only
                            </span>
                          ) : (
                            provSites.map(sk => (
                              <span
                                key={sk}
                                style={{
                                  fontSize: 11,
                                  fontWeight: 600,
                                  background: clay.pillActiveBg,
                                  color: clay.pillActiveText,
                                  padding: '2px 8px',
                                  borderRadius: 10,
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: 4,
                                }}
                              >
                                <span><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 11 }} /></span>
                                {siteNameMap[sk] || sk}
                              </span>
                            ))
                          )}
                        </div>
                      </td>
                      <td style={{ ...DMS_TD, textAlign: 'right' }}>
                        <button
                          onClick={() => handleOpenVesselModal(v)}
                          style={dmsRowBtn('plain')}
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
          <div style={{ background: 'var(--vdms-surface)', borderRadius: 16, padding: 24, width: 480, maxWidth: '95vw', boxShadow: clay.shadowRaised, border: '1px solid var(--vdms-line)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
              <div>
                <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: 'var(--vdms-text)' }}>
                  Manage Sites: {selectedVessel.name}
                </h3>
                <span style={{ fontSize: 12, color: 'var(--vdms-text-muted)' }}>IMO: {selectedVessel.imo || '—'}</span>
              </div>
              {!provisioning && (
                <button
                  onClick={() => setSelectedVessel(null)}
                  style={{ background: 'none', border: 'none', fontSize: 18, color: 'var(--vdms-text-faint)', cursor: 'pointer' }}
                ><Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 18 }} /></button>
              )}
            </div>

            {provMsg && <div style={{ background: clay.pillActiveBg, border: '1px solid var(--vdms-line)', color: clay.pillActiveText, padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{provMsg}</div>}
            {provErr && <div style={{ background: clay.pillDangerBg, border: '1px solid var(--vdms-line)', color: clay.pillDangerText, padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{provErr}</div>}

            <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginBottom: 12, lineHeight: 1.4 }}>
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
                      background: isAlreadyProvisioned ? 'var(--vdms-surface-alt)' : isChecked ? clay.accentSoft : 'var(--vdms-surface)',
                      border: `1px solid ${isChecked ? clay.accent : 'var(--vdms-line)'}`,
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
                        style={{ width: 16, height: 16, accentColor: clay.accent }}
                      />
                      <div>
                        <div style={{ fontSize: 13, fontWeight: 600, color: 'var(--vdms-text)' }}>{site.display_name}</div>
                        <div style={{ fontSize: 11, color: 'var(--vdms-text-muted)' }}>Key: {site.site_key}</div>
                      </div>
                    </div>

                    {isAlreadyProvisioned ? (
                      <span style={{ fontSize: 11, fontWeight: 700, color: clay.pillActiveText, background: clay.pillActiveBg, padding: '2px 8px', borderRadius: 10 }}>
                        <Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 11 }} /> Provisioned
                      </span>
                    ) : isChecked ? (
                      <span style={{ fontSize: 11, fontWeight: 600, color: clay.accentDark, background: clay.accentSoft, padding: '2px 8px', borderRadius: 10 }}>
<Icon iconName="Add" aria-hidden="true" style={{ fontSize: 11 }} /> Will Provision
                      </span>
                    ) : (
                      <span style={{ fontSize: 11, color: 'var(--vdms-text-faint)' }}>Not provisioned</span>
                    )}
                  </label>
                );
              })}
            </div>

            {/* Results breakdown if present */}
            {provResults && (
              <div style={{ background: 'var(--vdms-surface-alt)', border: '1px solid var(--vdms-line)', borderRadius: 8, padding: 10, marginBottom: 16, fontSize: 12 }}>
                <div style={{ fontWeight: 700, marginBottom: 6, color: 'var(--vdms-text)' }}>Provisioning Results:</div>
                {Object.entries(provResults).map(([sk, res]: [string, any]) => (
                  <div key={sk} style={{ display: 'flex', justifyContent: 'space-between', padding: '3px 0' }}>
                    <span>{siteNameMap[sk] || sk}:</span>
                    <span style={{ color: res.status === 'success' ? clay.pillActiveText : clay.pillDangerText, fontWeight: 600 }}>
                      {res.status === 'success' ? <><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 12 }} /> Provisioned</> : <><Icon iconName="ErrorBadge" aria-hidden="true" style={{ fontSize: 12 }} /> Failed: {res.error || 'Error'}</>}
                    </span>
                  </div>
                ))}
              </div>
            )}

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
              <button
                onClick={() => setSelectedVessel(null)}
                disabled={provisioning}
                style={dmsBtn('secondary', !provisioning)}
              >
                Close
              </button>
              <button
                onClick={handleProvisionSites}
                disabled={provisioning || targetSiteKeys.filter(k => !(selectedVessel.provisioned_site_ids || []).includes(k)).length === 0}
                style={dmsBtn('primary', !(provisioning || targetSiteKeys.filter(k => !(selectedVessel.provisioned_site_ids || []).includes(k)).length === 0))}
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

