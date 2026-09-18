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

export function renderSettingsPage(host: VesselEmail): React.ReactElement {
  return <SettingsPageView host={host} />;
}

function SettingsPageView({ host }: { host: VesselEmail }): React.ReactElement {
    const { settingsTab, settingsForm, settingsSavedMsg } = host.state;
  const [aiSuggestionMsg, setAiSuggestionMsg] = React.useState<string | null>(null);

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Settings</h2>
          <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Configure application settings and preferences.</p>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '200px 1fr', gap: 24, background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 24 }}>
          {/* Settings Left Nav */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4, borderRight: '1px solid #f1f5f9', paddingRight: 16 }}>
            {[
              'General', 'Site Selection', 'Vessel Site Provisioning', 'Document Settings', 'Notification Settings',
              'Permission Settings', 'Integration', 'Audit Logs'
            ].map(tab => (
              <button
                key={tab}
                onClick={() => host.setState({ settingsTab: tab as any })}
                style={{
                  border: 'none', background: settingsTab === tab ? '#eff6ff' : 'transparent',
                  color: settingsTab === tab ? '#0078d4' : '#475569', fontWeight: settingsTab === tab ? 700 : 500,
                  fontSize: 13, padding: '8px 12px', borderRadius: 6, textAlign: 'left', cursor: 'pointer',
                }}
              >
                {tab}
              </button>
            ))}
          </div>

          {/* Settings Content Area */}
          <div style={{ maxWidth: settingsTab === 'Vessel Site Provisioning' ? 950 : 500, width: '100%' }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' }}>{settingsTab}</h3>

            {settingsSavedMsg && (
              <div style={{ background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 13, marginBottom: 16 }}>
                ✓ Settings saved successfully.
              </div>
            )}

            {settingsTab === 'General' && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Site Title</label>
                  <input
                    type="text"
                    value={settingsForm.siteTitle}
                    onChange={e => host.setState({ settingsForm: { ...settingsForm, siteTitle: e.target.value } })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box' }}
                  />
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Site Description</label>
                  <textarea
                    value={settingsForm.siteDescription}
                    onChange={e => host.setState({ settingsForm: { ...settingsForm, siteDescription: e.target.value } })}
                    style={{ width: '100%', height: 60, padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box', resize: 'vertical' }}
                  />
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Date Format</label>
                  <select
                    value={settingsForm.dateFormat}
                    onChange={e => host.setState({ settingsForm: { ...settingsForm, dateFormat: e.target.value } })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', background: '#fff' }}
                  >
                    <option value="MM/DD/YYYY">MM/DD/YYYY</option>
                    <option value="DD/MM/YYYY">DD/MM/YYYY</option>
                    <option value="YYYY-MM-DD">YYYY-MM-DD</option>
                  </select>
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Time Zone</label>
                  <select
                    value={settingsForm.timeZone}
                    onChange={e => host.setState({ settingsForm: { ...settingsForm, timeZone: e.target.value } })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', background: '#fff' }}
                  >
                    <option value="(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi">(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi</option>
                    <option value="(UTC+00:00) UTC">(UTC+00:00) UTC</option>
                  </select>
                </div>

                <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 10, padding: 14 }}>
                  <div style={{ fontSize: 13, fontWeight: 700, color: '#0f172a', marginBottom: 8 }}>AI Suggestion Controls</div>
                  <div style={{ fontSize: 12, color: '#475569', lineHeight: 1.5, marginBottom: 10 }}>
                    Manage vessel suggestion memory from scanned files. Ignored false positives are stored in browser preferences.
                  </div>
                  <div style={{ background: '#ffffff', border: '1px solid #e2e8f0', borderRadius: 8, padding: '10px 12px', marginBottom: 10 }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
                      <div>
                        <div style={{ fontSize: 12, fontWeight: 700, color: '#0f172a' }}>Post-Upload OCR Navigation Prompt</div>
                        <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>
                          Ask whether to open Templates &amp; OCR when vessel name is not identified.
                        </div>
                      </div>
                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: 12, fontWeight: 600, color: '#334155' }}>
                        <input
                          type="checkbox"
                          checked={!host.state.suppressUploadNavigationPrompt}
                          onChange={(e) => host._setSuppressUploadNavigationPrompt(!e.target.checked)}
                        />
                        {host.state.suppressUploadNavigationPrompt ? 'Disabled' : 'Enabled'}
                      </label>
                    </div>
                  </div>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, marginBottom: 10 }}>
                    <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: '8px 10px' }}>
                      <div style={{ fontSize: 10, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>Pending Suggestions</div>
                      <div style={{ fontSize: 18, fontWeight: 800, color: '#0369a1' }}>{host.state.pendingVesselSuggestions.length}</div>
                    </div>
                    <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: '8px 10px' }}>
                      <div style={{ fontSize: 10, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>Ignored Suggestions</div>
                      <div style={{ fontSize: 18, fontWeight: 800, color: '#b45309' }}>{host.state.ignoredVesselSuggestionKeys.size}</div>
                    </div>
                  </div>
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                    <button
                      onClick={() => {
                        host._clearIgnoredVesselSuggestionKeys();
                        setAiSuggestionMsg('Ignored suggestions cleared.');
                        setTimeout(() => setAiSuggestionMsg(null), 2200);
                      }}
                      disabled={host.state.ignoredVesselSuggestionKeys.size === 0}
                      style={{
                        background: host.state.ignoredVesselSuggestionKeys.size > 0 ? '#fff7ed' : '#f1f5f9',
                        color: host.state.ignoredVesselSuggestionKeys.size > 0 ? '#c2410c' : '#94a3b8',
                        border: `1px solid ${host.state.ignoredVesselSuggestionKeys.size > 0 ? '#fdba74' : '#e2e8f0'}`,
                        borderRadius: 8,
                        minHeight: 36,
                        padding: '8px 12px',
                        fontSize: 12,
                        fontWeight: 700,
                        cursor: host.state.ignoredVesselSuggestionKeys.size > 0 ? 'pointer' : 'not-allowed',
                      }}
                    >
                      Clear Ignored Suggestions
                    </button>
                    <button
                      onClick={() => host._openSuggestedVesselsFromSidebar()}
                      style={{
                        background: '#eff6ff',
                        color: '#1d4ed8',
                        border: '1px solid #bfdbfe',
                        borderRadius: 8,
                        minHeight: 36,
                        padding: '8px 12px',
                        fontSize: 12,
                        fontWeight: 700,
                        cursor: 'pointer',
                      }}
                    >
                      Open Suggested Vessels
                    </button>
                  </div>
                  {aiSuggestionMsg && (
                    <div style={{ marginTop: 10, background: '#dcfce7', border: '1px solid #86efac', color: '#166534', borderRadius: 8, padding: '8px 10px', fontSize: 12, fontWeight: 600 }}>
                      ✓ {aiSuggestionMsg}
                    </div>
                  )}
                </div>

                <div style={{ marginTop: 12 }}>
                  <button
                    onClick={() => {
                      host.setState({ settingsSavedMsg: true });
                      setTimeout(() => host.setState({ settingsSavedMsg: false }), 2000);
                    }}
                    style={{
                      background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
                      padding: '8px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                    }}
                  >
                    Save Changes
                  </button>
                </div>
              </div>
            )}

            {settingsTab === 'Site Selection' && (
              <SiteSelectorComponent host={host} />
            )}

            {settingsTab === 'Vessel Site Provisioning' && (
              <VesselSiteProvisioningPanel host={host} />
            )}

            {settingsTab === 'Integration' && (
              <SiteIntegrationInfo host={host} />
            )}

            {settingsTab !== 'General' && settingsTab !== 'Integration' && settingsTab !== 'Site Selection' && settingsTab !== 'Vessel Site Provisioning' && (
              <div style={{ color: '#64748b', fontSize: 13 }}>
                Configuration settings for {settingsTab} are active with system default policies.
              </div>
            )}
          </div>
        </div>
      </div>
    );
}

function SiteIntegrationInfo({ host }: { host: VesselEmail }): React.ReactElement {
  const [info, setInfo] = React.useState<any>(null);
  const [adminConfig, setAdminConfig] = React.useState<any>(null);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [err, setErr] = React.useState<string | null>(null);
  const [isAdmin, setIsAdmin] = React.useState<boolean>(false);
  const [showSwitchDialog, setShowSwitchDialog] = React.useState<boolean>(false);
  const [selectedNewSite, setSelectedNewSite] = React.useState<string>('');
  // Site visibility toggle state
  const [showHiddenSites, setShowHiddenSites] = React.useState<boolean>(false);
  const [togglingSiteKey, setTogglingSiteKey] = React.useState<string | null>(null);
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
  const [switchReason, setSwitchReason] = React.useState<string>('');
  const [switching, setSwitching] = React.useState<boolean>(false);
  const [switchMsg, setSwitchMsg] = React.useState<string | null>(null);
  const [switchSuccess, setSwitchSuccess] = React.useState<boolean>(false);
  const [siteChanges, setSiteChanges] = React.useState<any[]>([]);
  const [showChanges, setShowChanges] = React.useState<boolean>(false);

  // Tenant-site discovery
  const [tenantSites, setTenantSites] = React.useState<any[]>([]);
  const [tenantSitesLoading, setTenantSitesLoading] = React.useState<boolean>(false);
  const [tenantSitesError, setTenantSitesError] = React.useState<string | null>(null);
  const [siteSearch, setSiteSearch] = React.useState<string>('');

  // Drive picker for unregistered tenant sites
  const [pendingSite, setPendingSite] = React.useState<any>(null); // the tenant site object
  const [pendingDrives, setPendingDrives] = React.useState<any[]>([]);
  const [pendingDrivesLoading, setPendingDrivesLoading] = React.useState<boolean>(false);
  const [selectedDrive, setSelectedDrive] = React.useState<any>(null);

  const loadTenantSites = React.useCallback(async (base: string, headers: any) => {
    setTenantSitesLoading(true);
    setTenantSitesError(null);
    try {
      const r = await fetch(`${base}/api/admin/discover-sites`, { headers });
      const d = await r.json();
      if (!r.ok) throw new Error(d.detail || 'Unable to load tenant sites');
      setTenantSites(d.sites || []);
    } catch (e: any) {
      setTenantSitesError(e?.message || 'Could not load tenant sites');
    } finally {
      setTenantSitesLoading(false);
    }
  }, []);

  const loadAdminConfig = React.useCallback(async (includeHidden: boolean = false) => {
    const base = host._base();
    try {
      const r = await fetch(`${base}/api/admin/site-configuration?include_hidden=${includeHidden}`, { headers: host._headers() });
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
      await loadAdminConfig(showHiddenSites);
    } catch (e: any) {
      setVisibilityMsg(`⚠️ ${e?.message || 'Failed to update visibility'}`);
      setTimeout(() => setVisibilityMsg(null), 4000);
    } finally {
      setTogglingSiteKey(null);
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
      fetch(`${base}/api/admin/site-configuration?include_hidden=false`, { headers: host._headers() })
        .then(r => {
          if (r.status === 403) { setIsAdmin(false); setLoading(false); return null; }
          return r.json();
        })
        .then(d => {
          if (d) {
            setAdminConfig(d);
            setIsAdmin(true);
            setSelectedNewSite(d.current_site);

            fetch(`${base}/api/admin/site-changes?limit=10`, { headers: host._headers() })
              .then(r => r.json())
              .then(ch => setSiteChanges(ch.changes || []))
              .catch(() => undefined);

            // Auto-load all tenant sites for the dropdown
            void loadTenantSites(base, host._headers());
          }
          setLoading(false);
        })
        .catch(() => { setIsAdmin(false); setLoading(false); });
    } else {
      setLoading(false);
    }
  }, [host, host.props.userEmail, loadTenantSites]);

  // When a site key is selected from the dropdown
  const handleSiteSelected = async (siteKey: string) => {
    setSelectedNewSite(siteKey);
    setPendingSite(null);
    setPendingDrives([]);
    setSelectedDrive(null);

    // Check if this is an unregistered tenant site (site_key starts with "tenant:")
    if (siteKey.startsWith('tenant:')) {
      const tenantId = siteKey.replace('tenant:', '');
      const tenantSiteObj = tenantSites.find(s => s.id === tenantId);
      if (!tenantSiteObj) return;
      setPendingSite(tenantSiteObj);
      setPendingDrivesLoading(true);
      try {
        const r = await fetch(`${host._base()}/api/admin/discover-sites/${encodeURIComponent(tenantSiteObj.id)}/drives`, { headers: host._headers() });
        const d = await r.json();
        setPendingDrives(r.ok ? (d.drives || []) : []);
        if (r.ok && (d.drives || []).length === 1) setSelectedDrive(d.drives[0]); // auto-pick if only one
      } catch { setPendingDrives([]); }
      finally { setPendingDrivesLoading(false); }
    }
  };

  const handleSwitchSite = async () => {
    if (!selectedNewSite || selectedNewSite === adminConfig?.current_site) {
      setSwitchMsg('Please select a different site');
      setTimeout(() => setSwitchMsg(null), 2000);
      return;
    }
    setSwitching(true);
    setSwitchMsg(null);
    setSwitchSuccess(false);

    try {
      const base = host._base();
      let response: Response;
      let data: any;

      if (selectedNewSite.startsWith('tenant:') && pendingSite && selectedDrive) {
        // Auto-register + switch for an undiscovered tenant site
        response = await fetch(`${base}/api/admin/switch-site-auto`, {
          method: 'POST',
          headers: { ...host._headers(), 'Content-Type': 'application/json' },
          body: JSON.stringify({
            site_id: pendingSite.id,
            site_name: pendingSite.name,
            drive_id: selectedDrive.id,
            display_name: pendingSite.name,
            web_url: pendingSite.web_url || '',
            reason: switchReason || undefined,
          }),
        });
        data = await response.json();
      } else {
        // Normal configured-site switch
        response = await fetch(`${base}/api/admin/switch-site`, {
          method: 'POST',
          headers: { ...host._headers(), 'Content-Type': 'application/json' },
          body: JSON.stringify({ site_name: selectedNewSite, reason: switchReason || undefined }),
        });
        data = await response.json();
      }

      if (!response.ok) throw new Error(data.detail || 'Failed to switch site');

      setSwitchSuccess(true);
      setSwitchMsg(`✅ Successfully switched to ${data.new_site_name}.`);
      setInfo((prev: any) => ({ ...(prev || {}), active_site: data.new_site, site_name: data.new_site_name }));
      setAdminConfig((prev: any) => prev ? { ...prev, current_site: data.new_site, current_site_name: data.new_site_name } : prev);
      setShowSwitchDialog(false);
      setSwitchReason('');
      setPendingSite(null);
      setPendingDrives([]);
      setSelectedDrive(null);
      setSwitching(false);

      void Promise.all([host._loadDocumentSites(), host._loadData()]).catch(() => {
        setSwitchMsg(`✅ Switched to ${data.new_site_name}, but view refresh failed — please reload.`);
      });
    } catch (e: any) {
      setSwitchMsg(`⚠️ ${e?.message || 'Failed to switch site'}`);
      setSwitching(false);
    }
  };

  const copyResolvedValue = (value: string) => { if (value) navigator.clipboard?.writeText(value); };

  // Build dropdown options: configured sites + unregistered tenant sites
  const configuredSiteKeys = new Set((adminConfig?.available_sites || []).map((s: any) => s.name));
  const filteredTenantSites = tenantSites.filter(ts => {
    const q = siteSearch.toLowerCase();
    return !q || `${ts.name} ${ts.web_url}`.toLowerCase().includes(q);
  });
  const filteredConfiguredSites = (adminConfig?.available_sites || []).filter((s: any) => {
    const q = siteSearch.toLowerCase();
    return !q || `${s.display_name} ${s.name}`.toLowerCase().includes(q);
  });

  // Resolve toSiteConfig for the confirm dialog
  const toSiteConfig = selectedNewSite.startsWith('tenant:') && selectedDrive
    ? { display_name: pendingSite?.name, drive_id: selectedDrive?.id }
    : adminConfig?.available_sites?.find((s: any) => s.name === selectedNewSite);

  const canSwitch = selectedNewSite
    && selectedNewSite !== adminConfig?.current_site
    && !switching
    && (!selectedNewSite.startsWith('tenant:') || (pendingSite && selectedDrive));

  if (loading) return <div style={{ color: '#64748b', fontSize: 13 }}>Loading site configuration...</div>;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
      <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16 }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: '#0f172a', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
          <span>🌐</span> Active Site: {info?.site_name || info?.active_site}
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '140px 1fr', rowGap: 8, fontSize: 12 }}>
          <span style={{ fontWeight: 600, color: '#64748b' }}>Site Key:</span>
          <span style={{ fontFamily: 'monospace', color: '#0f172a', fontWeight: 700 }}>{info?.active_site}</span>

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
        <>
          <div style={{ background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 8, padding: 16 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: '#166534', marginBottom: 6, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <span>👤 Admin: Switch Active Site</span>
              <button
                onClick={() => loadTenantSites(host._base(), host._headers())}
                disabled={tenantSitesLoading}
                title="Refresh all tenant sites"
                style={{ padding: '4px 10px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: 5, fontSize: 11, fontWeight: 600, cursor: tenantSitesLoading ? 'wait' : 'pointer' }}
              >
                {tenantSitesLoading ? '⏳ Loading…' : '🔄 Refresh Sites'}
              </button>
            </div>
            <p style={{ margin: '0 0 10px', fontSize: 12, color: '#166534', lineHeight: 1.5 }}>
              Select any configured site or any SharePoint site visible to this app from the dropdown below.
            </p>

            {/* Search filter */}
            <input
              value={siteSearch}
              onChange={e => setSiteSearch(e.target.value)}
              placeholder="🔍 Search site by name or URL…"
              style={{ width: '100%', boxSizing: 'border-box', padding: '7px 10px', border: '1px solid #bbf7d0', borderRadius: 6, fontSize: 12, marginBottom: 8, outline: 'none' }}
            />

            {visibilityMsg && (
              <div style={{ background: '#fee2e2', border: '1px solid #fca5a5', color: '#991b1b', borderRadius: 6, padding: '6px 10px', fontSize: 12, marginBottom: 8 }}>
                {visibilityMsg}
              </div>
            )}

            {/* ── Configured Sites card list ── */}
            <div style={{ marginBottom: 12 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
                <label style={{ fontSize: 12, fontWeight: 600, color: '#166534' }}>Configured Sites:</label>
                <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: '#166534', cursor: 'pointer', userSelect: 'none' }}>
                  <input
                    type="checkbox"
                    checked={showHiddenSites}
                    onChange={e => {
                      const checked = e.target.checked;
                      setShowHiddenSites(checked);
                      void loadAdminConfig(checked);
                    }}
                    style={{ accentColor: '#16a34a', cursor: 'pointer' }}
                  />
                  <span>Show hidden sites</span>
                </label>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {filteredConfiguredSites.map((site: any) => {
                  const isCurrent = site.name === adminConfig.current_site;
                  const isHidden = !!site.is_hidden;
                  const isToggling = togglingSiteKey === site.name;
                  return (
                    <div
                      key={`cfg-${site.name}`}
                      style={{
                        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                        padding: '8px 12px', borderRadius: 8,
                        border: `1px solid ${isCurrent ? '#86efac' : isHidden ? '#cbd5e1' : '#d1fae5'}`,
                        background: isCurrent ? '#dcfce7' : isHidden ? '#f8fafc' : '#f0fdf4',
                        opacity: isHidden && !isCurrent ? 0.75 : 1,
                        transition: 'all 0.15s',
                      }}
                    >
                      <div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                          <span style={{ fontSize: 13, fontWeight: 600, color: isHidden ? '#64748b' : '#166534' }}>
                            {site.display_name || site.name}
                          </span>
                          {isCurrent && (
                            <span style={{ fontSize: 10, fontWeight: 700, background: '#16a34a', color: '#fff', padding: '1px 6px', borderRadius: 8 }}>
                              ACTIVE
                            </span>
                          )}
                          {isHidden && (
                            <span style={{ fontSize: 10, fontWeight: 700, background: '#e2e8f0', color: '#475569', padding: '1px 6px', borderRadius: 8 }}>
                              HIDDEN
                            </span>
                          )}
                        </div>
                        <span style={{ display: 'block', fontSize: 10, color: '#64748b', fontFamily: 'monospace', marginTop: 2 }}>
                          {site.name}
                        </span>
                      </div>
                      <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                        {/* Switch button */}
                        {!isCurrent && !isHidden && (
                          <button
                            onClick={() => { setSelectedNewSite(site.name); setPendingSite(null); setSelectedDrive(null); setShowSwitchDialog(true); }}
                            disabled={switching || isToggling}
                            style={{ padding: '4px 10px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: 5, fontSize: 11, fontWeight: 600, cursor: 'pointer' }}
                          >
                            🔄 Switch
                          </button>
                        )}
                        {/* Hide / Show Toggle button */}
                        <button
                          onClick={() => handleToggleVisibility(site.name, isHidden)}
                          disabled={isCurrent || isToggling}
                          title={isCurrent ? 'Cannot hide the currently active site. Switch to a different site first.' : isHidden ? 'Make this site visible in pickers' : 'Hide this site from pickers'}
                          style={{
                            padding: '4px 10px',
                            background: isCurrent ? '#f1f5f9' : isHidden ? '#eff6ff' : '#fff',
                            color: isCurrent ? '#94a3b8' : isHidden ? '#1d4ed8' : '#475569',
                            border: `1px solid ${isCurrent ? '#e2e8f0' : isHidden ? '#bfdbfe' : '#cbd5e1'}`,
                            borderRadius: 5,
                            fontSize: 11,
                            fontWeight: 600,
                            cursor: isCurrent ? 'not-allowed' : isToggling ? 'wait' : 'pointer',
                          }}
                        >
                          {isToggling ? '⏳…' : isHidden ? '👁 Show' : '🙈 Hide'}
                        </button>
                      </div>
                    </div>
                  );
                })}
                {filteredConfiguredSites.length === 0 && (
                  <div style={{ fontSize: 12, color: '#64748b', padding: '8px 0' }}>
                    {showHiddenSites ? 'No configured sites match your search.' : 'No configured sites match your search. (Try checking "Show hidden sites" above)'}
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
                        // Refresh the configured-sites list by appending the new entry
                        setAdminConfig((prev: any) => prev ? {
                          ...prev,
                          available_sites: [
                            ...(prev.available_sites || []),
                            { name: addSiteKey.trim(), display_name: addSiteDisplayName.trim(), configured: true, site_id: addSiteSelected.id, drive_id: addSiteSelDrive.id },
                          ],
                        } : prev);
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

            {/* Drive picker for unregistered tenant sites */}
            {pendingSite && (
              <div style={{ marginBottom: 12, padding: 10, background: '#fff', border: '1px solid #86efac', borderRadius: 6, fontSize: 11 }}>
                <div style={{ fontWeight: 700, color: '#166534', marginBottom: 6 }}>
                  📂 Select Document Library for: <em>{pendingSite.name}</em>
                </div>
                {pendingDrivesLoading && <div style={{ color: '#64748b' }}>Loading document libraries…</div>}
                {!pendingDrivesLoading && pendingDrives.length === 0 && <div style={{ color: '#b45309' }}>No document libraries found on this site.</div>}
                {!pendingDrivesLoading && pendingDrives.map(drive => (
                  <button
                    key={drive.id}
                    onClick={() => setSelectedDrive(drive)}
                    style={{
                      display: 'block', width: '100%', textAlign: 'left', padding: '7px 10px', marginBottom: 4,
                      border: `1px solid ${selectedDrive?.id === drive.id ? '#16a34a' : '#dcfce7'}`,
                      borderRadius: 5, background: selectedDrive?.id === drive.id ? '#dcfce7' : '#fff',
                      cursor: 'pointer', fontSize: 11, color: '#166534',
                    }}
                  >
                    <strong>{drive.name}</strong>
                    <span style={{ display: 'block', fontSize: 10, color: '#64748b' }}>{drive.drive_type || 'documentLibrary'}</span>
                    <span style={{ display: 'block', fontSize: 9, color: '#94a3b8', fontFamily: 'monospace' }}>{drive.id.substring(0, 30)}…</span>
                  </button>
                ))}
                {selectedDrive && (
                  <div style={{ marginTop: 6, fontSize: 10, color: '#15803d', fontWeight: 600 }}>
                    ✅ Selected: {selectedDrive.name} — will be registered and switched to automatically.
                  </div>
                )}
              </div>
            )}

            <button
              onClick={() => setShowSwitchDialog(true)}
              disabled={!canSwitch}
              style={{
                padding: '10px 16px',
                background: canSwitch ? '#16a34a' : '#cbd5e1',
                color: '#fff', border: 'none', borderRadius: 6,
                fontSize: 12, fontWeight: 600,
                cursor: canSwitch ? 'pointer' : 'not-allowed',
              }}
            >
              {switching ? '⏳ Switching…' : 'Switch Site'}
            </button>

            {switchMsg && (
              <div style={{
                marginTop: 12,
                background: switchSuccess ? '#dcfce7' : '#fee2e2',
                border: `1px solid ${switchSuccess ? '#86efac' : '#fca5a5'}`,
                color: switchSuccess ? '#166534' : '#991b1b',
                borderRadius: 6, padding: '8px 12px', fontSize: 12, fontWeight: 600,
              }}>
                {switchMsg}
              </div>
            )}

            {showChanges && siteChanges.length > 0 && (
              <div style={{ marginTop: 16, borderTop: '1px solid #bbf7d0', paddingTop: 12 }}>
                <div style={{ fontSize: 12, fontWeight: 600, color: '#166534', marginBottom: 8 }}>Recent Changes:</div>
                <div style={{ maxHeight: 200, overflowY: 'auto', fontSize: 11, color: '#166534' }}>
                  {siteChanges.map((change: any) => (
                    <div key={change.id} style={{ padding: '6px 0', borderBottom: '1px solid #dcfce7' }}>
                      <div><strong>{change.previous_site} → {change.new_site}</strong> ({change.status})</div>
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
                style={{ marginTop: 8, background: 'transparent', border: 'none', color: '#16a34a', fontSize: 11, fontWeight: 600, cursor: 'pointer', textDecoration: 'underline' }}
              >
                {showChanges ? '▼ Hide' : '▶ Show'} Recent Changes
              </button>
            )}
          </div>

          {showSwitchDialog && (
            <AdminSwitchConfirmationDialog
              fromSite={adminConfig.current_site}
              fromSiteName={adminConfig.current_site_name}
              fromDb={adminConfig.current_db_name}
              fromDriveId={adminConfig.current_drive_id}
              toSite={pendingSite?.name || selectedNewSite}
              toSiteConfig={toSiteConfig}
              onConfirm={() => handleSwitchSite()}
              onCancel={() => { setShowSwitchDialog(false); setSwitchReason(''); }}
              reason={switchReason}
              onReasonChange={setSwitchReason}
              isProcessing={switching}
            />
          )}
        </>
      )}

      {err && (
        <div style={{ background: '#fee2e2', border: '1px solid #fca5a5', borderRadius: 6, padding: 12, fontSize: 12, color: '#991b1b' }}>
          ⚠️ {err}
        </div>
      )}

      <div style={{ fontSize: 11, color: '#64748b', lineHeight: 1.6, background: '#f8fafc', borderRadius: 6, padding: 12, border: '1px solid #e2e8f0' }}>
        <div style={{ fontWeight: 600, marginBottom: 6 }}>📖 How to switch sites:</div>
        {isAdmin ? (
          <>
            <p style={{ margin: '0 0 6px' }}>
              <strong>For admins:</strong> Use the "Switch Active Site" list above to change the site for this session. No server restart is needed.
            </p>
            <p style={{ margin: '0 0 6px' }}>
              Sites under <em>✅ Configured Sites</em> are already registered. Sites under <em>🌐 All Tenant Sites</em> will be auto-registered on first switch.
            </p>
          </>
        ) : (
          <p style={{ margin: 0 }}>Contact an administrator to switch to a different SharePoint site.</p>
        )}
      </div>
    </div>
  );
}

function AdminSwitchConfirmationDialog({
  fromSite, fromSiteName, fromDb, fromDriveId,
  toSite, toSiteConfig,
  onConfirm, onCancel,
  reason, onReasonChange,
  isProcessing
}: {
  fromSite: string;
  fromSiteName: string;
  fromDb: string;
  fromDriveId: string;
  toSite: string;
  toSiteConfig: any;
  onConfirm: () => void;
  onCancel: () => void;
  reason: string;
  onReasonChange: (r: string) => void;
  isProcessing: boolean;
}): React.ReactElement {
  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex',
      alignItems: 'center', justifyContent: 'center', zIndex: 9999,
    }} onClick={onCancel}>
      <div style={{
        background: '#fff', borderRadius: 12, padding: 24, maxWidth: 500,
        boxShadow: '0 20px 25px rgba(0,0,0,0.15)', maxHeight: '90vh', overflowY: 'auto',
      }} onClick={e => e.stopPropagation()}>
        <div style={{ fontSize: 16, fontWeight: 700, color: '#991b1b', marginBottom: 16, display: 'flex', alignItems: 'center', gap: 8 }}>
          <span>⚠️</span> Confirm Site Switch
        </div>

        <div style={{ background: '#fef2f2', border: '1px solid #fecaca', borderRadius: 8, padding: 12, marginBottom: 16, fontSize: 12, color: '#7f1d1d', lineHeight: 1.6 }}>
          <strong>Notice:</strong> Switching the active site changes this session's document and folder context. The page will refresh its data without restarting the application.
        </div>

        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: '#0f172a', marginBottom: 8 }}>Changes Summary:</div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 8 }}>
            <div style={{ background: '#f1f5f9', borderRadius: 6, padding: 10 }}>
              <div style={{ fontSize: 10, color: '#64748b', textTransform: 'uppercase', fontWeight: 600 }}>From</div>
              <div style={{ fontSize: 12, fontWeight: 700, color: '#0f172a', marginTop: 4 }}>{fromSite}</div>
              <div style={{ fontSize: 10, color: '#64748b', marginTop: 2 }}>{fromSiteName}</div>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <span style={{ fontSize: 20 }}>→</span>
            </div>
            <div style={{ background: '#f0fdf4', borderRadius: 6, padding: 10 }}>
              <div style={{ fontSize: 10, color: '#15803d', textTransform: 'uppercase', fontWeight: 600 }}>To</div>
              <div style={{ fontSize: 12, fontWeight: 700, color: '#15803d', marginTop: 4 }}>{toSite}</div>
              <div style={{ fontSize: 10, color: '#15803d', marginTop: 2 }}>{toSiteConfig?.display_name}</div>
            </div>
          </div>
        </div>

        <div style={{ marginBottom: 16, padding: 12, background: '#f8fafc', borderRadius: 8, fontSize: 11, lineHeight: 1.6 }}>
          <div style={{ fontWeight: 600, color: '#0f172a', marginBottom: 8 }}>Affected Configuration:</div>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <tbody>
              <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                <td style={{ padding: '6px 0', fontWeight: 600, color: '#475569' }}>Database:</td>
                <td style={{ padding: '6px 0 6px 12px', color: '#0f172a' }}>
                  <span style={{ color: '#64748b' }}>{fromDb}</span>
                  <span style={{ display: 'block', color: '#64748b', fontSize: 10 }}>(shared — same database for all sites)</span>
                </td>
              </tr>
              <tr>
                <td style={{ padding: '6px 0', fontWeight: 600, color: '#475569' }}>SharePoint Drive:</td>
                <td style={{ padding: '6px 0 6px 12px', color: '#0f172a', fontSize: 10, fontFamily: 'monospace' }}>
                  <span style={{ textDecoration: 'line-through', color: '#a1a1a1' }}>{fromDriveId.substring(0, 22)}…</span>
                  <span style={{ display: 'block', color: '#15803d', fontWeight: 600 }}>
                    → {toSiteConfig?.display_name || toSite}{toSiteConfig?.drive_id ? ` (${String(toSiteConfig.drive_id).substring(0, 22)}…)` : ''}
                  </span>
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        <div style={{ marginBottom: 16 }}>
          <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 6 }}>
            Reason for change (optional):
          </label>
          <textarea
            value={reason}
            onChange={e => onReasonChange(e.target.value)}
            placeholder="e.g., Testing new site, migration, maintenance..."
            style={{
              width: '100%', height: 60, padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1',
              fontSize: 12, outline: 'none', boxSizing: 'border-box', resize: 'vertical',
              fontFamily: 'inherit',
            }}
            disabled={isProcessing}
          />
          <div style={{ fontSize: 10, color: '#64748b', marginTop: 4 }}>
            This will be logged in the audit trail for traceability.
          </div>
        </div>

        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
          <button
            onClick={onCancel}
            disabled={isProcessing}
            style={{
              padding: '8px 16px', background: '#f1f5f9', color: '#334155', border: '1px solid #cbd5e1',
              borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: isProcessing ? 'not-allowed' : 'pointer',
              opacity: isProcessing ? 0.5 : 1,
            }}
          >
            Cancel
          </button>
          <button
            onClick={onConfirm}
            disabled={isProcessing}
            style={{
              padding: '8px 16px', background: '#dc2626', color: '#fff', border: 'none',
              borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: isProcessing ? 'not-allowed' : 'pointer',
              opacity: isProcessing ? 0.5 : 1,
            }}
          >
            {isProcessing ? '⏳ Switching...' : 'Confirm Switch'}
          </button>
        </div>
      </div>
    </div>
  );
}

function SiteSelectorComponent({ host }: { host: VesselEmail }): React.ReactElement {
  const [sites, setSites] = React.useState<any[]>([]);
  const [siteFilter, setSiteFilter] = React.useState('');
  const [showAll, setShowAll] = React.useState(false);
  const [selectedSite, setSelectedSite] = React.useState<any>(null);
  const [drives, setDrives] = React.useState<any[]>([]);
  const [selectedDrive, setSelectedDrive] = React.useState<any>(null);
  const [displayName, setDisplayName] = React.useState('');
  const [siteKey, setSiteKey] = React.useState('');
  const [loading, setLoading] = React.useState(true);
  const [loadingDrives, setLoadingDrives] = React.useState(false);
  const [saving, setSaving] = React.useState(false);
  const [isAdmin, setIsAdmin] = React.useState(false);
  const [message, setMessage] = React.useState<string | null>(null);
  const [error, setError] = React.useState<string | null>(null);

  const isSystemSite = (site: any) => /^(apps|all company|appcatalog|contenttypehub)$/i.test(site.name || '');
  const visibleSites = sites.filter(site => (showAll || !isSystemSite(site)) && `${site.name} ${site.web_url}`.toLowerCase().includes(siteFilter.toLowerCase()));

  React.useEffect(() => {
    const load = async () => {
      try {
        const adminResponse = await fetch(`${host._base()}/api/admin/site-configuration`, { headers: host._headers() });
        if (adminResponse.status === 403) throw new Error('Administrator access is required for Site Selection.');
        if (!adminResponse.ok) throw new Error('Unable to verify administrator access.');
        setIsAdmin(true);
        const response = await fetch(`${host._base()}/api/admin/discover-sites`, { headers: host._headers() });
        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || 'Unable to discover SharePoint sites. Admin consent may be required.');
        setSites(data.sites || []);
        if (!(data.sites || []).length) setMessage('No SharePoint sites were found for this app registration. Confirm Sites.Read.All admin consent and site access.');
      } catch (e: any) { setError(e.message || 'Unable to load SharePoint sites.'); }
      finally { setLoading(false); }
    };
    load();
  }, [host]);

  const selectSite = async (site: any) => {
    setSelectedSite(site); setSelectedDrive(null); setDrives([]); setLoadingDrives(true); setError(null); setMessage(null);
    try {
      const response = await fetch(`${host._base()}/api/admin/discover-sites/${encodeURIComponent(site.id)}/drives`, { headers: host._headers() });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Unable to load document libraries.');
      setDrives(data.drives || []);
      if (!(data.drives || []).length) setMessage('This site has no document libraries available to the app.');
    } catch (e: any) { setError(e.message || 'Unable to load document libraries.'); }
    finally { setLoadingDrives(false); }
  };

  const saveConfiguration = async () => {
    if (!selectedSite || !selectedDrive || !siteKey.trim() || !displayName.trim()) return;
    setSaving(true); setError(null); setMessage(null);
    try {
      const response = await fetch(`${host._base()}/api/admin/site-configurations`, {
        method: 'POST', headers: host._headers(),
        body: JSON.stringify({ site_key: siteKey.trim(), display_name: displayName.trim(), site_name: selectedSite.name, site_id: selectedSite.id, drive_id: selectedDrive.id }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Unable to save site configuration.');
      setMessage(`Saved ${data.site.display_name}. It is now available in Integration > Switch Active Site.`);
      setSiteKey(''); setDisplayName('');
    } catch (e: any) { setError(e.message || 'Unable to save site configuration.'); }
    finally { setSaving(false); }
  };

  if (loading) return <div style={{ color: '#64748b', fontSize: 13 }}>Discovering SharePoint sites...</div>;
  if (!isAdmin) return <div style={{ color: '#b91c1c', fontSize: 13 }}>Administrator access is required for Site Selection.</div>;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
      <div style={{ background: '#f0f9ff', border: '1px solid #bfdbfe', borderRadius: 8, padding: 12, fontSize: 12, color: '#1e40af', lineHeight: 1.5 }}>
        Sites are loaded from Microsoft Graph using the configured tenant app registration. Select a site, choose its document library, then save a reusable configuration.
      </div>
      {error && <div style={{ background: '#fef2f2', border: '1px solid #fecaca', borderRadius: 6, padding: 10, color: '#991b1b', fontSize: 12 }}>{error} {error.toLowerCase().includes('consent') && <a href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noreferrer">Open Entra app registrations</a>}</div>}
      {message && <div style={{ background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 6, padding: 10, color: '#166534', fontSize: 12 }}>{message}</div>}
      <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <input value={siteFilter} onChange={e => setSiteFilter(e.target.value)} placeholder="Search sites by name or URL" style={{ flex: 1, padding: '8px 10px', border: '1px solid #cbd5e1', borderRadius: 6, fontSize: 12 }} />
        <label style={{ fontSize: 11, color: '#475569', whiteSpace: 'nowrap' }}><input type="checkbox" checked={showAll} onChange={e => setShowAll(e.target.checked)} /> Show all sites</label>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1.3fr 1fr', gap: 12 }}>
        <div style={{ maxHeight: 360, overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: 8 }}>
          {visibleSites.map(site => <button key={site.id} onClick={() => selectSite(site)} style={{ display: 'block', width: '100%', textAlign: 'left', padding: 10, border: 0, borderBottom: '1px solid #f1f5f9', background: selectedSite?.id === site.id ? '#eff6ff' : '#fff', cursor: 'pointer' }}><strong style={{ display: 'block', fontSize: 12, color: '#0f172a' }}>{site.name}</strong><span style={{ display: 'block', marginTop: 3, fontSize: 10, color: '#64748b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{site.web_url}</span><span style={{ display: 'block', marginTop: 3, fontSize: 10, color: '#94a3b8' }}>Select site</span></button>)}
          {!visibleSites.length && <div style={{ padding: 14, fontSize: 12, color: '#64748b' }}>No matching sites.</div>}
        </div>
        <div style={{ border: '1px solid #e2e8f0', borderRadius: 8, padding: 12 }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: '#0f172a', marginBottom: 8 }}>2. Select document library</div>
          {!selectedSite && <div style={{ fontSize: 11, color: '#64748b' }}>Select a site first.</div>}
          {loadingDrives && <div style={{ fontSize: 11, color: '#64748b' }}>Loading libraries...</div>}
          {drives.map(drive => <button key={drive.id} onClick={() => setSelectedDrive(drive)} style={{ display: 'block', width: '100%', textAlign: 'left', padding: 9, marginBottom: 6, border: '1px solid #cbd5e1', borderRadius: 6, background: selectedDrive?.id === drive.id ? '#dcfce7' : '#fff', cursor: 'pointer', fontSize: 11 }}>{drive.name}<span style={{ display: 'block', color: '#64748b', fontSize: 10 }}>{drive.drive_type}</span></button>)}
          {selectedSite && !loadingDrives && !drives.length && <div style={{ fontSize: 11, color: '#64748b' }}>No document libraries found.</div>}
        </div>
      </div>
      {selectedSite && selectedDrive && <div style={{ background: '#f8fafc', border: '1px solid #cbd5e1', borderRadius: 8, padding: 14 }}>
        <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10, color: '#0f172a' }}>3. Confirm site configuration</div>
        <div style={{ fontSize: 11, color: '#475569', lineHeight: 1.6, marginBottom: 10 }}><strong>Site:</strong> {selectedSite.name}<br /><strong>Site ID:</strong> <code>{selectedSite.id}</code><br /><strong>Drive:</strong> {selectedDrive.name}<br /><strong>Drive ID:</strong> <code>{selectedDrive.id}</code></div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}><input value={displayName} onChange={e => setDisplayName(e.target.value)} placeholder="Display Name, e.g. Vessel Management - Production" style={{ padding: 8, border: '1px solid #cbd5e1', borderRadius: 6, fontSize: 11 }} /><input value={siteKey} onChange={e => setSiteKey(e.target.value)} placeholder="Internal Site Key, e.g. vessel_mgmt_prod" style={{ padding: 8, border: '1px solid #cbd5e1', borderRadius: 6, fontSize: 11 }} /></div>
        <button onClick={saveConfiguration} disabled={saving || !displayName.trim() || !siteKey.trim()} style={{ marginTop: 10, padding: '9px 14px', border: 0, borderRadius: 6, background: saving || !displayName.trim() || !siteKey.trim() ? '#cbd5e1' : '#0078d4', color: '#fff', fontSize: 12, fontWeight: 600, cursor: saving ? 'wait' : 'pointer' }}>{saving ? 'Saving...' : 'Save Site Configuration'}</button>
      </div>}
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
          <button
            onClick={() => host.setState({ settingsTab: 'Site Selection' })}
            style={{
              padding: '8px 14px',
              borderRadius: 6,
              border: '1px solid #bfdbfe',
              background: '#eff6ff',
              color: '#1d4ed8',
              fontSize: 12,
              fontWeight: 700,
              cursor: 'pointer',
            }}
          >
            Configure Connected Sites
          </button>
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

