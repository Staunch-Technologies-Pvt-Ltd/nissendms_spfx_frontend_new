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
  const [switchReason, setSwitchReason] = React.useState<string>('');
  const [switching, setSwitching] = React.useState<boolean>(false);
  const [switchMsg, setSwitchMsg] = React.useState<string | null>(null);
  const [switchSuccess, setSwitchSuccess] = React.useState<boolean>(false);
  const [siteChanges, setSiteChanges] = React.useState<any[]>([]);
  const [showChanges, setShowChanges] = React.useState<boolean>(false);
  const [discoveringSites, setDiscoveringSites] = React.useState<boolean>(false);
  const [discoveredSites, setDiscoveredSites] = React.useState<any[]>([]);
  const [siteSearch, setSiteSearch] = React.useState<string>('');
  const [selectedDiscoveredSite, setSelectedDiscoveredSite] = React.useState<any>(null);
  const [siteDrives, setSiteDrives] = React.useState<any[]>([]);
  const [loadingDrives, setLoadingDrives] = React.useState<boolean>(false);
  const [selectedDrive, setSelectedDrive] = React.useState<any>(null);
  const [discoverMsg, setDiscoverMsg] = React.useState<string | null>(null);

  React.useEffect(() => {
    const base = host._base();
    const userEmail = host.props.userEmail;
    
    // Fetch basic site info
    fetch(`${base}/api/config/site-info`)
      .then(r => r.json())
      .then(d => {
        setInfo(d);
      })
      .catch((e: any) => {
        setErr(e?.message || 'Failed to load site information');
      });

    // Fetch admin site configuration if user is admin
    if (userEmail) {
      fetch(`${base}/api/admin/site-configuration`, {
        headers: host._headers(),
      })
        .then(r => {
          if (r.status === 403) {
            setIsAdmin(false);
            setLoading(false);
            return null;
          }
          return r.json();
        })
        .then(d => {
          if (d) {
            const applySites = (sites: any[]) => {
              setAdminConfig({ ...d, available_sites: sites });
              setIsAdmin(true);
              setSelectedNewSite(d.current_site);
            };

            if (Array.isArray(d.available_sites) && d.available_sites.length > 0) {
              applySites(d.available_sites);
            } else {
              // Support the existing session endpoint if the backend is on an older build.
              fetch(`${base}/api/config/available-sites`, { headers: host._headers() })
                .then(r => r.ok ? r.json() : Promise.reject(new Error('Failed to load available sites')))
                .then(sitesResponse => applySites(sitesResponse.sites || []))
                .catch(() => applySites(d.available_sites || []));
            }
            
            // Fetch site changes for audit log
            fetch(`${base}/api/admin/site-changes?limit=10`, {
              headers: host._headers(),
            })
              .then(r => r.json())
              .then(changes => setSiteChanges(changes.changes || []))
              .catch(e => console.log('Failed to load changes:', e));
          }
          setLoading(false);
        })
        .catch((e: any) => {
          setIsAdmin(false);
          setLoading(false);
        });
    } else {
      setLoading(false);
    }
  }, [host, host.props.userEmail]);

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
      const response = await fetch(`${base}/api/admin/switch-site`, {
        method: 'POST',
        headers: {
          ...host._headers(),
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          site_name: selectedNewSite,
          reason: switchReason || undefined,
        }),
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.detail || 'Failed to switch site');
      }

      setSwitchSuccess(true);
      setSwitchMsg(`✓ Successfully switched to ${data.new_site_name}. Reloading page...`);
      setShowSwitchDialog(false);
      setSwitchReason('');

      // Reload the page after 2 seconds
      setTimeout(() => {
        window.location.reload();
      }, 2000);
    } catch (e: any) {
      setSwitchMsg(`⚠️ ${e?.message || 'Failed to switch site'}`);
      setSwitching(false);
    }
  };

  const discoverSites = async () => {
    setDiscoveringSites(true);
    setDiscoverMsg(null);
    try {
      const response = await fetch(`${host._base()}/api/admin/discover-sites`, { headers: host._headers() });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Unable to discover SharePoint sites');
      setDiscoveredSites(data.sites || []);
      if (!(data.sites || []).length) setDiscoverMsg('No SharePoint sites were found for this app registration. Check site access and admin consent.');
    } catch (e: any) {
      setDiscoverMsg(e?.message || 'Unable to discover SharePoint sites');
      setDiscoveredSites([]);
    } finally {
      setDiscoveringSites(false);
    }
  };

  const selectDiscoveredSite = async (site: any) => {
    setSelectedDiscoveredSite(site);
    setSelectedDrive(null);
    setSiteDrives([]);
    setLoadingDrives(true);
    setDiscoverMsg(null);
    try {
      const response = await fetch(`${host._base()}/api/admin/discover-sites/${encodeURIComponent(site.id)}/drives`, { headers: host._headers() });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Unable to load document libraries');
      setSiteDrives(data.drives || []);
      if (!(data.drives || []).length) setDiscoverMsg('No document libraries were found on this site.');
    } catch (e: any) {
      setDiscoverMsg(e?.message || 'Unable to load document libraries');
    } finally {
      setLoadingDrives(false);
    }
  };

  const copyResolvedValue = (value: string) => {
    if (value) navigator.clipboard?.writeText(value);
  };

  const filteredDiscoveredSites = discoveredSites.filter(site =>
    `${site.name} ${site.web_url}`.toLowerCase().includes(siteSearch.toLowerCase())
  );

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
            <div style={{ fontSize: 13, fontWeight: 700, color: '#166534', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span>👤</span> Admin: Switch Active Site
            </div>
            <p style={{ margin: '0 0 12px', fontSize: 12, color: '#166534', lineHeight: 1.5 }}>
              Select a different SharePoint site to make it the active site for all users. This affects which site's documents and folders are displayed in the application.
            </p>

            <div style={{ borderTop: '1px solid #bbf7d0', paddingTop: 12, marginBottom: 16 }}>
              <button
                onClick={discoverSites}
                disabled={discoveringSites}
                style={{ padding: '8px 12px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: discoveringSites ? 'wait' : 'pointer' }}
              >
                {discoveringSites ? 'Discovering...' : 'Discover Sites'}
              </button>
              <span style={{ marginLeft: 10, fontSize: 11, color: '#166534' }}>Browse sites and document libraries available to this app.</span>

              {(discoveredSites.length > 0 || discoveringSites) && (
                <div style={{ marginTop: 12, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                  <div>
                    <input
                      value={siteSearch}
                      onChange={e => setSiteSearch(e.target.value)}
                      placeholder="Filter sites by name or URL"
                      style={{ width: '100%', boxSizing: 'border-box', padding: '8px 10px', border: '1px solid #bbf7d0', borderRadius: 6, fontSize: 11 }}
                    />
                    <div style={{ marginTop: 6, maxHeight: 170, overflowY: 'auto', background: '#fff', border: '1px solid #dcfce7', borderRadius: 6 }}>
                      {filteredDiscoveredSites.map(site => (
                        <button
                          key={site.id}
                          onClick={() => selectDiscoveredSite(site)}
                          style={{ display: 'block', width: '100%', textAlign: 'left', padding: '8px 10px', border: 'none', borderBottom: '1px solid #f0fdf4', background: selectedDiscoveredSite?.id === site.id ? '#dcfce7' : '#fff', cursor: 'pointer' }}
                        >
                          <strong style={{ display: 'block', fontSize: 11, color: '#166534' }}>{site.name}</strong>
                          <span style={{ display: 'block', fontSize: 10, color: '#64748b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{site.web_url || 'No URL returned'}</span>
                        </button>
                      ))}
                      {!filteredDiscoveredSites.length && <div style={{ padding: 10, fontSize: 11, color: '#64748b' }}>No matching sites.</div>}
                    </div>
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 600, color: '#166534', marginBottom: 6 }}>Document Libraries</div>
                    {loadingDrives && <div style={{ fontSize: 11, color: '#64748b' }}>Loading libraries...</div>}
                    {!loadingDrives && selectedDiscoveredSite && !siteDrives.length && <div style={{ fontSize: 11, color: '#64748b' }}>Select a site to view its libraries.</div>}
                    {!loadingDrives && siteDrives.map(drive => (
                      <button
                        key={drive.id}
                        onClick={() => setSelectedDrive(drive)}
                        style={{ display: 'block', width: '100%', textAlign: 'left', padding: '8px 10px', marginBottom: 5, border: '1px solid #dcfce7', borderRadius: 6, background: selectedDrive?.id === drive.id ? '#dcfce7' : '#fff', color: '#166534', cursor: 'pointer', fontSize: 11 }}
                      >
                        <strong>{drive.name}</strong><span style={{ display: 'block', fontSize: 10, color: '#64748b' }}>{drive.drive_type || 'documentLibrary'}</span>
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {selectedDiscoveredSite && selectedDrive && (
                <div style={{ marginTop: 10, padding: 10, background: '#fff', border: '1px solid #86efac', borderRadius: 6, fontSize: 11 }}>
                  <div style={{ fontWeight: 700, color: '#166534', marginBottom: 6 }}>Resolved Site Configuration</div>
                  {[['Site Key', selectedDiscoveredSite.name.toLowerCase().replace(/[^a-z0-9]+/g, '_')], ['Site ID', selectedDiscoveredSite.id], ['Drive ID', selectedDrive.id]].map(([label, value]) => (
                    <div key={label} style={{ display: 'flex', alignItems: 'center', gap: 6, marginTop: 4 }}>
                      <strong style={{ width: 55, color: '#475569' }}>{label}:</strong><code style={{ flex: 1, overflowWrap: 'anywhere', color: '#0f172a' }}>{value}</code>
                      <button onClick={() => copyResolvedValue(String(value))} title={`Copy ${label}`} style={{ padding: '3px 7px', border: '1px solid #cbd5e1', borderRadius: 4, background: '#f8fafc', cursor: 'pointer', fontSize: 10 }}>Copy</button>
                    </div>
                  ))}
                </div>
              )}
              {discoverMsg && <div style={{ marginTop: 8, padding: 8, background: '#fff7ed', border: '1px solid #fed7aa', borderRadius: 6, color: '#9a3412', fontSize: 11 }}>{discoverMsg}</div>}
            </div>
            
            <div style={{ marginBottom: 12 }}>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#166534', marginBottom: 6 }}>
                Switch To:
              </label>
              <select
                value={selectedNewSite}
                onChange={e => setSelectedNewSite(e.target.value)}
                disabled={switching}
                style={{
                  width: '100%',
                  padding: '8px 12px',
                  borderRadius: 6,
                  border: '1px solid #bbf7d0',
                  fontSize: 13,
                  outline: 'none',
                  background: '#fff',
                  cursor: switching ? 'not-allowed' : 'pointer',
                  opacity: switching ? 0.6 : 1,
                }}
              >
                <option value="">-- Select a site --</option>
                {adminConfig.available_sites?.map((site: any) => (
                  <option key={site.name} value={site.name} disabled={site.name === adminConfig.current_site}>
                    {site.display_name} ({site.name}){site.name === adminConfig.current_site ? ' (current)' : ''}
                  </option>
                ))}
              </select>
            </div>

            <button
              onClick={() => setShowSwitchDialog(true)}
              disabled={!selectedNewSite || selectedNewSite === adminConfig.current_site || switching}
              style={{
                padding: '10px 16px',
                background: selectedNewSite && selectedNewSite !== adminConfig.current_site && !switching ? '#16a34a' : '#cbd5e1',
                color: '#fff',
                border: 'none',
                borderRadius: 6,
                fontSize: 12,
                fontWeight: 600,
                cursor: selectedNewSite && selectedNewSite !== adminConfig.current_site && !switching ? 'pointer' : 'not-allowed',
              }}
            >
              {switching ? '⏳ Switching...' : 'Switch Site'}
            </button>

            {switchMsg && (
              <div style={{
                marginTop: 12,
                background: switchSuccess ? '#dcfce7' : '#fee2e2',
                border: `1px solid ${switchSuccess ? '#86efac' : '#fca5a5'}`,
                color: switchSuccess ? '#166534' : '#991b1b',
                borderRadius: 6,
                padding: '8px 12px',
                fontSize: 12,
                fontWeight: 600,
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
                      <div>
                        <strong>{change.previous_site} → {change.new_site}</strong> ({change.status})
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
                style={{
                  marginTop: 8,
                  background: 'transparent',
                  border: 'none',
                  color: '#16a34a',
                  fontSize: 11,
                  fontWeight: 600,
                  cursor: 'pointer',
                  textDecoration: 'underline',
                }}
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
              toSite={selectedNewSite}
              toSiteConfig={adminConfig.available_sites?.find((s: any) => s.name === selectedNewSite)}
              onConfirm={() => handleSwitchSite()}
              onCancel={() => {
                setShowSwitchDialog(false);
                setSwitchReason('');
              }}
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
              <strong>For admins:</strong> Use the "Switch Active Site" dropdown above to change the site for all users. No server restarts needed.
            </p>
            <p style={{ margin: '0 0 6px' }}>
              <strong>Fallback method:</strong> Manually edit <code>ACTIVE_SITE=&lt;site_key&gt;</code> in the backend <code>.env</code> file and restart the server (requires server access).
            </p>
          </>
        ) : (
          <p style={{ margin: 0 }}>
            Contact an administrator to switch to a different SharePoint site. Changing the active site requires admin-level access.
          </p>
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
          <strong>Warning:</strong> Switching the active site affects all users in the application. This change is immediate and will redirect users' data contexts to the new site.
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
                  <span style={{ textDecoration: 'line-through', color: '#a1a1a1' }}>{fromDb}</span>
                  <span style={{ display: 'block', color: '#15803d', fontWeight: 600 }}>→ (new site DB)</span>
                </td>
              </tr>
              <tr>
                <td style={{ padding: '6px 0', fontWeight: 600, color: '#475569' }}>Drive ID:</td>
                <td style={{ padding: '6px 0 6px 12px', color: '#0f172a', fontSize: 10, fontFamily: 'monospace' }}>
                  <span style={{ textDecoration: 'line-through', color: '#a1a1a1' }}>{fromDriveId.substring(0, 20)}...</span>
                  <span style={{ display: 'block', color: '#15803d', fontWeight: 600 }}>→ (new site drive)</span>
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
                {sites.map(s => (
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
                                <span>✅</span> {siteNameMap[sk] || sk}
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
              {sites.filter(s => s.is_available_for_provisioning).map(site => {
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

