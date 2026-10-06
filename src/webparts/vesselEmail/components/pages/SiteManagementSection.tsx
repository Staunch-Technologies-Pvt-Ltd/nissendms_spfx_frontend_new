/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable @typescript-eslint/no-unused-expressions */
import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { Icon } from '@fluentui/react/lib/Icon';
import { clay } from '../clayTheme';
import { DMS_FONT_DISPLAY, DMS_ON_ACCENT, dmsBtn, dmsControlStyle, dmsTone } from '../dmsDesignSystem';

/**
 * Sites → Site Management.
 * Moved here (unchanged) from the former Settings → Site Management tab —
 * same component, same state, same handlers, same backend calls
 * (app/main.py's /api/admin/* site-configuration endpoints). Lives on the
 * main Sites page now so site browsing and site administration are in one
 * place; Settings no longer has a Site Management tab.
 */
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

  if (loading) return <div style={{ color: 'var(--vdms-text-muted)', fontSize: 13 }}>Loading site configuration...</div>;

  const successTone = dmsTone('success');
  const warnTone = dmsTone('warning');
  const dangerTone = dmsTone('danger');

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14, width: '100%' }}>
      <div style={{ background: 'var(--vdms-surface-alt)', border: '1px solid var(--vdms-line)', borderRadius: 8, padding: 16, width: '100%', boxSizing: 'border-box' }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: 'var(--vdms-text)', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
          <span><Icon iconName="Globe" aria-hidden="true" style={{ fontSize: 13 }} /></span> Default Site: {info?.site_name || info?.active_site}
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '170px minmax(0, 1fr)', columnGap: 16, rowGap: 8, fontSize: 12, width: '100%' }}>
          <span style={{ fontWeight: 600, color: 'var(--vdms-text-muted)' }}>Site Key:</span>
          <span style={{ fontFamily: 'monospace', color: 'var(--vdms-text)', fontWeight: 700, wordBreak: 'break-word' }}>{info?.active_site}</span>

          <span style={{ fontWeight: 600, color: 'var(--vdms-text-muted)' }}>Database:</span>
          <span style={{ fontFamily: 'monospace', color: info?.db_configured ? successTone.fg : warnTone.fg, fontWeight: 600 }}>
            {info?.db_name || (info?.db_configured ? 'Connected' : 'In-Memory Stub')}
          </span>

          <span style={{ fontWeight: 600, color: 'var(--vdms-text-muted)' }}>SharePoint Status:</span>
          <span style={{ color: info?.sp_configured ? successTone.fg : 'var(--vdms-text-muted)', fontWeight: 600 }}>
            {info?.sp_configured ? <><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 12 }} /> Configured</> : <><Icon iconName="StatusCircleRing" aria-hidden="true" style={{ fontSize: 12 }} /> Stub / Local Mode</>}
          </span>

          <span style={{ fontWeight: 600, color: 'var(--vdms-text-muted)' }}>Drive ID:</span>
          <span style={{ fontFamily: 'monospace', color: 'var(--vdms-text-muted)', fontSize: 11, wordBreak: 'break-all' }}>
            {info?.drive_id || '—'}
          </span>

          <span style={{ fontWeight: 600, color: 'var(--vdms-text-muted)' }}>Backend Mode:</span>
          <span style={{ color: clay.accent, fontWeight: 600 }}>{info?.mode}</span>
        </div>
      </div>

      {isAdmin && adminConfig && (
        <div style={{ background: successTone.bg, border: `1px solid ${successTone.bg}`, borderRadius: 8, padding: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: successTone.fg, marginBottom: 6, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <span><Icon iconName="Contact" aria-hidden="true" style={{ fontSize: 13 }} /> Configured SharePoint Sites & Default Site</span>
            <button
              onClick={() => {
                void loadAdminConfig();
                void loadTenantSites(host._base(), host._headers());
              }}
              disabled={tenantSitesLoading}
              title="Refresh all tenant and configured sites"
              style={dmsBtn('primary', !tenantSitesLoading)}
            >
              {tenantSitesLoading ? <><Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 11 }} /> Loading…</> : <><Icon iconName="Refresh" aria-hidden="true" style={{ fontSize: 11 }} /> Refresh Sites</>}
            </button>
          </div>
          <p style={{ margin: '0 0 10px', fontSize: 12, color: successTone.fg, lineHeight: 1.5 }}>
            Select the default site and manage site visibility. Hiding a site removes it from Sites, Documents, and Vessel Management modules. Hidden sites remain disabled here until you click Unhide.
          </p>

          {/* Search filter */}
          <input
            value={siteSearch}
            onChange={e => setSiteSearch(e.target.value)}
            placeholder="Search site by name or URL…"
            style={{ ...dmsControlStyle(), width: '100%', marginBottom: 8 }}
          />

          {defaultMsg && (
            <div style={{ background: defaultMsg.startsWith('✓') ? successTone.bg : dangerTone.bg, border: `1px solid ${defaultMsg.startsWith('✓') ? successTone.bg : dangerTone.bg}`, color: defaultMsg.startsWith('✓') ? successTone.fg : dangerTone.fg, borderRadius: 6, padding: '6px 10px', fontSize: 12, marginBottom: 8 }}>
              {defaultMsg}
            </div>
          )}

          {visibilityMsg && (
            <div style={{ background: dangerTone.bg, border: `1px solid ${dangerTone.bg}`, color: dangerTone.fg, borderRadius: 6, padding: '6px 10px', fontSize: 12, marginBottom: 8 }}>
              {visibilityMsg}
            </div>
          )}

          {/* ── Configured Sites card list ── */}
          <div style={{ marginBottom: 12 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: successTone.fg }}>Configured Sites:</label>
              <span style={{ fontSize: 11, color: 'var(--vdms-text-muted)' }}>Hidden sites remain listed here for unhide.</span>
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
                      border: `1px solid ${isCurrent ? successTone.bg : isHidden ? 'var(--vdms-line-strong)' : successTone.bg}`,
                      background: isCurrent ? successTone.bg : isHidden ? 'var(--vdms-surface-alt)' : successTone.bg,
                      opacity: isHidden && !isCurrent ? 0.65 : 1,
                      transition: 'all 0.15s',
                    }}
                  >
                    <div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <span style={{ fontSize: 13, fontWeight: 600, color: isHidden ? 'var(--vdms-text-muted)' : successTone.fg }}>
                          {site.display_name || site.name}
                        </span>
                        {isCurrent && (
                          <span style={{ fontSize: 10, fontWeight: 700, background: successTone.fg, color: DMS_ON_ACCENT, padding: '2px 7px', borderRadius: 8 }}>
                            <Icon iconName="FavoriteStar" aria-hidden="true" style={{ fontSize: 10 }} /> DEFAULT
                          </span>
                        )}
                        {isHidden && (
                          <span style={{ fontSize: 10, fontWeight: 700, background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text-muted)', padding: '2px 7px', borderRadius: 8 }}>
                            HIDDEN
                          </span>
                        )}
                      </div>
                      <span style={{ display: 'block', fontSize: 10, color: 'var(--vdms-text-muted)', fontFamily: 'monospace', marginTop: 2 }}>
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
                          style={dmsBtn('primary', !(isSettingDefault || isToggling))}
                        >
                          {isSettingDefault ? <><Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 11 }} /> Setting…</> : <><Icon iconName="FavoriteStar" aria-hidden="true" style={{ fontSize: 11 }} /> Set as Default</>}
                        </button>
                      )}
                      {/* Hidden rows stay visible here, but are disabled everywhere else. */}
                      {isHidden && (
                        <button
                          type="button"
                          disabled
                          style={dmsBtn('secondary', false)}
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
                        style={dmsBtn('secondary', !isCurrent)}
                      >
                        {isToggling ? <Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 11 }} /> : isHidden ? 'Unhide' : 'Hide'}
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
                        style={dmsBtn('danger', !isCurrent)}
                      >
                        {isRemoving ? <><Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 11 }} /> Removing…</> : 'Remove'}
                      </button>
                    </div>
                  </div>
                );
              })}
              {filteredConfiguredSites.length === 0 && (
                <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', padding: '8px 0' }}>
                  No configured sites match your search.
                </div>
              )}
            </div>
          </div>

          {/* ── Add New Site section ── */}
          <div style={{ marginBottom: 12, border: `1px solid ${successTone.bg}`, borderRadius: 8, overflow: 'hidden' }}>
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
              style={{ width: '100%', padding: '9px 14px', background: successTone.bg, border: 'none', color: successTone.fg, fontSize: 12, fontWeight: 700, cursor: 'pointer', textAlign: 'left' }}
            >
              {showAddSite ? <><Icon iconName="Cancel" aria-hidden="true" style={{ fontSize: 12 }} /> Cancel Add Site</> : <><Icon iconName="Add" aria-hidden="true" style={{ fontSize: 12 }} /> Add New Site</>}
            </button>

            {showAddSite && (
              <div style={{ padding: 14, background: 'var(--vdms-surface)' }}>
                <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginBottom: 10 }}>Search and select a SharePoint site to register as a new configured site.</div>
                <input
                  value={addSiteSearch}
                  onChange={e => setAddSiteSearch(e.target.value)}
                  placeholder="Search tenant sites…"
                  style={{ ...dmsControlStyle(), width: '100%', marginBottom: 8 }}
                />
                {addSiteTenantLoading && <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)', marginBottom: 8 }}><Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 12 }} /> Loading tenant sites…</div>}
                <div style={{ maxHeight: 160, overflowY: 'auto', border: '1px solid var(--vdms-line)', borderRadius: 6, marginBottom: 10 }}>
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
                          border: 0, borderBottom: '1px solid var(--vdms-border-soft)',
                          background: addSiteSelected?.id === ts.id ? clay.accentSoft : 'var(--vdms-surface)',
                          cursor: 'pointer', fontSize: 12, color: 'var(--vdms-text)',
                        }}
                      >
                        <strong>{ts.name}</strong>
                        <span style={{ display: 'block', fontSize: 10, color: 'var(--vdms-text-muted)' }}>{ts.web_url}</span>
                      </button>
                    ))}
                </div>

                {/* Drive picker */}
                {addSiteSelected && (
                  <div style={{ marginBottom: 10 }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--vdms-text)', marginBottom: 6 }}>Document Library:</div>
                    {addSiteDrivesLoading && <div style={{ fontSize: 11, color: 'var(--vdms-text-muted)' }}>Loading libraries…</div>}
                    {!addSiteDrivesLoading && addSiteDrives.map(drv => (
                      <button
                        key={drv.id}
                        onClick={() => setAddSiteSelDrive(drv)}
                        style={{
                          display: 'block', width: '100%', textAlign: 'left', padding: '7px 10px', marginBottom: 4,
                          border: `1px solid ${addSiteSelDrive?.id === drv.id ? successTone.fg : 'var(--vdms-line)'}`,
                          borderRadius: 5, background: addSiteSelDrive?.id === drv.id ? successTone.bg : 'var(--vdms-surface)',
                          cursor: 'pointer', fontSize: 11,
                        }}
                      >
                        <strong>{drv.name}</strong>
                        <span style={{ display: 'block', fontSize: 10, color: 'var(--vdms-text-muted)' }}>{drv.drive_type}</span>
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
                      style={dmsControlStyle()}
                    />
                    <input
                      value={addSiteKey}
                      onChange={e => setAddSiteKey(e.target.value.toLowerCase().replace(/[^a-z0-9_]/g, ''))}
                      placeholder="Internal site key (e.g. vessel_prod)"
                      style={{ ...dmsControlStyle(), fontFamily: 'monospace' }}
                    />
                  </div>
                )}

                {addSiteMsg && (
                  <div style={{ background: addSiteMsg.startsWith('✓') ? successTone.bg : dangerTone.bg, border: `1px solid ${addSiteMsg.startsWith('✓') ? successTone.bg : dangerTone.bg}`, color: addSiteMsg.startsWith('✓') ? successTone.fg : dangerTone.fg, borderRadius: 6, padding: '6px 10px', fontSize: 12, marginBottom: 8 }}>
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
                  style={dmsBtn('primary', !(!addSiteSelected || !addSiteSelDrive || !addSiteDisplayName.trim() || !addSiteKey.trim() || addSiteSaving))}
                >
                  {addSiteSaving ? <><Icon iconName="Sync" aria-hidden="true" style={{ fontSize: 12 }} /> Saving…</> : <><Icon iconName="CheckMark" aria-hidden="true" style={{ fontSize: 12 }} /> Save & Register Site</>}
                </button>
              </div>
            )}
          </div>

          {showChanges && siteChanges.length > 0 && (
            <div style={{ marginTop: 16, borderTop: `1px solid ${successTone.bg}`, paddingTop: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: successTone.fg, marginBottom: 8 }}>Recent Changes:</div>
              <div style={{ maxHeight: 200, overflowY: 'auto', fontSize: 11, color: successTone.fg }}>
                {siteChanges.map((change: any) => (
                  <div key={change.id} style={{ padding: '6px 0', borderBottom: `1px solid ${successTone.bg}` }}>
                    <div>
                      <strong>
                        {change.action === 'hidden' ? 'Hidden' : change.action === 'unhidden' ? 'Unhidden' : change.action === 'removed' ? 'Removed' : `${change.previous_site} → ${change.new_site}`}
                      </strong>
                      {change.action && change.action !== 'success' ? ` ${change.new_site_name || change.new_site}` : ` (${change.status})`}
                      {change.action && change.action !== 'success' && <span> ({change.status})</span>}
                    </div>
                    <div style={{ fontSize: 10, color: successTone.fg }}>
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
              style={{ marginTop: 8, background: 'transparent', border: 'none', color: successTone.fg, fontSize: 11, fontWeight: 600, cursor: 'pointer', textDecoration: 'underline' }}
            >
              {showChanges ? <Icon iconName="ChevronDown" aria-hidden="true" style={{ fontSize: 10 }} /> : <Icon iconName="ChevronRight" aria-hidden="true" style={{ fontSize: 10 }} />} Hide / Show Recent Changes
            </button>
          )}
        </div>
      )}

      {err && (
        <div style={{ background: dangerTone.bg, border: `1px solid ${dangerTone.bg}`, borderRadius: 6, padding: 12, fontSize: 12, color: dangerTone.fg }}>
          <Icon iconName="Warning" aria-hidden="true" style={{ fontSize: 12 }} /> {err}
        </div>
      )}
    </div>
  );
}

/**
 * Wraps the (unchanged) site-administration panel above in a card that
 * matches the rest of the Sites page — a heading plus a white, rounded,
 * shadowed container — instead of the bare block it used to sit in under
 * Settings' own left-nav shell.
 */
export function SiteManagementSection({ host }: { host: VesselEmail }): React.ReactElement {
  return (
    <div style={{ marginTop: 24 }}>
      <h2 style={{ margin: '0 0 12px', fontSize: 20, fontWeight: 800, color: 'var(--vdms-text)', fontFamily: DMS_FONT_DISPLAY }}>Site Management</h2>
      <div style={{ background: 'var(--vdms-glass)', borderRadius: 18, boxShadow: clay.shadowRaised, backdropFilter: 'blur(18px) saturate(1.3)', WebkitBackdropFilter: 'blur(18px) saturate(1.3)', padding: 18, boxSizing: 'border-box' }}>
        <SiteIntegrationInfo host={host} />
      </div>
    </div>
  );
}
