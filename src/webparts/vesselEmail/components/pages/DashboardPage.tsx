import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';
import type { VesselRecord } from '../types/rows';
import type { AlertItem } from '../types/ui';
import { dashboardHeroShipImage } from '../dashboardHeroShip';

// ─────────────────────────────────────────────────────────────────────────────
// Home / Dashboard
//
// An at-a-glance summary, not another browsing surface: live totals (GET
// /api/dashboard/stats), fleet status (GET /api/vessels, already loaded
// app-wide), monthly document activity (facets from GET
// /api/dashboard/documents, the same endpoint the Documents module itself
// uses) and real events/anomalies (GET /api/alerts/all, already polled
// app-wide for the header alert bell). Every document, every vessel and
// every site in full detail still lives in the Documents / Vessels / Sites
// modules — this page never re-renders those as tables.
//
// Plain inline styles + React 17 hooks only (SPFx / Fluent UI 8 constraint).
// ─────────────────────────────────────────────────────────────────────────────

export interface RealDashboardDoc {
  id: string;
  name: string;
  vessel: string;
  type: string;
  status?: 'Valid' | 'Expiring Soon' | 'Expired' | string;
  modified: string;
  modifiedEpoch: number;
  createdEpoch?: number;
  modifiedBy?: string;
  createdBy?: string;
  fileSize?: string;
  sizeBytes?: number;
  subFolderPath?: string;
  webUrl?: string;
  ext?: string;
  fileType?: string;
  site?: string;
  siteName?: string;
}

export interface DashboardSiteSummary {
  site_key: string;
  site_name: string;
  drive_id?: string;
  web_url?: string;
  files: number;
  folders: number;
  last_modified_epoch?: number | null;
  truncated?: boolean;
  /** Set when the backend couldn't scan this site (permission, throttling,
   * missing drive_id, ...) — files/folders are 0 in that case too, but this
   * field is what distinguishes "failed to fetch" from "actually empty". */
  error?: string | null;
  /** Set for a site that was just added (or just became visible) and whose
   * real file/folder counts haven't been scanned yet — files/folders are 0
   * placeholders here, not an actual empty count. The backend fills this in
   * with a background scan within moments; no user action needed. */
  stats_pending?: boolean;
  /** Backend attached a live running total (files/folders read so far) while
   * the full scan of this site is still in progress. */
  counting?: boolean;
  /** Live vessel count for this site — root-level (or one-level-nested)
   * folders whose name matches a Term Store vessel term, computed as part
   * of the same scan as files/folders (see backend
   * _term_store_vessel_folder_count). This is what "Total Vessels" sums. */
  vessels?: number;
  /** The official Term Store labels behind `vessels` — real, current
   * SharePoint vessel folder names for this site, not read from the
   * vessels DB table. */
  vessel_names?: string[];
}

export interface DashboardStats {
  total_documents: number;
  total_files?: number;
  total_folders?: number;
  total_sites?: number;
  total_vessels: number;
  sites?: DashboardSiteSummary[];
  truncated?: boolean;
  recent_documents: RealDashboardDoc[];
  documents?: RealDashboardDoc[];
  /** When these figures were computed (epoch ms). Counts are served from a
   * short-lived server cache, kept warm by a background job, rather than
   * recomputed on every page load — this is what "Refresh" invalidates. */
  last_refreshed_epoch?: number | null;
}

/**
 * All documents currently loaded for the dashboard (full list when the
 * stats endpoint returned it, otherwise the recent-documents sample).
 * Used by ReportsPage.
 */
export function extractRealDocuments(host: VesselEmail): RealDashboardDoc[] {
  const stats = host.state.dashboardStats;
  if (!stats) return [];
  if (stats.documents && stats.documents.length > 0) return stats.documents;
  return stats.recent_documents || [];
}

// Kept across visits to Home in the same page load, so leaving Home and
// coming back doesn't lose the last refresh.
let refreshNonce = 0;

const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function formatNumber(value: number | undefined | null): string {
  return typeof value === 'number' && isFinite(value) ? value.toLocaleString() : '—';
}

function formatDateTime(epoch?: number | null): string {
  if (!epoch) return '—';
  const d = new Date(epoch);
  return d.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' }) +
    ', ' + d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
}

/** "vessel_unrecognised" → "Vessel unrecognised" (generic alert-type label). */
function humanize(value: string): string {
  const t = String(value || '').replace(/_/g, ' ').trim();
  return t ? t.charAt(0).toUpperCase() + t.slice(1) : 'Alert';
}

function timeAgo(epochMs: number | null | undefined): string {
  if (!epochMs) return '';
  const diff = Date.now() - epochMs;
  if (diff < 60_000) return 'Just now';
  if (diff < 3_600_000) return `${Math.floor(diff / 60_000)}m ago`;
  if (diff < 86_400_000) return `${Math.floor(diff / 3_600_000)}h ago`;
  if (diff < 7 * 86_400_000) return `${Math.floor(diff / 86_400_000)}d ago`;
  return new Date(epochMs).toLocaleDateString(undefined, { day: '2-digit', month: 'short' });
}

function alertEpoch(a: AlertItem): number {
  const t = a.created_at ? new Date(a.created_at).getTime() : NaN;
  return isFinite(t) ? t : 0;
}

/** Anomaly-type alerts are the real, actionable "something needs a look"
 *  signal — as opposed to routine activity (folder created, email sent, …). */
const ANOMALY_ALERT_TYPES = new Set(['vessel_unrecognised', 'file_outside_structure', 'subfolder_anomaly']);

function iconForAlert(a: AlertItem): string {
  switch (a.alert_type) {
    case 'vessel_provisioned': return 'Ferry';
    case 'vessel_deleted':
    case 'document_deleted': return 'Delete';
    case 'folder_created': return 'FabricFolder';
    case 'crud_operation': return 'Sync';
    case 'vessel_unrecognised':
    case 'file_outside_structure':
    case 'subfolder_anomaly': return 'Warning';
    default: return 'History';
  }
}

interface Palette { surface: string; border: string; text: string; muted: string; }

const controlStyle = (border: string, surface: string, text: string): React.CSSProperties => ({
  padding: '7px 10px', borderRadius: 8, border: `1px solid ${border}`, fontSize: 12,
  background: surface, color: text, outline: 'none', minWidth: 0,
});

// ── Small building blocks shared by the chart/list cards below ──────────────

function SectionCard(props: { palette: Palette; children: React.ReactNode }): React.ReactElement {
  const { palette, children } = props;
  return (
    <div style={{
      background: palette.surface, borderRadius: 18, border: `1px solid ${palette.border}`,
      padding: '14px 16px', boxShadow: clay.shadowRaised,
      backdropFilter: 'blur(18px) saturate(1.3)', WebkitBackdropFilter: 'blur(18px) saturate(1.3)',
      display: 'flex', flexDirection: 'column', minWidth: 0,
    }}>
      {children}
    </div>
  );
}

function EmptyNote(props: { palette: Palette; children: React.ReactNode }): React.ReactElement {
  return (
    <div style={{ padding: '26px 10px', textAlign: 'center', color: props.palette.muted, fontSize: 12.5 }}>
      {props.children}
    </div>
  );
}

/** Flat SVG donut — no charting library installed, and one isn't worth
 *  adding for three static segments. */
function DonutChart(props: {
  segments: Array<{ value: number; color: string }>;
  centerValue: string;
  centerLabel: string;
  palette: Palette;
  size?: number;
}): React.ReactElement {
  const { segments, centerValue, centerLabel, palette, size = 118 } = props;
  const strokeWidth = Math.round(size * 0.145);
  const r = (size - strokeWidth) / 2;
  const c = size / 2;
  const circumference = 2 * Math.PI * r;
  const total = segments.reduce((sum, s) => sum + s.value, 0);
  let drawn = 0;

  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ flexShrink: 0 }}>
      <g transform={`rotate(-90 ${c} ${c})`}>
        <circle cx={c} cy={c} r={r} fill="none" stroke={palette.border} strokeWidth={strokeWidth} />
        {total > 0 && segments.filter(s => s.value > 0).map((s, i) => {
          const dash = (s.value / total) * circumference;
          const el = (
            <circle
              key={i} cx={c} cy={c} r={r} fill="none" stroke={s.color} strokeWidth={strokeWidth}
              strokeDasharray={`${dash} ${circumference - dash}`} strokeDashoffset={-drawn}
            />
          );
          drawn += dash;
          return el;
        })}
      </g>
      <text x={c} y={c - 4} textAnchor="middle" dominantBaseline="middle" fill={palette.text} style={{ fontSize: size * 0.2, fontWeight: 800, fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif" }}>
        {centerValue}
      </text>
      <text x={c} y={c + size * 0.14} textAnchor="middle" dominantBaseline="middle" fill={palette.muted} style={{ fontSize: size * 0.082, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.04em' }}>
        {centerLabel}
      </text>
    </svg>
  );
}

function DonutLegendRow(props: { label: string; count: number; total: number; color: string; palette: Palette }): React.ReactElement {
  const { label, count, total, color, palette } = props;
  const pct = total > 0 ? Math.round((count / total) * 100) : 0;
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12.5 }}>
      <span style={{ width: 9, height: 9, borderRadius: '50%', background: color, flexShrink: 0 }} />
      <span style={{ flex: 1, color: palette.text, fontWeight: 600, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{label}</span>
      <span style={{ fontWeight: 800, color: palette.text }}>{count}</span>
      <span style={{ color: palette.muted, fontSize: 11, width: 34, textAlign: 'right' }}>{pct}%</span>
    </div>
  );
}

function MiniBarChart(props: { data: Array<{ label: string; value: number }>; palette: Palette }): React.ReactElement {
  const { data, palette } = props;
  const max = Math.max(1, ...data.map(d => d.value));
  return (
    <div style={{ display: 'flex', alignItems: 'flex-end', gap: 8, height: 128, padding: '4px 2px 0' }}>
      {data.map(d => (
        <div key={d.label} style={{ flex: '1 1 0', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6, minWidth: 0 }}>
          <div style={{ fontSize: 10.5, fontWeight: 700, color: palette.text, minHeight: 13 }}>{d.value > 0 ? d.value : ''}</div>
          <div style={{ width: '100%', maxWidth: 30, height: Math.max(3, (d.value / max) * 82), background: clay.accentGradient, borderRadius: 5 }} />
          <div style={{ fontSize: 10.5, fontWeight: 700, color: palette.muted }}>{d.label}</div>
        </div>
      ))}
    </div>
  );
}

// ── Document Activity — real monthly volume from the Documents module's
// own facets (GET /api/dashboard/documents), not an invented expiry figure
// the backend doesn't compute. ───────────────────────────────────────────────
function DocumentActivityTrend(props: { host: VesselEmail; siteKey: string; nonce: number; palette: Palette }): React.ReactElement {
  const { host, siteKey, nonce, palette } = props;
  const [months, setMonths] = React.useState<Array<{ month: string; count: number }> | null>(null);
  const [error, setError] = React.useState<string>('');

  React.useEffect(() => {
    let cancelled = false;
    const base = host._base();
    if (!base) return undefined;
    const params = new URLSearchParams();
    if (siteKey && siteKey !== 'all') params.set('site_key', siteKey);
    params.set('date_field', 'modified');
    params.set('page', '1');
    params.set('page_size', '5'); // facets cover the whole site-scoped set regardless of page size
    setError('');
    host._fetchJson(`${base}/api/dashboard/documents?${params.toString()}`)
      .then((res: any) => {
        if (cancelled) return;
        setMonths(Array.isArray(res?.facets?.months) ? res.facets.months : []);
      })
      .catch((err: any) => { if (!cancelled) setError(err?.message || 'Could not load activity'); });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [siteKey, nonce]);

  const last6 = React.useMemo(() => {
    const byMonth = new Map((months || []).map(m => [m.month, m.count]));
    const now = new Date();
    const out: Array<{ label: string; value: number }> = [];
    for (let i = 5; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      out.push({ label: MONTH_NAMES[d.getMonth()], value: byMonth.get(key) || 0 });
    }
    return out;
  }, [months]);

  return (
    <>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: 4 }}>
        <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: palette.text }}>Document Activity</h3>
      </div>
      <p style={{ margin: '0 0 10px', fontSize: 11.5, color: palette.muted, fontWeight: 600 }}>Files added or modified, last 6 months</p>
      {error ? (
        <EmptyNote palette={palette}>Couldn't load activity — {error}</EmptyNote>
      ) : months === null ? (
        <EmptyNote palette={palette}>Loading…</EmptyNote>
      ) : last6.every(m => m.value === 0) ? (
        <EmptyNote palette={palette}>No document activity recorded yet</EmptyNote>
      ) : (
        <MiniBarChart data={last6} palette={palette} />
      )}
    </>
  );
}

// ── Compact activity/attention row ───────────────────────────────────────────
function ListRow(props: { icon: string; iconBg: string; iconColor: string; title: string; sub: string; when?: string }): React.ReactElement {
  const { icon, iconBg, iconColor, title, sub, when } = props;
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
      <div style={{ width: 28, height: 28, borderRadius: 9, background: iconBg, color: iconColor, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
        <Icon iconName={icon} style={{ fontSize: 12.5 }} />
      </div>
      <div style={{ minWidth: 0, flex: 1 }}>
        <div style={{ fontSize: 12.5, fontWeight: 700, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{title}</div>
        <div style={{ fontSize: 11, opacity: 0.75, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{sub}</div>
      </div>
      {when && <div style={{ fontSize: 10.5, opacity: 0.6, flexShrink: 0, whiteSpace: 'nowrap' }}>{when}</div>}
    </div>
  );
}

// ── Page ─────────────────────────────────────────────────────────────────────
export function renderDashboard(host: VesselEmail): React.ReactElement {
  const { vessels, dashboardStats } = host.state;
  const cardSurface = 'var(--vdms-glass)';
  const cardBorder = 'var(--vdms-line)';
  const primaryText = 'var(--vdms-text)';
  const mutedText = 'var(--vdms-text-muted)';
  const now = new Date();
  const hour = now.getHours();
  const greeting = hour < 12 ? 'Good morning' : hour < 18 ? 'Good afternoon' : 'Good evening';
  const firstName = String(host.props.userDisplayName || '').trim().split(/\s+/)[0] || 'there';
  const unreadAlerts = host._unreadAlertCount();
  const dateLabel = now.toLocaleDateString(undefined, { weekday: 'long', day: '2-digit', month: 'short', year: 'numeric' });
  const palette: Palette = { surface: cardSurface, border: cardBorder, text: primaryText, muted: mutedText };
  const statsLoading = !dashboardStats;
  const siteFilter = host.state.dashboardSiteFilter || 'all';

  // SharePoint site filter for the dashboard. Same "managed sites" heuristic
  // as the Vessels page's site filter, so the two dropdowns list the same
  // sites — but this one drives its own dashboardSiteFilter state, kept
  // separate so it doesn't disturb the Vessels/Documents site selections.
  const documentSites = host.state.documentSites || [];
  const filteredDocumentSites = documentSites.filter(site => {
    const siteKey = String(site.site_key || site.site_id || site.sp_site_name || '').toLowerCase();
    const siteName = String(site.sp_site_name || site.site_key || site.site_id || '').toLowerCase();
    const activeKey = String(host.state.activeDocumentSite || '').toLowerCase();
    const isDefaultSite = siteKey === activeKey || site.is_primary === true;
    return (
      isDefaultSite ||
      siteKey.includes('dev') || siteKey.includes('communication') || siteKey.includes('local') ||
      siteKey.includes('docman') || siteKey.includes('nks') || siteKey.includes('external') ||
      siteName.includes('dev') || siteName.includes('communication') || siteName.includes('local') ||
      siteName.includes('docman') || siteName.includes('nks') || siteName.includes('external')
    );
  });
  const dashboardSiteOptions = Array.from(new Map(
    filteredDocumentSites.map(site => [String(site.site_key || site.site_id || site.sp_site_name || '').toLowerCase(), site])
  ).values()).map(site => ({
    key: String(site.site_key || site.site_id || site.sp_site_name || ''),
    label: site.sp_site_name || site.site_key || site.site_id || 'SharePoint Site',
  }));

  const totalFiles = dashboardStats ? (dashboardStats.total_files ?? dashboardStats.total_documents) : undefined;
  const siteRows: DashboardSiteSummary[] = dashboardStats?.sites || [];
  const totalVessels = dashboardStats?.total_vessels ?? (vessels || []).length;

  // Fleet status — real, from vessels registered in the DMS (GET /api/vessels,
  // already loaded app-wide). Unknown/missing status defaults to "Active",
  // the same convention used everywhere else this field is read.
  const vesselList: VesselRecord[] = vessels || [];
  let activeVesselCount = 0, maintenanceVesselCount = 0, otherVesselCount = 0;
  vesselList.forEach(v => {
    const s = v.status || 'Active';
    if (s === 'Active') activeVesselCount++;
    else if (s === 'In Maintenance') maintenanceVesselCount++;
    else otherVesselCount++;
  });

  // Requires attention — real signals only: sites the backend failed to scan,
  // and unread anomaly alerts (unrecognised vessel folders, files uploaded
  // outside the DMS tree, unexpected subfolders). No invented "expiring /
  // expired document" counts: the backend doesn't track document expiry.
  const siteIssues = siteRows.filter(s => !!s.error);
  const unreadAnomalies = (host.state.alertsList || [])
    .filter(a => ANOMALY_ALERT_TYPES.has(a.alert_type) && !a.read)
    .sort((a, b) => alertEpoch(b) - alertEpoch(a));
  const attentionItems: Array<{ key: string; icon: string; title: string; sub: string; when?: string }> = [
    ...siteIssues.map(s => ({
      key: `site:${s.site_key}`, icon: 'Warning',
      title: s.site_name || s.site_key, sub: s.error || 'Unable to load this site',
    })),
    ...unreadAnomalies.map(a => ({
      key: a.id, icon: iconForAlert(a),
      title: a.folder_name || a.vessel_name || humanize(a.alert_type),
      sub: humanize(a.alert_type), when: timeAgo(alertEpoch(a)),
    })),
  ];
  const attentionCount = attentionItems.length;
  const attentionPreview = attentionItems.slice(0, 5);

  const recentActivity = [...(host.state.alertsList || [])]
    .sort((a, b) => alertEpoch(b) - alertEpoch(a))
    .slice(0, 6);

  const refresh = (): void => {
    void host._loadDashboardStats(true).then(() => {
      refreshNonce += 1;
      host.forceUpdate();
    });
  };

  const kpiCard = (opts: { label: string; value: number | undefined; icon: string; tone?: 'neutral' | 'success' | 'warning'; sub?: React.ReactNode }): React.ReactElement => {
    const tone = opts.tone || 'neutral';
    const iconBg = tone === 'success' ? clay.pillActiveBg : tone === 'warning' ? clay.pillWarnBg : clay.accentSoft;
    const iconColor = tone === 'success' ? clay.pillActiveText : tone === 'warning' ? clay.pillWarnText : clay.accentDark;
    return (
      <div style={{ background: cardSurface, borderRadius: 16, padding: '12px 14px', border: `1px solid ${cardBorder}`, boxShadow: clay.shadowRaised, display: 'flex', flexDirection: 'column', gap: 8, backdropFilter: 'blur(18px) saturate(1.3)', WebkitBackdropFilter: 'blur(18px) saturate(1.3)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <span style={{ fontSize: 12.5, fontWeight: 700, color: mutedText, textTransform: 'uppercase', letterSpacing: '0.04em' }}>{opts.label}</span>
          <div style={{ width: 30, height: 30, borderRadius: 9, background: iconBg, color: iconColor, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
            <Icon iconName={opts.icon} style={{ fontSize: 14 }} />
          </div>
        </div>
        <div style={{ fontSize: 30, fontWeight: 800, fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif", letterSpacing: '-0.02em', color: primaryText, lineHeight: 1 }}>
          {statsLoading && opts.value === undefined ? <span style={{ fontSize: 18, color: mutedText, fontWeight: 600 }}>…</span> : formatNumber(opts.value)}
        </div>
        {opts.sub}
      </div>
    );
  };

  const sectionHeader = (title: string, right?: React.ReactNode): React.ReactElement => (
    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
      <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: primaryText }}>{title}</h3>
      {right}
    </div>
  );

  const viewAllAlertsLink = (show: boolean): React.ReactElement | undefined => show ? (
    <button type="button" onClick={() => { void host._goToView('alerts'); }} style={{ background: 'none', border: 'none', color: clay.accentDark, fontSize: 12, fontWeight: 700, cursor: 'pointer', padding: 0 }}>
      View all →
    </button>
  ) : undefined;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      {/* Header: page title + subtitle (left), site filter + refresh (right) */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: 12 }}>
        <div>
          <h1 style={{ margin: 0, fontSize: 'clamp(22px, 2.2vw, 28px)', fontWeight: 800, letterSpacing: '-0.02em', color: primaryText, fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif" }}>
            Dashboard
          </h1>
          <p style={{ margin: '3px 0 0', fontSize: 13, fontWeight: 600, color: mutedText }}>Vessel compliance and document overview</p>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
          {dashboardStats?.last_refreshed_epoch && (
            <span style={{ fontSize: 11.5, color: mutedText }}>Last refreshed {formatDateTime(dashboardStats.last_refreshed_epoch)}</span>
          )}
          <select
            value={siteFilter}
            onChange={e => host._handleDashboardSiteChange(e.target.value)}
            style={{ ...controlStyle(cardBorder, cardSurface, primaryText), padding: '8px 12px', minWidth: 200, boxShadow: clay.shadowRaised, cursor: 'pointer' }}
          >
            <option value="all">All SharePoint Sites</option>
            {dashboardSiteOptions.map(site => <option key={site.key} value={site.key}>{site.label}</option>)}
          </select>
          <button
            type="button"
            onClick={refresh}
            title="Re-count files, folders and vessels"
            style={{ padding: '8px 12px', borderRadius: 8, border: `1px solid ${cardBorder}`, background: cardSurface, color: primaryText, fontSize: 12.5, fontWeight: 600, cursor: 'pointer', boxShadow: clay.shadowRaised, display: 'inline-flex', alignItems: 'center', gap: 6 }}
          >
            <Icon iconName="Refresh" style={{ fontSize: 12 }} /> Refresh
          </button>
        </div>
      </div>

      {/* Slim marine banner — keeps the brand identity without dominating the page */}
      <div style={{ position: 'relative', borderRadius: 18, overflow: 'hidden', minHeight: 72, display: 'flex', alignItems: 'center', padding: '0 22px', border: `1px solid ${cardBorder}`, boxShadow: clay.shadowRaised }}>
        <div aria-hidden="true" style={{ position: 'absolute', inset: 0, backgroundImage: `url("${dashboardHeroShipImage}")`, backgroundSize: 'cover', backgroundPosition: 'center 55%', opacity: 0.18 }} />
        <div aria-hidden="true" style={{ position: 'absolute', inset: 0, background: `linear-gradient(100deg, ${cardSurface} 32%, transparent 90%)` }} />
        <div style={{ position: 'relative', zIndex: 1 }}>
          <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.1em', textTransform: 'uppercase', color: clay.accent }}>{dateLabel}</div>
          <div style={{ fontSize: 14.5, fontWeight: 700, color: primaryText, marginTop: 2 }}>
            {greeting}, {firstName} — {unreadAlerts > 0 ? `${unreadAlerts} alert${unreadAlerts === 1 ? '' : 's'} need${unreadAlerts === 1 ? 's' : ''} attention.` : 'the fleet is all caught up.'}
          </div>
        </div>
      </div>

      {/* KPI row */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(190px, 1fr))', gap: 10 }}>
        {kpiCard({ label: 'Total Vessels', value: totalVessels, icon: 'Ferry' })}
        {kpiCard({
          label: 'Active Vessels', value: activeVesselCount, icon: 'CheckMark', tone: 'success',
          sub: vesselList.length > 0 ? <span style={{ fontSize: 11, color: mutedText, fontWeight: 600 }}>of {vesselList.length} in DMS</span> : undefined,
        })}
        {kpiCard({ label: 'Total Documents', value: totalFiles, icon: 'Page' })}
        {kpiCard({
          label: 'Needs Attention', value: attentionCount, icon: 'Warning', tone: attentionCount > 0 ? 'warning' : 'success',
          sub: viewAllAlertsLink(attentionCount > 0),
        })}
      </div>

      {/* Fleet status + document activity */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))', gap: 10 }}>
        <SectionCard palette={palette}>
          {sectionHeader('Fleet Status', <span style={{ fontSize: 11, fontWeight: 700, color: clay.accentDark, background: clay.accentSoft, borderRadius: 12, padding: '1px 8px' }}>{vesselList.length} vessels</span>)}
          {vesselList.length === 0 ? (
            <EmptyNote palette={palette}>No vessel records yet</EmptyNote>
          ) : (
            <div style={{ display: 'flex', gap: 20, alignItems: 'center', flexWrap: 'wrap' }}>
              <DonutChart
                palette={palette}
                centerValue={String(vesselList.length)}
                centerLabel="Vessels"
                segments={[
                  { value: activeVesselCount, color: clay.pillActiveText },
                  { value: maintenanceVesselCount, color: clay.pillWarnText },
                  { value: otherVesselCount, color: clay.pillDangerText },
                ]}
              />
              <div style={{ display: 'flex', flexDirection: 'column', gap: 9, flex: 1, minWidth: 150 }}>
                <DonutLegendRow label="Active" count={activeVesselCount} total={vesselList.length} color={clay.pillActiveText} palette={palette} />
                <DonutLegendRow label="In Maintenance" count={maintenanceVesselCount} total={vesselList.length} color={clay.pillWarnText} palette={palette} />
                <DonutLegendRow label="Inactive / Other" count={otherVesselCount} total={vesselList.length} color={clay.pillDangerText} palette={palette} />
              </div>
            </div>
          )}
          <div style={{ fontSize: 11, color: mutedText, marginTop: 12 }}>Based on vessels registered in the DMS.</div>
        </SectionCard>

        <SectionCard palette={palette}>
          <DocumentActivityTrend host={host} siteKey={siteFilter} nonce={refreshNonce} palette={palette} />
        </SectionCard>
      </div>

      {/* Requires attention + recent activity */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))', gap: 10 }}>
        <SectionCard palette={palette}>
          {sectionHeader('Requires Attention', viewAllAlertsLink(attentionCount > 0))}
          {attentionPreview.length === 0 ? (
            <EmptyNote palette={palette}>Nothing needs attention right now</EmptyNote>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12, color: primaryText }}>
              {attentionPreview.map(item => (
                <ListRow key={item.key} icon={item.icon} iconBg={clay.pillWarnBg} iconColor={clay.pillWarnText} title={item.title} sub={item.sub} when={item.when} />
              ))}
            </div>
          )}
        </SectionCard>

        <SectionCard palette={palette}>
          {sectionHeader('Recent Activity', viewAllAlertsLink(recentActivity.length > 0))}
          {recentActivity.length === 0 ? (
            <EmptyNote palette={palette}>No recent activity yet</EmptyNote>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12, color: primaryText }}>
              {recentActivity.map(a => (
                <ListRow
                  key={a.id}
                  icon={iconForAlert(a)}
                  iconBg={clay.accentSoft}
                  iconColor={clay.accentDark}
                  title={a.folder_name || a.vessel_name || humanize(a.alert_type)}
                  sub={humanize(a.alert_type)}
                  when={timeAgo(alertEpoch(a))}
                />
              ))}
            </div>
          )}
        </SectionCard>
      </div>
    </div>
  );
}
