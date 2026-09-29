import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';
import type { VesselRecord } from '../types/rows';

// Home Vessels panel: rows per page (see the "Pagination footer" below).
const VESSELS_PAGE_SIZE = 10;

// ─────────────────────────────────────────────────────────────────────────────
// Home / Dashboard
//
//   • Counters: Total Files, Total Folders, Total Sites, Total Vessels
//     (GET /api/dashboard/stats — backend pages through each site's whole
//     library, so these are real totals, not the first 500 items).
//   • Sites: one row per SharePoint site with its own file / folder counts.
//   • Vessels: one row per vessel (name, hull no., ship type, IMO, shipyard,
//     which site it came from, status), scoped by the same SharePoint Site
//     filter as everything else on Home, via GET /api/vessels?site_key=…
//     (the same endpoint the Vessels module itself uses). When "All
//     SharePoint Sites" is selected, a vessel that exists on more than one
//     site can legitimately appear more than once — the "From Site" column
//     is what tells those rows apart, there is no dedup here.
//   • Documents: browsed and paged right here on Home (no redirect to the
//     Documents module) via GET /api/dashboard/documents, with filters for
//     site, vessel, file type, person, text, and date — day-wise (today,
//     yesterday, a specific day), month-wise (this/last month, any month that
//     has files) and ranges (last 7/30/90 days, this/last year, custom) — on
//     either the modified or created date, plus group-by Day / Month.
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
  /** Live vessel count for this site — root-level (or one-level-nested)
   * folders whose name matches a Term Store vessel term, computed as part
   * of the same scan as files/folders (see backend
   * _term_store_vessel_folder_count). This is what "Total Vessels" sums. */
  vessels?: number;
  /** The official Term Store labels behind `vessels` — real, current
   * SharePoint vessel folder names for this site, not read from the
   * vessels DB table. Used to build the Home Vessels section so it always
   * matches "Total Vessels" and never depends on a vessel's (sometimes
   * stale) provisioned_site_key. */
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

interface DocumentsPageResponse {
  items: RealDashboardDoc[];
  total: number;
  page: number;
  page_size: number;
  total_pages: number;
  facets: {
    months: Array<{ month: string; count: number }>;
    vessels: Array<{ name: string; count: number }>;
    file_types: Array<{ type: string; count: number }>;
    people: string[];
  };
}

type DatePreset =
  | 'all' | 'today' | 'yesterday' | 'day'
  | 'last7' | 'last30' | 'last90'
  | 'this_month' | 'last_month' | 'month'
  | 'this_year' | 'last_year' | 'custom';

type GroupBy = 'none' | 'day' | 'month';

interface DocFilters {
  q: string;
  vessel: string;
  fileType: string;
  person: string;
  dateField: 'modified' | 'created';
  preset: DatePreset;
  day: string;        // yyyy-mm-dd, for preset 'day'
  month: string;      // yyyy-mm, for preset 'month'
  from: string;       // yyyy-mm-dd, for preset 'custom'
  to: string;         // yyyy-mm-dd (inclusive), for preset 'custom'
  groupBy: GroupBy;
  sort: string;
  pageSize: number;
}

const DEFAULT_FILTERS: DocFilters = {
  q: '', vessel: 'all', fileType: 'all', person: 'all',
  dateField: 'modified', preset: 'all', day: '', month: '', from: '', to: '',
  groupBy: 'none', sort: 'newest', pageSize: 20,
};

// Kept across visits to Home in the same page load, so leaving Home and
// coming back doesn't reset the filters or the page the user was on.
let rememberedFilters: DocFilters = { ...DEFAULT_FILTERS };
let rememberedPage = 1;
let refreshNonce = 0;

const FILE_TYPE_LABELS: Record<string, string> = {
  pdf: 'PDF', word: 'Word', excel: 'Excel / CSV', powerpoint: 'PowerPoint', image: 'Images',
  drawing: 'Drawings (DWG/DXF)', email: 'Emails', archive: 'Archives', text: 'Text', other: 'Other',
};

const DATE_PRESET_LABELS: Array<{ value: DatePreset; label: string }> = [
  { value: 'all', label: 'Any time' },
  { value: 'today', label: 'Today' },
  { value: 'yesterday', label: 'Yesterday' },
  { value: 'day', label: 'Specific day…' },
  { value: 'last7', label: 'Last 7 days' },
  { value: 'last30', label: 'Last 30 days' },
  { value: 'last90', label: 'Last 90 days' },
  { value: 'this_month', label: 'This month' },
  { value: 'last_month', label: 'Last month' },
  { value: 'month', label: 'Specific month…' },
  { value: 'this_year', label: 'This year' },
  { value: 'last_year', label: 'Last year' },
  { value: 'custom', label: 'Custom range…' },
];

const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function monthLabel(ym: string): string {
  const [y, m] = ym.split('-').map(Number);
  return y && m ? `${MONTH_NAMES[m - 1]} ${y}` : ym;
}

function parseLocalDate(value: string): Date | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value || '');
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

/** [from, to) in epoch ms for a preset, in the browser's own time zone. */
function dateRangeFor(f: DocFilters): { from?: number; to?: number } {
  const now = new Date();
  const startOfDay = (d: Date): Date => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const addDays = (d: Date, n: number): Date => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
  const today = startOfDay(now);
  switch (f.preset) {
    case 'today': return { from: today.getTime(), to: addDays(today, 1).getTime() };
    case 'yesterday': return { from: addDays(today, -1).getTime(), to: today.getTime() };
    case 'day': {
      const d = parseLocalDate(f.day);
      return d ? { from: d.getTime(), to: addDays(d, 1).getTime() } : {};
    }
    case 'last7': return { from: addDays(today, -6).getTime(), to: addDays(today, 1).getTime() };
    case 'last30': return { from: addDays(today, -29).getTime(), to: addDays(today, 1).getTime() };
    case 'last90': return { from: addDays(today, -89).getTime(), to: addDays(today, 1).getTime() };
    case 'this_month': return {
      from: new Date(now.getFullYear(), now.getMonth(), 1).getTime(),
      to: new Date(now.getFullYear(), now.getMonth() + 1, 1).getTime(),
    };
    case 'last_month': return {
      from: new Date(now.getFullYear(), now.getMonth() - 1, 1).getTime(),
      to: new Date(now.getFullYear(), now.getMonth(), 1).getTime(),
    };
    case 'month': {
      const m = /^(\d{4})-(\d{2})$/.exec(f.month || '');
      if (!m) return {};
      const y = Number(m[1]);
      const mo = Number(m[2]) - 1;
      return { from: new Date(y, mo, 1).getTime(), to: new Date(y, mo + 1, 1).getTime() };
    }
    case 'this_year': return { from: new Date(now.getFullYear(), 0, 1).getTime(), to: new Date(now.getFullYear() + 1, 0, 1).getTime() };
    case 'last_year': return { from: new Date(now.getFullYear() - 1, 0, 1).getTime(), to: new Date(now.getFullYear(), 0, 1).getTime() };
    case 'custom': {
      const a = parseLocalDate(f.from);
      const b = parseLocalDate(f.to);
      return { from: a ? a.getTime() : undefined, to: b ? addDays(b, 1).getTime() : undefined };
    }
    default: return {};
  }
}

function getFileIcon(name: string): { icon: string; color: string } {
  const low = (name || '').toLowerCase();
  if (low.endsWith('.pdf')) return { icon: '📄', color: '#ef4444' };
  if (low.endsWith('.docx') || low.endsWith('.doc')) return { icon: '📝', color: '#2563eb' };
  if (low.endsWith('.xlsx') || low.endsWith('.xls') || low.endsWith('.csv')) return { icon: '📊', color: '#10b981' };
  if (low.endsWith('.pptx') || low.endsWith('.ppt')) return { icon: '📽️', color: '#ea580c' };
  if (low.endsWith('.dwg') || low.endsWith('.dxf')) return { icon: '📐', color: '#8b5cf6' };
  if (/\.(png|jpe?g|gif|bmp|tiff?|webp|heic)$/.test(low)) return { icon: '🖼️', color: '#f59e0b' };
  if (/\.(msg|eml)$/.test(low)) return { icon: '✉️', color: '#0284c7' };
  if (/\.(zip|rar|7z)$/.test(low)) return { icon: '🗜️', color: '#64748b' };
  return { icon: '📄', color: 'var(--vdms-text-muted)' };
}

function formatNumber(value: number | undefined | null): string {
  return typeof value === 'number' && isFinite(value) ? value.toLocaleString() : '—';
}

function formatDateTime(epoch?: number | null): string {
  if (!epoch) return '—';
  const d = new Date(epoch);
  return d.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' }) +
    ', ' + d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
}

function dayKey(epoch: number): string {
  const d = new Date(epoch);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function dayHeading(key: string): string {
  const d = parseLocalDate(key);
  if (!d) return key;
  const today = new Date();
  const t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  if (d.getTime() === t0) return 'Today';
  if (d.getTime() === t0 - 86400000) return 'Yesterday';
  return d.toLocaleDateString(undefined, { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' });
}

// Folder names that are never real vessels even when they slip into a
// site's Term Store vessel-term matching — same exclusion list the
// Documents/Vessels Graph-fallback discovery already applies (see
// VesselEmail.tsx's vessel-node filter: pool-* slots and the two "common"
// wrapper folders). Kept here too because the live per-site scan behind
// the Home Vessels section (below) matches purely against the Term Store,
// with no name-shape check of its own.
function _isKnownNonVesselFolderName(name: string): boolean {
  const n = (name || '').trim().toLowerCase();
  return /^pool-/i.test(n) || n === 'common for all ships' || n === 'common (not ship specific)';
}

// ── Shared small styles ──────────────────────────────────────────────────────
const controlStyle = (border: string, surface: string, text: string): React.CSSProperties => ({
  padding: '7px 10px', borderRadius: 8, border: `1px solid ${border}`, fontSize: 12,
  background: surface, color: text, outline: 'none', minWidth: 0,
});

const thStyle: React.CSSProperties = {
  padding: '9px 12px', textAlign: 'left', fontSize: 11, fontWeight: 700,
  textTransform: 'uppercase', letterSpacing: '0.04em', color: 'var(--vdms-text-muted)', whiteSpace: 'nowrap',
};

const tdStyle: React.CSSProperties = { padding: '10px 12px', verticalAlign: 'top', fontSize: 13 };

interface Palette { surface: string; border: string; text: string; muted: string; }

// ── Documents panel (paged, filtered) ────────────────────────────────────────
function DashboardDocumentsPanel(props: {
  host: VesselEmail;
  siteKey: string;
  nonce: number;
  statsEpoch?: number | null;
  palette: Palette;
}): React.ReactElement {
  const { host, siteKey, nonce, statsEpoch, palette } = props;
  const [filters, setFiltersState] = React.useState<DocFilters>(rememberedFilters);
  const [page, setPageState] = React.useState<number>(rememberedPage);
  const [qInput, setQInput] = React.useState<string>(rememberedFilters.q);
  const [data, setData] = React.useState<DocumentsPageResponse | null>(null);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');
  const [retry, setRetry] = React.useState<number>(0);

  const setFilters = (patch: Partial<DocFilters>): void => {
    const next = { ...filters, ...patch };
    rememberedFilters = next;
    rememberedPage = 1;
    setFiltersState(next);
    setPageState(1);
  };
  const setPage = (p: number): void => {
    rememberedPage = p;
    setPageState(p);
  };

  // Debounce the search box so typing doesn't fire a request per key.
  React.useEffect(() => {
    if (qInput === filters.q) return undefined;
    const t = window.setTimeout(() => setFilters({ q: qInput }), 350);
    return () => window.clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [qInput]);

  // A different site on the Home site switcher starts from page 1.
  const lastSite = React.useRef(siteKey);
  React.useEffect(() => {
    if (lastSite.current !== siteKey) {
      lastSite.current = siteKey;
      setPage(1);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [siteKey]);

  React.useEffect(() => {
    let cancelled = false;
    const base = host._base();
    if (!base) return undefined;
    const range = dateRangeFor(filters);
    const params = new URLSearchParams();
    if (siteKey && siteKey !== 'all') params.set('site_key', siteKey);
    if (filters.q.trim()) params.set('q', filters.q.trim());
    if (filters.vessel !== 'all') params.set('vessel', filters.vessel);
    if (filters.fileType !== 'all') params.set('file_type', filters.fileType);
    if (filters.person !== 'all') params.set('person', filters.person);
    params.set('date_field', filters.dateField);
    if (range.from !== undefined) params.set('from_ms', String(range.from));
    if (range.to !== undefined) params.set('to_ms', String(range.to));
    params.set('tz_offset_min', String(new Date().getTimezoneOffset()));
    params.set('sort', filters.sort);
    params.set('page', String(page));
    params.set('page_size', String(filters.pageSize));
    setLoading(true);
    setError('');
    host._fetchJson(`${base}/api/dashboard/documents?${params.toString()}`)
      .then((res: any) => {
        if (cancelled) return;
        if (!res || !Array.isArray(res.items)) throw new Error('Unexpected response');
        setData(res as DocumentsPageResponse);
        if (res.page && res.page !== page) setPage(res.page);
      })
      .catch((err: any) => {
        if (cancelled) return;
        setError(err?.message || 'Could not load documents');
      })
      .then(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [siteKey, filters, page, nonce, retry, statsEpoch]);

  const { surface, border, text, muted } = palette;
  const ctl = controlStyle(border, surface, text);
  const facets = data?.facets;
  const items = data?.items || [];
  const total = data?.total || 0;
  const totalPages = data?.total_pages || 1;
  const startIdx = total === 0 ? 0 : (page - 1) * filters.pageSize + 1;
  const endIdx = Math.min(total, page * filters.pageSize);
  const epochOf = (d: RealDashboardDoc): number =>
    (filters.dateField === 'created' ? (d.createdEpoch || d.modifiedEpoch) : d.modifiedEpoch) || 0;
  const activeFilterCount = [
    filters.q, filters.vessel !== 'all', filters.fileType !== 'all', filters.person !== 'all', filters.preset !== 'all',
  ].filter(Boolean).length;

  // Group rows on the current page into Day / Month sections.
  const groups: Array<{ key: string; label: string; rows: RealDashboardDoc[] }> = [];
  if (filters.groupBy === 'none') {
    groups.push({ key: 'all', label: '', rows: items });
  } else {
    items.forEach(doc => {
      const ep = epochOf(doc);
      const key = ep ? (filters.groupBy === 'day' ? dayKey(ep) : dayKey(ep).slice(0, 7)) : 'unknown';
      let g = groups.find(x => x.key === key);
      if (!g) {
        g = { key, label: key === 'unknown' ? 'No date' : (filters.groupBy === 'day' ? dayHeading(key) : monthLabel(key)), rows: [] };
        groups.push(g);
      }
      g.rows.push(doc);
    });
  }

  const pageButtons: number[] = [];
  {
    const windowStart = Math.max(1, Math.min(page - 2, totalPages - 4));
    for (let p = windowStart; p <= Math.min(totalPages, windowStart + 4); p++) pageButtons.push(p);
  }
  const pageBtn = (active: boolean, disabled = false): React.CSSProperties => ({
    minWidth: 32, height: 30, padding: '0 10px', borderRadius: 8, fontSize: 12, fontWeight: 700,
    border: `1px solid ${active ? clay.accentDark : border}`,
    background: active ? clay.accent : surface, color: active ? '#fff' : (disabled ? 'var(--vdms-text-faint)' : text),
    cursor: disabled ? 'not-allowed' : 'pointer',
  });

  const openDoc = (doc: RealDashboardDoc): void => {
    if (doc.webUrl) window.open(doc.webUrl, '_blank', 'noopener');
  };

  return (
    <div style={{ background: surface, borderRadius: 16, border: `1px solid ${border}`, padding: 22, boxShadow: clay.shadowRaised }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12, flexWrap: 'wrap', marginBottom: 14 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: text }}>Documents</h3>
          <span style={{ fontSize: 11, fontWeight: 700, color: clay.accentDark, background: clay.accentSoft, borderRadius: 12, padding: '1px 8px' }}>
            {loading && !data ? '…' : `${formatNumber(total)} ${activeFilterCount > 0 ? 'matching' : 'files'}`}
          </span>
          {loading && data && <span style={{ fontSize: 11, color: muted }}>Updating…</span>}
        </div>
        {activeFilterCount > 0 && (
          <button
            type="button"
            onClick={() => { setQInput(''); setFilters({ ...DEFAULT_FILTERS, groupBy: filters.groupBy, sort: filters.sort, pageSize: filters.pageSize }); }}
            style={{ background: 'none', border: 'none', color: clay.accentDark, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}
          >
            ✕ Clear filters ({activeFilterCount})
          </button>
        )}
      </div>

      {/* Filters */}
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 14 }}>
        <input
          value={qInput}
          onChange={e => setQInput(e.target.value)}
          placeholder="Search file name (partial ok), folder, vessel, person…"
          aria-label="Search documents"
          style={{ ...ctl, flex: '1 1 240px' }}
        />
        <select aria-label="Date" value={filters.preset} onChange={e => setFilters({ preset: e.target.value as DatePreset })} style={ctl}>
          {DATE_PRESET_LABELS.map(p => <option key={p.value} value={p.value}>{p.label}</option>)}
        </select>
        {filters.preset === 'day' && (
          <input type="date" aria-label="Day" value={filters.day} onChange={e => setFilters({ day: e.target.value })} style={ctl} />
        )}
        {filters.preset === 'month' && (
          <select aria-label="Month" value={filters.month} onChange={e => setFilters({ month: e.target.value })} style={ctl}>
            <option value="">Select month</option>
            {(facets?.months || []).map(m => (
              <option key={m.month} value={m.month}>{monthLabel(m.month)} ({m.count})</option>
            ))}
          </select>
        )}
        {filters.preset === 'custom' && (
          <>
            <input type="date" aria-label="From" value={filters.from} onChange={e => setFilters({ from: e.target.value })} style={ctl} />
            <span style={{ alignSelf: 'center', fontSize: 12, color: muted }}>to</span>
            <input type="date" aria-label="To" value={filters.to} onChange={e => setFilters({ to: e.target.value })} style={ctl} />
          </>
        )}
        <select aria-label="Date field" value={filters.dateField} onChange={e => setFilters({ dateField: e.target.value as 'modified' | 'created' })} style={ctl}>
          <option value="modified">by Modified date</option>
          <option value="created">by Uploaded date</option>
        </select>
        <select aria-label="Vessel" value={filters.vessel} onChange={e => setFilters({ vessel: e.target.value })} style={{ ...ctl, maxWidth: 190 }}>
          <option value="all">All vessels</option>
          {(facets?.vessels || []).map(v => <option key={v.name} value={v.name}>{v.name} ({v.count})</option>)}
        </select>
        <select aria-label="File type" value={filters.fileType} onChange={e => setFilters({ fileType: e.target.value })} style={ctl}>
          <option value="all">All file types</option>
          {(facets?.file_types || []).map(t => <option key={t.type} value={t.type}>{FILE_TYPE_LABELS[t.type] || t.type} ({t.count})</option>)}
        </select>
        <select aria-label="Person" value={filters.person} onChange={e => setFilters({ person: e.target.value })} style={{ ...ctl, maxWidth: 190 }}>
          <option value="all">Anyone</option>
          {(facets?.people || []).map(p => <option key={p} value={p}>{p}</option>)}
        </select>
        <select aria-label="Group by" value={filters.groupBy} onChange={e => setFilters({ groupBy: e.target.value as GroupBy })} style={ctl}>
          <option value="none">No grouping</option>
          <option value="day">Group by day</option>
          <option value="month">Group by month</option>
        </select>
        <select aria-label="Sort" value={filters.sort} onChange={e => setFilters({ sort: e.target.value })} style={ctl}>
          <option value="newest">Newest first</option>
          <option value="oldest">Oldest first</option>
          <option value="created_newest">Recently uploaded</option>
          <option value="name_az">Name A–Z</option>
          <option value="name_za">Name Z–A</option>
          <option value="size_desc">Largest first</option>
          <option value="size_asc">Smallest first</option>
        </select>
      </div>

      {error ? (
        <div style={{ padding: '32px 16px', textAlign: 'center', color: '#b91c1c', border: '1px dashed #fca5a5', borderRadius: 12, background: '#fef2f2' }}>
          <div style={{ fontWeight: 700 }}>Couldn't load documents</div>
          <div style={{ fontSize: 12, marginTop: 4 }}>{error}. If the backend was just updated, restart it so the new Home endpoint is available.</div>
          <button type="button" onClick={() => setRetry(r => r + 1)} style={{ marginTop: 10, border: '1px solid #fca5a5', background: '#fff', color: '#b91c1c', borderRadius: 8, padding: '5px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>↻ Retry</button>
        </div>
      ) : (
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 760 }}>
            <thead>
              <tr style={{ borderBottom: '2px solid var(--vdms-border)' }}>
                <th style={thStyle}>Name</th>
                <th style={thStyle}>Site</th>
                <th style={thStyle}>Vessel</th>
                <th style={thStyle}>{filters.dateField === 'created' ? 'Uploaded' : 'Modified'}</th>
                <th style={thStyle}>Size</th>
                <th style={{ ...thStyle, textAlign: 'right' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {!data && loading ? (
                <tr><td colSpan={6} style={{ ...tdStyle, padding: '36px 16px', textAlign: 'center', color: muted }}>Loading documents…</td></tr>
              ) : items.length === 0 ? (
                <tr>
                  <td colSpan={6} style={{ ...tdStyle, padding: '36px 16px', textAlign: 'center', color: 'var(--vdms-text-faint)' }}>
                    <div style={{ fontSize: 28, marginBottom: 6 }}>📂</div>
                    <div style={{ fontWeight: 600, color: 'var(--vdms-text-muted)' }}>
                      {activeFilterCount > 0 ? 'No documents match these filters' : 'No documents found'}
                    </div>
                  </td>
                </tr>
              ) : groups.map(group => (
                <React.Fragment key={group.key}>
                  {filters.groupBy !== 'none' && (
                    <tr>
                      <td colSpan={6} style={{ padding: '10px 12px 6px', fontSize: 12, fontWeight: 800, color: clay.accentDark, background: 'var(--vdms-surface-alt)' }}>
                        {group.label} <span style={{ fontWeight: 600, color: muted }}>· {group.rows.length} on this page</span>
                      </td>
                    </tr>
                  )}
                  {group.rows.map(doc => {
                    const { icon, color } = getFileIcon(doc.name);
                    const who = filters.dateField === 'created' ? doc.createdBy : doc.modifiedBy;
                    return (
                      <tr key={doc.id} style={{ borderBottom: '1px solid var(--vdms-border-soft)' }}>
                        <td style={{ ...tdStyle, maxWidth: 360 }}>
                          <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
                            <span style={{ fontSize: 16, color, flexShrink: 0 }}>{icon}</span>
                            <div style={{ minWidth: 0 }}>
                              <div
                                onClick={() => openDoc(doc)}
                                title={doc.name}
                                style={{ fontWeight: 600, color: doc.webUrl ? '#0284c7' : 'var(--vdms-text)', cursor: doc.webUrl ? 'pointer' : 'default', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                              >
                                {doc.name}
                              </div>
                              {doc.subFolderPath && (
                                <div title={doc.subFolderPath} style={{ fontSize: 11, color: muted, marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                  {doc.subFolderPath.split(/\s*>\s*/).join(' › ')}
                                </div>
                              )}
                            </div>
                          </div>
                        </td>
                        <td style={{ ...tdStyle, whiteSpace: 'nowrap', color: 'var(--vdms-text-secondary)' }}>{doc.siteName || doc.site || '—'}</td>
                        <td style={{ ...tdStyle, whiteSpace: 'nowrap' }}>
                          {doc.vessel && doc.vessel !== 'Not Listed' ? (
                            <span style={{ background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text)', borderRadius: 6, padding: '2px 8px', fontSize: 11, fontWeight: 600 }}>{doc.vessel}</span>
                          ) : (
                            <span style={{ fontSize: 11, color: 'var(--vdms-text-faint)' }}>Not listed</span>
                          )}
                        </td>
                        <td style={{ ...tdStyle, whiteSpace: 'nowrap', color: 'var(--vdms-text-muted)', fontSize: 12 }}>
                          <div>{formatDateTime(epochOf(doc))}</div>
                          {who && <div style={{ fontSize: 11, color: 'var(--vdms-text-faint)' }}>{who}</div>}
                        </td>
                        <td style={{ ...tdStyle, whiteSpace: 'nowrap', color: 'var(--vdms-text-muted)', fontSize: 12 }}>{doc.fileSize || '—'}</td>
                        <td style={{ ...tdStyle, textAlign: 'right' }}>
                          <button
                            type="button"
                            disabled={!doc.webUrl}
                            onClick={() => openDoc(doc)}
                            style={{ background: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: 6, padding: '4px 10px', fontSize: 12, fontWeight: 600, color: '#0284c7', cursor: doc.webUrl ? 'pointer' : 'not-allowed', opacity: doc.webUrl ? 1 : 0.5 }}
                          >
                            Open
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Pagination */}
      {!error && total > 0 && (
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginTop: 14 }}>
          <div style={{ fontSize: 12, color: muted }}>
            Showing <strong style={{ color: text }}>{formatNumber(startIdx)}–{formatNumber(endIdx)}</strong> of <strong style={{ color: text }}>{formatNumber(total)}</strong>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
            <button type="button" disabled={page <= 1} onClick={() => setPage(1)} style={pageBtn(false, page <= 1)} aria-label="First page">«</button>
            <button type="button" disabled={page <= 1} onClick={() => setPage(page - 1)} style={pageBtn(false, page <= 1)}>‹ Prev</button>
            {pageButtons.map(p => (
              <button key={p} type="button" onClick={() => setPage(p)} style={pageBtn(p === page)} aria-current={p === page ? 'page' : undefined}>{p}</button>
            ))}
            <button type="button" disabled={page >= totalPages} onClick={() => setPage(page + 1)} style={pageBtn(false, page >= totalPages)}>Next ›</button>
            <button type="button" disabled={page >= totalPages} onClick={() => setPage(totalPages)} style={pageBtn(false, page >= totalPages)} aria-label="Last page">»</button>
            <select
              aria-label="Rows per page"
              value={filters.pageSize}
              onChange={e => setFilters({ pageSize: Number(e.target.value) })}
              style={{ ...ctl, marginLeft: 6 }}
            >
              {[10, 20, 50, 100].map(n => <option key={n} value={n}>{n} / page</option>)}
            </select>
          </div>
        </div>
      )}
    </div>
  );
}

interface DashboardVesselRow {
  name: string;
  siteKey: string;
  meta?: VesselRecord;
}

// ── Vessels panel (site-wise, mirrors the Documents panel) ──────────────────
//
// Source of truth: dashboardStats.sites[].vessel_names — the SAME live,
// per-site Term Store folder scan that "Total Vessels" is summed from (see
// backend _term_store_vessel_folder_count). This is deliberately NOT built
// from GET /api/vessels (the vessels DB table): a vessel's stored
// provisioned_site_key can be stale (e.g. left over as "local" from early
// testing, long after the vessel's real folder moved / was reconciled to a
// production site), and trusting it both mis-attributes vessels to the
// wrong site AND silently drops real, currently-existing vessels whenever
// their name collides with an unrelated, wrongly-tagged DB row. The
// vessels table is still queried once (unscoped) purely to enrich a
// matching name with hull no. / IMO / ship type / shipyard when a DMS
// record exists for it — never to decide which vessels or sites exist.
// The Vessel Name cell shows only the name; whether it has a DMS record
// is a Status-column concern, not something to clutter the name with.
function DashboardVesselsPanel(props: {
  host: VesselEmail;
  siteKey: string;
  nonce: number;
  dashboardStats: DashboardStats | null;
  palette: Palette;
}): React.ReactElement {
  const { host, siteKey, nonce, dashboardStats, palette } = props;
  const [metaByName, setMetaByName] = React.useState<Record<string, VesselRecord>>({});
  const [metaError, setMetaError] = React.useState<string>('');
  const [search, setSearch] = React.useState<string>('');
  const [retry, setRetry] = React.useState<number>(0);
  const [page, setPage] = React.useState<number>(0);

  // Jump back to page 1 whenever the site filter changes or the vessel list
  // is refreshed, so pagination never gets stuck past the end of a shorter,
  // re-filtered list.
  React.useEffect(() => { setPage(0); }, [siteKey, nonce]);

  React.useEffect(() => {
    let cancelled = false;
    const base = host._base();
    if (!base) return undefined;
    host._fetchJson(`${base}/api/vessels`)
      .then((res: any) => {
        if (cancelled || !Array.isArray(res)) return;
        const map: Record<string, VesselRecord> = {};
        res.forEach((v: any) => {
          const key = String(v.name || '').trim().toLowerCase();
          // Prefer a real DMS record over a SharePoint-only discovered one
          // if both happen to be present under the same name.
          if (key && (!map[key] || (map[key].source === 'sharepoint' && v.source !== 'sharepoint'))) {
            map[key] = v;
          }
        });
        setMetaByName(map);
      })
      .catch((err: any) => { if (!cancelled) setMetaError(err?.message || ''); });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nonce, retry]);

  const { surface, border, text, muted } = palette;
  const ctl = controlStyle(border, surface, text);
  const sites = dashboardStats?.sites || [];
  const loading = !dashboardStats;

  const siteNameFor = (key?: string | null): string => {
    if (!key) return '—';
    const match = (host.state.documentSites || []).find(s =>
      String(s.site_key || '').toLowerCase() === key.toLowerCase() ||
      String(s.sp_site_name || '').toLowerCase() === key.toLowerCase()
    );
    return match?.sp_site_name || match?.site_key || key;
  };

  const rows: DashboardVesselRow[] = [];
  sites.forEach(site => {
    if (siteKey !== 'all' && String(site.site_key || '').toLowerCase() !== siteKey.toLowerCase()) return;
    (site.vessel_names || []).forEach(name => {
      if (_isKnownNonVesselFolderName(name)) return;
      rows.push({ name, siteKey: site.site_key, meta: metaByName[name.trim().toLowerCase()] });
    });
  });

  const q = search.trim().toLowerCase();
  const filtered = q
    ? rows.filter(r =>
        r.name.toLowerCase().includes(q) ||
        (r.meta?.imo || '').toLowerCase().includes(q) ||
        (r.meta?.hull_number || '').toLowerCase().includes(q) ||
        (r.meta?.vessel_type || '').toLowerCase().includes(q) ||
        (r.meta?.shipyard || '').toLowerCase().includes(q)
      )
    : rows;

  const totalPages = Math.max(1, Math.ceil(filtered.length / VESSELS_PAGE_SIZE));
  const safePage = Math.min(page, totalPages - 1);
  const pageRows = filtered.slice(safePage * VESSELS_PAGE_SIZE, (safePage + 1) * VESSELS_PAGE_SIZE);

  const statusPill = (row: DashboardVesselRow): React.ReactElement => {
    const hasDbRecord = !!row.meta && row.meta.source !== 'sharepoint';
    if (!hasDbRecord) {
      return (
        <span
          title="A real SharePoint folder matching this site's Term Store vessel list — no DMS vessel record for it yet"
          style={{ display: 'inline-block', borderRadius: 12, padding: '2px 10px', fontSize: 11, fontWeight: 700, background: clay.pillWarnBg, color: clay.pillWarnText, boxShadow: clay.pillWarnShadow, whiteSpace: 'nowrap' }}
        >
          Not in DMS yet
        </span>
      );
    }
    const s = row.meta?.status || 'Active';
    const bg = s === 'Active' ? clay.pillActiveBg : s === 'In Maintenance' ? clay.pillWarnBg : clay.pillDangerBg;
    const fg = s === 'Active' ? clay.pillActiveText : s === 'In Maintenance' ? clay.pillWarnText : clay.pillDangerText;
    const shadow = s === 'Active' ? clay.pillActiveShadow : s === 'In Maintenance' ? clay.pillWarnShadow : clay.pillDangerShadow;
    return (
      <span style={{ display: 'inline-block', borderRadius: 12, padding: '2px 10px', fontSize: 11, fontWeight: 700, background: bg, color: fg, boxShadow: shadow, whiteSpace: 'nowrap' }}>
        {s === 'Active' ? '● Active' : s}
      </span>
    );
  };

  return (
    <div style={{ background: surface, borderRadius: 16, border: `1px solid ${border}`, padding: 22, boxShadow: clay.shadowRaised }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12, flexWrap: 'wrap', marginBottom: 12 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: text }}>Vessels</h3>
          <span style={{ fontSize: 11, fontWeight: 700, color: clay.accentDark, background: clay.accentSoft, borderRadius: 12, padding: '1px 8px' }}>
            {loading ? '…' : `${formatNumber(filtered.length)} ${q ? 'matching' : 'vessels'}`}
          </span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
          <input
            value={search}
            onChange={e => { setSearch(e.target.value); setPage(0); }}
            placeholder="Search vessel, IMO, hull no…"
            aria-label="Search vessels"
            style={{ ...ctl, minWidth: 220 }}
          />
          <button
            type="button"
            onClick={() => host._goToView('vessels')}
            style={{ background: 'none', border: 'none', color: clay.accentDark, fontSize: 12, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}
          >
            Manage vessels →
          </button>
        </div>
      </div>

      <div style={{ fontSize: 11, color: muted, marginBottom: 10 }}>
        Live count from each site's actual SharePoint folders — the same figure as the <strong style={{ color: text }}>Total Vessels</strong> tile above, not the vessels database.
        {siteKey === 'all' && rows.length > 0 && (
          <> The same vessel can appear more than once if its folder genuinely exists on more than one site — check <strong style={{ color: text }}>From Site</strong> to tell those rows apart.</>
        )}
      </div>

      {metaError && (
        <div style={{ fontSize: 11, color: '#b91c1c', marginBottom: 8 }}>
          Vessel details (hull no. / IMO / type) couldn't be loaded — names and sites below are still live. {metaError}{' '}
          <button type="button" onClick={() => { setMetaError(''); setRetry(r => r + 1); }} style={{ background: 'none', border: 'none', color: '#b91c1c', fontWeight: 700, textDecoration: 'underline', cursor: 'pointer', fontSize: 11, padding: 0 }}>Retry</button>
        </div>
      )}

      <div style={{ overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 880 }}>
          <thead>
            <tr style={{ borderBottom: '2px solid var(--vdms-border)' }}>
              <th style={thStyle}>Vessel Name</th>
              <th style={thStyle}>Hull Number</th>
              <th style={thStyle}>Ship Type</th>
              <th style={thStyle}>IMO Number</th>
              <th style={thStyle}>Shipyard</th>
              <th style={thStyle}>From Site</th>
              <th style={thStyle}>Status</th>
              <th style={{ ...thStyle, textAlign: 'right' }}>Action</th>
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <tr><td colSpan={8} style={{ ...tdStyle, textAlign: 'center', color: muted, padding: '24px 12px' }}>Counting vessels from SharePoint…</td></tr>
            ) : filtered.length === 0 ? (
              <tr>
                <td colSpan={8} style={{ ...tdStyle, padding: '36px 16px', textAlign: 'center', color: 'var(--vdms-text-faint)' }}>
                  <div style={{ fontSize: 28, marginBottom: 6 }}>⚓</div>
                  <div style={{ fontWeight: 600, color: 'var(--vdms-text-muted)' }}>
                    {q ? 'No vessels match this search' : 'No vessels found for this site'}
                  </div>
                </td>
              </tr>
            ) : pageRows.map((row, idx) => (
              <tr key={`${row.siteKey}:${row.name}:${idx}`} style={{ borderBottom: '1px solid var(--vdms-border-soft)' }}>
                <td style={{ ...tdStyle, fontWeight: 700, color: text, whiteSpace: 'nowrap' }}>⚓ {row.name}</td>
                <td style={{ ...tdStyle, color: 'var(--vdms-text-secondary)' }}>{row.meta?.hull_number || '—'}</td>
                <td style={{ ...tdStyle, color: 'var(--vdms-text-secondary)' }}>{row.meta?.vessel_type || '—'}</td>
                <td style={{ ...tdStyle, color: 'var(--vdms-text-secondary)', whiteSpace: 'nowrap' }}>{row.meta?.imo || '—'}</td>
                <td style={{ ...tdStyle, color: 'var(--vdms-text-secondary)' }}>{row.meta?.shipyard || '—'}</td>
                <td style={{ ...tdStyle, whiteSpace: 'nowrap' }}>
                  <span style={{ background: 'var(--vdms-surface-alt)', color: 'var(--vdms-text)', borderRadius: 6, padding: '2px 8px', fontSize: 11, fontWeight: 600 }}>
                    🌐 {siteNameFor(row.siteKey)}
                  </span>
                </td>
                <td style={{ ...tdStyle, whiteSpace: 'nowrap' }}>{statusPill(row)}</td>
                <td style={{ ...tdStyle, textAlign: 'right' }}>
                  <button
                    type="button"
                    onClick={() => host.setState({ vesselSiteFilter: row.siteKey || 'all' }, () => { void host._goToView('vessels'); })}
                    style={{ background: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: 6, padding: '4px 10px', fontSize: 12, fontWeight: 600, color: '#0284c7', cursor: 'pointer' }}
                  >
                    View
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {!loading && filtered.length > 0 && (
        <div style={{ padding: '10px 2px 0', color: muted, fontSize: 11, display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
          <span>
            Showing {safePage * VESSELS_PAGE_SIZE + 1}–{Math.min((safePage + 1) * VESSELS_PAGE_SIZE, filtered.length)} of {filtered.length} vessels
          </span>
          <div style={{ display: 'flex', gap: 3, alignItems: 'center' }}>
            <button
              type="button"
              onClick={() => setPage(Math.max(0, safePage - 1))}
              disabled={safePage === 0}
              style={{ border: `1px solid ${border}`, background: surface, borderRadius: 6, padding: '3px 8px', fontSize: 11, cursor: safePage === 0 ? 'not-allowed' : 'pointer', opacity: safePage === 0 ? 0.4 : 1 }}
            >‹</button>
            {Array.from({ length: totalPages }, (_, i) => (
              <button
                key={i}
                type="button"
                onClick={() => setPage(i)}
                style={{ border: `1px solid ${i === safePage ? 'transparent' : border}`, background: i === safePage ? clay.accentGradient : surface, color: i === safePage ? '#fff' : text, borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: i === safePage ? 700 : 400, cursor: 'pointer', minWidth: 26 }}
              >{i + 1}</button>
            ))}
            <button
              type="button"
              onClick={() => setPage(Math.min(totalPages - 1, safePage + 1))}
              disabled={safePage >= totalPages - 1}
              style={{ border: `1px solid ${border}`, background: surface, borderRadius: 6, padding: '3px 8px', fontSize: 11, cursor: safePage >= totalPages - 1 ? 'not-allowed' : 'pointer', opacity: safePage >= totalPages - 1 ? 0.4 : 1 }}
            >›</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Page ─────────────────────────────────────────────────────────────────────
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

export function renderDashboard(host: VesselEmail): React.ReactElement {
  const { vessels, loading, dashboardStats } = host.state;
  const isNight = host.state.themeMode === 'night';
  const cardSurface = isNight ? '#302219' : clay.surface;
  const cardBorder = isNight ? '#614331' : '#ead7c6';
  const primaryText = isNight ? '#f8eee6' : clay.text;
  const mutedText = isNight ? '#c7a58d' : clay.textMuted;
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
  const totalFolders = dashboardStats?.total_folders;
  const siteRows: DashboardSiteSummary[] = dashboardStats?.sites && dashboardStats.sites.length > 0
    ? dashboardStats.sites
    : [];
  const totalSites = dashboardStats?.total_sites ?? (siteRows.length || (siteFilter === 'all' ? dashboardSiteOptions.length : 1));
  const totalVessels = dashboardStats?.total_vessels ?? (vessels || []).length;

  const refresh = (): void => {
    void host._loadDashboardStats(true).then(() => {
      refreshNonce += 1;
      host.forceUpdate();
    });
  };

  const statCard = (label: string, value: number | undefined, icon: string, sub?: React.ReactNode): React.ReactElement => (
    <div style={{ background: cardSurface, borderRadius: 16, padding: 20, border: `1px solid ${cardBorder}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', boxShadow: clay.shadowRaised }}>
      <div>
        <div style={{ fontSize: 28, fontWeight: 800, color: primaryText, minHeight: 34 }}>
          {statsLoading || (loading && value === undefined) ? <span style={{ color: mutedText, fontSize: 22 }}>…</span> : formatNumber(value)}
        </div>
        <div style={{ fontSize: 13, color: mutedText, fontWeight: 600, marginTop: 2 }}>{label}</div>
        {sub}
      </div>
      <div style={{ width: 44, height: 44, borderRadius: '50%', background: clay.accentSoft, color: clay.accentDark, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>{icon}</div>
    </div>
  );

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      {/* SharePoint Site Filter — scopes every figure and list below. */}
      <div style={{ display: 'flex', justifyContent: 'flex-end', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
        {dashboardStats?.last_refreshed_epoch && (
          <span style={{ fontSize: 12, color: mutedText }}>
            Last refreshed {formatDateTime(dashboardStats.last_refreshed_epoch)}
          </span>
        )}
        <span style={{ fontSize: 13, color: mutedText, fontWeight: 600 }}>SharePoint Site:</span>
        <select
          value={siteFilter}
          onChange={e => host._handleDashboardSiteChange(e.target.value)}
          style={{
            padding: '9px 14px', borderRadius: 10, border: `1px solid ${cardBorder}`, fontSize: 13,
            background: cardSurface, outline: 'none', minWidth: 220, color: primaryText,
            boxShadow: clay.shadowRaised, cursor: 'pointer',
          }}
        >
          <option value="all">All SharePoint Sites</option>
          {dashboardSiteOptions.map(site => <option key={site.key} value={site.key}>{site.label}</option>)}
        </select>
        <button
          type="button"
          onClick={refresh}
          title="Re-count files and folders from SharePoint"
          style={{ padding: '9px 14px', borderRadius: 10, border: `1px solid ${cardBorder}`, background: cardSurface, color: primaryText, fontSize: 13, fontWeight: 600, cursor: 'pointer', boxShadow: clay.shadowRaised }}
        >
          ↻ Refresh
        </button>
      </div>

      {/* Counters */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 16 }}>
        {statCard('Total Files', totalFiles, '📄')}
        {statCard('Total Folders', totalFolders, '📁')}
        {statCard('Total Sites', totalSites, '🌐')}
        {statCard('Total Vessels', totalVessels, '⚓',
          <button onClick={() => host._goToView('vessels')} style={{ background: 'none', border: 'none', color: clay.accentDark, fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>Manage vessels →</button>
        )}
      </div>
      {/* Sites */}
      <div style={{ background: cardSurface, borderRadius: 16, border: `1px solid ${cardBorder}`, padding: 22, boxShadow: clay.shadowRaised }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: primaryText }}>SharePoint Sites</h3>
          {siteFilter !== 'all' && (
            <button type="button" onClick={() => host._handleDashboardSiteChange('all')} style={{ background: 'none', border: 'none', color: clay.accentDark, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Show all sites</button>
          )}
        </div>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 560 }}>
            <thead>
              <tr style={{ borderBottom: '2px solid var(--vdms-border)' }}>
                <th style={thStyle}>Site</th>
                <th style={{ ...thStyle, textAlign: 'right' }}>Files</th>
                <th style={{ ...thStyle, textAlign: 'right' }}>Folders</th>
                <th style={thStyle}>Last activity</th>
                <th style={{ ...thStyle, textAlign: 'right' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {statsLoading ? (
                <tr><td colSpan={5} style={{ ...tdStyle, textAlign: 'center', color: mutedText, padding: '24px 12px' }}>Counting files and folders…</td></tr>
              ) : siteRows.length === 0 ? (
                <tr><td colSpan={5} style={{ ...tdStyle, textAlign: 'center', color: 'var(--vdms-text-faint)', padding: '24px 12px' }}>No site details available yet — restart the backend to pick up the new Home counters.</td></tr>
              ) : siteRows.map(site => {
                const isSelected = siteFilter !== 'all' && (siteFilter.toLowerCase() === (site.site_key || '').toLowerCase());
                return (
                  <tr key={site.site_key + (site.drive_id || '')} style={{ borderBottom: '1px solid var(--vdms-border-soft)', background: isSelected ? 'var(--vdms-surface-alt)' : undefined }}>
                    <td style={{ ...tdStyle, fontWeight: 700, color: primaryText }}>
                      🌐 {site.site_name || site.site_key}
                      {site.truncated && <span style={{ marginLeft: 6, fontSize: 10, color: '#92400e' }}>(partial)</span>}
                      {site.stats_pending && (
                        <span
                          title="This site was just added — its file/folder counts are being scanned now and will fill in shortly."
                          style={{ marginLeft: 6, fontSize: 10, fontWeight: 700, color: '#0a66d0', background: '#eef4ff', border: '1px solid #cfe0fb', borderRadius: 6, padding: '1px 6px' }}
                        >
                          ⏳ Counting…
                        </span>
                      )}
                      {site.error && (
                        <span
                          title={site.error}
                          style={{ marginLeft: 6, fontSize: 10, fontWeight: 700, color: '#b91c1c', background: '#fef2f2', border: '1px solid #fecaca', borderRadius: 6, padding: '1px 6px' }}
                        >
                          ⚠ Unable to load
                        </span>
                      )}
                      {site.error && (
                        <div style={{ marginTop: 2, fontSize: 11, fontWeight: 400, color: '#b91c1c' }}>
                          {site.error}
                        </div>
                      )}
                    </td>
                    <td style={{ ...tdStyle, textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: site.error ? '#b91c1c' : undefined }}>
                      {site.error ? <span title={site.error}>Error</span> : site.stats_pending ? <span style={{ color: 'var(--vdms-text-muted)' }}>…</span> : formatNumber(site.files)}
                    </td>
                    <td style={{ ...tdStyle, textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: site.error ? '#b91c1c' : undefined }}>
                      {site.error ? <span title={site.error}>Error</span> : site.stats_pending ? <span style={{ color: 'var(--vdms-text-muted)' }}>…</span> : formatNumber(site.folders)}
                    </td>
                    <td style={{ ...tdStyle, color: 'var(--vdms-text-muted)', fontSize: 12 }}>{site.error ? '—' : site.stats_pending ? 'Just added' : formatDateTime(site.last_modified_epoch)}</td>
                    <td style={{ ...tdStyle, textAlign: 'right', whiteSpace: 'nowrap' }}>
                      {!isSelected && (
                        <button
                          type="button"
                          onClick={() => host._handleDashboardSiteChange(site.site_key)}
                          style={{ background: clay.accentSoft, border: 'none', borderRadius: 6, padding: '4px 10px', fontSize: 12, fontWeight: 700, color: clay.accentDeep, cursor: 'pointer', marginRight: 6 }}
                        >
                          Show documents
                        </button>
                      )}
                      {site.web_url && (
                        <button
                          type="button"
                          onClick={() => window.open(site.web_url, '_blank', 'noopener')}
                          style={{ background: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: 6, padding: '4px 10px', fontSize: 12, fontWeight: 600, color: '#0284c7', cursor: 'pointer' }}
                        >
                          Open site
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Vessels — site-wise, same SharePoint Site filter as everything else on Home */}
      <DashboardVesselsPanel host={host} siteKey={siteFilter} nonce={refreshNonce} dashboardStats={dashboardStats} palette={palette} />

      {/* Documents — browsed and paged here on Home */}
      <DashboardDocumentsPanel host={host} siteKey={siteFilter} nonce={refreshNonce} statsEpoch={dashboardStats?.last_refreshed_epoch} palette={palette} />
    </div>
  );
}
