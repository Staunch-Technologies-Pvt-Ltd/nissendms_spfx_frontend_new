import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';

/**
 * Document-management summary shown on each Vessels-page card: how many
 * documents the vessel has, their size, when they were last updated, and how
 * they split across Drawings / Manuals / To Be Classified.
 *
 * Data: GET /api/dashboard/vessel-summary?site_key= (built from the cached
 * dashboard scan — no extra SharePoint calls). One request per site is
 * shared by every card on the page and reused for a minute.
 */

export interface VesselDocSummary {
  name: string;
  total: number;
  size_bytes: number;
  last_modified_epoch: number;
  last_modified_by: string | null;
  drawings: number;
  manuals: number;
  to_be_classified: number;
  other: number;
}

interface SummaryResponse { scan_pending: boolean; vessels: Record<string, VesselDocSummary> }

const TTL_MS = 60_000;
const cache = new Map<string, { at: number; promise: Promise<SummaryResponse> }>();

function loadSummary(host: VesselEmail, siteKey: string): Promise<SummaryResponse> {
  const hit = cache.get(siteKey);
  if (hit && Date.now() - hit.at < TTL_MS) return hit.promise;
  const qs = siteKey && siteKey !== 'all' ? `?site_key=${encodeURIComponent(siteKey)}` : '';
  const promise = fetch(`${host._base()}/api/dashboard/vessel-summary${qs}`, { headers: host._headers() })
    .then((r) => (r.ok ? r.json() : Promise.reject(new Error(`HTTP ${r.status}`))))
    .catch((e) => { cache.delete(siteKey); throw e; });
  cache.set(siteKey, { at: Date.now(), promise });
  return promise;
}

/** Drop cached figures (e.g. after the page's refresh button). */
export function clearVesselSummaryCache(): void { cache.clear(); }

function formatBytes(n: number): string {
  if (!n) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.min(units.length - 1, Math.floor(Math.log(n) / Math.log(1024)));
  return `${(n / Math.pow(1024, i)).toFixed(i >= 3 ? 1 : 0)} ${units[i]}`;
}

function formatWhen(epoch: number): string {
  if (!epoch) return '—';
  const days = Math.floor((Date.now() - epoch) / 86_400_000);
  if (days <= 0) return 'today';
  if (days === 1) return 'yesterday';
  if (days < 30) return `${days} days ago`;
  return new Date(epoch).toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
}

const label: React.CSSProperties = {
  fontSize: 10, color: 'var(--vdms-text-3, #52708a)', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.12em',
};

export function VesselDocumentPanel({ host, vesselName, siteKey }: { host: VesselEmail; vesselName: string; siteKey: string }): React.ReactElement {
  const [state, setState] = React.useState<{ loading: boolean; pending: boolean; row: VesselDocSummary | null; error: boolean }>(
    { loading: true, pending: false, row: null, error: false },
  );

  React.useEffect(() => {
    let cancelled = false;
    setState((s) => ({ ...s, loading: true }));
    loadSummary(host, siteKey)
      .then((d) => {
        if (cancelled) return;
        const row = d.vessels[vesselName.trim().toLowerCase()] || null;
        setState({ loading: false, pending: d.scan_pending, row, error: false });
      })
      .catch(() => { if (!cancelled) setState({ loading: false, pending: false, row: null, error: true }); });
    return () => { cancelled = true; };
  }, [host, vesselName, siteKey]);

  const box: React.CSSProperties = {
    background: 'var(--vdms-field, rgba(255,255,255,0.72))', borderRadius: 12, padding: '10px 12px',
    border: '1px solid var(--vdms-line, rgba(16,84,138,0.13))', display: 'flex', flexDirection: 'column', gap: 8,
  };

  if (state.loading || state.pending) {
    return (
      <div style={box}>
        <span style={label}>Documents</span>
        <span style={{ fontSize: 12.5, color: 'var(--vdms-text-muted)' }}>
          <Icon iconName="Sync" style={{ fontSize: 11, marginRight: 6 }} />
          {state.pending ? 'Counting documents in SharePoint…' : 'Loading…'}
        </span>
      </div>
    );
  }
  if (state.error) {
    return (
      <div style={box}>
        <span style={label}>Documents</span>
        <span style={{ fontSize: 12.5, color: 'var(--vdms-text-muted)' }}>Document figures are unavailable right now.</span>
      </div>
    );
  }

  const row = state.row;
  if (!row || row.total === 0) {
    return (
      <div style={box}>
        <span style={label}>Documents</span>
        <span style={{ fontSize: 13, color: 'var(--vdms-text-muted)' }}>
          <Icon iconName="Page" style={{ fontSize: 12, marginRight: 6 }} />No documents yet
        </span>
      </div>
    );
  }

  const groups: { key: keyof VesselDocSummary; name: string; color: string; icon: string }[] = [
    { key: 'drawings', name: 'Drawings', color: clay.accent, icon: 'Design' },
    { key: 'manuals', name: 'Manuals', color: '#7c3aed', icon: 'ReadingMode' },
    { key: 'to_be_classified', name: 'To Be Classified', color: '#d97706', icon: 'Inbox' },
    { key: 'other', name: 'Other', color: '#94a3b8', icon: 'Page' },
  ];
  const shown = groups.filter((g) => (row[g.key] as number) > 0);

  return (
    <div style={box}>
      <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', gap: 8 }}>
        <span style={label}>Documents</span>
        <span style={{ fontSize: 11.5, color: 'var(--vdms-text-muted)' }} title={row.last_modified_by ? `Last change by ${row.last_modified_by}` : undefined}>
          Updated {formatWhen(row.last_modified_epoch)} · {formatBytes(row.size_bytes)}
        </span>
      </div>
      {/* Proportion bar */}
      <div style={{ display: 'flex', height: 6, borderRadius: 999, overflow: 'hidden', background: 'var(--vdms-surface-alt, #eef2f6)' }}>
        {shown.map((g) => (
          <div key={g.key} title={`${g.name}: ${row[g.key]}`} style={{ width: `${((row[g.key] as number) / row.total) * 100}%`, background: g.color }} />
        ))}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: '4px 16px' }}>
        {groups.filter((g) => g.key !== 'other' || row.other > 0).map((g) => {
          const n = row[g.key] as number;
          const attention = g.key === 'to_be_classified' && n > 0;
          return (
            <div key={g.key} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: attention ? '#b45309' : 'var(--vdms-text)' }}
              title={attention ? 'Documents waiting to be sorted into Drawings or Manuals' : undefined}>
              <span style={{ width: 8, height: 8, borderRadius: 2, background: g.color, flexShrink: 0 }} />
              <span style={{ fontWeight: attention ? 700 : 500, whiteSpace: 'nowrap' }}>{g.name}</span>
              <span style={{ marginLeft: 'auto', fontWeight: 700 }}>{n.toLocaleString()}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}
