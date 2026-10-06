import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';
import { dmsBtn, dmsControlStyle } from '../dmsDesignSystem';

/**
 * Settings → Vessel Settings → Folder Structure Mode.
 * Backend: backend/app/folder_structure_api.py (+ services/folder_structure.py).
 * Flow: pick mode + vessel(s) → Preview (dry run) → Confirm → Apply → summary.
 * Admin-only writes are enforced server-side; this UI just disables controls.
 */

type ModeId = 'empty_pool' | 'full_template' | 'adopt_existing' | 'adopt_create';

interface ModeInfo { id: ModeId; label: string; description: string; }
interface VesselInfo { id: string; name: string; folder_structure_mode: ModeId; is_provisioned: boolean; }
interface SiteInfo { site_key: string; site_name: string; site_url: string; drive_id: string; }
interface GuardInfo { production: boolean; environment: string; blocked: boolean; reason: string; }
interface ConfigResponse {
  modes: ModeInfo[];
  default_mode: ModeId;
  is_admin: boolean;
  site: SiteInfo;
  guard: GuardInfo;
  template: { version: number; root: string; mains: Array<{ name: string }> };
  vessels: VesselInfo[];
}
interface PlanItem {
  path: string; name: string; level: string; scope: string; vessel: string | null;
  sp: string; dms: string; outcome: 'created' | 'reused' | 'custom' | 'skipped' | 'failed'; reason: string;
}
interface Summary { created_dms: number; created_sp: number; reused: number; custom_kept: number; skipped: number; failed: number; }
interface RunResult { mode: ModeId; mode_label: string; dry_run: boolean; vessels: string[]; site: SiteInfo; summary: Summary; items: PlanItem[]; }

// Token-backed palette — every value resolves through `clay.*` / `var(--vdms-*)`
// so this section matches Dashboard's look (and re-themes with it).
const C = {
  text: 'var(--vdms-text)', sub: 'var(--vdms-text-muted)', border: 'var(--vdms-line)', blue: clay.accent, blueBg: clay.accentSoft,
  amberBg: clay.pillWarnBg, amber: clay.pillWarnText, redBg: clay.pillDangerBg, red: clay.pillDangerText, greenBg: clay.pillActiveBg, green: clay.pillActiveText,
};

const OUTCOME_STYLE: Record<string, { bg: string; fg: string; label: string }> = {
  created: { bg: C.greenBg, fg: C.green, label: 'Created' },
  reused: { bg: C.blueBg, fg: C.blue, label: 'Reused' },
  custom: { bg: 'var(--vdms-surface-alt)', fg: 'var(--vdms-text-muted)', label: 'Custom kept' },
  skipped: { bg: 'var(--vdms-surface-alt)', fg: 'var(--vdms-text-muted)', label: 'Skipped' },
  failed: { bg: C.redBg, fg: C.red, label: 'Failed' },
};

// Thin wrapper over the shared `dmsBtn` so every call site below (unchanged)
// now renders with Dashboard's button spec instead of this section's own.
const btn = (primary: boolean, disabled: boolean): React.CSSProperties =>
  dmsBtn(primary ? 'primary' : 'secondary', !disabled);

async function readError(r: Response): Promise<string> {
  try {
    const j = await r.json();
    return (j && (j.detail || j.message)) || `HTTP ${r.status}`;
  } catch {
    return `HTTP ${r.status}`;
  }
}

export function FolderStructureModeSection({ host }: { host: VesselEmail }): React.ReactElement {
  const [config, setConfig] = React.useState<ConfigResponse | null>(null);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState('');
  const [mode, setMode] = React.useState<ModeId>('empty_pool');
  const [vesselId, setVesselId] = React.useState('');
  const [allVessels, setAllVessels] = React.useState(false);
  const [busy, setBusy] = React.useState<'' | 'preview' | 'apply' | 'default'>('');
  const [preview, setPreview] = React.useState<RunResult | null>(null);
  const [previewKey, setPreviewKey] = React.useState('');
  const [result, setResult] = React.useState<RunResult | null>(null);
  const [confirming, setConfirming] = React.useState(false);
  const [filter, setFilter] = React.useState<string>('all');
  const [limit, setLimit] = React.useState(300);
  const [notice, setNotice] = React.useState('');

  const selectionKey = `${mode}|${allVessels ? '*' : vesselId}`;

  const load = React.useCallback(async (): Promise<void> => {
    setLoading(true); setError('');
    try {
      const r = await fetch(`${host._base()}/api/folder-structure/config`, { headers: host._headers() });
      if (!r.ok) throw new Error(await readError(r));
      const data: ConfigResponse = await r.json();
      setConfig(data);
      setMode(data.default_mode);
      if (data.vessels.length && !vesselId) setVesselId(data.vessels[0].id);
    } catch (e) {
      setError((e as Error).message || 'Could not load folder structure settings.');
    } finally {
      setLoading(false);
    }
  }, [host]);

  React.useEffect(() => { load().catch(() => undefined); }, [load]);

  // A preview only authorises Apply for the exact selection it was run for.
  React.useEffect(() => { setConfirming(false); }, [selectionKey]);

  const selectedVessel = config?.vessels.find(v => v.id === vesselId);
  const canWrite = !!config?.is_admin && !config?.guard.blocked;
  const targetReady = allVessels ? (config?.vessels.length || 0) > 0 : !!vesselId;

  const run = async (apply: boolean): Promise<void> => {
    setBusy(apply ? 'apply' : 'preview'); setError(''); setNotice('');
    try {
      const body = JSON.stringify({ mode, vessel_ids: allVessels ? [] : [vesselId], all_vessels: allVessels, confirm: apply });
      const r = await fetch(`${host._base()}/api/folder-structure/${apply ? 'apply' : 'preview'}`, {
        method: 'POST', headers: host._headers(), body,
      });
      if (!r.ok) throw new Error(await readError(r));
      const data: RunResult = await r.json();
      setFilter('all'); setLimit(300);
      if (apply) {
        setResult(data); setPreview(null); setPreviewKey(''); setConfirming(false);
        await load();
      } else {
        setPreview(data); setPreviewKey(selectionKey); setResult(null);
      }
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy('');
    }
  };

  const saveDefault = async (): Promise<void> => {
    setBusy('default'); setError(''); setNotice('');
    try {
      const r = await fetch(`${host._base()}/api/folder-structure/default-mode`, {
        method: 'PUT', headers: host._headers(), body: JSON.stringify({ mode }),
      });
      if (!r.ok) throw new Error(await readError(r));
      setNotice('Saved. New vessels will use this mode.');
      await load();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy('');
    }
  };

  if (loading && !config) return <div style={{ fontSize: 13, color: C.sub }}>Loading folder structure settings…</div>;
  if (!config) {
    return (
      <div style={{ padding: 12, borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>
        {error || 'Folder structure settings are unavailable.'}{' '}
        <button style={btn(false, false)} onClick={() => { load().catch(() => undefined); }}>Retry</button>
      </div>
    );
  }

  const shown = result || (previewKey === selectionKey ? preview : null);
  const items = shown ? shown.items.filter(i => filter === 'all' || i.outcome === filter) : [];
  const modeLabel = config.modes.find(m => m.id === mode)?.label || mode;
  const targetLabel = allVessels ? `all ${config.vessels.length} vessels` : (selectedVessel?.name || '—');

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14, maxWidth: 980 }}>
      <div>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>Folder Structure Mode</div>
        <div style={{ fontSize: 12, color: C.sub, marginTop: 2 }}>
          Site: <b>{config.site.site_name || config.site.site_key}</b> · library drive {config.site.drive_id ? `${config.site.drive_id.slice(0, 10)}…` : '—'} · template v{config.template.version}
        </div>
      </div>

      {config.guard.blocked && (
        <div style={{ padding: '10px 12px', borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>
          <b>Blocked on this site.</b> {config.guard.reason}
        </div>
      )}
      <div style={{ padding: '10px 12px', borderRadius: 8, background: C.amberBg, color: C.amber, fontSize: 13 }}>
        Switching modes never deletes, renames or moves existing folders or files — in the DMS or in SharePoint. Existing names and custom folders are kept.
      </div>
      {!config.is_admin && (
        <div style={{ fontSize: 13, color: C.sub }}>Only administrators can change this setting.</div>
      )}

      <div role="radiogroup" aria-label="Folder structure mode" style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {config.modes.map((m, idx) => (
          <label key={m.id} style={{
            display: 'flex', gap: 10, alignItems: 'flex-start', padding: '10px 12px', borderRadius: 8, cursor: config.is_admin ? 'pointer' : 'default',
            border: `1px solid ${mode === m.id ? C.blue : C.border}`, background: mode === m.id ? C.blueBg : 'var(--vdms-surface)',
          }}>
            <input type="radio" name="folder-structure-mode" value={m.id} checked={mode === m.id}
              disabled={!config.is_admin} onChange={() => setMode(m.id)} style={{ marginTop: 3 }} />
            <span>
              <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>
                Mode {idx + 1} – {m.label}
                {config.default_mode === m.id && <span style={{ marginLeft: 8, fontSize: 11, fontWeight: 600, color: C.blue }}>Default for new vessels</span>}
              </span>
              <span style={{ display: 'block', fontSize: 12, color: C.sub, marginTop: 2 }}>{m.description}</span>
            </span>
          </label>
        ))}
      </div>

      <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
        <button style={btn(false, !config.is_admin || busy !== '' || mode === config.default_mode)}
          disabled={!config.is_admin || busy !== '' || mode === config.default_mode} onClick={() => { saveDefault().catch(() => undefined); }}>
          {busy === 'default' ? 'Saving…' : 'Use as default for new vessels'}
        </button>
        <span style={{ fontSize: 12, color: C.sub }}>New vessels read this default when they are created.</span>
      </div>

      <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 14, display: 'flex', flexDirection: 'column', gap: 10 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>Apply to existing vessels</div>
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
          <select value={vesselId} disabled={allVessels || !config.vessels.length} onChange={e => setVesselId(e.target.value)}
            style={{ ...dmsControlStyle(), minWidth: 280 }}>
            {!config.vessels.length && <option value="">No vessels on this site</option>}
            {config.vessels.map(v => (
              <option key={v.id} value={v.id}>
                {v.name} — {config.modes.find(m => m.id === v.folder_structure_mode)?.label || v.folder_structure_mode}
              </option>
            ))}
          </select>
          <label style={{ fontSize: 13, color: C.text, display: 'flex', gap: 6, alignItems: 'center' }}>
            <input type="checkbox" checked={allVessels} onChange={e => setAllVessels(e.target.checked)} />
            Apply to all vessels ({config.vessels.length})
          </label>
        </div>
        {selectedVessel && !allVessels && (
          <div style={{ fontSize: 12, color: C.sub }}>
            Current mode for {selectedVessel.name}: <b>{config.modes.find(m => m.id === selectedVessel.folder_structure_mode)?.label}</b>
          </div>
        )}
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
          <button style={btn(false, !canWrite || !targetReady || busy !== '')} disabled={!canWrite || !targetReady || busy !== ''}
            onClick={() => { run(false).catch(() => undefined); }}>
            {busy === 'preview' ? 'Scanning…' : 'Preview (dry run)'}
          </button>
          <button style={btn(true, !canWrite || previewKey !== selectionKey || busy !== '')}
            disabled={!canWrite || previewKey !== selectionKey || busy !== ''} onClick={() => setConfirming(true)}
            title={previewKey !== selectionKey ? 'Run a preview for this selection first' : ''}>
            Apply…
          </button>
        </div>

        {confirming && preview && (
          <div style={{ padding: 12, borderRadius: 8, border: `1px solid ${C.blue}`, background: C.blueBg, fontSize: 13, color: C.text }}>
            Apply <b>{modeLabel}</b> to <b>{targetLabel}</b> on <b>{config.site.site_name || config.site.site_key}</b>? This will create {preview.summary.created_sp} SharePoint folder(s) and {preview.summary.created_dms} DMS link(s). Nothing is deleted, renamed or moved.
            <div style={{ display: 'flex', gap: 8, marginTop: 10 }}>
              <button style={btn(true, busy !== '')} disabled={busy !== ''} onClick={() => { run(true).catch(() => undefined); }}>
                {busy === 'apply' ? 'Applying…' : 'Confirm and apply'}
              </button>
              <button style={btn(false, busy !== '')} disabled={busy !== ''} onClick={() => setConfirming(false)}>Cancel</button>
            </div>
          </div>
        )}
      </div>

      {error && <div style={{ padding: '10px 12px', borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>{error}</div>}
      {notice && <div style={{ padding: '10px 12px', borderRadius: 8, background: C.greenBg, color: C.green, fontSize: 13 }}>{notice}</div>}

      {shown && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>
            {shown.dry_run ? 'Preview — nothing has been changed yet' : 'Result'} · {shown.mode_label} · {shown.vessels.length} vessel(s)
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            {[
              ['all', `All ${shown.items.length}`],
              ['created', `${shown.dry_run ? 'To create' : 'Created'} — SharePoint ${shown.summary.created_sp} / DMS ${shown.summary.created_dms}`],
              ['reused', `Reused ${shown.summary.reused}`],
              ['custom', `Custom kept ${shown.summary.custom_kept}`],
              ['skipped', `Skipped ${shown.summary.skipped}`],
              ['failed', `Failed ${shown.summary.failed}`],
            ].map(([k, label]) => (
              <button key={k} onClick={() => { setFilter(k); setLimit(300); }} style={{
                border: `1px solid ${filter === k ? C.blue : C.border}`, background: filter === k ? C.blueBg : 'var(--vdms-surface)',
                color: k === 'failed' && shown.summary.failed ? C.red : C.text, fontSize: 12, fontWeight: 600,
                padding: '5px 10px', borderRadius: 999, cursor: 'pointer',
              }}>{label}</button>
            ))}
          </div>
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: 'auto', maxHeight: 460 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ background: 'var(--vdms-surface-alt)', position: 'sticky', top: 0 }}>
                  {['Folder', 'Vessel', 'SharePoint', 'DMS', 'Result', 'Note'].map(h => (
                    <th key={h} style={{ textAlign: 'left', padding: '7px 10px', color: C.sub, fontWeight: 600, borderBottom: `1px solid ${C.border}` }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {items.slice(0, limit).map((it, i) => {
                  const st = OUTCOME_STYLE[it.outcome];
                  const depth = Math.max(0, it.path.split('/').length - 1);
                  return (
                    <tr key={`${it.path}-${i}`} style={{ borderBottom: '1px solid var(--vdms-border-soft)' }}>
                      <td style={{ padding: '6px 10px', paddingLeft: 10 + depth * 12, color: C.text }} title={it.path}>{it.name}</td>
                      <td style={{ padding: '6px 10px', color: C.sub }}>{it.vessel || 'Common'}</td>
                      <td style={{ padding: '6px 10px', color: C.sub }}>{it.sp}</td>
                      <td style={{ padding: '6px 10px', color: C.sub }}>{it.dms}</td>
                      <td style={{ padding: '6px 10px' }}>
                        <span style={{ background: st.bg, color: st.fg, borderRadius: 999, padding: '2px 8px', fontWeight: 600 }}>{st.label}</span>
                      </td>
                      <td style={{ padding: '6px 10px', color: C.sub }}>{it.reason}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {items.length > limit && (
            <button style={btn(false, false)} onClick={() => setLimit(limit + 300)}>Show more ({items.length - limit} left)</button>
          )}
        </div>
      )}
    </div>
  );
}
