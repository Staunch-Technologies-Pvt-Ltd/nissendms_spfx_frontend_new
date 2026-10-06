import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';
import { dmsBtn } from '../dmsDesignSystem';

/**
 * Sites → Site Management → (per site) Vessel folders.
 *
 * Which folders of the site's library hold the vessel folders. The Vessels
 * page ("Found in SharePoint"), the Dashboard's vessel count and the
 * SharePoint vessel sync only look inside them. Backend:
 * /api/vessel-roots/{site_key} (services/vessel_roots.py).
 */

type Mode = 'auto' | 'folders' | 'none';
interface RootsState { mode: Mode; paths: string[]; is_admin?: boolean; updated_by?: string | null }
interface BrowseFolder { id: string; name: string; path: string }

async function readError(r: Response): Promise<string> {
  try { const j = await r.json(); return (j && (j.detail || j.message)) || `HTTP ${r.status}`; } catch { return `HTTP ${r.status}`; }
}

export function VesselRootsControl({ host, siteKey, onSaved }: { host: VesselEmail; siteKey: string; onSaved?: () => void }): React.ReactElement {
  const [saved, setSaved] = React.useState<RootsState | null>(null);
  const [loadError, setLoadError] = React.useState('');
  const [open, setOpen] = React.useState(false);
  const [mode, setMode] = React.useState<Mode>('auto');
  const [paths, setPaths] = React.useState<string[]>([]);
  const [browsePath, setBrowsePath] = React.useState('');
  const [folders, setFolders] = React.useState<BrowseFolder[]>([]);
  const [browsing, setBrowsing] = React.useState(false);
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState('');

  const load = React.useCallback(async (): Promise<void> => {
    setLoadError('');
    try {
      const r = await fetch(`${host._base()}/api/vessel-roots/${encodeURIComponent(siteKey)}`, { headers: host._headers() });
      if (!r.ok) throw new Error(await readError(r));
      setSaved(await r.json());
    } catch (e) {
      setLoadError((e as Error).message);
    }
  }, [host, siteKey]);

  React.useEffect(() => { load().catch(() => undefined); }, [load]);

  const browse = React.useCallback(async (path: string): Promise<void> => {
    setBrowsing(true); setError('');
    try {
      const r = await fetch(`${host._base()}/api/admin/sites/${encodeURIComponent(siteKey)}/folders?path=${encodeURIComponent(path)}`, { headers: host._headers() });
      if (!r.ok) throw new Error(await readError(r));
      const d = await r.json();
      setFolders(d.folders || []);
      setBrowsePath(d.path || '');
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBrowsing(false);
    }
  }, [host, siteKey]);

  const startEdit = (): void => {
    setMode(saved?.mode || 'auto');
    setPaths(saved?.paths || []);
    setError('');
    setOpen(true);
    void browse('');
  };

  const toggle = (path: string): void => setPaths((prev) => (prev.includes(path) ? prev.filter((p) => p !== path) : [...prev, path]));

  const save = async (): Promise<void> => {
    if (mode === 'folders' && paths.length === 0) { setError('Tick at least one folder, or choose another option.'); return; }
    setSaving(true); setError('');
    try {
      const r = await fetch(`${host._base()}/api/vessel-roots/${encodeURIComponent(siteKey)}`, {
        method: 'PUT', headers: host._headers(), body: JSON.stringify({ mode, paths: mode === 'folders' ? paths : [] }),
      });
      if (!r.ok) throw new Error(await readError(r));
      await load();
      setOpen(false);
      onSaved?.();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setSaving(false);
    }
  };

  const crumbs = browsePath.split('/').filter(Boolean);
  const summary = !saved ? (loadError ? 'Couldn’t load' : 'Loading…')
    : saved.mode === 'none' ? 'This site has no vessel folders'
    : saved.mode === 'folders' ? saved.paths.join(' · ')
    : 'Automatic (whole library)';

  const radio = (value: Mode, title: string, hint: string): React.ReactElement => (
    <label style={{ display: 'flex', gap: 8, alignItems: 'flex-start', padding: '6px 8px', borderRadius: 8, cursor: 'pointer', background: mode === value ? clay.accentSoft : 'transparent' }}>
      <input type="radio" name={`roots-${siteKey}`} checked={mode === value} onChange={() => setMode(value)} style={{ marginTop: 3 }} />
      <span>
        <span style={{ fontSize: 12.5, fontWeight: 600, color: 'var(--vdms-text)' }}>{title}</span>
        <span style={{ display: 'block', fontSize: 11.5, color: 'var(--vdms-text-muted)' }}>{hint}</span>
      </span>
    </label>
  );

  return (
    <div style={{ margin: '-2px 0 4px 12px', padding: '6px 10px', borderLeft: `2px solid ${clay.accentSoft}` }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
        <Icon iconName="Boat" aria-hidden="true" style={{ fontSize: 12, color: clay.accent }} />
        <span style={{ fontSize: 11.5, fontWeight: 700, color: 'var(--vdms-text-muted)' }}>Vessel folders:</span>
        <span style={{ fontSize: 12, color: 'var(--vdms-text)' }} title="Where the Vessels page and Dashboard look for this site's vessel folders">{summary}</span>
        {saved?.is_admin && !open && (
          <button type="button" onClick={startEdit} style={{ ...dmsBtn('ghost'), height: 26, padding: '0 8px', fontSize: 12 }}>
            <Icon iconName="Edit" aria-hidden="true" style={{ fontSize: 11 }} /> Change
          </button>
        )}
      </div>

      {open && (
        // The editor opens as a dialog so it never pushes the page content
        // (vessel cards, site rows) out of view.
        <div role="dialog" aria-modal="true" aria-label={`Vessel folders for ${siteKey}`}
          onClick={(e) => { if (e.target === e.currentTarget && !saving) setOpen(false); }}
          style={{ position: 'fixed', inset: 0, zIndex: 1000, background: 'rgba(8,18,32,0.42)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}>
        <div style={{ width: 'min(680px, 100%)', maxHeight: '86vh', overflowY: 'auto', padding: 18, borderRadius: 14, border: '1px solid var(--vdms-line)', background: 'var(--vdms-surface)', boxShadow: '0 20px 50px rgba(8,18,32,0.28)', display: 'flex', flexDirection: 'column', gap: 10 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <Icon iconName="Boat" aria-hidden="true" style={{ fontSize: 16, color: clay.accent }} />
            <span style={{ fontSize: 15, fontWeight: 800, color: 'var(--vdms-text)' }}>Vessel folders</span>
            <span style={{ fontSize: 12, color: 'var(--vdms-text-muted)' }}>· {siteKey}</span>
            <button type="button" aria-label="Close" onClick={() => setOpen(false)} style={{ marginLeft: 'auto', border: 'none', background: 'transparent', cursor: 'pointer', color: 'var(--vdms-text-muted)' }}>
              <Icon iconName="Cancel" style={{ fontSize: 13 }} />
            </button>
          </div>
          <div style={{ fontSize: 12, color: 'var(--vdms-text-muted)' }}>
            Where are this site’s vessel folders? Every folder inside the folders you pick is treated as a vessel —
            the Vessels page suggests them and the Dashboard counts them. Nothing else in the library is considered.
          </div>
          {radio('folders', 'Only inside these folders', 'Pick the folder(s) that contain one folder per vessel, e.g. “Technical and Crewing New”.')}
          {mode === 'folders' && (
            <div style={{ marginLeft: 28, display: 'flex', flexDirection: 'column', gap: 6 }}>
              <div style={{ fontSize: 11.5, color: 'var(--vdms-text-muted)', display: 'flex', gap: 4, flexWrap: 'wrap', alignItems: 'center' }}>
                <a onClick={() => void browse('')} style={{ cursor: 'pointer', color: clay.accent }}>Library</a>
                {crumbs.map((c, i) => (
                  <span key={i} style={{ display: 'inline-flex', gap: 4, alignItems: 'center' }}>
                    <Icon iconName="ChevronRight" style={{ fontSize: 9 }} />
                    <a onClick={() => void browse(crumbs.slice(0, i + 1).join('/'))} style={{ cursor: 'pointer', color: clay.accent }}>{c}</a>
                  </span>
                ))}
              </div>
              <div style={{ maxHeight: 220, overflowY: 'auto', border: '1px solid var(--vdms-border-soft)', borderRadius: 8 }}>
                {browsing ? <div style={{ padding: 10, fontSize: 12, color: 'var(--vdms-text-muted)' }}>Loading…</div>
                  : folders.length === 0 ? <div style={{ padding: 10, fontSize: 12, color: 'var(--vdms-text-muted)' }}>No folders here.</div>
                  : folders.map((f) => (
                    <div key={f.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 10px', borderBottom: '1px solid var(--vdms-border-soft)' }}>
                      <input type="checkbox" checked={paths.includes(f.path)} onChange={() => toggle(f.path)} aria-label={`Use ${f.name} as a vessel folder`} />
                      <Icon iconName="FabricFolder" style={{ fontSize: 13, color: clay.accent }} />
                      <span style={{ fontSize: 12.5, color: 'var(--vdms-text)', flex: 1 }}>{f.name}</span>
                      <button type="button" onClick={() => void browse(f.path)} title="Open this folder" style={{ border: 'none', background: 'transparent', cursor: 'pointer', color: 'var(--vdms-text-muted)' }}>
                        <Icon iconName="ChevronRight" style={{ fontSize: 11 }} />
                      </button>
                    </div>
                  ))}
              </div>
              {paths.length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                  {paths.map((p) => (
                    <span key={p} style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: clay.accentSoft, color: clay.accent, borderRadius: 14, padding: '3px 9px', fontSize: 12 }}>
                      {p}
                      <button type="button" onClick={() => toggle(p)} aria-label={`Remove ${p}`} style={{ border: 0, background: 'transparent', color: clay.accent, cursor: 'pointer', padding: 0 }}>×</button>
                    </span>
                  ))}
                </div>
              )}
            </div>
          )}
          {radio('none', 'This site has no vessel folders', 'E.g. a migration source site. The Vessels page won’t suggest anything from it.')}
          {radio('auto', 'Automatic', 'Guess from the whole library (folders at the top level or one level down whose names match a vessel). This is how it worked before.')}
          {error && <div style={{ fontSize: 12, color: clay.pillDangerText }}>{error}</div>}
          <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
            <button type="button" onClick={() => setOpen(false)} style={dmsBtn('ghost')}>Cancel</button>
            <button type="button" onClick={() => void save()} disabled={saving} style={dmsBtn('primary', !saving)}>{saving ? 'Saving…' : 'Save'}</button>
          </div>
        </div>
        </div>
      )}
    </div>
  );
}
