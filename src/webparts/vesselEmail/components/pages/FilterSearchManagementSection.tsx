import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';
import { dmsBtn } from '../dmsDesignSystem';

/**
 * Settings → Filter Search Management.
 * Backend: backend/app/filter_settings_api.py (key/value in app_settings,
 * same table folder_structure_default_mode / hidden_modules use — no
 * schema change).
 * Lets an admin choose, app-wide, how the Documents page's filter toolbar
 * is presented to every user:
 *   - "dropdown": the existing inline row of <select> filters.
 *   - "panel": a "Filters" button that opens a slide-out panel with the
 *     same filters grouped as sections (Keyword, Main folder, Vessel,
 *     Category, Sub-folder, Attachment status, Sort) — see FilterPanel.tsx.
 * VesselEmail.state.filterUiMode is loaded once in componentDidMount
 * (_loadFilterSettings) and updated immediately after a save here, so the
 * Documents page (DocumentsPage.tsx) picks up the change without a reload.
 */

interface ConfigResponse { mode: 'dropdown' | 'panel'; is_admin: boolean; }

const C = {
  text: 'var(--vdms-text)', sub: 'var(--vdms-text-muted)', border: 'var(--vdms-line)', blue: clay.accent, blueBg: clay.accentSoft,
  redBg: clay.pillDangerBg, red: clay.pillDangerText, greenBg: clay.pillActiveBg, green: clay.pillActiveText,
};

async function readError(r: Response): Promise<string> {
  try {
    const j = await r.json();
    return (j && (j.detail || j.message)) || `HTTP ${r.status}`;
  } catch {
    return `HTTP ${r.status}`;
  }
}

interface ModeOption { value: 'dropdown' | 'panel'; title: string; description: string; }
const MODE_OPTIONS: ModeOption[] = [
  {
    value: 'dropdown',
    title: 'Dropdown filters (default)',
    description: 'The current inline row of dropdown filters above the document list (Main folder, Vessel, Category, Sub-folder, Attachment status, Sort).',
  },
  {
    value: 'panel',
    title: 'Filter panel',
    description: 'A "Filters" button opens a slide-out panel with the same filters grouped into sections — closer to a keyword + checklist filter panel, with Clear All / Apply.',
  },
];

export function FilterSearchManagementSection({ host }: { host: VesselEmail }): React.ReactElement {
  const [config, setConfig] = React.useState<ConfigResponse | null>(null);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState('');
  const [mode, setMode] = React.useState<'dropdown' | 'panel'>('dropdown');
  const [saving, setSaving] = React.useState(false);
  const [notice, setNotice] = React.useState('');

  const load = React.useCallback(async (): Promise<void> => {
    setLoading(true); setError('');
    try {
      const r = await fetch(`${host._base()}/api/filter-settings/config`, { headers: host._headers() });
      if (!r.ok) throw new Error(await readError(r));
      const data: ConfigResponse = await r.json();
      setConfig(data);
      setMode(data.mode === 'panel' ? 'panel' : 'dropdown');
    } catch (e) {
      setError((e as Error).message || 'Could not load filter settings.');
    } finally {
      setLoading(false);
    }
  }, [host]);

  React.useEffect(() => { load().catch(() => undefined); }, [load]);

  const dirty = !!config && config.mode !== mode;

  const save = async (): Promise<void> => {
    setSaving(true); setError(''); setNotice('');
    try {
      const r = await fetch(`${host._base()}/api/filter-settings/mode`, {
        method: 'PUT', headers: host._headers(), body: JSON.stringify({ mode }),
      });
      if (!r.ok) throw new Error(await readError(r));
      const data: { mode: 'dropdown' | 'panel' } = await r.json();
      setMode(data.mode);
      setConfig(prev => (prev ? { ...prev, mode: data.mode } : prev));
      // Refresh every open Documents page immediately — not just this tab —
      // by pushing the same value into VesselEmail's own state.
      host.setState({ filterUiMode: data.mode });
      setNotice('Saved. The Documents page updates immediately for everyone.');
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setSaving(false);
    }
  };

  if (loading && !config) return <div style={{ fontSize: 13, color: C.sub }}>Loading filter settings…</div>;
  if (!config) {
    return (
      <div style={{ padding: 12, borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>
        {error || 'Filter settings are unavailable.'}{' '}
        <button style={dmsBtn('secondary', true)} onClick={() => { load().catch(() => undefined); }}>Retry</button>
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14, maxWidth: 720 }}>
      <div>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>Filter Search Management</div>
        <div style={{ fontSize: 12, color: C.sub, marginTop: 2 }}>
          Choose how the Documents page's filters are presented to every user.
        </div>
      </div>

      {!config.is_admin && (
        <div style={{ fontSize: 13, color: C.sub }}>Only administrators can change this setting.</div>
      )}

      <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        {MODE_OPTIONS.map(opt => {
          const selected = mode === opt.value;
          return (
            <label
              key={opt.value}
              style={{
                display: 'flex', alignItems: 'flex-start', gap: 10, padding: '12px 14px',
                border: `1px solid ${selected ? C.blue : C.border}`, borderRadius: 8,
                background: selected ? C.blueBg : 'var(--vdms-surface)',
                cursor: config.is_admin && !saving ? 'pointer' : 'not-allowed',
              }}
            >
              <input
                type="radio"
                name="filter-ui-mode"
                checked={selected}
                disabled={!config.is_admin || saving}
                onChange={() => setMode(opt.value)}
                style={{ marginTop: 3 }}
              />
              <div>
                <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>{opt.title}</div>
                <div style={{ fontSize: 12, color: C.sub, marginTop: 2 }}>{opt.description}</div>
              </div>
            </label>
          );
        })}
      </div>

      <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
        <button style={dmsBtn('primary', !(!config.is_admin || !dirty || saving))} disabled={!config.is_admin || !dirty || saving} onClick={() => { save().catch(() => undefined); }}>
          {saving ? 'Saving…' : 'Save changes'}
        </button>
        {dirty && !saving && <span style={{ fontSize: 12, color: C.sub }}>You have unsaved changes.</span>}
      </div>

      {error && <div style={{ padding: '10px 12px', borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>{error}</div>}
      {notice && <div style={{ padding: '10px 12px', borderRadius: 8, background: C.greenBg, color: C.green, fontSize: 13 }}>{notice}</div>}
    </div>
  );
}
