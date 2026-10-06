import * as React from 'react';
import type VesselEmail from '../VesselEmail';
import {
  applyColorTheme, CLAY_COLOR_PRESETS, DEFAULT_CLAY_COLORS, clay,
} from '../clayTheme';
import type { ClayColorSet, ClayColorTheme } from '../clayTheme';
import { dmsBtn } from '../dmsDesignSystem';

/**
 * Settings → Color Management.
 * Backend: backend/app/color_settings_api.py (key/value in app_settings,
 * same table module_settings_api / filter_settings_api use — no schema
 * change). Lets an admin recolor the whole app — background, text,
 * design/accent color and hover color — set independently for light mode
 * and night mode. Every signed-in user gets the same colors
 * (VesselEmail.state.colorTheme, loaded in componentDidMount via
 * _loadColorSettings), the same way Module Management's hidden list is
 * shared app-wide.
 *
 * Editing here calls applyColorTheme() on every keystroke for an instant
 * live preview across the whole app (buttons, top bar, cards, dialogs —
 * everything built from the clay.* tokens in clayTheme.ts), before the
 * change is ever saved. "Save changes" persists it for everyone; leaving
 * the page without saving reverts the preview back to the saved colors.
 */

const C = {
  text: 'var(--vdms-text)', sub: 'var(--vdms-text-muted)', border: 'var(--vdms-line)', blue: clay.accent, blueBg: clay.accentSoft,
  redBg: clay.pillDangerBg, red: clay.pillDangerText, greenBg: clay.pillActiveBg, green: clay.pillActiveText,
};

interface ConfigResponse { colors: ClayColorTheme; defaults: ClayColorTheme; presets: { id: string; label: string; colors: ClayColorTheme }[]; is_admin: boolean; }

const HEX_RE = /^#[0-9a-fA-F]{6}$/;

async function readError(r: Response): Promise<string> {
  try {
    const j = await r.json();
    return (j && (j.detail || j.message)) || `HTTP ${r.status}`;
  } catch {
    return `HTTP ${r.status}`;
  }
}

const FIELDS: { key: keyof ClayColorSet; label: string; hint: string }[] = [
  { key: 'bg', label: 'Background color', hint: 'The app’s base background.' },
  { key: 'text', label: 'Text color', hint: 'Main body and heading text.' },
  { key: 'accent', label: 'Design color', hint: 'Primary/accent color — buttons, top bar, highlights.' },
  { key: 'hover', label: 'Hover color', hint: 'Shown when hovering buttons and clickable cards.' },
];

function ColorField({ label, hint, value, disabled, onChange }: {
  label: string; hint: string; value: string; disabled: boolean; onChange: (hex: string) => void;
}): React.ReactElement {
  const [text, setText] = React.useState(value);
  React.useEffect(() => { setText(value); }, [value]);
  const invalid = !HEX_RE.test(text);

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 0' }}>
      <input
        type="color"
        value={HEX_RE.test(value) ? value : '#000000'}
        disabled={disabled}
        onChange={e => { setText(e.target.value); onChange(e.target.value); }}
        style={{ width: 40, height: 32, border: `1px solid ${C.border}`, borderRadius: 6, padding: 0, cursor: disabled ? 'not-allowed' : 'pointer', background: 'none' }}
      />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 13, fontWeight: 600, color: C.text }}>{label}</div>
        <div style={{ fontSize: 11, color: C.sub, marginTop: 1 }}>{hint}</div>
      </div>
      <input
        type="text"
        value={text}
        disabled={disabled}
        onChange={e => {
          const v = e.target.value;
          setText(v);
          if (HEX_RE.test(v)) onChange(v);
        }}
        placeholder="#RRGGBB"
        style={{
          width: 96, fontSize: 12, fontFamily: 'monospace', padding: '6px 8px', borderRadius: 6,
          border: `1px solid ${invalid ? clay.pillDangerText : C.border}`, color: C.text, textTransform: 'uppercase',
          background: 'var(--vdms-glass)',
        }}
      />
    </div>
  );
}

function ModeGroup({ title, colors, disabled, onChange }: {
  title: string; colors: ClayColorSet; disabled: boolean; onChange: (key: keyof ClayColorSet, hex: string) => void;
}): React.ReactElement {
  return (
    <div style={{ border: `1px solid ${C.border}`, borderRadius: clay.radiusCard, padding: '4px 14px', background: 'var(--vdms-surface)', boxShadow: clay.shadowRaised }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: C.sub, textTransform: 'uppercase', letterSpacing: 0.4, padding: '10px 0 0' }}>{title}</div>
      <div style={{ display: 'flex', flexDirection: 'column' }}>
        {FIELDS.map((f, idx) => (
          <div key={f.key} style={{ borderTop: idx === 0 ? 'none' : `1px solid ${C.border}` }}>
            <ColorField label={f.label} hint={f.hint} value={colors[f.key]} disabled={disabled} onChange={hex => onChange(f.key, hex)} />
          </div>
        ))}
      </div>
    </div>
  );
}

function PresetSwatch({ label, colors, active, disabled, onClick }: {
  label: string; colors: ClayColorTheme; active: boolean; disabled: boolean; onClick: () => void;
}): React.ReactElement {
  return (
    <button
      type="button"
      disabled={disabled}
      onClick={onClick}
      title={label}
      style={{
        display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6, padding: 8, borderRadius: 8,
        border: active ? `2px solid ${C.blue}` : `1px solid ${C.border}`, background: active ? C.blueBg : 'var(--vdms-surface)',
        cursor: disabled ? 'not-allowed' : 'pointer', width: 92,
      }}
    >
      <div style={{ display: 'flex', borderRadius: 6, overflow: 'hidden', width: 60, height: 24, border: `1px solid ${C.border}` }}>
        <div style={{ flex: 1, background: colors.light.bg }} />
        <div style={{ flex: 1, background: colors.light.accent }} />
        <div style={{ flex: 1, background: colors.night.bg }} />
        <div style={{ flex: 1, background: colors.night.accent }} />
      </div>
      <div style={{ fontSize: 11, fontWeight: 600, color: C.text, textAlign: 'center', lineHeight: 1.2 }}>{label}</div>
    </button>
  );
}

export function ColorManagementSection({ host }: { host: VesselEmail }): React.ReactElement {
  const [config, setConfig] = React.useState<ConfigResponse | null>(null);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState('');
  const [saving, setSaving] = React.useState(false);
  const [notice, setNotice] = React.useState('');
  const [draft, setDraft] = React.useState<ClayColorTheme | null>(null);
  const savedRef = React.useRef<ClayColorTheme | null>(null);

  const load = React.useCallback(async (): Promise<void> => {
    setLoading(true); setError('');
    try {
      const r = await fetch(`${host._base()}/api/color-settings/config`, { headers: host._headers() });
      if (!r.ok) throw new Error(await readError(r));
      const data: ConfigResponse = await r.json();
      setConfig(data);
      setDraft(data.colors);
      savedRef.current = data.colors;
    } catch (e) {
      setError((e as Error).message || 'Could not load color settings.');
    } finally {
      setLoading(false);
    }
  }, [host]);

  React.useEffect(() => { load().catch(() => undefined); }, [load]);

  // Live preview: whenever the draft changes, apply it app-wide immediately
  // (not yet saved). Revert to the saved colors if this section unmounts
  // without saving, so navigating away without clicking Save doesn't leave
  // the app on an un-persisted preview.
  React.useEffect(() => {
    if (draft) applyColorTheme(draft);
    return () => {
      if (savedRef.current) applyColorTheme(savedRef.current);
    };
  }, [draft]);

  const dirty = React.useMemo(() => {
    if (!config || !draft) return false;
    return JSON.stringify(draft) !== JSON.stringify(config.colors);
  }, [config, draft]);

  const setField = (mode: 'light' | 'night', key: keyof ClayColorSet, hex: string): void => {
    setDraft(prev => (prev ? { ...prev, [mode]: { ...prev[mode], [key]: hex } } : prev));
  };

  const applyPreset = (colors: ClayColorTheme): void => setDraft(colors);
  const resetToDefault = (): void => setDraft(config ? config.defaults : DEFAULT_CLAY_COLORS);

  const save = async (): Promise<void> => {
    if (!draft) return;
    setSaving(true); setError(''); setNotice('');
    try {
      const r = await fetch(`${host._base()}/api/color-settings/colors`, {
        method: 'PUT', headers: host._headers(), body: JSON.stringify(draft),
      });
      if (!r.ok) throw new Error(await readError(r));
      const data: { colors: ClayColorTheme } = await r.json();
      setDraft(data.colors);
      savedRef.current = data.colors;
      setConfig(prev => (prev ? { ...prev, colors: data.colors } : prev));
      // Refresh every user's live app immediately — not just this tab — by
      // pushing the saved colors into VesselEmail's own state, the same way
      // Module Management pushes hiddenModules into host.setState.
      host.setState({ colorTheme: data.colors });
      setNotice('Saved. The new colors apply immediately for everyone.');
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setSaving(false);
    }
  };

  if (loading && !config) return <div style={{ fontSize: 13, color: C.sub }}>Loading color settings…</div>;
  if (!config || !draft) {
    return (
      <div style={{ padding: 12, borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>
        {error || 'Color settings are unavailable.'}{' '}
        <button style={dmsBtn('secondary', true)} onClick={() => { load().catch(() => undefined); }}>Retry</button>
      </div>
    );
  }

  const disabled = !config.is_admin || saving;
  const isDefault = JSON.stringify(draft) === JSON.stringify(config.defaults);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16, maxWidth: 760 }}>
      <div>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>Color Management</div>
        <div style={{ fontSize: 12, color: C.sub, marginTop: 2 }}>
          Change the background, text, design (accent) and hover colors used across the whole app, for light mode and
          night mode independently. Changes preview instantly here; click Save to apply them for every user.
        </div>
      </div>

      {!config.is_admin && (
        <div style={{ fontSize: 13, color: C.sub }}>Only administrators can change this setting.</div>
      )}

      <div>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.sub, textTransform: 'uppercase', letterSpacing: 0.4, marginBottom: 8 }}>
          Preset palettes
        </div>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
          {config.presets.map(p => (
            <PresetSwatch
              key={p.id}
              label={p.label}
              colors={p.colors}
              disabled={disabled}
              active={JSON.stringify(draft) === JSON.stringify(p.colors)}
              onClick={() => applyPreset(p.colors)}
            />
          ))}
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: 14 }}>
        <ModeGroup title="Light mode" colors={draft.light} disabled={disabled} onChange={(k, hex) => setField('light', k, hex)} />
        <ModeGroup title="Night mode" colors={draft.night} disabled={disabled} onChange={(k, hex) => setField('night', k, hex)} />
      </div>

      <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
        <button style={dmsBtn('primary', !(disabled || !dirty))} disabled={disabled || !dirty} onClick={() => { save().catch(() => undefined); }}>
          {saving ? 'Saving…' : 'Save changes'}
        </button>
        <button style={dmsBtn('secondary', !(disabled || isDefault))} disabled={disabled || isDefault} onClick={resetToDefault}>
          Reset to default
        </button>
        {dirty && !saving && <span style={{ fontSize: 12, color: C.sub }}>You have unsaved changes — previewing live.</span>}
      </div>

      {error && <div style={{ padding: '10px 12px', borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>{error}</div>}
      {notice && <div style={{ padding: '10px 12px', borderRadius: 8, background: C.greenBg, color: C.green, fontSize: 13 }}>{notice}</div>}
    </div>
  );
}
