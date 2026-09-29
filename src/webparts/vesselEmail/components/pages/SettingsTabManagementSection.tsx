import * as React from 'react';
import type VesselEmail from '../VesselEmail';

/**
 * Settings → Settings Management.
 * Backend: backend/app/settings_tab_api.py (key/value in app_settings,
 * same table hidden_modules and folder_structure_default_mode use — no
 * schema change).
 * Lets an admin show/hide the Settings page's own left-nav tabs app-wide.
 * Every signed-in user's Settings left nav reads the same hidden list
 * (VesselEmail.state.hiddenSettingsTabs, loaded in componentDidMount via
 * _loadSettingsTabSettings) so a hidden tab disappears for everyone, not
 * just the admin. "Settings Management" can't be hidden (server strips it)
 * so there's always a way back in to un-hide the rest.
 */

interface SettingsTabInfo { id: string; label: string; }
interface ConfigResponse { tabs: SettingsTabInfo[]; hidden: string[]; is_admin: boolean; }

const C = {
  text: '#0f172a', sub: '#64748b', border: '#e2e8f0', blue: '#0a66d0', blueBg: '#e8f1ff',
  redBg: '#fef2f2', red: '#b91c1c', greenBg: '#ecfdf5', green: '#047857',
};

const btn = (primary: boolean, disabled: boolean): React.CSSProperties => ({
  border: primary ? 'none' : `1px solid ${C.border}`, background: disabled ? '#cbd5e1' : primary ? C.blue : '#fff',
  color: primary ? '#fff' : C.text, fontSize: 13, fontWeight: 600, padding: '8px 14px', borderRadius: 6,
  cursor: disabled ? 'not-allowed' : 'pointer',
});

async function readError(r: Response): Promise<string> {
  try {
    const j = await r.json();
    return (j && (j.detail || j.message)) || `HTTP ${r.status}`;
  } catch {
    return `HTTP ${r.status}`;
  }
}

function Toggle({ on, disabled, onChange }: { on: boolean; disabled: boolean; onChange: () => void }): React.ReactElement {
  const trackW = 40, trackH = 22, knob = 16, pad = 3;
  return (
    <button
      type="button"
      role="switch"
      aria-checked={on}
      disabled={disabled}
      onClick={onChange}
      style={{
        position: 'relative', flexShrink: 0, width: trackW, height: trackH, borderRadius: trackH,
        border: `1px solid ${on ? C.blue : C.border}`, background: on ? C.blue : '#e2e8f0',
        cursor: disabled ? 'not-allowed' : 'pointer', padding: 0, opacity: disabled ? 0.6 : 1,
        transition: 'background 0.2s ease, border-color 0.2s ease',
      }}
    >
      <span style={{
        position: 'absolute', top: pad - 1, left: on ? trackW - knob - pad - 1 : pad - 1,
        width: knob, height: knob, borderRadius: '50%', background: '#fff',
        boxShadow: '0 1px 3px rgba(0,0,0,0.3)', transition: 'left 0.2s ease',
      }} />
    </button>
  );
}

export function SettingsTabManagementSection({ host }: { host: VesselEmail }): React.ReactElement {
  const [config, setConfig] = React.useState<ConfigResponse | null>(null);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState('');
  const [hidden, setHidden] = React.useState<Set<string>>(new Set());
  const [saving, setSaving] = React.useState(false);
  const [notice, setNotice] = React.useState('');

  const load = React.useCallback(async (): Promise<void> => {
    setLoading(true); setError('');
    try {
      const r = await fetch(`${host._base()}/api/settings-tab-settings/config`, { headers: host._headers() });
      if (!r.ok) throw new Error(await readError(r));
      const data: ConfigResponse = await r.json();
      setConfig(data);
      setHidden(new Set(data.hidden));
    } catch (e) {
      setError((e as Error).message || 'Could not load settings tab settings.');
    } finally {
      setLoading(false);
    }
  }, [host]);

  React.useEffect(() => { load().catch(() => undefined); }, [load]);

  const dirty = React.useMemo(() => {
    if (!config) return false;
    const original = new Set(config.hidden);
    if (original.size !== hidden.size) return true;
    // Not `for (const id of hidden)` — iterating a Set directly needs
    // --downlevelIteration or an ES2015+ target, neither of which this
    // SPFx build is set up for (and changing tsconfig/build config is out
    // of scope here). Array.from() first, then iterate the array instead.
    return Array.from(hidden).some(id => !original.has(id));
  }, [config, hidden]);

  const toggle = (id: string): void => {
    setHidden(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  const save = async (): Promise<void> => {
    setSaving(true); setError(''); setNotice('');
    try {
      const r = await fetch(`${host._base()}/api/settings-tab-settings/hidden`, {
        method: 'PUT', headers: host._headers(), body: JSON.stringify({ hidden: Array.from(hidden) }),
      });
      if (!r.ok) throw new Error(await readError(r));
      const data: { hidden: string[] } = await r.json();
      setHidden(new Set(data.hidden));
      setConfig(prev => (prev ? { ...prev, hidden: data.hidden } : prev));
      // Refresh every user's live Settings left nav immediately — not just
      // this tab — by pushing the same list into VesselEmail's own state.
      host.setState({ hiddenSettingsTabs: data.hidden });
      setNotice('Saved. The Settings menu updates immediately for everyone.');
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setSaving(false);
    }
  };

  if (loading && !config) return <div style={{ fontSize: 13, color: C.sub }}>Loading settings management…</div>;
  if (!config) {
    return (
      <div style={{ padding: 12, borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>
        {error || 'Settings management is unavailable.'}{' '}
        <button style={btn(false, false)} onClick={() => { load().catch(() => undefined); }}>Retry</button>
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14, maxWidth: 720 }}>
      <div>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>Settings Management</div>
        <div style={{ fontSize: 12, color: C.sub, marginTop: 2 }}>
          Turn a tab off to hide it from the Settings menu for every user. Settings Management is always available.
        </div>
      </div>

      {!config.is_admin && (
        <div style={{ fontSize: 13, color: C.sub }}>Only administrators can change this setting.</div>
      )}

      <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: 'hidden' }}>
        {config.tabs.map((t, idx) => {
          const isHidden = hidden.has(t.id);
          return (
            <div key={t.id} style={{
              display: 'flex', alignItems: 'center', justifyContent: 'space-between',
              padding: '12px 14px', borderTop: idx === 0 ? 'none' : `1px solid ${C.border}`,
              background: '#fff',
            }}>
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.text }}>{t.label}</div>
                <div style={{ fontSize: 11, color: isHidden ? C.red : C.green, marginTop: 2, fontWeight: 600 }}>
                  {isHidden ? 'Hidden' : 'Active'}
                </div>
              </div>
              <Toggle on={!isHidden} disabled={!config.is_admin || saving} onChange={() => toggle(t.id)} />
            </div>
          );
        })}
      </div>

      <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
        <button style={btn(true, !config.is_admin || !dirty || saving)} disabled={!config.is_admin || !dirty || saving} onClick={() => { save().catch(() => undefined); }}>
          {saving ? 'Saving…' : 'Save changes'}
        </button>
        {dirty && !saving && <span style={{ fontSize: 12, color: C.sub }}>You have unsaved changes.</span>}
      </div>

      {error && <div style={{ padding: '10px 12px', borderRadius: 8, background: C.redBg, color: C.red, fontSize: 13 }}>{error}</div>}
      {notice && <div style={{ padding: '10px 12px', borderRadius: 8, background: C.greenBg, color: C.green, fontSize: 13 }}>{notice}</div>}
    </div>
  );
}
