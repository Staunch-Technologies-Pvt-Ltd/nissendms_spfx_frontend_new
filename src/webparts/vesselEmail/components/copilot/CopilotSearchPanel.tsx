import * as React from 'react';
import type VesselEmail from '../VesselEmail';

/**
 * Documents Copilot — ask a plain-English question and get matching
 * documents back, instead of building a query with the filter dropdowns.
 *
 * Rendered as a floating action button pinned to the bottom-right corner of
 * the viewport; clicking it opens the chat popup above the button. The
 * button/popup are fixed-position, so this component can be mounted anywhere
 * in the tree without affecting page layout.
 *
 * Self-contained: this is the only file the Documents module imports for
 * this feature. Backend: backend/app/copilot/ (api.py + service.py +
 * config.py) — mounted from main.py, nothing else in the app is affected.
 *
 *   GET  /api/copilot/status  -> { configured: boolean }
 *   POST /api/copilot/query   -> { answer, mode, filters, results }
 *
 * `configured: false` just means no Azure OpenAI key is set yet in
 * backend/.env (see backend/app/copilot/config.py) — the box still works,
 * it falls back to a plain keyword/vessel-name search server-side.
 */

interface CopilotTrailItem { id: string; name: string; }
interface CopilotResult {
  id: string;
  name: string;
  kind: 'file' | 'folder' | string;
  path: string;
  trail: CopilotTrailItem[];
}
interface CopilotResponse {
  answer: string;
  mode: 'ai' | 'keyword';
  filters: { keywords?: string; vessel_name?: string | null; group?: string | null; category?: string | null };
  results: CopilotResult[];
}

const EXAMPLES: string[] = [
  'crew certificates for MV Horizon',
  'drawings for MV Aurora',
  'insurance documents',
];

async function readError(r: Response): Promise<string> {
  try {
    const j = await r.json();
    return (j && j.detail) || `HTTP ${r.status}`;
  } catch {
    return `HTTP ${r.status}`;
  }
}

export function CopilotSearchPanel({ host }: { host: VesselEmail }): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const [configured, setConfigured] = React.useState<boolean | null>(null);
  const [question, setQuestion] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState('');
  const [response, setResponse] = React.useState<CopilotResponse | null>(null);

  React.useEffect(() => {
    let cancelled = false;
    fetch(`${host._base()}/api/copilot/status`, { headers: host._headers() })
      .then(r => (r.ok ? r.json() : { configured: false }))
      .then(data => { if (!cancelled) setConfigured(!!data.configured); })
      .catch(() => { if (!cancelled) setConfigured(false); });
    return () => { cancelled = true; };
  }, [host]);

  const ask = async (override?: string): Promise<void> => {
    const text = (override ?? question).trim();
    if (!text || busy) return;
    setBusy(true); setError(''); setResponse(null);
    try {
      const r = await fetch(`${host._base()}/api/copilot/query`, {
        method: 'POST',
        headers: host._headers(),
        body: JSON.stringify({ question: text }),
      });
      if (!r.ok) throw new Error(await readError(r));
      setResponse(await r.json());
    } catch (e) {
      setError((e as Error).message || 'Copilot could not answer that.');
    } finally {
      setBusy(false);
    }
  };

  const openResult = (item: CopilotResult): void => {
    if (item.kind !== 'file') return;
    const folderTrail = item.trail.slice(0, -1).map(t => t.name).join('/');
    void host._openDocumentFile(item.id, item.name, folderTrail);
  };

  return (
    <>
      {/* Floating toggle button — bottom-right of the viewport */}
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        title={open ? 'Close Copilot' : 'Ask Copilot'}
        style={{
          position: 'fixed', right: 24, bottom: 24, zIndex: 1000,
          width: 52, height: 52, borderRadius: '50%', border: 'none',
          background: open ? '#5b21b6' : 'linear-gradient(135deg, #7c3aed 0%, #4c1d95 100%)',
          color: '#fff', fontSize: 22, cursor: 'pointer',
          boxShadow: '0 6px 18px rgba(76, 29, 149, 0.4)',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}
      >
        {open ? '✕' : '✨'}
      </button>

      {/* Popup chat panel */}
      {open && (
        <div
          style={{
            position: 'fixed', right: 24, bottom: 88, zIndex: 1000,
            width: 360, maxWidth: 'calc(100vw - 48px)', maxHeight: 'calc(100vh - 140px)',
            overflowY: 'auto',
            background: 'linear-gradient(135deg, #eef2ff 0%, #f5f3ff 100%)',
            border: '1px solid #ddd6fe', borderRadius: 12, padding: 14,
            boxShadow: '0 12px 32px rgba(76, 29, 149, 0.28)',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <span style={{ fontSize: 16 }}>✨</span>
            <span style={{ fontWeight: 700, fontSize: 13, color: '#4c1d95' }}>Ask Copilot</span>
          </div>
          <div style={{ fontSize: 11, color: '#7c3aed', marginTop: 2 }}>
            {configured === false
              ? 'basic mode — set an Azure OpenAI key in backend/.env for smarter answers'
              : 'ask a question about your documents, in plain English'}
          </div>

          <div style={{ marginTop: 10 }}>
            <div style={{ display: 'flex', gap: 8 }}>
              <input
                type="text"
                autoFocus
                value={question}
                onChange={e => setQuestion(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter') void ask(); }}
                placeholder='e.g. "crew certificates for MV Horizon"'
                style={{
                  flex: 1, padding: '8px 12px', borderRadius: 8, border: '1px solid #ddd6fe',
                  fontSize: 13, outline: 'none', boxSizing: 'border-box',
                }}
              />
              <button
                type="button"
                onClick={() => void ask()}
                disabled={busy || !question.trim()}
                style={{
                  padding: '8px 16px', borderRadius: 8, border: 'none',
                  background: busy || !question.trim() ? '#c4b5fd' : '#7c3aed', color: '#fff',
                  fontWeight: 600, fontSize: 13, cursor: busy || !question.trim() ? 'not-allowed' : 'pointer',
                  whiteSpace: 'nowrap',
                }}
              >
                {busy ? 'Asking…' : 'Ask'}
              </button>
            </div>

            {!question && !response && (
              <div style={{ marginTop: 8, display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                {EXAMPLES.map(ex => (
                  <button
                    key={ex}
                    type="button"
                    onClick={() => { setQuestion(ex); void ask(ex); }}
                    style={{
                      fontSize: 11, padding: '4px 10px', borderRadius: 999, border: '1px solid #ddd6fe',
                      background: '#fff', color: '#6d28d9', cursor: 'pointer',
                    }}
                  >
                    {ex}
                  </button>
                ))}
              </div>
            )}

            {error && <div style={{ marginTop: 10, fontSize: 12, color: '#b91c1c' }}>{error}</div>}

            {response && (
              <div style={{ marginTop: 12 }}>
                <div style={{ fontSize: 13, color: '#3730a3', marginBottom: 8 }}>{response.answer}</div>
                {response.results.length > 0 && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4, maxHeight: 280, overflowY: 'auto' }}>
                    {response.results.map(item => (
                      <div
                        key={item.id}
                        onClick={() => openResult(item)}
                        title={item.path}
                        style={{
                          display: 'flex', alignItems: 'center', gap: 8, padding: '6px 10px',
                          background: '#fff', borderRadius: 6, border: '1px solid #ede9fe',
                          cursor: item.kind === 'file' ? 'pointer' : 'default', fontSize: 12,
                        }}
                      >
                        <span>{item.kind === 'file' ? '📄' : '📁'}</span>
                        <span style={{ flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {item.name}
                        </span>
                        <span style={{
                          color: '#94a3b8', fontSize: 10, overflow: 'hidden',
                          textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 220,
                        }}>
                          {item.trail.slice(0, -1).map(t => t.name).join(' › ')}
                        </span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      )}
    </>
  );
}
