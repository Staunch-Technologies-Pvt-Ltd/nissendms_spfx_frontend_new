import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';

/**
 * Documents Copilot — a chat that finds documents from a plain-English
 * question ("latest manuals for Belle Lune", "how many drawings do we
 * have?", "PDFs uploaded last week") across every site in Site Management.
 *
 * A floating button pinned bottom-right on every page (mounted once by the
 * app shell); it opens a chat panel. Each answer can be refined with a
 * follow-up ("only drawings") — the previous answer's filters are sent back
 * as context — and every result opens the file in SharePoint.
 *
 * Backend: app/copilot/ —
 *   GET  /api/copilot/status  -> { configured }
 *   POST /api/copilot/query   -> { answer, mode, filters, results, total,
 *                                  breakdown, top_vessels, suggestions }
 * Works without an AI key (rule-based); with Azure OpenAI configured in the
 * backend .env it understands questions more freely.
 */

interface CopilotResult {
  id: string;
  name: string;
  vessel: string | null;
  group: string;
  file_type: string;
  site: string;
  site_name: string;
  path: string;
  modified: string;
  modified_by: string | null;
  size: string;
  web_url: string | null;
}
interface CopilotFilters {
  vessel: string | null; group: string | null; file_type: string | null;
  period: string | null; person: string | null; intent: string; keywords: string; site?: string | null;
}
interface CopilotResponse {
  answer: string;
  mode: 'ai' | 'keyword';
  filters: CopilotFilters;
  results: CopilotResult[];
  total: number;
  breakdown: { drawings: number; manuals: number; to_be_classified: number; other: number } | null;
  top_vessels: Array<{ name: string; count: number }>;
  by_site?: Array<{ name: string; count: number; vessels: number }>;
  suggestions: string[];
}
interface Turn { id: number; question: string; response?: CopilotResponse; error?: string }

const STARTERS: string[] = [
  'Which sites have our documents?',
  'How many drawings do we have?',
  'Documents uploaded this week',
  'What still needs classifying?',
  'Largest files',
];

const LINE = 'var(--vdms-line)';
const BRAND = 'linear-gradient(135deg, #0e7490 0%, #0b2a4a 100%)';

/** Minimal **bold** / *italic* rendering for the answer line. */
function renderRich(text: string): React.ReactNode[] {
  const out: React.ReactNode[] = [];
  const re = /\*\*([^*]+)\*\*|\*([^*]+)\*/g;
  let last = 0;
  let m: RegExpExecArray | null;
  let i = 0;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(text.slice(last, m.index));
    out.push(m[1] ? <strong key={i++}>{m[1]}</strong> : <em key={i++}>{m[2]}</em>);
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(text.slice(last));
  return out;
}

const FILE_ICONS: Record<string, { icon: string; color: string }> = {
  pdf: { icon: 'PDF', color: '#dc2626' },
  word: { icon: 'WordDocument', color: '#2563eb' },
  excel: { icon: 'ExcelDocument', color: '#16a34a' },
  powerpoint: { icon: 'PowerPointDocument', color: '#ea580c' },
  image: { icon: 'FileImage', color: '#9333ea' },
  drawing: { icon: 'FileCAD', color: '#0e7490' },
  email: { icon: 'Mail', color: '#0891b2' },
  archive: { icon: 'ZipFolder', color: '#a16207' },
};

async function readError(r: Response): Promise<string> {
  try {
    const j = await r.json();
    return (j && (j.detail || j.message)) || `HTTP ${r.status}`;
  } catch {
    return `HTTP ${r.status}`;
  }
}

function Chip(props: { label: React.ReactNode; onClick?: () => void; title?: string }): React.ReactElement {
  return (
    <button
      type="button"
      onClick={props.onClick}
      title={props.title}
      style={{
        fontSize: 11.5, padding: '4px 10px', borderRadius: 999, border: `1px solid ${LINE}`,
        background: clay.surface, color: clay.accentDark, cursor: props.onClick ? 'pointer' : 'default',
        fontWeight: 600, whiteSpace: 'nowrap',
      }}
    >
      {props.label}
    </button>
  );
}

function ResultCard({ item, onOpen }: { item: CopilotResult; onOpen: (r: CopilotResult) => void }): React.ReactElement {
  const fi = FILE_ICONS[item.file_type] || { icon: 'Page', color: clay.textMuted };
  const meta = [item.vessel, item.group !== 'Other' ? item.group : null, item.site_name].filter(Boolean).join(' · ');
  return (
    <div
      onClick={() => onOpen(item)}
      title={item.path}
      style={{
        display: 'flex', alignItems: 'center', gap: 10, padding: '8px 10px', background: clay.surface,
        borderRadius: 10, border: `1px solid ${LINE}`, cursor: item.web_url ? 'pointer' : 'default',
      }}
    >
      <Icon iconName={fi.icon} aria-hidden="true" style={{ fontSize: 18, color: fi.color, flexShrink: 0 }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: clay.text, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
          {item.name}
        </div>
        <div style={{ fontSize: 11, color: clay.textMuted, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
          {meta}{meta ? ' · ' : ''}{item.modified}{item.modified_by ? ` · ${item.modified_by}` : ''}
        </div>
      </div>
      {item.web_url && <Icon iconName="OpenInNewWindow" aria-hidden="true" style={{ fontSize: 13, color: clay.accent, flexShrink: 0 }} />}
    </div>
  );
}

export function CopilotSearchPanel({ host }: { host: VesselEmail }): React.ReactElement | React.ReactPortal {
  const [open, setOpen] = React.useState(false);
  const [configured, setConfigured] = React.useState<boolean | null>(null);
  const [m365Url, setM365Url] = React.useState<string>('https://m365.cloud.microsoft/chat');
  const [question, setQuestion] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const [turns, setTurns] = React.useState<Turn[]>([]);
  const [siteKey, setSiteKey] = React.useState<string>(host.state.dashboardSiteFilter || 'all');
  const [expanded, setExpanded] = React.useState<Record<number, boolean>>({});
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const inputRef = React.useRef<HTMLInputElement>(null);
  const nextId = React.useRef(1);

  React.useEffect(() => {
    if (!open || configured !== null) return;
    let cancelled = false;
    fetch(`${host._base()}/api/copilot/status`, { headers: host._headers() })
      .then(r => (r.ok ? r.json() : { configured: false }))
      .then(data => {
        if (cancelled) return;
        setConfigured(!!data.configured);
        if (data.m365_url) setM365Url(data.m365_url);
      })
      .catch(() => { if (!cancelled) setConfigured(false); });
    return () => { cancelled = true; };
  }, [open, configured, host]);

  React.useEffect(() => {
    if (scrollRef.current) scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
  }, [turns, busy]);

  React.useEffect(() => {
    if (open && inputRef.current) inputRef.current.focus();
  }, [open]);

  const sites = React.useMemo(() => {
    const seen = new Set<string>();
    return (host.state.documentSites || [])
      .filter(s => {
        const k = String(s.site_key || '').toLowerCase();
        if (!k || seen.has(k)) return false;
        seen.add(k);
        return true;
      })
      .map(s => ({ key: s.site_key, label: s.sp_site_name || s.site_key }));
  }, [host.state.documentSites]);

  const lastFilters = (): CopilotFilters | undefined => {
    for (let i = turns.length - 1; i >= 0; i--) if (turns[i].response) return turns[i].response!.filters;
    return undefined;
  };

  const ask = async (override?: string): Promise<void> => {
    const text = (override ?? question).trim();
    if (!text || busy) return;
    const id = nextId.current++;
    setTurns(t => [...t, { id, question: text }]);
    setQuestion('');
    setBusy(true);
    try {
      const r = await fetch(`${host._base()}/api/copilot/query`, {
        method: 'POST',
        headers: host._headers(),
        body: JSON.stringify({
          question: text,
          site_key: siteKey,
          context: lastFilters() || null,
          history: turns.filter(t => t.response).slice(-4).map(t => ({ question: t.question, answer: t.response!.answer })),
        }),
      });
      if (!r.ok) throw new Error(await readError(r));
      const data: CopilotResponse = await r.json();
      setTurns(t => t.map(x => (x.id === id ? { ...x, response: data } : x)));
    } catch (e) {
      const msg = (e as Error).message || 'Copilot could not answer that.';
      setTurns(t => t.map(x => (x.id === id ? { ...x, error: msg } : x)));
    } finally {
      setBusy(false);
    }
  };

  const openResult = (item: CopilotResult): void => {
    if (item.web_url) window.open(item.web_url, '_blank', 'noopener');
  };

  const renderTurn = (turn: Turn): React.ReactElement => {
    const res = turn.response;
    const showAll = !!expanded[turn.id];
    const shown = res ? (showAll ? res.results : res.results.slice(0, 6)) : [];
    return (
      <div key={turn.id} style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        <div style={{ alignSelf: 'flex-end', maxWidth: '85%', background: BRAND, color: '#fff', padding: '8px 12px', borderRadius: '14px 14px 4px 14px', fontSize: 13 }}>
          {turn.question}
        </div>
        {turn.error && (
          <div style={{ alignSelf: 'flex-start', fontSize: 12.5, color: '#b91c1c', background: '#fef2f2', border: '1px solid #fecaca', padding: '8px 12px', borderRadius: 12 }}>
            {turn.error}
          </div>
        )}
        {res && (
          <div style={{ alignSelf: 'stretch', background: clay.surfaceRaised, border: `1px solid ${LINE}`, borderRadius: '4px 14px 14px 14px', padding: 12, display: 'flex', flexDirection: 'column', gap: 10 }}>
            <div style={{ fontSize: 13, color: clay.text, lineHeight: 1.5 }}>{renderRich(res.answer)}</div>

            {res.breakdown && (res.breakdown.drawings + res.breakdown.manuals + res.breakdown.to_be_classified) > 0 && (
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                {res.breakdown.drawings > 0 && res.filters.group !== 'drawings' && (
                  <Chip label={`Drawings · ${res.breakdown.drawings.toLocaleString()}`} onClick={() => void ask('only drawings')} />
                )}
                {res.breakdown.manuals > 0 && res.filters.group !== 'manuals' && (
                  <Chip label={`Manuals · ${res.breakdown.manuals.toLocaleString()}`} onClick={() => void ask('only manuals')} />
                )}
                {res.breakdown.to_be_classified > 0 && res.filters.group !== 'to_be_classified' && (
                  <Chip label={`To be classified · ${res.breakdown.to_be_classified.toLocaleString()}`} onClick={() => void ask('only to be classified')} />
                )}
              </div>
            )}

            {res.by_site && res.by_site.length > 1 && (
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                {res.by_site.map(b => (
                  <Chip key={b.name}
                    label={<><Icon iconName="Globe" style={{ fontSize: 11, marginRight: 4 }} />{b.name} · {b.count.toLocaleString()}</>}
                    title={b.vessels ? `${b.vessels} vessels` : undefined}
                    onClick={() => void ask(`overview of ${b.name}`)} />
                ))}
              </div>
            )}

            {res.top_vessels.length > 1 && (
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                {res.top_vessels.map(v => (
                  <Chip key={v.name} label={<><Icon iconName="Ferry" style={{ fontSize: 11, marginRight: 4 }} />{v.name} · {v.count.toLocaleString()}</>}
                    onClick={() => void ask(`${turn.question} for ${v.name}`)} />
                ))}
              </div>
            )}

            {shown.length > 0 && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {shown.map(item => <ResultCard key={item.id} item={item} onOpen={openResult} />)}
                {res.results.length > shown.length && (
                  <button type="button" onClick={() => setExpanded(e => ({ ...e, [turn.id]: true }))}
                    style={{ alignSelf: 'flex-start', border: 'none', background: 'none', color: clay.accent, fontWeight: 700, fontSize: 12, cursor: 'pointer', padding: 0 }}>
                    Show {res.results.length - shown.length} more
                  </button>
                )}
              </div>
            )}

            {res.suggestions.length > 0 && (
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', borderTop: `1px dashed ${LINE}`, paddingTop: 8 }}>
                {res.suggestions.map(s => <Chip key={s} label={s} onClick={() => void ask(s)} />)}
              </div>
            )}
          </div>
        )}
      </div>
    );
  };

  // Rendered into <body>: SharePoint wraps web parts in transformed
  // containers, which turn position:fixed into "bottom of the web part" —
  // on a long page the button ended up far below the screen. Above the
  // app shell, which is itself a body-level overlay (AppLayout, z 99999).
  const content = (
    <div
      className="vdms-copilot-root"
      data-vessel-theme={host.state.themeMode}
      style={{ fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif" }}
    >
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        title={open ? 'Close Copilot' : 'Ask Copilot about your documents'}
        aria-label={open ? 'Close Copilot' : 'Open Copilot'}
        style={{
          position: 'fixed', right: 24, bottom: 24, zIndex: 100000,
          height: 52, minWidth: 52, padding: open ? 0 : '0 18px 0 14px', borderRadius: 26, border: 'none',
          background: BRAND, color: '#fff', cursor: 'pointer', gap: 8,
          boxShadow: '0 8px 22px rgba(11, 42, 74, 0.35)', fontWeight: 700, fontSize: 14,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}
      >
        <Icon iconName={open ? 'Cancel' : 'Robot'} aria-hidden="true" style={{ fontSize: 18 }} />
        {!open && <span>Ask Copilot</span>}
      </button>

      {open && (
        <div
          role="dialog"
          aria-label="Documents Copilot"
          style={{
            position: 'fixed', right: 24, bottom: 88, zIndex: 100000,
            width: 440, maxWidth: 'calc(100vw - 32px)', height: 620, maxHeight: 'calc(100vh - 120px)',
            display: 'flex', flexDirection: 'column', overflow: 'hidden',
            background: clay.surface, border: `1px solid ${LINE}`, borderRadius: 18,
            boxShadow: '0 18px 48px rgba(11, 42, 74, 0.28)',
          }}
        >
          {/* Header */}
          <div style={{ background: BRAND, color: '#fff', padding: '12px 14px', display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{ width: 34, height: 34, borderRadius: 10, background: 'rgba(255,255,255,0.16)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <Icon iconName="Robot" aria-hidden="true" style={{ fontSize: 18 }} />
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 800, fontSize: 14.5 }}>Documents Copilot</div>
              <div style={{ fontSize: 11, opacity: 0.85 }}>
                Finds documents across your vessels and sites
              </div>
            </div>
            {turns.length > 0 && (
              <button type="button" onClick={() => { setTurns([]); setExpanded({}); }} title="New conversation"
                style={{ border: 'none', background: 'rgba(255,255,255,0.16)', color: '#fff', borderRadius: 8, padding: '6px 8px', cursor: 'pointer' }}>
                <Icon iconName="Refresh" aria-hidden="true" style={{ fontSize: 13 }} />
              </button>
            )}
          </div>

          {/* Site scope */}
          <div style={{ padding: '8px 14px', borderBottom: `1px solid ${LINE}`, display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, color: clay.textMuted }}>
            <Icon iconName="Globe" aria-hidden="true" style={{ fontSize: 12 }} />
            <span>Search in</span>
            <select value={siteKey} onChange={e => setSiteKey(e.target.value)}
              style={{ flex: 1, fontSize: 12, padding: '4px 6px', borderRadius: 6, border: `1px solid ${LINE}`, background: clay.surface, color: clay.text }}>
              <option value="all">All SharePoint Sites</option>
              {sites.map(s => <option key={s.key} value={s.key}>{s.label}</option>)}
            </select>
          </div>

          {/* Microsoft 365 Copilot (company licence) for questions about document contents */}
          <a href={m365Url} target="_blank" rel="noopener noreferrer"
            title="Opens Microsoft 365 Copilot with your company account — answers from document contents, within your permissions"
            style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '8px 14px', borderBottom: `1px solid ${LINE}`, background: clay.surfaceRaised, color: clay.accentDark, fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>
            <Icon iconName="Robot" aria-hidden="true" style={{ fontSize: 13 }} />
            <span style={{ flex: 1 }}>Ask about what's <em>inside</em> documents in Microsoft 365 Copilot</span>
            <Icon iconName="OpenInNewWindow" aria-hidden="true" style={{ fontSize: 12 }} />
          </a>

          {/* Conversation */}
          <div ref={scrollRef} style={{ flex: 1, overflowY: 'auto', padding: 14, display: 'flex', flexDirection: 'column', gap: 14 }}>
            {turns.length === 0 && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                <div style={{ fontSize: 13.5, color: clay.text, lineHeight: 1.5 }}>
                  Hi! Ask me about your documents in plain English — a vessel, a type of document, a date or a person.
                </div>
                <div style={{ fontSize: 11.5, color: clay.textMuted }}>
                  e.g. <em>“latest manuals for Belle Lune”</em>, <em>“PDFs uploaded last month”</em>, <em>“overview of Bow Fighter”</em>
                </div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                  {STARTERS.map(s => <Chip key={s} label={s} onClick={() => void ask(s)} />)}
                </div>
              </div>
            )}
            {turns.map(renderTurn)}
            {busy && (
              <div style={{ alignSelf: 'flex-start', fontSize: 12.5, color: clay.textMuted, display: 'flex', alignItems: 'center', gap: 6 }}>
                <Icon iconName="Search" aria-hidden="true" /> Searching documents…
              </div>
            )}
          </div>

          {/* Input */}
          <div style={{ padding: 12, borderTop: `1px solid ${LINE}`, display: 'flex', gap: 8 }}>
            <input
              ref={inputRef}
              type="text"
              value={question}
              onChange={e => setQuestion(e.target.value)}
              onKeyDown={e => { if (e.key === 'Enter') void ask(); }}
              placeholder={turns.length ? 'Ask a follow-up, e.g. "only drawings"' : 'Ask about your documents…'}
              maxLength={500}
              style={{
                flex: 1, padding: '10px 12px', borderRadius: 10, border: `1px solid ${LINE}`,
                fontSize: 13, outline: 'none', background: clay.surface, color: clay.text, boxSizing: 'border-box',
              }}
            />
            <button
              type="button"
              onClick={() => void ask()}
              disabled={busy || !question.trim()}
              aria-label="Send"
              style={{
                width: 42, borderRadius: 10, border: 'none', background: busy || !question.trim() ? clay.accentSoft : BRAND,
                color: '#fff', cursor: busy || !question.trim() ? 'not-allowed' : 'pointer',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
              }}
            >
              <Icon iconName="Send" aria-hidden="true" style={{ fontSize: 15 }} />
            </button>
          </div>
        </div>
      )}
    </div>
  );
  return typeof document !== 'undefined' ? ReactDOM.createPortal(content, document.body) : content;
}
