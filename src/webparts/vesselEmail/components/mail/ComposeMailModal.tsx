import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Icon } from '@fluentui/react/lib/Icon';
import type VesselEmail from '../VesselEmail';
import { clay } from '../clayTheme';
import { SharePointPicker, PickedItem, fileIcon, formatSize } from './SharePointPicker';

/**
 * Outlook-style "New message" for Vessel DMS (left menu → Send Email).
 *
 *  - To / Cc / Bcc with company-directory search (GET /api/mail/people)
 *    plus any typed address;
 *  - formatted body (bold, italic, underline, lists, links);
 *  - attachments from any SharePoint site in Site Management (files as
 *    copies or links, folders as links) and from the user's computer;
 *  - POST /api/mail/send — sent from the user's own mailbox when the
 *    Mail.Send permission is approved, otherwise from the DMS mailbox with
 *    the user as Reply-To (the result says which).
 */

const LINE = 'var(--vdms-line)';
const BRAND = 'linear-gradient(135deg, #0e7490 0%, #0b2a4a 100%)';
const MAX_TOTAL = 150 * 1024 * 1024;
const EMAIL_RE = /^[^@\s<>]+@[^@\s<>]+\.[^@\s<>]+$/;

interface Person { name: string; email: string; job_title?: string | null; department?: string | null }
interface LocalAttachment { id: string; name: string; size: number; content_type: string; content_b64: string }

function readAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result || '').split(',')[1] || '');
    r.onerror = () => reject(r.error);
    r.readAsDataURL(file);
  });
}

function initials(name: string): string {
  const parts = name.trim().split(/\s+/);
  return ((parts[0] || '')[0] || '') + ((parts[1] || '')[0] || '');
}

function RecipientField(props: {
  host: VesselEmail;
  label: string;
  values: Person[];
  onChange: (v: Person[]) => void;
  autoFocus?: boolean;
  right?: React.ReactNode;
}): React.ReactElement {
  const { host, values, onChange } = props;
  const [text, setText] = React.useState('');
  const [hits, setHits] = React.useState<Person[]>([]);
  const [active, setActive] = React.useState(0);
  const [open, setOpen] = React.useState(false);
  const inputRef = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => {
    const q = text.trim();
    if (!q) { setHits([]); return; }
    let cancelled = false;
    const t = window.setTimeout(() => {
      fetch(`${host._base()}/api/mail/people?q=${encodeURIComponent(q)}&limit=8`, { headers: host._headers() })
        .then(r => (r.ok ? r.json() : { people: [] }))
        .then(d => {
          if (cancelled) return;
          const taken = new Set(values.map(v => v.email.toLowerCase()));
          setHits((d.people || []).filter((p: Person) => !taken.has(p.email.toLowerCase())));
          setActive(0);
        })
        .catch(() => { if (!cancelled) setHits([]); });
    }, 180);
    return () => { cancelled = true; window.clearTimeout(t); };
  }, [text]);

  const add = (p: Person): void => {
    if (!values.some(v => v.email.toLowerCase() === p.email.toLowerCase())) onChange([...values, p]);
    setText(''); setHits([]);
    if (inputRef.current) inputRef.current.focus();
  };
  const addTyped = (): boolean => {
    const parts = text.split(/[;,\s]+/).map(s => s.trim()).filter(Boolean);
    const good = parts.filter(p => EMAIL_RE.test(p));
    if (!good.length) return false;
    const taken = new Set(values.map(v => v.email.toLowerCase()));
    onChange([...values, ...good.filter(e => !taken.has(e.toLowerCase())).map(e => ({ name: e, email: e }))]);
    setText(''); setHits([]);
    return true;
  };

  const onKey = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'ArrowDown' && hits.length) { e.preventDefault(); setActive((active + 1) % hits.length); return; }
    if (e.key === 'ArrowUp' && hits.length) { e.preventDefault(); setActive((active - 1 + hits.length) % hits.length); return; }
    if (e.key === 'Enter' || e.key === 'Tab' || e.key === ',' || e.key === ';') {
      if (hits.length && open) { e.preventDefault(); add(hits[active]); return; }
      if (text.trim() && addTyped()) e.preventDefault();
      return;
    }
    if (e.key === 'Backspace' && !text && values.length) onChange(values.slice(0, -1));
    if (e.key === 'Escape') setOpen(false);
  };

  return (
    <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, padding: '6px 16px', borderBottom: `1px solid ${LINE}`, position: 'relative' }}>
      <span style={{ width: 34, paddingTop: 7, fontSize: 13, color: clay.textMuted, fontWeight: 600 }}>{props.label}</span>
      <div style={{ flex: 1, display: 'flex', flexWrap: 'wrap', gap: 5, alignItems: 'center', minHeight: 32 }} onClick={() => inputRef.current && inputRef.current.focus()}>
        {values.map(v => {
          const bad = !EMAIL_RE.test(v.email);
          return (
            <span key={v.email} title={v.email} style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '3px 6px 3px 3px', borderRadius: 999, background: bad ? '#fee2e2' : clay.accentSoft, color: bad ? '#b91c1c' : clay.accentDark, fontSize: 12.5, fontWeight: 600 }}>
              <span style={{ width: 20, height: 20, borderRadius: '50%', background: '#0e7490', color: '#fff', fontSize: 9.5, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', textTransform: 'uppercase' }}>{initials(v.name)}</span>
              {v.name}
              <button type="button" aria-label={`Remove ${v.name}`} onClick={e => { e.stopPropagation(); onChange(values.filter(x => x.email !== v.email)); }}
                style={{ border: 'none', background: 'none', cursor: 'pointer', color: 'inherit', padding: 0, display: 'inline-flex' }}>
                <Icon iconName="Cancel" style={{ fontSize: 9 }} />
              </button>
            </span>
          );
        })}
        <input ref={inputRef} autoFocus={props.autoFocus} value={text}
          onChange={e => { setText(e.target.value); setOpen(true); }}
          onKeyDown={onKey}
          onBlur={() => window.setTimeout(() => { setOpen(false); if (text.trim()) addTyped(); }, 150)}
          placeholder={values.length ? '' : 'Type a name or email'}
          style={{ flex: 1, minWidth: 140, border: 'none', outline: 'none', fontSize: 13.5, padding: '6px 2px', background: 'transparent', color: clay.text }} />
      </div>
      {props.right}
      {open && hits.length > 0 && (
        <div style={{ position: 'absolute', left: 60, top: '100%', zIndex: 3, width: 360, maxWidth: 'calc(100% - 76px)', background: clay.surface, border: `1px solid ${LINE}`, borderRadius: 10, boxShadow: '0 10px 28px rgba(11,42,74,0.2)', overflow: 'hidden' }}>
          {hits.map((p, i) => (
            <div key={p.email} onMouseDown={e => { e.preventDefault(); add(p); }} onMouseEnter={() => setActive(i)}
              style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', cursor: 'pointer', background: i === active ? clay.surfaceHover : 'transparent' }}>
              <span style={{ width: 30, height: 30, borderRadius: '50%', background: BRAND, color: '#fff', fontSize: 11.5, fontWeight: 700, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', textTransform: 'uppercase', flexShrink: 0 }}>{initials(p.name)}</span>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: clay.text }}>{p.name}</div>
                <div style={{ fontSize: 11.5, color: clay.textMuted, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {p.email}{p.job_title ? ` · ${p.job_title}` : ''}
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export function ComposeMailModal(props: { host: VesselEmail; onClose: () => void; initialItems?: PickedItem[] }): React.ReactElement | React.ReactPortal {
  const { host } = props;
  const [to, setTo] = React.useState<Person[]>([]);
  const [cc, setCc] = React.useState<Person[]>([]);
  const [bcc, setBcc] = React.useState<Person[]>([]);
  const [showCc, setShowCc] = React.useState(false);
  const [subject, setSubject] = React.useState('');
  const [important, setImportant] = React.useState(false);
  const [items, setItems] = React.useState<PickedItem[]>(props.initialItems || []);
  const [files, setFiles] = React.useState<LocalAttachment[]>([]);
  const [picker, setPicker] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState('');
  const [done, setDone] = React.useState<string>('');
  const [senders, setSenders] = React.useState<Array<{ email: string; label: string; kind: string }>>([]);
  const [fromAddr, setFromAddr] = React.useState('');

  React.useEffect(() => {
    fetch(`${host._base()}/api/mail/senders`, { headers: host._headers() })
      .then(r => (r.ok ? r.json() : { senders: [] }))
      .then(d => setSenders(d.senders || []))
      .catch(() => setSenders([]));
  }, []);

  const refreshHistory = (): void => {
    host._fetchJson(`${host._base()}/api/email-logs`).then(d => host.setState({ bentoLogs: d || [] })).catch(() => undefined);
  };
  const bodyRef = React.useRef<HTMLDivElement>(null);
  const fileRef = React.useRef<HTMLInputElement>(null);

  const total = items.filter(i => i.mode === 'attach' && !i.is_folder).reduce((s, i) => s + i.size, 0)
    + files.reduce((s, f) => s + f.size, 0);
  const tooBig = total > MAX_TOTAL;

  const cmd = (name: string, value?: string): void => {
    if (bodyRef.current) bodyRef.current.focus();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (document as any).execCommand(name, false, value);
  };

  const addLink = (): void => {
    const url = window.prompt('Link address (https://…)');
    if (url && /^https?:\/\//i.test(url.trim())) cmd('createLink', url.trim());
  };

  const addLocal = async (list: FileList | null): Promise<void> => {
    if (!list) return;
    const out: LocalAttachment[] = [];
    for (let i = 0; i < list.length; i++) {
      const f = list[i];
      out.push({ id: `${f.name}-${f.size}-${Date.now()}-${i}`, name: f.name, size: f.size, content_type: f.type || 'application/octet-stream', content_b64: await readAsBase64(f) });
    }
    setFiles(prev => [...prev, ...out]);
  };

  const send = async (): Promise<void> => {
    if (busy) return;
    if (!to.length && !cc.length && !bcc.length) { setError('Add at least one recipient.'); return; }
    if (tooBig) { setError('Attachments are over 150 MB — switch some files to "Link".'); return; }
    if (!subject.trim() && !window.confirm('Send this message without a subject?')) return;
    setBusy(true); setError('');
    try {
      const r = await fetch(`${host._base()}/api/mail/send`, {
        method: 'POST',
        headers: host._headers(),
        body: JSON.stringify({
          to: to.map(p => p.email), cc: cc.map(p => p.email), bcc: bcc.map(p => p.email),
          subject: subject.trim(), body_html: bodyRef.current ? bodyRef.current.innerHTML : '',
          importance: important ? 'high' : 'normal',
          items: items.map(i => ({ drive_id: i.drive_id, item_id: i.item_id, name: i.name, is_folder: i.is_folder, mode: i.is_folder ? 'link' : i.mode })),
          local_files: files.map(f => ({ name: f.name, content_type: f.content_type, content_b64: f.content_b64 })),
          from_address: fromAddr || null,
        }),
      });
      const data = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error(data.detail || data.message || `Sending failed (${r.status}).`);
      setDone(data.sent_as === 'user'
        ? 'Sent from your mailbox — it is in your Outlook Sent Items.'
        : `Sent from ${data.from} (replies come to you).`);
      refreshHistory();
      window.setTimeout(props.onClose, 1800);
    } catch (e) {
      setError((e as Error).message || 'Sending failed.');
      refreshHistory();
    } finally {
      setBusy(false);
    }
  };

  const toolBtn = (icon: string, title: string, onClick: () => void): React.ReactElement => (
    <button type="button" title={title} aria-label={title} onMouseDown={e => e.preventDefault()} onClick={onClick}
      style={{ width: 30, height: 28, border: 'none', borderRadius: 6, background: 'transparent', color: clay.text, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}>
      <Icon iconName={icon} style={{ fontSize: 13 }} />
    </button>
  );

  const content = (
    <div data-vessel-theme={host.state.themeMode}
      style={{ position: 'fixed', inset: 0, zIndex: 100001, background: 'rgba(11,42,74,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16, fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif" }}>
      <div role="dialog" aria-label="New message"
        style={{ position: 'relative', width: 860, maxWidth: '100%', height: '90vh', background: clay.surface, borderRadius: 16, boxShadow: '0 24px 60px rgba(11,42,74,0.35)', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
        {/* Header */}
        <div style={{ background: BRAND, color: '#fff', padding: '12px 16px', display: 'flex', alignItems: 'center', gap: 10 }}>
          <Icon iconName="Mail" style={{ fontSize: 18 }} />
          <div style={{ flex: 1, fontWeight: 800, fontSize: 15 }}>New message</div>
          <button type="button" onClick={props.onClose} aria-label="Close" style={{ border: 'none', background: 'rgba(255,255,255,0.16)', color: '#fff', borderRadius: 8, padding: '6px 8px', cursor: 'pointer' }}>
            <Icon iconName="Cancel" />
          </button>
        </div>

        {/* Toolbar */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 16px', borderBottom: `1px solid ${LINE}`, flexWrap: 'wrap' }}>
          <button type="button" onClick={() => void send()} disabled={busy || !!done}
            style={{ display: 'inline-flex', alignItems: 'center', gap: 8, padding: '8px 18px', borderRadius: 8, border: 'none', background: BRAND, color: '#fff', fontWeight: 700, fontSize: 13.5, cursor: busy ? 'wait' : 'pointer', opacity: busy ? 0.7 : 1 }}>
            <Icon iconName="Send" /> {busy ? 'Sending…' : 'Send'}
          </button>
          <button type="button" onClick={() => setPicker(true)}
            style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '8px 12px', borderRadius: 8, border: `1px solid ${LINE}`, background: clay.surface, color: clay.text, fontWeight: 600, fontSize: 13, cursor: 'pointer' }}>
            <Icon iconName="SharepointLogo" style={{ color: '#0e7490' }} /> Attach from SharePoint
          </button>
          <button type="button" onClick={() => fileRef.current && fileRef.current.click()}
            style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '8px 12px', borderRadius: 8, border: `1px solid ${LINE}`, background: clay.surface, color: clay.text, fontWeight: 600, fontSize: 13, cursor: 'pointer' }}>
            <Icon iconName="Attach" /> From computer
          </button>
          <input ref={fileRef} type="file" multiple style={{ display: 'none' }} onChange={e => { void addLocal(e.target.files); e.target.value = ''; }} />
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 12.5, color: important ? '#b91c1c' : clay.textMuted, cursor: 'pointer', marginLeft: 'auto', fontWeight: 600 }}>
            <input type="checkbox" checked={important} onChange={e => setImportant(e.target.checked)} />
            <Icon iconName="Important" /> High importance
          </label>
        </div>

        {/* Addressing */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 16px', borderBottom: `1px solid ${LINE}`, fontSize: 13 }}>
          <span style={{ width: 34, color: clay.textMuted, fontWeight: 600 }}>From</span>
          {senders.length > 1 ? (
            <select value={fromAddr} onChange={e => setFromAddr(e.target.value)} title="Send from your mailbox or an approved shared mailbox"
              style={{ flex: 1, maxWidth: 460, padding: '6px 8px', borderRadius: 8, border: `1px solid ${LINE}`, background: clay.surface, color: clay.text, fontSize: 13 }}>
              {senders.map(s => (
                <option key={s.email} value={s.kind === 'me' ? '' : s.email}>
                  {s.kind === 'me' ? `${host.props.userDisplayName || 'Me'} <${s.email}>` : `${s.email} (shared mailbox)`}
                </option>
              ))}
            </select>
          ) : (
            <span style={{ color: clay.text }}>{host.props.userDisplayName || host.props.userEmail} <span style={{ color: clay.textMuted }}>&lt;{host.props.userEmail}&gt;</span></span>
          )}
        </div>
        <RecipientField host={host} label="To" values={to} onChange={setTo} autoFocus
          right={!showCc ? <button type="button" onClick={() => setShowCc(true)} style={{ border: 'none', background: 'none', color: clay.accent, fontWeight: 700, fontSize: 12.5, cursor: 'pointer', paddingTop: 7 }}>Cc Bcc</button> : undefined} />
        {showCc && <RecipientField host={host} label="Cc" values={cc} onChange={setCc} />}
        {showCc && <RecipientField host={host} label="Bcc" values={bcc} onChange={setBcc} />}
        <div style={{ padding: '4px 16px', borderBottom: `1px solid ${LINE}` }}>
          <input value={subject} onChange={e => setSubject(e.target.value)} placeholder="Add a subject" maxLength={900}
            style={{ width: '100%', border: 'none', outline: 'none', fontSize: 15, fontWeight: 600, padding: '8px 0', background: 'transparent', color: clay.text, boxSizing: 'border-box' }} />
        </div>

        {/* Attachments */}
        {(items.length > 0 || files.length > 0) && (
          <div style={{ padding: '10px 16px', borderBottom: `1px solid ${LINE}`, display: 'flex', flexWrap: 'wrap', gap: 8, maxHeight: 150, overflowY: 'auto' }}>
            {items.map(it => {
              const fi = fileIcon(it.name, it.is_folder);
              return (
                <div key={it.drive_id + it.item_id} title={`${it.site_name} · ${it.name}`}
                  style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px', border: `1px solid ${LINE}`, borderRadius: 10, background: clay.surfaceRaised, maxWidth: 330 }}>
                  <Icon iconName={fi.icon} style={{ fontSize: 18, color: fi.color }} />
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontSize: 12.5, fontWeight: 600, color: clay.text, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 190 }}>{it.name}</div>
                    <div style={{ fontSize: 11, color: clay.textMuted }}>{it.is_folder ? 'Folder · shared as link' : `${formatSize(it.size)} · ${it.site_name}`}</div>
                  </div>
                  {!it.is_folder && (
                    <select value={it.mode} onChange={e => setItems(prev => prev.map(p => (p === it ? { ...p, mode: e.target.value as 'attach' | 'link' } : p)))}
                      title="Send a copy, or a SharePoint link" style={{ fontSize: 11.5, border: `1px solid ${LINE}`, borderRadius: 6, padding: '2px 4px', background: clay.surface, color: clay.text }}>
                      <option value="attach">Copy</option>
                      <option value="link">Link</option>
                    </select>
                  )}
                  <button type="button" aria-label={`Remove ${it.name}`} onClick={() => setItems(prev => prev.filter(p => p !== it))}
                    style={{ border: 'none', background: 'none', cursor: 'pointer', color: clay.textMuted }}><Icon iconName="Cancel" style={{ fontSize: 10 }} /></button>
                </div>
              );
            })}
            {files.map(f => {
              const fi = fileIcon(f.name, false);
              return (
                <div key={f.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px', border: `1px solid ${LINE}`, borderRadius: 10, background: clay.surfaceRaised, maxWidth: 300 }}>
                  <Icon iconName={fi.icon} style={{ fontSize: 18, color: fi.color }} />
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontSize: 12.5, fontWeight: 600, color: clay.text, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 200 }}>{f.name}</div>
                    <div style={{ fontSize: 11, color: clay.textMuted }}>{formatSize(f.size)} · from computer</div>
                  </div>
                  <button type="button" aria-label={`Remove ${f.name}`} onClick={() => setFiles(prev => prev.filter(p => p.id !== f.id))}
                    style={{ border: 'none', background: 'none', cursor: 'pointer', color: clay.textMuted }}><Icon iconName="Cancel" style={{ fontSize: 10 }} /></button>
                </div>
              );
            })}
            <div style={{ width: '100%', fontSize: 11.5, color: tooBig ? '#b91c1c' : clay.textMuted }}>
              {formatSize(total) || '0 B'} attached{tooBig ? ' — over the 150 MB limit, switch some files to "Link"' : ''}
            </div>
          </div>
        )}

        {/* Formatting */}
        <div style={{ display: 'flex', gap: 2, padding: '6px 12px', borderBottom: `1px solid ${LINE}` }}>
          {toolBtn('Bold', 'Bold', () => cmd('bold'))}
          {toolBtn('Italic', 'Italic', () => cmd('italic'))}
          {toolBtn('Underline', 'Underline', () => cmd('underline'))}
          {toolBtn('BulletedList', 'Bulleted list', () => cmd('insertUnorderedList'))}
          {toolBtn('NumberedList', 'Numbered list', () => cmd('insertOrderedList'))}
          {toolBtn('Link', 'Insert link', addLink)}
          {toolBtn('ClearFormatting', 'Clear formatting', () => cmd('removeFormat'))}
        </div>

        {/* Body */}
        <div ref={bodyRef} contentEditable suppressContentEditableWarning role="textbox" aria-multiline="true" aria-label="Message body"
          style={{ flex: 1, overflowY: 'auto', padding: '14px 16px', fontSize: 14, lineHeight: 1.6, color: clay.text, outline: 'none' }} />

        {/* Status */}
        {(error || done) && (
          <div style={{ padding: '10px 16px', borderTop: `1px solid ${LINE}`, fontSize: 13, fontWeight: 600, color: done ? '#15803d' : '#b91c1c', background: done ? '#ecfdf5' : '#fef2f2', display: 'flex', alignItems: 'center', gap: 8 }}>
            <Icon iconName={done ? 'CompletedSolid' : 'ErrorBadge'} /> {done || error}
          </div>
        )}

        {picker && (
          <SharePointPicker host={host} onClose={() => setPicker(false)}
            onAdd={picked => setItems(prev => {
              const have = new Set(prev.map(p => p.drive_id + p.item_id));
              return [...prev, ...picked.filter(p => !have.has(p.drive_id + p.item_id))];
            })} />
        )}
      </div>
    </div>
  );
  return typeof document !== 'undefined' ? ReactDOM.createPortal(content, document.body) : content;
}
